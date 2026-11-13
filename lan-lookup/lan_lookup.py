#!/usr/bin/env python3
"""
lan_lookup.py - read LAN IDs from an Excel sheet, look each one up on the
internal person page (wab/person.asp), and write Department + Reviewer back
to a new Excel file.

Run it on your work PC (it needs the corporate network and your Windows login).

    pip install pandas openpyxl requests beautifulsoup4 requests-negotiate-sspi
    # optional, if SSO doesn't work:      pip install requests-ntlm
    # optional, for corp TLS certs:       pip install truststore

    # 1) First, see what fields the page returns for one LAN ID:
    python lan_lookup.py --dump CJ2M716

    # 2) Then process the whole sheet:
    python lan_lookup.py users.xlsx
    python lan_lookup.py users.xlsx --column "LAN ID" --out users_with_dept.xlsx

    # Lookups are spaced out with random 2-6s waits and a 30-90s break every
    # 25 lookups; tune with --min-delay/--max-delay/--break-every/--break-min/--break-max.

    # Run several lookups at once (each worker keeps its own random pace):
    python lan_lookup.py users.xlsx --workers 4
"""
import argparse
import getpass
import re
import sys
import random
import threading
import time
from urllib.parse import quote

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ---------------------------------------------------------------- settings --
# Copy the host from your browser's address bar (the part before /wab/...).
BASE_URL = "https://njdirector.bbh.com/wab/person.asp?anr={anr}"

# The page is addressed by the user's LDAP path; {lan} is replaced by the LAN ID.
LDAP_PATH = "LDAP://bbhdom01.bbh.com/CN={lan},CN=Users,DC=bbhdom01,DC=bbh,DC=com"

# Labels on the page that hold the values we want (first match wins,
# case-insensitive). Run with --dump to see the real labels and adjust.
DEPARTMENT_LABELS = ["Department", "Dept", "Business Unit", "Division", "Group"]
REVIEWER_LABELS = ["Manager", "Reports To", "Supervisor", "Reviewer", "Approver"]

# Column names we auto-detect in the Excel sheet if --column isn't given.
LAN_COLUMN_GUESSES = ["lan id", "lanid", "lan", "user id", "userid", "username", "login", "id"]
# ------------------------------------------------------------------------------


def make_session_factory(user=None, ca_bundle=None):
    """Return a function that builds a logged-in requests.Session.
    Each worker thread gets its own session (login handshakes aren't thread-safe)."""
    try:  # use the Windows certificate store for internal TLS certs
        import truststore
        truststore.inject_into_ssl()
    except ImportError:
        if not ca_bundle:
            print("truststore not installed: internal certificates may fail "
                  "(pip install truststore)", file=sys.stderr)

    make_auth = None
    if user:  # explicit NTLM login (asks for the password once, never stores it on disk)
        from requests_ntlm import HttpNtlmAuth
        password = getpass.getpass(f"Password for {user}: ")
        make_auth = lambda: HttpNtlmAuth(user, password)
    else:     # Windows single sign-on: uses the account you're logged in with
        try:
            from requests_negotiate_sspi import HttpNegotiateAuth
            make_auth = HttpNegotiateAuth
        except ImportError:
            print("requests-negotiate-sspi not installed; trying without auth "
                  "(or pass --user DOMAIN\\you)", file=sys.stderr)

    def factory():
        s = requests.Session()
        if ca_bundle:
            s.verify = ca_bundle
        if make_auth:
            s.auth = make_auth()
        return s
    return factory


def build_url(lan):
    return BASE_URL.format(anr=quote(LDAP_PATH.format(lan=lan.upper()), safe=""))


def clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


def parse_fields(html):
    """Return {label: value} from the person page.

    Handles the usual layouts: <tr><td>Label</td><td>Value</td></tr>,
    <th>Label</th><td>Value</td>, <dt>/<dd>, and 'Label: Value' lines.
    Values with a link (e.g. the manager's name) use the link text.
    """
    soup = BeautifulSoup(html, "html.parser")
    fields = {}

    def add(label, value):
        label = clean(label).rstrip(":").strip()
        value = clean(value)
        if label and value and label not in fields and len(label) < 60:
            fields[label] = value

    for tr in soup.find_all("tr"):
        cells = tr.find_all(["td", "th"], recursive=False)
        if len(cells) >= 2:
            add(cells[0].get_text(" "), cells[1].get_text(" "))
    for dt in soup.find_all("dt"):
        dd = dt.find_next_sibling("dd")
        if dd:
            add(dt.get_text(" "), dd.get_text(" "))
    for line in soup.get_text("\n").splitlines():
        m = re.match(r"\s*([A-Za-z][\w /&.-]{1,40}?)\s*:\s+(.+)", line)
        if m:
            add(m.group(1), m.group(2))
    return fields


def pick(fields, labels):
    lower = {k.lower(): v for k, v in fields.items()}
    for want in labels:
        if want.lower() in lower:
            return lower[want.lower()]
    for want in labels:  # fall back to a partial match, e.g. "Department Name"
        for k, v in lower.items():
            if want.lower() in k:
                return v
    return ""


def tidy_person(value):
    """'1988 DiPlacido, Olimpio' -> 'DiPlacido, Olimpio' (drops a leading ID)."""
    return re.sub(r"^\d+\s+", "", value)


def human_pause(a, n_done, who=""):
    """Random wait between lookups, plus a longer break every so often,
    so requests come at a gentle, uneven pace instead of a rapid burst."""
    wait = random.uniform(a.min_delay, a.max_delay)
    if a.break_every and n_done and n_done % a.break_every == 0:
        wait += random.uniform(a.break_min, a.break_max)
        print(f"  ...{who}taking a {wait:.0f}s break after {n_done} lookups")
    time.sleep(wait)


def lookup(session, lan, timeout=30, retries=3):
    for attempt in range(retries + 1):
        r = session.get(build_url(lan), timeout=timeout)
        # Server says slow down / busy: back off (longer each time) and retry.
        if r.status_code in (429, 503) and attempt < retries:
            retry_after = r.headers.get("Retry-After", "")
            wait = int(retry_after) if retry_after.isdigit() else 30 * 2 ** attempt
            wait += random.uniform(0, 10)
            print(f"  server busy ({r.status_code}), waiting {wait:.0f}s before retrying")
            time.sleep(wait)
            continue
        break
    if r.status_code == 401:
        raise RuntimeError("401 Unauthorized - try --user DOMAIN\\yourid")
    r.raise_for_status()
    return parse_fields(r.text), r.text


def find_lan_column(df, wanted):
    if wanted:
        if wanted not in df.columns:
            sys.exit(f"Column '{wanted}' not found. Columns: {list(df.columns)}")
        return wanted
    norm = {str(c).strip().lower(): c for c in df.columns}
    for g in LAN_COLUMN_GUESSES:
        if g in norm:
            return norm[g]
    print(f"No LAN ID column recognised, using the first one: '{df.columns[0]}'",
          file=sys.stderr)
    return df.columns[0]


def main():
    global BASE_URL
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel", nargs="?", help="input .xlsx with a LAN ID column")
    ap.add_argument("--sheet", default=0, help="sheet name or index (default: first)")
    ap.add_argument("--column", help="name of the LAN ID column")
    ap.add_argument("--out", help="output .xlsx (default: <input>_with_dept.xlsx)")
    ap.add_argument("--base-url", help="override BASE_URL (must contain {anr})")
    ap.add_argument("--user", help="DOMAIN\\user for NTLM login instead of SSO")
    ap.add_argument("--ca-bundle", help="path to the corporate CA .pem if TLS fails")
    ap.add_argument("--min-delay", type=float, default=2.0,
                    help="shortest random wait between lookups, seconds (default 2)")
    ap.add_argument("--max-delay", type=float, default=6.0,
                    help="longest random wait between lookups, seconds (default 6)")
    ap.add_argument("--break-every", type=int, default=25,
                    help="take a longer break after this many lookups (0 = never; default 25)")
    ap.add_argument("--break-min", type=float, default=30.0, help="shortest break, seconds (default 30)")
    ap.add_argument("--break-max", type=float, default=90.0, help="longest break, seconds (default 90)")
    ap.add_argument("--workers", type=int, default=1,
                    help="lookups to run at the same time (default 1; 3-4 is a sensible max)")
    ap.add_argument("--dump", metavar="LANID",
                    help="look up one LAN ID, print every field found, save the HTML")
    a = ap.parse_args()

    if a.base_url:
        BASE_URL = a.base_url
    if "CHANGE-ME" in BASE_URL:
        sys.exit("Set BASE_URL at the top of the script (or pass --base-url) to the "
                 "host shown in your browser, e.g. https://xxx.bbh.com/wab/person.asp?anr={anr}")

    new_session = make_session_factory(a.user, a.ca_bundle)

    if a.dump:
        fields, html = lookup(new_session(), a.dump)
        with open(f"dump_{a.dump}.html", "w", encoding="utf-8") as f:
            f.write(html)
        print(f"URL: {build_url(a.dump)}\nSaved raw page to dump_{a.dump}.html\n")
        for k, v in fields.items():
            print(f"  {k!r:35} -> {v}")
        print(f"\nDepartment -> {pick(fields, DEPARTMENT_LABELS)!r}")
        print(f"Reviewer   -> {tidy_person(pick(fields, REVIEWER_LABELS))!r}")
        print("\nIf those are wrong, edit DEPARTMENT_LABELS / REVIEWER_LABELS "
              "using the labels listed above.")
        return

    if not a.excel:
        ap.error("give an Excel file, or use --dump LANID")

    sheet = int(a.sheet) if str(a.sheet).isdigit() else a.sheet
    df = pd.read_excel(a.excel, sheet_name=sheet, dtype=str)
    col = find_lan_column(df, a.column)
    out = a.out or re.sub(r"\.xls[xm]?$", "", a.excel, flags=re.I) + "_with_dept.xlsx"

    if a.max_delay < a.min_delay:
        a.max_delay = a.min_delay
    a.workers = max(1, a.workers)
    run_sheet(df, col, out, a, new_session)


def lookup_one(session, lan):
    """Look up one LAN ID -> (department, reviewer, status). Raises on errors."""
    fields, _ = lookup(session, lan)
    if not fields:
        return ("", "", "not found")
    d = pick(fields, DEPARTMENT_LABELS)
    rv = tidy_person(pick(fields, REVIEWER_LABELS))
    return (d, rv, "ok" if d or rv else "fields not matched")


def fix_hint(err):
    if "CERTIFICATE_VERIFY_FAILED" in err:
        return "\nFix: pip install truststore   (or pass --ca-bundle corp-ca.pem)"
    if "401" in err:
        return "\nFix: pip install requests-negotiate-sspi   (or pass --user DOMAIN\\you)"
    return ""


def run_sheet(df, col, out, a, new_session):
    lans = [str(x).strip() for x in df[col].fillna("")]
    total = len(lans)
    todo = list(dict.fromkeys(l.upper() for l in lans if l))  # unique IDs, in sheet order
    results = {}             # LAN ID (upper) -> (department, reviewer, status)
    lock = threading.Lock()
    stop = threading.Event()
    state = {"errors_in_a_row": 0, "done": 0, "last_error": ""}

    def save():
        """Write results so far; IDs not looked up yet are marked 'not done yet'."""
        rows = [("", "", "blank") if not l else results.get(l.upper(), ("", "", "not done yet"))
                for l in lans]
        out_df = df.copy()
        out_df["Department"] = [r[0] for r in rows]
        out_df["Reviewer"] = [r[1] for r in rows]
        out_df["Lookup Status"] = [r[2] for r in rows]
        try:
            out_df.to_excel(out, index=False)
        except PermissionError:
            print(f"  could not save {out} - is it open in Excel? Close it.", file=sys.stderr)

    def record(key, res):
        with lock:
            results[key] = res
            state["done"] += 1
            if res[2].startswith("error"):
                state["errors_in_a_row"] += 1
                state["last_error"] = res[2]
                if state["errors_in_a_row"] >= 3:
                    stop.set()
            else:
                state["errors_in_a_row"] = 0
            print(f"[{state['done']}/{len(todo)}] {key:12} {res[2]:10} {res[0]} | {res[1]}")
            if state["done"] % 20 == 0:  # save progress regularly
                save()

    local = threading.local()

    def worker(w, keys):
        who = f"worker {w}: " if a.workers > 1 else ""
        n = 0
        for key in keys:
            if stop.is_set():
                return
            if n:  # each worker keeps its own random pace and breaks
                human_pause(a, n, who)
                if stop.is_set():
                    return
            if not hasattr(local, "session"):
                local.session = new_session()
            try:
                res = lookup_one(local.session, key)
            except Exception as e:  # record the error in the sheet and carry on
                res = ("", "", f"error: {e}")
            record(key, res)
            n += 1

    print(f"{total} rows, {len(todo)} unique LAN IDs, {a.workers} worker(s)")
    try:
        if todo:  # try the first ID alone so a setup problem shows up at once
            worker(1, todo[:1])
            res = results[todo[0]]
            if res[2].startswith("error") and fix_hint(res[2]):
                stop.set()
                state["last_error"] = res[2]
        rest = todo[1:]
        if not stop.is_set() and rest:
            # deal the IDs out round-robin, one list per worker
            chunks = [rest[w::a.workers] for w in range(a.workers)]
            threads = [threading.Thread(target=worker, args=(w + 1, c), daemon=True)
                       for w, c in enumerate(chunks) if c]
            for t in threads:
                t.start()
            for t in threads:  # join with a timeout so Ctrl+C still works on Windows
                while t.is_alive():
                    t.join(0.5)
        if stop.is_set():
            err = state["last_error"]
            print(f"Stopping: lookups are failing. Last error:\n{err}{fix_hint(err)}",
                  file=sys.stderr)
    except KeyboardInterrupt:
        stop.set()
        print("\nStopped by you (Ctrl+C). Saving what's done so far...")
    finally:
        with lock:
            save()
            found = sum(1 for l in lans if l and results.get(l.upper(), ("", "", ""))[2] == "ok")
        print(f"\n{found}/{total} rows found. Saved to {out}")


if __name__ == "__main__":
    main()
