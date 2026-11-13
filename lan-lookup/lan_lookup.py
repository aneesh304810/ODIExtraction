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
"""
import argparse
import getpass
import re
import sys
import random
import time
from urllib.parse import quote

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ---------------------------------------------------------------- settings --
# Copy the host from your browser's address bar (the part before /wab/...).
BASE_URL = "https://CHANGE-ME.bbh.com/wab/person.asp?anr={anr}"

# The page is addressed by the user's LDAP path; {lan} is replaced by the LAN ID.
LDAP_PATH = "LDAP://bbhdom01.bbh.com/CN={lan},CN=Users,DC=bbhdom01,DC=bbh,DC=com"

# Labels on the page that hold the values we want (first match wins,
# case-insensitive). Run with --dump to see the real labels and adjust.
DEPARTMENT_LABELS = ["Department", "Dept", "Business Unit", "Division", "Group"]
REVIEWER_LABELS = ["Manager", "Reports To", "Supervisor", "Reviewer", "Approver"]

# Column names we auto-detect in the Excel sheet if --column isn't given.
LAN_COLUMN_GUESSES = ["lan id", "lanid", "lan", "user id", "userid", "username", "login", "id"]
# ------------------------------------------------------------------------------


def make_session(user=None, ca_bundle=None):
    try:  # use the Windows certificate store for internal TLS certs
        import truststore
        truststore.inject_into_ssl()
    except ImportError:
        pass

    s = requests.Session()
    if ca_bundle:
        s.verify = ca_bundle
    if user:  # explicit NTLM login (asks for the password, never stores it)
        from requests_ntlm import HttpNtlmAuth
        s.auth = HttpNtlmAuth(user, getpass.getpass(f"Password for {user}: "))
    else:     # Windows single sign-on: uses the account you're logged in with
        try:
            from requests_negotiate_sspi import HttpNegotiateAuth
            s.auth = HttpNegotiateAuth()
        except ImportError:
            print("requests-negotiate-sspi not installed; trying without auth "
                  "(or pass --user DOMAIN\\you)", file=sys.stderr)
    return s


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


def human_pause(a, n_done):
    """Random wait between lookups, plus a longer break every so often,
    so requests come at a gentle, uneven pace instead of a rapid burst."""
    wait = random.uniform(a.min_delay, a.max_delay)
    if a.break_every and n_done and n_done % a.break_every == 0:
        wait += random.uniform(a.break_min, a.break_max)
        print(f"  ...taking a {wait:.0f}s break after {n_done} lookups")
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
    ap.add_argument("--dump", metavar="LANID",
                    help="look up one LAN ID, print every field found, save the HTML")
    a = ap.parse_args()

    if a.base_url:
        BASE_URL = a.base_url
    if "CHANGE-ME" in BASE_URL:
        sys.exit("Set BASE_URL at the top of the script (or pass --base-url) to the "
                 "host shown in your browser, e.g. https://xxx.bbh.com/wab/person.asp?anr={anr}")

    session = make_session(a.user, a.ca_bundle)

    if a.dump:
        fields, html = lookup(session, a.dump)
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
    cache, depts, reviewers, statuses = {}, [], [], []
    total = len(df)
    for i, raw in enumerate(df[col].fillna(""), 1):
        lan = str(raw).strip()
        if not lan:
            depts.append(""); reviewers.append(""); statuses.append("blank")
            continue
        key = lan.upper()
        if key not in cache:
            if cache:  # pause before every lookup except the first
                human_pause(a, len(cache))
            try:
                fields, _ = lookup(session, lan)
                if not fields:
                    cache[key] = ("", "", "not found")
                else:
                    d = pick(fields, DEPARTMENT_LABELS)
                    rv = tidy_person(pick(fields, REVIEWER_LABELS))
                    cache[key] = (d, rv, "ok" if d or rv else "fields not matched")
            except Exception as e:  # keep going; record the error in the sheet
                cache[key] = ("", "", f"error: {e}")
        d, rv, st = cache[key]
        depts.append(d); reviewers.append(rv); statuses.append(st)
        print(f"[{i}/{total}] {lan:12} {st:10} {d} | {rv}")

    df["Department"], df["Reviewer"], df["Lookup Status"] = depts, reviewers, statuses
    df.to_excel(out, index=False)
    ok = statuses.count("ok")
    print(f"\nDone: {ok}/{total} found. Saved to {out}")


if __name__ == "__main__":
    main()
