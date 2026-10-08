#!/usr/bin/env python3
"""ODI 360 codebase profiler - produces ONE redacted markdown report describing the
structure of the PBDW repos, so the odi360_extract.py parser can be tuned to them.

It does not copy code wholesale. For each repo it reports:
  - git summary (branch, commit count, date range, Jira key prefixes, sample subjects)
  - folder tree (depth-limited) with file counts, extension counts, file-name patterns
  - ODI XML: root tags, <Object class> counts, every <Field name> per class with fill-rate
    and a short REDACTED sample value, plus a skeleton of one scenario / load plan file
  - SQL: CREATE statement types per folder, encodings, a short redacted head per folder
  - scripts: keyword hit counts (startscen, startloadplan, sqlplus, insert_job, ...) with
    a few redacted example lines

Secrets are masked: password/key fields, IDENTIFIED BY, JDBC hosts, host names, IPs,
e-mails, long encrypted/base64 strings. Review the report before sharing it.

Usage (Windows):
    python odi360_profile.py --root "C:\\Users\\me\\bitbucket\\PBDW" --out odi360_profile.md
    python odi360_profile.py --root ... --no-samples      # structure only, no code lines
"""
import argparse
import datetime
import os
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict

SKIP_DIRS = {".git", "node_modules", "__pycache__", ".idea", ".vscode", "target", ".svn"}
SQL_EXT = {".sql", ".pks", ".pkb", ".pck", ".prc", ".fnc", ".vw", ".trg", ".tab", ".seq", ".idx", ".ddl"}
KEYWORDS = {
    "startscen": r"startscen",
    "startloadplan": r"startloadplan",
    "OdiStartScen": r"OdiStartScen|OdiStartLoadPlan",
    "odiRef": r"odiRef\.",
    "odi_agent_url": r"oraclediagent|AGENT_URL",
    "sqlplus": r"\bsqlplus\b",
    "sqlldr": r"\bsqlldr\b",
    "autosys_jil": r"\binsert_job\s*:",
    "autosys_cmd": r"\b(sendevent|autorep|jil)\b",
    "jdbc": r"jdbc:",
    "create_table": r"\bcreate\s+(global\s+temporary\s+)?table\b",
    "create_or_replace": r"\bcreate\s+or\s+replace\b",
    "merge_into": r"\bmerge\s+into\b",
    "insert_into": r"\binsert\s+(/\*.*?\*/\s*)?into\b",
    "truncate": r"\btruncate\s+table\b",
    "exec_proc": r"\bexec(ute)?\s+\w+\.",
    "db_link": r"@\w+[\s;)]",
    "curl_http": r"\bcurl\b|https?://",
    "sftp_ftp": r"\b(sftp|ftp|scp)\b",
    "python": r"\bpython\d?\b",
    "java": r"\bjava\s+-",
    "cron": r"\bcrontab\b",
    "spring_boot": r"@SpringBootApplication|@Scheduled",
    "kafka_mq": r"\b(kafka|jms|ibm\.mq|MQQueue)\b",
}

# ---------------------------------------------------------------- redaction
_R = [
    (re.compile(r"(?i)(password|passwd|pwd|secret|token|apikey|api_key)(\s*[=:]\s*|\s+)(['\"]?)[^\s'\";,<>]+"), r"\1\2\3<REDACTED>"),
    (re.compile(r"(?i)identified\s+by\s+\S+"), "IDENTIFIED BY <REDACTED>"),
    (re.compile(r"(?i)(jdbc:[\w:]+(?:thin:)?)(@|//)[^\s'\"<>;]+"), r"\1\2<HOST>:<PORT>/<SERVICE>"),
    (re.compile(r"(?i)(-u|--user)\s+\S+:\S+"), r"\1 <USER>:<REDACTED>"),
    (re.compile(r"(?i)\b[\w.+-]+@[\w-]+\.[\w.]+\b"), "<EMAIL>"),
    (re.compile(r"\b\d{1,3}(?:\.\d{1,3}){3}\b"), "<IP>"),
    (re.compile(r"(?i)\b(?:[a-z0-9-]+\.){2,}(?:com|net|org|local|corp|internal|int)\b"), "<HOST>"),
    (re.compile(r"(?i)\\\\[\w.-]+\\"), r"\\\\<HOST>\\"),
    (re.compile(r"\b[A-Za-z0-9+/]{32,}={0,2}"), "<ENC>"),
]
SENSITIVE_FIELD = re.compile(r"(?i)pass|pwd|secret|token|cred|encrypt|hash|^user(name)?$|^user|key(value|hash)?$")


def redact(s):
    for rx, rep in _R:
        s = rx.sub(rep, s)
    return s


def short(s, n=90):
    s = re.sub(r"\s+", " ", (s or "")).strip()
    return s if len(s) <= n else s[: n - 3] + "..."


# ---------------------------------------------------------------- helpers
def read_text(path, max_bytes=2_000_000):
    try:
        with open(path, "rb") as f:
            raw = f.read(max_bytes)
    except OSError:
        return "", "unreadable"
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16", "replace"), "utf-16"
    if b"\x00" in raw[:4096]:
        return "", "binary"
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            pass
    return raw.decode("latin-1", "replace"), "latin-1"


def walk(root):
    for dp, dns, fns in os.walk(root):
        dns[:] = sorted(d for d in dns if d not in SKIP_DIRS)
        for fn in sorted(fns):
            yield os.path.join(dp, fn)


def rel(p, root):
    return os.path.relpath(p, root).replace("\\", "/")


def name_pattern(fn):
    """SCEN_SCN_FOO_BAR_Version_001.xml -> SCEN_SCN_<X>_Version_<N>.xml"""
    b, e = os.path.splitext(fn)
    b = re.sub(r"\d+", "<N>", b)
    parts = b.split("_")
    if len(parts) > 3:
        b = "_".join(parts[:2]) + "_<X>" + ("_" + "_".join(parts[-2:]) if "Version" in parts else "")
    return b + e


def git_summary(root):
    def run(*a):
        try:
            return subprocess.run(["git", "-C", root, *a], capture_output=True, text=True,
                                  encoding="utf-8", errors="replace", timeout=300).stdout.strip()
        except (OSError, subprocess.SubprocessError):
            return ""
    if not run("rev-parse", "--is-inside-work-tree"):
        return None
    log = run("log", "--format=%ad\x1f%s", "--date=short")
    rows = [l.split("\x1f", 1) for l in log.splitlines() if "\x1f" in l]
    subjects = [r[1] for r in rows]
    jira = Counter(m.split("-")[0] for s in subjects for m in re.findall(r"\b[A-Z][A-Z0-9]{1,9}-\d+\b", s))
    return {
        "branch": run("rev-parse", "--abbrev-ref", "HEAD"),
        "branches": len(run("branch", "-a").splitlines()),
        "shallow": run("rev-parse", "--is-shallow-repository"),
        "commits": len(rows),
        "first": rows[-1][0] if rows else "", "last": rows[0][0] if rows else "",
        "jira_prefixes": dict(jira.most_common(8)),
        "with_jira_pct": round(100 * sum(1 for s in subjects if re.search(r"\b[A-Z][A-Z0-9]{1,9}-\d+\b", s)) / max(len(subjects), 1)),
        "sample_subjects": [redact(short(s, 80)) for s in dict.fromkeys(subjects[:40])][:12],
        "authors": len(set(run("log", "--format=%an").splitlines())),
    }


# ---------------------------------------------------------------- XML profiling
class XmlProfile:
    def __init__(self):
        self.roots = Counter()
        self.classes = Counter()                 # class -> object count
        self.class_files = defaultdict(set)      # class -> folders seen
        self.fields = defaultdict(Counter)       # class -> field -> count
        self.field_types = defaultdict(dict)     # class -> field -> type attr
        self.samples = defaultdict(dict)         # class -> field -> sample
        self.max_len = defaultdict(Counter)      # class -> field -> max value length
        self.other_tags = Counter()
        self.skeletons = {}                      # kind -> (file, [lines])
        self.errors = []
        self.files = 0

    def add(self, path, relpath, folder):
        self.files += 1
        seq, cur, n_obj = [], None, 0
        try:
            for ev, el in ET.iterparse(path, events=("start", "end")):
                tag = el.tag.rsplit("}", 1)[-1]
                if ev == "start":
                    if tag == "Object":
                        cur = (el.get("class") or "?").rsplit(".", 1)[-1]
                        n_obj += 1
                        self.classes[cur] += 1
                        self.class_files[cur].add(folder)
                        seq.append([cur, []])
                    elif tag not in ("Field",) and cur is None:
                        self.other_tags[tag] += 1
                    continue
                if tag == "Field" and cur:
                    name = el.get("name") or "?"
                    val = el.text or ""
                    self.fields[cur][name] += 1
                    self.field_types[cur].setdefault(name, (el.get("type") or "").rsplit(".", 1)[-1])
                    self.max_len[cur][name] = max(self.max_len[cur][name], len(val))
                    if val.strip() and name not in self.samples[cur]:
                        self.samples[cur][name] = "<REDACTED>" if SENSITIVE_FIELD.search(name) else redact(short(val, 70))
                    if len(seq[-1][1]) < 40:
                        seq[-1][1].append((name, "<REDACTED>" if SENSITIVE_FIELD.search(name) else redact(short(val, 50))))
                elif tag == "Object":
                    cur = None
                    el.clear()
        except ET.ParseError as e:
            self.errors.append(f"{relpath}: {e}")
            return
        except OSError as e:
            self.errors.append(f"{relpath}: {e}")
            return
        # remember root tag
        try:
            with open(path, "rb") as f:
                head = f.read(600).decode("utf-8", "replace")
            m = re.search(r"<([A-Za-z_][\w.:-]*)[\s>]", re.sub(r"<\?.*?\?>|<!--.*?-->", "", head, flags=re.S))
            self.roots[m.group(1) if m else "?"] += 1
        except OSError:
            pass
        classes = {c for c, _ in seq}
        kind = ("load_plan" if any("LoadPlan" in c for c in classes) else
                "scenario" if "SnpScen" in classes else
                "mapping" if "SnpMapping" in classes else
                "package" if "SnpPackage" in classes else
                "procedure_or_km" if "SnpTrt" in classes else
                "topology" if classes & {"SnpConnect", "SnpContext", "SnpPschema", "SnpLschema", "SnpAgent"} else
                "other_odi" if classes else "non_odi_xml")
        if kind not in self.skeletons or len(seq) > len(self.skeletons[kind][1]) and len(seq) < 60:
            self.skeletons[kind] = (relpath, seq[:60])


# ---------------------------------------------------------------- repo profiling
def profile_repo(root, args, out):
    name = os.path.basename(os.path.abspath(root))
    w = out.append
    w(f"\n---\n\n# REPO `{name}`\n")
    gs = None if args.no_git else git_summary(root)
    if gs:
        w(f"- git: branch `{gs['branch']}`, {gs['commits']} commits ({gs['first']} -> {gs['last']}), "
          f"{gs['authors']} authors, {gs['branches']} branches, shallow={gs['shallow']}")
        w(f"- Jira keys in {gs['with_jira_pct']}% of commit subjects; prefixes: {gs['jira_prefixes']}")
        w("- sample commit subjects: " + " | ".join(f"`{s}`" for s in gs["sample_subjects"]))
    else:
        w("- git: not a git repo / git not on PATH")

    ext = Counter()
    size = 0
    dir_files = Counter()
    dir_ext = defaultdict(Counter)
    dir_names = defaultdict(list)
    xmlp = XmlProfile()
    sql_create = defaultdict(Counter)
    sql_heads = {}
    encodings = Counter()
    kw_hits = Counter()
    kw_files = Counter()
    kw_examples = defaultdict(list)
    kw_rx = {k: re.compile(v, re.I) for k, v in KEYWORDS.items()}
    big = []
    n = 0
    for path in walk(root):
        n += 1
        r = rel(path, root)
        parts = r.split("/")
        folder = "/".join(parts[: min(len(parts) - 1, args.depth)]) or "."
        e = os.path.splitext(path)[1].lower() or "(none)"
        ext[e] += 1
        try:
            sz = os.path.getsize(path)
        except OSError:
            sz = 0
        size += sz
        big.append((sz, r))
        dir_files[folder] += 1
        dir_ext[folder][e] += 1
        if len(dir_names[folder]) < 400:
            dir_names[folder].append(os.path.basename(path))
        if e == ".xml":
            xmlp.add(path, r, folder)
            continue
        if sz > args.max_file_mb * 1_000_000 or e in (".jar", ".zip", ".gz", ".class", ".png", ".jpg",
                                                          ".xlsx", ".xls", ".pdf", ".docx", ".exe", ".dll"):
            continue
        text, enc = read_text(path)
        encodings[enc] += 1
        if not text:
            continue
        if e in SQL_EXT or (e in (".txt", "(none)") and re.search(r"(?i)\bcreate\s", text[:5000])):
            for m in re.finditer(r"(?i)\bcreate\s+(?:or\s+replace\s+)?(?:(?:non)?editionable\s+)?(?:force\s+)?"
                                 r"(?:global\s+temporary\s+)?(?:unique\s+|bitmap\s+)?(table|materialized\s+view|view|"
                                 r"procedure|function|package\s+body|package|sequence|index|trigger|synonym|type)\b", text):
                sql_create[folder][re.sub(r"\s+", " ", m.group(1).upper())] += 1
            if not re.search(r"(?i)\bcreate\s", text):
                sql_create[folder]["(no CREATE)"] += 1
            if folder not in sql_heads and not args.no_samples:
                lines = [redact(l.rstrip())[:160] for l in text.splitlines() if l.strip()][: args.head_lines]
                sql_heads[folder] = (r, lines)
        for k, rx in kw_rx.items():
            c = len(rx.findall(text))
            if c:
                kw_hits[k] += c
                kw_files[k] += 1
                if len(kw_examples[k]) < args.examples and not args.no_samples:
                    for line in text.splitlines():
                        if rx.search(line):
                            kw_examples[k].append(f"{r}: {redact(short(line, 150))}")
                            break

    w(f"- files: {n}, size: {size / 1e6:.1f} MB")
    w(f"- extensions: {dict(ext.most_common(20))}")
    if encodings:
        w(f"- text encodings: {dict(encodings)}")
    w("- largest files: " + ", ".join(f"`{r}` ({s / 1e6:.1f} MB)" for s, r in sorted(big, reverse=True)[:5]))

    w(f"\n## Folder tree (depth {args.depth})\n")
    w("| folder | files | extensions | file-name patterns (examples) |\n|---|---|---|---|")
    for folder in sorted(dir_files):
        pats = Counter(name_pattern(x) for x in dir_names[folder])
        ex = "; ".join(f"{p} x{c}" for p, c in pats.most_common(4))
        names = ", ".join(dir_names[folder][:3])
        w(f"| `{folder}` | {dir_files[folder]} | {dict(dir_ext[folder].most_common(4))} | {ex}<br>e.g. {redact(names)} |")

    if xmlp.files:
        w(f"\n## XML ({xmlp.files} files)\n")
        w(f"- root tags: {dict(xmlp.roots.most_common(10))}")
        if xmlp.other_tags:
            w(f"- non-Object tags: {dict(xmlp.other_tags.most_common(15))}")
        if xmlp.errors:
            w(f"- parse errors ({len(xmlp.errors)}): " + "; ".join(xmlp.errors[:5]))
        if xmlp.classes:
            w("\n### ODI object classes and fields\n")
            w("Fields listed as `name:type(count, maxlen)=sample`.\n")
            for cls, cnt in xmlp.classes.most_common():
                w(f"**{cls}**: {cnt} objects, in folders: {', '.join(sorted(xmlp.class_files[cls])[:6])}")
                fl = []
                for f, c in xmlp.fields[cls].most_common():
                    s = xmlp.samples[cls].get(f, "")
                    if args.no_samples and s:
                        s = "..."
                    fl.append(f"`{f}`:{xmlp.field_types[cls].get(f, '')}({c},{xmlp.max_len[cls][f]})"
                              + (f"=`{s}`" if s else ""))
                w("  - " + " · ".join(fl) + "\n")
            w("\n### Skeletons (one representative file per kind)\n")
            for kind, (f, seq) in sorted(xmlp.skeletons.items()):
                if kind == "non_odi_xml":
                    continue
                w(f"**{kind}**: `{f}`\n```")
                for cls, fields in seq:
                    body = ", ".join(f"{k}={'' if args.no_samples else v}" for k, v in fields[:25])
                    w(f"{cls}: {body}")
                w("```")

    if sql_create:
        w("\n## SQL / DDL\n")
        w("| folder | CREATE statements by type |\n|---|---|")
        for folder in sorted(sql_create):
            w(f"| `{folder}` | {dict(sql_create[folder].most_common())} |")
        for folder, (f, lines) in sorted(sql_heads.items())[: args.max_sql_samples]:
            w(f"\n`{f}` (first {len(lines)} non-blank lines, redacted)\n```sql")
            w("\n".join(lines))
            w("```")

    if kw_hits:
        w("\n## Keyword scan\n")
        w("| keyword | hits | files |\n|---|---|---|")
        for k, c in kw_hits.most_common():
            w(f"| {k} | {c} | {kw_files[k]} |")
        if not args.no_samples:
            w("\nExamples (redacted):\n```")
            for k in kw_hits:
                for x in kw_examples[k]:
                    w(f"[{k}] {x}")
            w("```")


def main():
    ap = argparse.ArgumentParser(description="Redacted structure profile of the PBDW repos")
    ap.add_argument("--root", required=True, help="folder that holds the cloned repos")
    ap.add_argument("--repos", nargs="*", help="only these sub-folders (default: all)")
    ap.add_argument("--out", default="odi360_profile.md")
    ap.add_argument("--depth", type=int, default=3, help="folder depth for the tree table")
    ap.add_argument("--examples", type=int, default=3, help="example lines per keyword")
    ap.add_argument("--head-lines", type=int, default=25, help="lines shown per SQL sample")
    ap.add_argument("--max-sql-samples", type=int, default=12)
    ap.add_argument("--max-file-mb", type=float, default=20)
    ap.add_argument("--no-samples", action="store_true", help="structure only - no code lines or values")
    ap.add_argument("--no-git", action="store_true")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    repos = [os.path.join(root, r) for r in (args.repos or sorted(os.listdir(root)))
             if os.path.isdir(os.path.join(root, r)) and r not in SKIP_DIRS]
    if not repos:
        repos = [root]
    out = ["# ODI 360 codebase profile",
           f"Generated {datetime.datetime.now():%Y-%m-%d %H:%M}, python {sys.version.split()[0]}, "
           f"samples={'off' if args.no_samples else 'on (redacted)'}",
           f"Repos: {', '.join(os.path.basename(r) for r in repos)}"]
    for r in repos:
        print(f"[profile] {os.path.basename(r)} ...", file=sys.stderr)
        profile_repo(r, args, out)
    text = "\n".join(out) + "\n"
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"[profile] wrote {args.out} ({len(text) / 1000:.0f} KB). Review it before sharing.", file=sys.stderr)


if __name__ == "__main__":
    main()
