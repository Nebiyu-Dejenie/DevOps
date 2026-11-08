"""Scan EVERY commit on every local branch for private data, including inside .docx/.xlsx/.pdf files.
Usage (from the repository root): python3 -I SCHOLARSHIP_PROJECT/DASHBOARD/_build/privacy_audit_history.py <out.md>
Findings are MASKED in the report so the report itself never repeats the private value."""
import io, os, re, subprocess, sys, tempfile, zipfile
from collections import defaultdict

OUT = sys.argv[1]
PATTERNS = {
    "birth date / age text": [r"19\s*(?:th\s*)?(?:feb|february)\.?\s*,?\s*1997", r"1997[-/.]0?2[-/.]19", r"19[-/.]0?2[-/.]1997",
                              r"(?:feb|february)\s*19\s*,?\s*1997", r"turns?\s+30\s+on", r"\bborn\s+on\b", r"date\s+of\s+birth\s*[:\-]\s*\S+",
                              r"\bDOB\s*[:\-]\s*\S+", r"19\s*(?:feb|february)\s*2027"],
    "passport number": [r"passport\s*(?:no\.?|number|#)\s*[:#]?\s*[A-Z]{1,2}\d{6,8}", r"\bEP\d{7}\b"],
    "phone number": [r"\+251[\s-]?\d{2,3}[\s-]?\d{3}[\s-]?\d{3,4}", r"(?<!\d)09\d{8}(?!\d)", r"(?<!\d)07\d{8}(?!\d)"],
    "personal e-mail": [r"[\w.+-]+@(?:gmail|yahoo|hotmail|outlook)\.com"],
    "student / national ID": [r"(?:student\s*id|id\s*no\.?|registration\s*no\.?|fayda)\s*[:#]?\s*[A-Z0-9/-]{5,}"],
    "home address": [r"\bkebele\s*\d+", r"\bhouse\s*no\.?\s*\d+", r"\bworeda\s*\d+"],
    "bank account": [r"account\s*(?:no\.?|number)\s*[:#]?\s*\d{6,}"],
}
CRX = {k: [re.compile(p, re.I) for p in v] for k, v in PATTERNS.items()}


def git(*a, binary=False):
    r = subprocess.run(["git", *a], capture_output=True)
    return r.stdout if binary else r.stdout.decode("utf-8", "replace")


def text_of(path, data):
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".docx", ".xlsx", ".pptx"):
            z = zipfile.ZipFile(io.BytesIO(data))
            parts = [n for n in z.namelist() if n.endswith(".xml")]
            return " ".join(re.sub(r"<[^>]+>", " ", z.read(n).decode("utf-8", "replace")) for n in parts)
        if ext == ".pdf":
            with tempfile.NamedTemporaryFile(suffix=".pdf") as t:
                t.write(data); t.flush()
                return subprocess.run(["pdftotext", t.name, "-"], capture_output=True).stdout.decode("utf-8", "replace")
        if ext in (".png", ".jpg", ".jpeg", ".gif", ".zip"):
            return ""
        return data.decode("utf-8", "replace")
    except Exception:
        return ""


def mask(s):
    return re.sub(r"[A-Za-z0-9]", "•", s[:3]) + "…" if s else s


commits = git("rev-list", "--all").split()
blob_paths, blob_commits = defaultdict(set), defaultdict(set)
for c in commits:
    for line in git("ls-tree", "-r", c).splitlines():
        meta, path = line.split("\t", 1)
        _, typ, sha = meta.split()
        if typ == "blob":
            blob_paths[sha].add(path); blob_commits[sha].add(c[:7])
pushed = set(git("rev-list", "--remotes").split())
findings = []
for sha, paths in blob_paths.items():
    p = sorted(paths)[0]
    txt = text_of(p, git("cat-file", "-p", sha, binary=True))
    if not txt:
        continue
    for kind, rxs in CRX.items():
        n = sum(len(rx.findall(txt)) for rx in rxs)
        if n:
            first = next((m.group(0) for rx in rxs for m in [rx.search(txt)] if m), "")
            findings.append((kind, n, sorted(paths), sorted(blob_commits[sha]), mask(first)))
in_pushed = lambda cs: any(any(full.startswith(x) for full in pushed) for x in cs)
L = ["# Git history privacy audit", "",
     f"Commits scanned: **{len(commits)}** (all local branches). Unique file versions scanned: **{len(blob_paths)}**, including text inside .docx, .xlsx and .pdf files.",
     "Matched values are masked in this report.", ""]
if not findings:
    L.append("**No private data patterns found in any commit.**")
else:
    L += ["| Type | Matches | File(s) | In commits | Already on the remote? | Masked sample |", "|---|---|---|---|---|---|"]
    for kind, n, paths, cs, sample in sorted(findings):
        L.append(f"| {kind} | {n} | {'; '.join(paths)[:160]} | {', '.join(cs)[:80]} | {'YES' if in_pushed(cs) else 'no (local only)'} | {sample} |")
open(OUT, "w").write("\n".join(L) + "\n")
print(len(commits), "commits;", len(blob_paths), "blobs;", len(findings), "findings")
for f in sorted(findings):
    print(f[0], "|", f[1], "|", "; ".join(f[2])[:120], "|", ",".join(f[3])[:60], "| pushed:", in_pushed(f[3]))
