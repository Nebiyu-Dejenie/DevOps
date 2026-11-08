"""Sensitive-data and secret scan used before every commit/push (the repository is PUBLIC).
Usage:
  python3 -I sensitive_scan.py --staged          # scan files staged for commit
  python3 -I sensitive_scan.py --tree <commit>   # scan every file in a commit's tree
Exit code 0 = PASS, 1 = FAIL. Findings are printed MASKED (the private value is never echoed)."""
import io, os, re, subprocess, sys, tempfile, zipfile

RULES = {
    "applicant birth date": r"19\s*(?:th\s*)?(?:feb|february)\.?\s*,?\s*(?:1997|2027)|1997[-/.]0?2[-/.]19|19[-/.]0?2[-/.]1997|(?:feb|february)\s*19\s*,?\s*1997|turns?\s+30\s+on",
    "passport number": r"passport\s*(?:no\.?|number|#)\s*[:#]?\s*[A-Z]{1,2}\d{6,8}|\bEP\d{7}\b",
    "phone number": r"\+251[\s-]?\d{2,3}[\s-]?\d{3}[\s-]?\d{3,4}|(?<!\d)09\d{8}(?!\d)|(?<!\d)07\d{8}(?!\d)",
    "personal e-mail": r"[\w.+-]+@(?:gmail|yahoo|hotmail|outlook|icloud|proton(?:mail)?)\.(?:com|me)",
    "national/student ID": r"(?:fayda|national\s*id|student\s*id)\s*(?:no\.?|number|#)?\s*[:#]\s*[A-Z0-9/-]{5,}",
    "bank account": r"account\s*(?:no\.?|number)\s*[:#]?\s*\d{8,}",
    "third-party name (sanitised list)": r"\bKasahun\b|\bBiruk\b",
    "GitHub token": r"\b(?:ghp|gho|ghu|ghs|ghr)_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{30,}",
    "AWS key": r"\bAKIA[0-9A-Z]{16}\b",
    "private key": r"-----BEGIN (?:RSA |EC |OPENSSH |DSA |PGP )?PRIVATE KEY-----",
    "Slack token": r"\bxox[baprs]-[A-Za-z0-9-]{10,}",
    "Google API key": r"\bAIza[0-9A-Za-z_-]{35}\b",
    "OpenAI/Anthropic key": r"\bsk-(?:ant-)?[A-Za-z0-9_-]{32,}",
    "password assignment": r"(?i)\b(?:password|passwd|pwd)\s*[:=]\s*['\"]?[^\s'\"\[<]{6,}",
}
BLOCKED_PATHS = re.compile(r"(^|/)PRIVATE/|passport|_ID_scan|private_overlay\.yaml$|output_private/", re.I)
CRX = {k: re.compile(v, re.I) for k, v in RULES.items()}


def git(*a, binary=False):
    r = subprocess.run(["git", *a], capture_output=True)
    return r.stdout if binary else r.stdout.decode("utf-8", "replace")


def text_of(path, data):
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".docx", ".xlsx", ".pptx"):
            z = zipfile.ZipFile(io.BytesIO(data))
            return " ".join(re.sub(r"<[^>]+>", " ", z.read(n).decode("utf-8", "replace")) for n in z.namelist() if n.endswith(".xml"))
        if ext == ".pdf":
            with tempfile.NamedTemporaryFile(suffix=".pdf") as t:
                t.write(data); t.flush()
                return subprocess.run(["pdftotext", t.name, "-"], capture_output=True).stdout.decode("utf-8", "replace")
        if ext in (".png", ".jpg", ".jpeg", ".gif", ".zip", ".ico"):
            return ""
        return data.decode("utf-8", "replace")
    except Exception:
        return ""


def mask(s):
    return (s[:2] + re.sub(r"[A-Za-z0-9]", "•", s[2:12])) if s else s


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "--staged"
    files = []
    if mode == "--staged":
        for p in git("diff", "--cached", "--name-only", "--diff-filter=ACMR").splitlines():
            files.append((p, git("show", f":{p}", binary=True)))
    elif mode == "--tree":
        c = sys.argv[2]
        for line in git("ls-tree", "-r", c).splitlines():
            meta, p = line.split("\t", 1)
            files.append((p, git("cat-file", "-p", meta.split()[2], binary=True)))
    findings = []
    for p, data in files:
        if p.endswith("_sync/sensitive_scan.py"):
            continue  # the rule definitions themselves
        if BLOCKED_PATHS.search(p):
            findings.append((p, "blocked path (private file)", "path"))
            continue
        t = text_of(p, data)
        for name, rx in CRX.items():
            m = rx.search(t)
            if m:
                findings.append((p, name, mask(m.group(0))))
    print(f"scanned {len(files)} file(s)")
    for p, name, sample in findings:
        print(f"FAIL | {name} | {p} | {sample}")
    print("SENSITIVE-DATA SCAN:", "PASS" if not findings else "FAIL")
    sys.exit(1 if findings else 0)


if __name__ == "__main__":
    main()
