#!/usr/bin/env python3
"""Profile-driven document generator for the scholarship project.

Commands (run from this folder):
  python3 generate.py status                 # which documents are outdated, pending facts, review state
  python3 generate.py build <doc_id|all>     # render drafts (never overwrites reviewed/final versions)
  python3 generate.py build <doc_id> --private   # merge private_overlay.yaml; output goes to output_private/ (git-ignored)
  python3 generate.py review <doc_id>        # mark the latest version as reviewed (protected from overwriting)
  python3 generate.py final <doc_id>         # mark as final (refused while pending facts remain)
  python3 generate.py diff <doc_id>          # list profile facts that changed since the latest version was built

Rules built in:
  * Facts are {v, ev, note}. E5 facts are never rendered. Pending facts render as "[⏳ note]".
  * Every build records which profile facts the document used (with hashes), the template hash and the version.
  * A reviewed or final version is never overwritten: rebuilding creates the next version number.
"""
import hashlib, json, os, re, signal, subprocess, sys
from datetime import datetime, timezone

import yaml
from jinja2 import Environment, FileSystemLoader, StrictUndefined

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE = os.path.join(HERE, "..", "profile", "master_profile.yaml")
PRIVATE = os.path.join(HERE, "..", "profile", "private_overlay.yaml")
REGISTRY = os.path.join(HERE, "documents.yaml")
MANIFEST = os.path.join(HERE, "manifest.json")
TEMPLATES = os.path.join(HERE, "templates")


def sha(obj):
    return hashlib.sha256(json.dumps(obj, sort_keys=True, ensure_ascii=False, default=str).encode()).hexdigest()[:16]


def load_yaml(p):
    with open(p, encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def deep_merge(a, b):
    out = dict(a)
    for k, v in b.items():
        out[k] = deep_merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


def resolve(data, path):
    cur = data
    for part in re.findall(r"[^.\[\]]+|\[\d+\]", path):
        if part.startswith("["):
            i = int(part[1:-1])
            if not isinstance(cur, list) or i >= len(cur):
                return "<missing>"
            cur = cur[i]
        else:
            if not isinstance(cur, dict) or part not in cur:
                return "<missing>"
            cur = cur[part]
    return cur


class T:
    """Wraps profile data and records which paths a template reads."""

    def __init__(self, data, path, log):
        self._d, self._p, self._log = data, path, log

    def __getattr__(self, k):
        if k.startswith("_"):
            raise AttributeError(k)
        try:
            return self[k]
        except (KeyError, IndexError, TypeError):
            raise AttributeError(k)

    def __getitem__(self, k):
        d = self._d
        if isinstance(d, dict):
            if k not in d:
                raise KeyError(k)
            return wrap(d[k], f"{self._p}.{k}" if self._p else str(k), self._log)
        if isinstance(d, list):
            return wrap(d[k], f"{self._p}[{k}]", self._log)
        raise KeyError(k)

    def __iter__(self):
        self._log.add(self._p)
        if isinstance(self._d, list):
            for i, x in enumerate(self._d):
                yield wrap(x, f"{self._p}[{i}]", self._log)
        else:
            for k in self._d:
                yield k

    def __len__(self):
        self._log.add(self._p)
        return len(self._d)

    def __bool__(self):
        return bool(self._d)

    def __str__(self):
        self._log.add(self._p)
        return str(self._d)


def wrap(x, p, log):
    if isinstance(x, (dict, list)):
        return T(x, p, log)
    log.add(p)
    return x


class Render:
    def __init__(self):
        self.pending, self.e4 = [], []

    def fact(self, x):
        """Render one fact {v, ev, note}; record pending and E4 items; never render E5."""
        if isinstance(x, T):
            x._log.add(x._p)
            path, raw = x._p, x._d
        else:
            path, raw = "?", x
        if isinstance(raw, dict) and "ev" in raw:
            ev, v, note = str(raw.get("ev")), raw.get("v"), raw.get("note", "")
            if ev == "E5":
                return ""
            if ev == "Pending" or v in (None, ""):
                self.pending.append(f"{path}: {note or 'missing'}")
                return f"[⏳ {note or path}]"
            if ev == "E4":
                self.e4.append(f"{path}: {v}" + (f" ({note})" if note else ""))
            return str(v)
        return "" if raw is None else str(raw)

    def usable(self, x):
        """True if a fact may appear (not E5)."""
        raw = x._d if isinstance(x, T) else x
        return not (isinstance(raw, dict) and str(raw.get("ev")) == "E5")


def load_manifest():
    if os.path.exists(MANIFEST):
        with open(MANIFEST, encoding="utf-8") as fh:
            return json.load(fh)
    return {}


def save_manifest(m):
    with open(MANIFEST, "w", encoding="utf-8") as fh:
        json.dump(m, fh, indent=2, ensure_ascii=False)


def template_sha(name):
    with open(os.path.join(TEMPLATES, name), encoding="utf-8") as fh:
        return hashlib.sha256(fh.read().encode()).hexdigest()[:16]


def render_doc(doc, profile):
    log = set()
    r = Render()
    env = Environment(loader=FileSystemLoader(TEMPLATES), undefined=StrictUndefined, keep_trailing_newline=True)
    env.filters["f"] = r.fact
    env.filters["lcfirst"] = lambda s: (s[:1].lower() + s[1:]) if s and not s.startswith("[") else s
    env.filters["ucfirst"] = lambda s: (s[:1].upper() + s[1:]) if s else s
    env.tests["usable"] = r.usable
    ctx = {"profile": T(profile, "", log), "doc": doc, "today": datetime.now().strftime("%d %B %Y")}
    if doc.get("programme"):
        ctx["prog"] = T(profile["programmes"][doc["programme"]], f"programmes.{doc['programme']}", log)
    body = env.get_template(doc["template"]).render(**ctx)
    deps = {p: sha(resolve(profile, p)) for p in sorted(log) if p}
    return body, deps, r


def footer(r):
    lines = ["", "---", "", "**Generator check (delete before submitting):**", ""]
    if r.pending:
        lines.append(f"- Pending facts ({len(r.pending)}): " + "; ".join(sorted(set(r.pending))))
    else:
        lines.append("- No pending facts.")
    if r.e4:
        lines.append(f"- Based on your confirmation only (E4), check before use ({len(set(r.e4))}): " + "; ".join(sorted(set(r.e4))))
    lines.append("- E5 facts are never included.")
    return "\n".join(lines) + "\n"


def cmd_build(target, private=False, force_new=False):
    registry = load_yaml(REGISTRY)["documents"]
    profile = load_yaml(PROFILE)
    if private:
        if not os.path.exists(PRIVATE):
            sys.exit("private_overlay.yaml not found (copy private_overlay.example.yaml first).")
        profile = deep_merge(profile, load_yaml(PRIVATE))
    out_root = os.path.join(HERE, "output_private" if private else "output")
    man = load_manifest()
    ids = [d["id"] for d in registry] if target == "all" else [target]
    for did in ids:
        doc = next((d for d in registry if d["id"] == did), None)
        if not doc:
            print(f"unknown document: {did}")
            continue
        body, deps, r = render_doc(doc, profile)
        body = body.rstrip() + "\n" + footer(r)
        key = did + ("__private" if private else "")
        hist = man.setdefault(key, {"versions": []})["versions"]
        latest = hist[-1] if hist else None
        if latest and latest["state"] == "draft" and not force_new:
            n = latest["v"]
        else:
            n = (latest["v"] + 1) if latest else 1
        folder = os.path.join(out_root, did)
        os.makedirs(folder, exist_ok=True)
        md_path = os.path.join(folder, f"{did}_v{n}.md")
        with open(md_path, "w", encoding="utf-8") as fh:
            fh.write(body)
        docx_path = md_path[:-3] + ".docx"
        try:
            ref = os.path.join(HERE, "reference_compact.docx")
            subprocess.run(["pandoc", md_path, "-o", docx_path] + (["--reference-doc", ref] if os.path.exists(ref) else []), check=True, capture_output=True)
        except Exception:
            docx_path = None
        entry = {"v": n, "file": os.path.relpath(md_path, HERE), "docx": os.path.relpath(docx_path, HERE) if docx_path else None,
                 "built_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ"), "state": "draft",
                 "template": doc["template"], "template_sha": template_sha(doc["template"]), "deps": deps,
                 "pending": sorted(set(r.pending)), "e4_count": len(set(r.e4))}
        if latest and latest["state"] == "draft" and latest["v"] == n:
            entry["rebuilds"] = latest.get("rebuilds", 0) + 1
            entry["exports"] = latest.get("exports", {})
            hist[-1] = entry
        else:
            entry["exports"] = (latest or {}).get("exports", {})
            hist.append(entry)
        if doc.get("export") and not private:
            export_copy(doc, entry, md_path, docx_path)
        print(f"built {did} v{n} (draft) | pending {len(entry['pending'])} | E4 {entry['e4_count']} -> {entry['file']}")
    save_manifest(man)


def file_sha(p):
    with open(p, "rb") as fh:
        return hashlib.sha256(fh.read()).hexdigest()[:16]


def export_copy(doc, entry, md_path, docx_path):
    """Copy the latest draft into the application workspace, but never over a file you have edited there."""
    import shutil
    base = os.path.normpath(os.path.join(HERE, doc["export"]))
    os.makedirs(os.path.dirname(base), exist_ok=True)
    known = entry.setdefault("exports", {})
    for src, ext in ((md_path, ".md"), (docx_path, ".docx")):
        if not src:
            continue
        dst = base + ext
        if os.path.exists(dst) and known.get(dst) and file_sha(dst) != known[dst]:
            alt = base + f"_generated_v{entry['v']}{ext}"
            shutil.copyfile(src, alt)
            print(f"  kept your edited {os.path.basename(dst)}; new draft saved as {os.path.basename(alt)}")
            continue
        if os.path.exists(dst) and not known.get(dst):
            alt = base + f"_generated_v{entry['v']}{ext}"
            shutil.copyfile(src, alt)
            print(f"  {os.path.basename(dst)} exists but was not created by the generator; new draft saved as {os.path.basename(alt)}")
            continue
        shutil.copyfile(src, dst)
        known[dst] = file_sha(dst)


def outdated(entry, profile):
    reasons = []
    if not os.path.exists(os.path.join(TEMPLATES, entry["template"])):
        return ["template missing"]
    if template_sha(entry["template"]) != entry["template_sha"]:
        reasons.append("template changed")
    changed = [p for p, h in entry["deps"].items() if sha(resolve(profile, p)) != h]
    if changed:
        reasons.append("profile changed: " + ", ".join(changed[:6]) + (" …" if len(changed) > 6 else ""))
    return reasons


def cmd_status():
    registry = load_yaml(REGISTRY)["documents"]
    profile = load_yaml(PROFILE)
    man = load_manifest()
    print(f"{'document':28s} {'ver':>4s}  {'state':9s} {'up to date?':12s} {'pending':>7s}  note")
    for d in registry:
        hist = man.get(d["id"], {}).get("versions", [])
        if not hist:
            print(f"{d['id']:28s} {'-':>4s}  {'-':9s} {'NOT BUILT':12s} {'-':>7s}")
            continue
        e = hist[-1]
        why = outdated(e, profile)
        state = e["state"]
        flag = "yes" if not why else "NO"
        note = "; ".join(why)
        if why and state in ("reviewed", "final"):
            note += f" -> rebuild creates v{e['v'] + 1}; v{e['v']} is kept"
        print(f"{d['id']:28s} {('v' + str(e['v'])):>4s}  {state:9s} {flag:12s} {len(e['pending']):>7d}  {note}")


def cmd_mark(did, state):
    man = load_manifest()
    hist = man.get(did, {}).get("versions", [])
    if not hist:
        sys.exit(f"{did} has not been built")
    e = hist[-1]
    if state == "final" and e["pending"]:
        sys.exit(f"refused: {did} v{e['v']} still has {len(e['pending'])} pending facts")
    e["state"] = state
    e[f"{state}_at"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ")
    save_manifest(man)
    print(f"{did} v{e['v']} marked {state}")


def cmd_diff(did):
    man = load_manifest()
    hist = man.get(did, {}).get("versions", [])
    if not hist:
        sys.exit(f"{did} has not been built")
    profile = load_yaml(PROFILE)
    e = hist[-1]
    changed = [p for p, h in e["deps"].items() if sha(resolve(profile, p)) != h]
    print("\n".join(f"changed: {p} -> {resolve(profile, p)}" for p in changed) or "no profile changes affect this document")


if __name__ == "__main__":
    signal.signal(signal.SIGPIPE, signal.SIG_DFL)
    a = sys.argv[1:]
    if not a or a[0] == "status":
        cmd_status()
    elif a[0] == "build" and len(a) >= 2:
        cmd_build(a[1], private="--private" in a, force_new="--new" in a)
    elif a[0] in ("review", "final") and len(a) == 2:
        cmd_mark(a[1], "reviewed" if a[0] == "review" else "final")
    elif a[0] == "diff" and len(a) == 2:
        cmd_diff(a[1])
    else:
        print(__doc__)
