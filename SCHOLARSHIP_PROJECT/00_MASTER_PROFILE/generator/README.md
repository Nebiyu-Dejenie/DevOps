# Document generator (profile-driven)

One profile → many documents. Change a fact once in `../profile/master_profile.yaml`, then let the generator tell you which documents need updating.

## Use

```bash
cd SCHOLARSHIP_PROJECT/00_MASTER_PROFILE/generator
python3 generate.py status            # what is outdated, what is pending, what is reviewed
python3 generate.py build all         # (re)build drafts as Markdown + Word
python3 generate.py diff cv_academic  # which facts changed since the last build
python3 generate.py review cv_academic   # protect a version you have reviewed
python3 generate.py final cv_academic    # refused while any fact is still pending
```

Requirements: Python 3 with PyYAML and Jinja2; pandoc for Word output (optional).

## Rules it enforces

| Rule | How |
|---|---|
| Never invent | Pending facts print as `[⏳ note]`; every draft ends with a list of pending facts and of facts based only on your confirmation (E4) |
| Never use E5 | Facts marked `ev: E5` (e.g. Plan International) are skipped in every document |
| Never overwrite reviewed work | Rebuilding a reviewed/final document creates the next version (`_v2`, `_v3` …); the reviewed file is kept |
| Know what changed | Each build stores the exact profile facts it used (hashed) and the template hash; `status` flags any change |
| Keep private data private | Name, e-mail and phone live in `../profile/private_overlay.yaml` (git-ignored). `build <doc> --private` merges it and writes to `output_private/` (git-ignored) |

## Documents (see `documents.yaml`)

CV (academic, professional) · motivation letters (AISS, EDISS, CMU-Africa, SI) · statement of purpose · CSC study plan · recommender briefings (academic, work) · interview preparation (AISS, CMU-Africa).

To add a document: write a template in `templates/` and add one line to `documents.yaml`.

## Version history

`manifest.json` keeps every version with build time, state (draft/reviewed/final), template hash, the facts used and the pending list.
