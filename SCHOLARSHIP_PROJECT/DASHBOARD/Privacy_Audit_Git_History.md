# Git History Privacy Audit

**Run:** 8 Nov 2026, from this session. **Scope:** every commit on every local branch, including the text inside .docx, .xlsx and .pdf files. Matched values are masked here; this report does not repeat any private value.

Commits scanned: **12** (all local branches). Unique file versions scanned: **371**, including text inside .docx, .xlsx and .pdf files.

## Result

| Finding | Where | Pushed to GitHub? |
|---|---|---|
| **Birth-date text** (a sentence stating the date you turn 30, which reveals your full birth date) | **Only in local commit a5c7983** (Phase 4): master profile v3, opportunity register v3 (.md/.docx/.pdf), Command Center workbook and build script (7 file versions) | **No.** The only pushed commit, 277ef6c, contains no private data |
| Passport numbers, phone numbers, personal e-mail addresses, student/national IDs, home address, bank details | – | None found in any commit |
| False positives (no private value) | 21 matches: a sentence saying that files "contained" your birth date, and programmes' published age rules (HUST "born on/after 1 Sep 1992"; Italy IYT "born on or after 1 January 2000") | – |

The text was removed from the working files in commit 2ca1aed. **Every later commit is clean**; only the historical commit a5c7983 still contains it.

## Protections now in place

1. **No push from this session.** Nothing has been or will be pushed by me.
2. **Local push guard** (`.git/hooks/pre-push`, not versioned): any push from this machine that includes a5c7983 or birth-date text is refused. It was tested on the current branch: blocked.
3. **Clean branch `claude/amazing-pasteur-3yuhhm-clean`**: your current files as **one commit on top of the already-pushed commit 277ef6c**, so commit a5c7983 is not in its history. If you ever push from here, push that branch. The working branch is **not** rewritten.
4. **Ignore rules** (`.gitignore`): `SCHOLARSHIP_PROJECT/PRIVATE/`, passport scans and `*_ID_scan*` files are never committed, and neither is the generator's private overlay. Keep personal documents in `SCHOLARSHIP_PROJECT/PRIVATE/` or outside the repository.

## Masked findings (machine output)

| Type | Matches | File(s) | In commits | Already on the remote? | Masked sample |
|---|---|---|---|---|---|
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.docx | 294a00d, 7262b87, a6e0290, da3e96f | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.docx | 2ca1aed, c165f2c | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.docx | 59ca281, ced8e85 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.md | 294a00d, 7262b87, a6e0290, da3e96f | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.md | 2ca1aed, c165f2c | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.md | 59ca281, ced8e85 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.pdf | 294a00d, 7262b87, a6e0290, da3e96f | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.pdf | 2ca1aed, c165f2c | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.pdf | 59ca281, ced8e85 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/00_DASHBOARD.md | 59ca281 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/00_DASHBOARD.pdf | 59ca281 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/00_VERIFICATION_REGISTER.xlsx | 59ca281 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/02_TOP_SCHOLARSHIPS.xlsx | 59ca281 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/11_CHINA/China_Application_Operation.xlsx | 59ca281, ced8e85 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Italy_IYT_SUBMISSION_PACK.docx | 59ca281 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Italy_IYT_SUBMISSION_PACK.md | 59ca281 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Italy_IYT_SUBMISSION_PACK.pdf | 59ca281 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/_build/build_china.py | 59ca281, ced8e85 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/_build/build_core.py | 59ca281 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/_build/build_packs.py | 59ca281 | no (local only) | •••… |
| birth date / age text | 1 | SCHOLARSHIP_PROJECT/APPLICATION_EXECUTION_CENTER/_build/build_verification.py | 59ca281 | no (local only) | •••… |
| birth date / age text | 2 | SCHOLARSHIP_PROJECT/00_MASTER_PROFILE/Master_Profile_and_Discrepancy_Register_v3.docx | a5c7983 | no (local only) | •••… |
| birth date / age text | 2 | SCHOLARSHIP_PROJECT/00_MASTER_PROFILE/Master_Profile_and_Discrepancy_Register_v3.md | a5c7983 | no (local only) | •••… |
| birth date / age text | 2 | SCHOLARSHIP_PROJECT/DASHBOARD/_build/build_command_center.py | a5c7983 | no (local only) | •••… |
| birth date / age text | 4 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/Verified_Opportunity_Register_v3.docx | a5c7983 | no (local only) | •••… |
| birth date / age text | 4 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/Verified_Opportunity_Register_v3.md | a5c7983 | no (local only) | •••… |
| birth date / age text | 4 | SCHOLARSHIP_PROJECT/03_SCHOLARSHIP_RESEARCH/Verified_Opportunity_Register_v3.pdf | a5c7983 | no (local only) | •••… |
| birth date / age text | 4 | SCHOLARSHIP_PROJECT/DASHBOARD/Scholarship_Command_Center.xlsx | a5c7983 | no (local only) | •••… |
