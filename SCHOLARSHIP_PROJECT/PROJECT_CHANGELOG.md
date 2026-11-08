# Project Changelog

**Two branches exist. Only one is synchronised.**

| Branch | Content | On GitHub? |
|---|---|---|
| `claude/amazing-pasteur-3yuhhm` | Original working history, including private commit a5c7983 | **Only up to 277ef6c.** Later commits stay local and are **never pushed** |
| `claude/amazing-pasteur-3yuhhm-clean` | Same project files, history without a5c7983 | **Synchronisation branch from Phase 8 onward** |

**Convention:** the sync script records each push in the table at the bottom of this file **after** the push. Each record is therefore committed with the next update. The final push of a session is reported in that session's summary, and appears here at the start of the next session.

## History

| Date (UTC) | Commit | Branch | Change |
|---|---|---|---|
| 5 Nov 06:28 | 277ef6c | both (pushed) | Phase 1–2 report, tracker, master profile |
| 5 Nov 17:50 | d3255fb | original (local) | Phase 3: verification, discrepancies, eligibility v2 |
| 7 Nov 21:17 | eb30730 | original (local) | Phase 4: security audit |
| 7 Nov 22:10 | **a5c7983** | original (local) | Phase 4: Command Center, kit, packages. **Contains birth-date text: never push** |
| 7 Nov 22:38 | 2ca1aed | original (local) | Phase 5 P1: corrections; birth-date text removed from files |
| 7 Nov 22:49 – 8 Nov 01:55 | c165f2c, 7262b87, a6e0290, da3e96f, 294a00d | original (local) | Phase 5 P2–P12 |
| 8 Nov 06:38 | ced8e85 | original (local) | Phase 6: execution workspace |
| 8 Nov 16:26–16:27 | 59ca281, ed79693 | original (local) | Phase 7: submission packs, privacy audit |
| 8 Nov | e290619 | clean (local, never pushed) | Phases 3–7 as one commit on top of 277ef6c, without a5c7983. **Replaced before its first push** by the sanitised commit below, so third-party names in archived files never reach GitHub |
| 8 Nov | (see sync log) | clean | Phase 8 updates (below) |

## Phase 8 changes

- Third-party names sanitised in archived files (Markdown, scripts, Word, Excel; PDFs regenerated).
- Added `_sync/` (sensitive-data scan + mandatory sync workflow), and this changelog, the conversation log and the status file.
- **GPA 2.62 everywhere:** the remaining references to the earlier unsupported figure were removed from active files:
  - profile; discrepancy register DR-01; Personal Information Confirmation Form
  - Command Center DR-01; conversation log
  - `14_ARCHIVE/README.md` warns that archived files may still show it.
- **Experience calculation:** `APPLICATION_EXECUTION_CENTER/15_EMPLOYMENT_EVIDENCE/Experience_Calculation.xlsx`, with a `.md`/`.docx`/`.pdf` summary.
  - Employer by employer, split before and after graduation.
  - Three separate tests: SI 3,000 h (E1); Chevening 2,800 h after graduation (E1 chevening.org, new snapshot); EPOS 24 months (E2/E3).
  - Each month counted once; an illustrative scenario that is not evidence.
  - Tested with sample data against an independent calculation: identical; 0 formula errors.
  - Replaces `Employment_Evidence_Collection_Form_v2.xlsx`, moved to `14_ARCHIVE/phase8_superseded/`.
- **Build fixes:**
  - The render tools moved from the session scratchpad into `SCHOLARSHIP_PROJECT/_tools/`, so builds are reproducible from GitHub.
  - Stale paths to Phase 6 files archived in Phase 7 were repointed (completion audit: 0 missing; manifest).
  - The kit builder no longer recreates the superseded HR requests.

## Synchronisation log (written by `_sync/sync.sh`)

| Time (UTC) | Local commit | Remote commit | Result | Message |
|---|---|---|---|---|
| 2026-11-08 17:21 UTC | ec6aaa8 | ec6aaa8 | PUSHED + VERIFIED (314 files) | Scholarship project Phases 3-8: sanitised history-safe snapshot + sync tooling and project records

- All Phase 3-7 work (application packs, evidence kit, research, dashboards) without private commit a5c7983
- Third-party names removed from archived files (all formats)
- _sync/: sensitive-data scan + mandatory sync workflow
- PROJECT_CONVERSATION_LOG, PROJECT_CHANGELOG, PROJECT_STATUS; START_HERE updated |
