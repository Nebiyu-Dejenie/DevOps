# Project Changelog

**Two branches exist. Only one is synchronised.**

| Branch | Content | On GitHub? |
|---|---|---|
| `claude/amazing-pasteur-3yuhhm` | Original working history, including private commit a5c7983 | **Only up to 277ef6c.** Later commits stay local and are **never pushed** |
| `claude/amazing-pasteur-3yuhhm-clean` | Same project files, history without a5c7983 | **Synchronisation branch from Phase 8 onward** |

**Convention:**

- After each verified push, `_sync/sync.sh` adds a row to the table at the bottom of this file.
- It then pushes that row in a small follow-up commit ("Record verified sync of …"), which is scanned and verified the same way.
- Follow-up commits are not listed themselves; that would never end.
- **Before this change** (rows up to commit c220eae), each record was committed with the next update instead.

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
- **Register corrections (Phase 8.3):**
  - Command Center: "Country" is now "Country of study", plus a new "Funder (and where it is based)" column for all 57 entries.
  - New entries: S61 ASU Mastercard Foundation (USA; closed for 2027/28; monitor), S62 McGill.
  - Corrected entries: S17 is now UBC with its actual eligibility and closed stage; S13 Chevening is a future-cycle target with the E1 2,800-hour rule; S22 EPOS is conditional, next cycle.
  - Verification register: +6 facts.
  - Deadlines: ASU (Jun 2027 watch) and EPOS (Aug 2027).
  - `Germany_USA_Canada_UK_Assessment` rewritten with a funder column.
- **DAAD EPOS:** `DAAD_EPOS_Upper_Third_Evidence`.
  - The rule (E2/E3); the indicative German grade 3.07; what DTU can issue.
  - A rank-letter paragraph (not sent); the decision rule; the IT-course search.
- **Evidence:** Mastercard Foundation pages for ASU, UBC and McGill saved (E1) in `03_SCHOLARSHIP_RESEARCH/Official_Evidence_2026-11-08`.
- **Submission readiness (Phase 8.4):** a section 0 at the top of the CMU-Africa, SI, SMACCs and AISS packs. It shows:
  - can it be submitted today; what's complete or drafted
  - what you must provide; what to verify
  - the exact portal link; the deadline with days left; the final checks
- **AISS official name:** "Artificial Intelligence for Sustainable Societies (SITeS-JM)" in the pack title, the dashboards, the Command Center, the deadlines and the portfolio.
- **Dashboard:** new "Evidence first" table (HR letters, experience proof, DTU Registrar, referees).
- **Word copies:** the `_tools/br.lua` filter keeps line breaks inside table cells.
- **Session close:**
  - PROJECT_STATUS rewritten: applications; next-cycle targets; evidence; your next 5 actions.
  - Conversation log: Session 7 completed.
  - START_HERE updated.
  - `sync.sh` now pushes its own record row (step 9), so no local change is left after a sync.

## Synchronisation log (written by `_sync/sync.sh`)

| Time (UTC) | Local commit | Remote commit | Result | Message |
|---|---|---|---|---|
| 2026-11-08 17:21 UTC | ec6aaa8 | ec6aaa8 | PUSHED + VERIFIED (314 files) | Scholarship project Phases 3-8: sanitised history-safe snapshot + sync tooling and project records

- All Phase 3-7 work (application packs, evidence kit, research, dashboards) without private commit a5c7983
- Third-party names removed from archived files (all formats)
- _sync/: sensitive-data scan + mandatory sync workflow
- PROJECT_CONVERSATION_LOG, PROJECT_CHANGELOG, PROJECT_STATUS; START_HERE updated |
| 2026-11-08 17:38 UTC | 031d62f | 031d62f | PUSHED + VERIFIED (33 files) | Phase 8.2: GPA 2.62 everywhere; employer-by-employer experience calculator (SI/Chevening/EPOS)

- Remove the outdated self-reported GPA figure from active files (profile, register DR-01, confirmation form, Command Center, log)
- New 15_EMPLOYMENT_EVIDENCE/Experience_Calculation.xlsx (+ md/docx/pdf): three separate tests, each month counted once, scenario clearly not evidence
- Chevening work-experience rule saved as E1 evidence (chevening.org, 8 Nov 2026)
- Archive the v2 employment form (superseded); 14_ARCHIVE/README.md
- Render tools moved into SCHOLARSHIP_PROJECT/_tools; stale archived paths repointed |
| 2026-11-08 17:46 UTC | 06cb151 | 06cb151 | PUSHED + VERIFIED (25 files) | Phase 8.3: register corrections (ASU/UBC/McGill Mastercard Foundation, Chevening, EPOS) + country of study vs funder

- Command Center: Country of study + Funder column (57 entries); S61 ASU (closed 2027/28, monitor), S17 UBC, S62 McGill (closed, no CS route), S13 Chevening future target, S22 EPOS conditional
- DAAD_EPOS_Upper_Third_Evidence (rule, indicative German grade 3.07, DTU rank-letter paragraph, decision rule)
- Germany/USA/Canada/UK assessment rewritten; verification register +6 facts; deadlines (ASU Jun 2027, EPOS Aug 2027)
- E1 snapshots: Mastercard Foundation ASU/UBC/McGill pages |
| 2026-11-08 17:50 UTC | c220eae | c220eae | PUSHED + VERIFIED (26 files) | Phase 8.4: submission readiness blocks (CMU-Africa, SI, SMACCs, AISS); official AISS name; evidence-first dashboard

- Section 0 in each priority pack: submittable today?, complete/drafted, you must provide, verify, exact portal link, deadline + days left, final checks
- Official name 'Artificial Intelligence for Sustainable Societies (SITeS-JM)' in pack title, dashboards, Command Center, deadlines, portfolio
- 00_DASHBOARD: evidence-first table (HR letters, experience proof, DTU Registrar, referees) + Phase 8 completed list
- SI audit checklist points to Experience_Calculation.xlsx; Word line-break filter (_tools/br.lua) |
