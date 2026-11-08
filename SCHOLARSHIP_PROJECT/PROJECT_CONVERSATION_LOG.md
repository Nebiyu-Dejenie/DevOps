# Project Conversation Log (sanitised)

**Purpose:** a faithful, sanitised record of every project-relevant work session: your instructions and decisions, the work done, the files produced, the findings (with verification status), problems, corrections, outstanding tasks and next actions.

**Not recorded:**

- raw chat transcripts
- passport details
- full date of birth
- contact details
- authentication information
- confidential correspondence

Times are UTC, taken from commit times where available.

---

## Session 1: 5 Nov 2026 (Phases 1–2)

- **Your instructions:** find fully funded master's scholarships for the 2027 intake. Never invent facts. Never create accounts, pay, sign or submit without approval. You gave a self-reported GPA figure that the transcript later corrected (see Session 2).
- **Work done:** Phase 1–2 Scholarship Intelligence Report; 12-sheet tracker; master profile.
- **Files:** Phase 1–2 report, tracker v1, Master Applicant Profile v1.
- **Git:** commit 277ef6c, **pushed 5 Nov 06:28**. It is the only commit that was on GitHub before Phase 8.
- **Problems found later:** that commit contains your name, GPA and course grades, the employer analysis, and one lecturer's name with a critique of his letter (third-party data). It contains no date of birth, passport, phone or ID.

## Session 2: 5 Nov 2026 (Phase 3)

- **Your instructions:** verify every fact and resolve discrepancies (GPA, dates, employers, recommendation letter). Protect personal information. **No pushes without approval.**
- **Work done:** discrepancy resolution report; eligibility v2; tracker v2; CV and essay drafts.
- **Key finding:** the official transcript gives a cumulative GPA of **2.62/4.00** (E0). The earlier self-reported figure is not supported by any document. In Phase 8 you confirmed 2.62 as final.
- **Git:** commit d3255fb (local only).

## Session 3: 7 Nov 2026 (Phase 4)

- **Your instructions:** security first (stop pushes; audit what was published); a Command Center; a document collection kit; a dedicated research environment limited to official domains.
- **Work done:** security audit (repository found **public**); Command Center; HR, registrar and referee request kit; official-source verification (46 snapshots).
- **Git:** commits eb30730 and **a5c7983** (local only). a5c7983 was later found to contain a sentence revealing your birth date.

## Session 4: 7–8 Nov 2026 (Phase 5)

- **Your instructions:**
  - correct AISS, EDISS, CMU-Africa and SI
  - find 15+ new opportunities
  - a CSC deep-dive
  - packages and a portfolio
  - a supervisor system
  - portal guides
  - a document generator
  - an English strategy
  - readiness scoring
  - an action centre and a manifest

  **Do not reproduce your full date of birth; do not decide eligibility on unverified age.**
- **Work done:** all 12 priorities drafted. Birth-date text removed from the working files (commit 2ca1aed).
- **Git:** commits 2ca1aed → 294a00d (local).

## Session 5: 8 Nov 2026 (Phase 6)

- **Your instructions:** move to execution; audit Phase 5 honestly; build the APPLICATION_EXECUTION_CENTER workspace; do not discuss GitHub visibility.
- **Work done:**
  - completion audit
  - priority strategy
  - SMACCs facts
  - CMU-Africa draft
  - China route analysis (DUT first; NPU excluded; HUST on hold)
  - Sweden programme shortlist from the official catalogue
  - AISS portfolio
  - workspace files 00–14
- **Git:** commit ced8e85 (local).

## Session 6: 8 Nov 2026 (Phase 7)

- **Your instructions:**
  - re-verify every scholarship, separating verified from search-indexed facts
  - classify CMU-Africa as Kigali, Rwanda (an African opportunity)
  - prioritise CMU-Africa, SI, SMACCs, AISS, Italy IYT and Germany/USA/Canada/UK
  - assess eligibility honestly
  - employment letters and recommendation drafts
  - CMU-Africa first
  - audit the whole Git history
  - submission-ready packs and a dashboard
- **Work done:**
  - 17 official pages re-verified
  - AISS identity confirmed as SITeS-JM
  - Italy IYT age rule found
  - submission packs; employment requests; recommendation drafts; verification register
- **Privacy:** the history audit found birth-date text **only in local commit a5c7983**. A local push guard was added, and a clean branch created.
- **Git:** commits 59ca281 and ed79693 (local). Clean branch `claude/amazing-pasteur-3yuhhm-clean` = e290619 (local).

## Session 7: 8 Nov 2026 (Phase 8)

- **Your instructions:**
  - **mandatory continuous GitHub sync** after every meaningful update, through the clean branch (never the original branch while it contains a5c7983)
  - verify every push against the remote
  - keep a sanitised conversation log, changelog, status and START_HERE
  - strict privacy rules
  - GPA **2.62** everywhere
  - employer-by-employer experience calculation (SI 3,000 h; Chevening 2,800 h after graduation; DAAD EPOS 2 years after graduation)
  - resolve DAAD EPOS's upper-third rule
  - register corrections: ASU, UBC and McGill Mastercard Foundation; Chevening as a future target; country of study vs provider
  - finish the CMU-Africa, SI, SMACCs and AISS applications
  - no fabricated achievements
  - no submissions, e-mails, signatures or payments without authorisation
- **Findings:**
  - The repository is **public** (GitHub API, 8 Nov).
  - Third-party names (a former lecturer; another person's CV filename) were present in archived files.
  - a5c7983 is not an ancestor of the clean branch.
  - **Chevening:** ≥ 2,800 h after the undergraduate degree, over ≥ 2 years; internships and voluntary work count (**E1**, chevening.org).
  - **Mastercard Foundation (E1 pages + E2 dates):**
    - ASU: Phase III "Innovation and Technology" includes an MS in IT; 2027/28 closed (27 Sep 2026)
    - UBC: forestry/agriculture/water only; closed 6 Sep 2026
    - McGill: nutrition/public health/policy/agriculture only; closed for Fall 2027
  - **DAAD EPOS (E2/E3):**
    - "far above average (upper third)", and a university ranking letter is required
    - indicative German grade for 2.62 = 3.07
    - degree turns 6 on 14 Jul 2028, so only the Aug–Oct 2027 deadlines are realistic
    - no ICT course confirmed
  - **Experience:** 0 hours documented; your 42-month statement would give about 6,720 h under stated assumptions (not evidence).
- **Corrections made:**
  - third-party names replaced in every archived format
  - the outdated GPA figure removed from active files
  - the v2 employment form superseded by the experience calculator (archived)
  - stale paths to archived Phase 6 files repointed
  - render tools moved into the repository
- **Files produced:**
  - `15_EMPLOYMENT_EVIDENCE/Experience_Calculation` (.xlsx/.md/.docx/.pdf)
  - `12_OTHER_SCHOLARSHIPS/DAAD_EPOS_Upper_Third_Evidence`
  - rewritten `Germany_USA_Canada_UK_Assessment`
  - Command Center with a funder column (57 entries)
  - verification register (42 facts)
  - readiness blocks in four packs
  - evidence-first dashboard
  - `14_ARCHIVE/README.md`
  - `_sync/`, `_tools/`
  - Official_Evidence_2026-11-08 (4 snapshots)
- **Git:** every update was pushed to the clean branch and verified (local SHA = remote SHA). The commits are in the changelog's synchronisation log.
- **Outstanding, and next actions:** see `PROJECT_STATUS.md` ("Your next 5 actions").
