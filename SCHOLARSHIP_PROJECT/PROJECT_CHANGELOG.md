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

## Synchronisation log (written by `_sync/sync.sh`)

| Time (UTC) | Local commit | Remote commit | Result | Message |
|---|---|---|---|---|
