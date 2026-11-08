#!/bin/bash
# Mandatory sync workflow (Phase 8 instruction): status → diff → deletion check → sensitive scan → commit →
# ancestry check → push → verify remote SHA → record in PROJECT_CHANGELOG.md.
# Usage: SCHOLARSHIP_PROJECT/_sync/sync.sh "Commit message"   (ALLOW_DELETIONS=1 to permit deletions/moves)
set -u
SYNC_BRANCH="claude/amazing-pasteur-3yuhhm-clean"
PRIVATE_COMMIT="a5c7983"
REPO_ROOT=$(git rev-parse --show-toplevel) || exit 1
cd "$REPO_ROOT" || exit 1
LOG="SCHOLARSHIP_PROJECT/PROJECT_CHANGELOG.md"
MSG="${1:?commit message required}"
now() { date -u +"%Y-%m-%d %H:%M UTC"; }
record() { printf '| %s | %s | %s | %s | %s |\n' "$(now)" "$1" "$2" "$3" "$4" >> "$LOG"; }

branch=$(git branch --show-current)
if [ "$branch" != "$SYNC_BRANCH" ]; then echo "STOP: on '$branch', expected '$SYNC_BRANCH' (the original branch contains private commit $PRIVATE_COMMIT)."; exit 1; fi

echo "== 1. status"; git status --short | head -40
echo "== 2. diff (stat)"; git diff --stat | tail -3
deleted=$(git status --porcelain | awk '$1 ~ /D/ {print $2}')
if [ -n "$deleted" ] && [ "${ALLOW_DELETIONS:-0}" != "1" ]; then echo "STOP: deletions found (re-run with ALLOW_DELETIONS=1 if intended):"; echo "$deleted"; exit 1; fi

echo "== 3. stage project files only"; git add -A SCHOLARSHIP_PROJECT .gitignore
if git diff --cached --quiet; then echo "Nothing to commit."; exit 0; fi
nfiles=$(git diff --cached --name-only | wc -l)

echo "== 4. sensitive-data and secret scan (staged)"
if ! python3 -I SCHOLARSHIP_PROJECT/_sync/sensitive_scan.py --staged; then
  git reset -q; record "not committed" "-" "SCAN FAIL: nothing committed or pushed" "$MSG"; echo "STOP: scan failed; files unstaged; nothing committed."; exit 1; fi

echo "== 5. commit"
git commit -q -m "$MSG

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_011wY2tVFvWKD2S8fVsN7Aih" || exit 1
local_sha=$(git rev-parse HEAD)

echo "== 6. ancestry check"
if git merge-base --is-ancestor "$PRIVATE_COMMIT" HEAD 2>/dev/null; then record "${local_sha:0:7}" "-" "BLOCKED: private commit in ancestry" "$MSG"; echo "STOP: $PRIVATE_COMMIT is an ancestor. Not pushing."; exit 1; fi
echo "ok: $PRIVATE_COMMIT is not an ancestor"

echo "== 7. push (retry 2/4/8/16 s on network errors)"
ok=0
for wait in 0 2 4 8 16; do
  [ "$wait" -gt 0 ] && sleep "$wait"
  if git push -u origin "$SYNC_BRANCH" 2>&1 | tail -3; then
    remote_sha=$(git ls-remote origin "refs/heads/$SYNC_BRANCH" | cut -f1)
    if [ "$remote_sha" = "$local_sha" ]; then ok=1; break; fi
  fi
done

echo "== 8. verify remote"
remote_sha=$(git ls-remote origin "refs/heads/$SYNC_BRANCH" | cut -f1)
echo "local : $local_sha"; echo "remote: ${remote_sha:-<none>}"
if [ "$ok" = 1 ] && [ "$remote_sha" = "$local_sha" ]; then
  record "${local_sha:0:7}" "${remote_sha:0:7}" "PUSHED + VERIFIED ($nfiles files)" "$MSG"; echo "SYNC RESULT: PUSHED AND VERIFIED"
else
  record "${local_sha:0:7}" "${remote_sha:0:7}" "PUSH NOT VERIFIED: committed locally only" "$MSG"; echo "SYNC RESULT: NOT VERIFIED (work kept locally)"; exit 2
fi
