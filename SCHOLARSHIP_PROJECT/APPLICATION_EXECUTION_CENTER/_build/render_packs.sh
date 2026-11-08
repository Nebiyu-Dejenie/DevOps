#!/bin/bash
# Renders the submission packs and Phase 7 documents to Word (editable) and PDF (reading). Run from anywhere.
set -e
cd "$(dirname "$0")/.."
REF=../00_MASTER_PROFILE/generator/reference_compact.docx
TOOLS="$(cd ../_tools && pwd)"
for f in 07_CMU_AFRICA/CMU-Africa_SUBMISSION_PACK.md 08_SWEDEN_SI/Sweden_SI_SUBMISSION_PACK.md 09_AISS/AISS_SUBMISSION_PACK.md \
         10_SMACCs/SMACCs_SUBMISSION_PACK.md 11_CHINA/China_DUT_SUBMISSION_PACK.md 12_OTHER_SCHOLARSHIPS/Italy_IYT_SUBMISSION_PACK.md \
         12_OTHER_SCHOLARSHIPS/Germany_USA_Canada_UK_Assessment.md 13_INTERVIEW_PREPARATION/Interview_Bank.md 14_SUBMISSION_CHECKLIST.md; do
  pandoc "$f" -o "${f%.md}.docx" --reference-doc "$REF"; zip -q -d "${f%.md}.docx" '*/' >/dev/null 2>&1 || true
  "$TOOLS/md2pdf.sh" "$f" landscape >/dev/null && echo "ok $f"
done
