#!/bin/bash
# Renders the workspace: Word copies of editable documents, PDFs of the read-only checklists and dashboard.
# Run from APPLICATION_EXECUTION_CENTER/_build
set -e
cd "$(dirname "$0")/.."
REF=../00_MASTER_PROFILE/generator/reference_compact.docx
TOOLS=/tmp/claude-0/-home-user-DevOps/76fe69da-fa2f-51c3-baf0-e814e33d9b97/scratchpad/tools
DOCX_LIST="05_READY_TO_SEND_EMAILS.md 07_CMU_AFRICA/CMU-Africa_Application_Draft.md 08_SWEDEN_SI/SI_Application_Package.md \
 09_AISS/AISS_Application_Package.md 09_AISS/AISS_Portfolio_Summaries.md 10_SMACCs/SMACCs_Application_Strategy.md \
 11_CHINA/China_Study_Plan_and_Emails.md 12_OTHER_SCHOLARSHIPS/Backups_and_Financial_Gap.md \
 12_OTHER_SCHOLARSHIPS/EDISS_Application_Package.md 12_OTHER_SCHOLARSHIPS/Stipendium_Hungaricum_Package.md \
 12_OTHER_SCHOLARSHIPS/Research_MSc_Canada/Supervisor_Contact_System.md 13_INTERVIEW_PREPARATION/Interview_Bank.md \
 13_INTERVIEW_PREPARATION/English_Test_Plan.md 06_FINAL_CV/Technical_Project_Portfolio.md 14_SUBMISSION_GUIDES.md 01_MY_PROFILE.md"
for f in $DOCX_LIST; do
  pandoc "$f" -o "${f%.md}.docx" --reference-doc "$REF"
  zip -q -d "${f%.md}.docx" '*/' >/dev/null 2>&1 || true
  echo "docx ${f%.md}.docx"
done
"$TOOLS/md2pdf.sh" 04_MY_MISSING_DOCUMENTS.md landscape
"$TOOLS/md2pdf.sh" 14_SUBMISSION_CHECKLIST.md
"$TOOLS/md2pdf.sh" 00_DASHBOARD.md landscape
"$TOOLS/md2pdf.sh" 07_CMU_AFRICA/CMU-Africa_Application_Draft.md
"$TOOLS/md2pdf.sh" 10_SMACCs/SMACCs_Application_Strategy.md
