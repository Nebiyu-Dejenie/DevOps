"""Phase 5 completion audit -> 00_PHASE5_COMPLETION_AUDIT.xlsx (file existence is checked when the script runs)."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))  # SCHOLARSHIP_PROJECT
OUT = os.path.join(ROOT, "APPLICATION_EXECUTION_CENTER", "00_PHASE5_COMPLETION_AUDIT.xlsx")
LADDER = ["1 Research completed", "2 Document drafted", "3 Document reviewed (self-check)", "4 Facts verified (E1)",
          "5 Ready for your approval", "6 Application ready for submission"]
# id, requirement, deliverable, file (relative to SCHOLARSHIP_PROJECT), how Claude reviewed it, verification level, ladder level reached, ready for application use?, remaining work
R = [
 ("P1a", "AISS: correct the GPA statement; re-evaluate competitiveness; plan from real projects", "Corrections report, section 1",
  "03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.md", "Written; rendered PDF page 3 viewed", "E1 (aissprogram.eu snapshot 7 Nov)", 4, "Research: yes. Application: no", "None for research. Portfolio, video and IELTS are application work (09_AISS)"),
 ("P1b", "EDISS: course-by-course comparison; prerequisites; English; realistic or low probability", "Corrections report, section 2",
  "03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.md", "Written; averages computed by script from the transcript", "E0 transcript + E1 master-ediss.eu", 4, "Research: yes", "None (verdict C; optional application)"),
 ("P1c", "CMU-Africa: separate admission, financial aid, Mastercard Foundation selection", "Corrections report, section 3",
  "03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.md", "Written; not re-read line by line", "E2 (site fails TLS here) + E1 Mastercard Foundation", 1, "Research: partly", "Open the CMU-Africa pages in your browser to confirm E2 facts"),
 ("P1d", "Sweden SI: hours without double counting; employer forms; leadership evidence", "Employment form v2 with 'SI hours - no double count' sheet",
  "09_APPLICATION_DOCUMENTS/00_Document_Collection_Kit/Employment_Evidence_Collection_Form_v2.xlsx", "Formula test with sample data (8,360 vs 8,840 h); recalc 0 errors", "SI rules E1; overlap rule is a conservative assumption (SI is silent)", 5, "Tool: yes. Hours: no data yet", "Your HR letters (dates, hours/week) must be entered"),
 ("P1e", "Personal information: investigate age discrepancy; no full DOB in reports; no age-based decisions", "Redaction + age policy",
  "03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1.md", "Scan of every current md/docx/xlsx/pdf: 0 birth-date hits", "Scan result", 4, "Yes (files)", "Birth-date text remains in local git commit a5c7983 (your decision)"),
 ("P2", "At least 15 new quality-controlled opportunities", "16 new + rejected list",
  "03_SCHOLARSHIP_RESEARCH/New_Opportunities_2027_v1.md", "Written; rendered PDF page 1 viewed", "Mostly E2/E3 (official pages via search)", 1, "Research: yes (E2 level)", "Browser-check any option before spending money"),
 ("P3", "CSC deep-dive: Ethiopia call, route, agency number, age, pre-admission, programmes, coverage, deadline", "Deep-dive report",
  "03_SCHOLARSHIP_RESEARCH/China_CSC_2027-28_Deep_Dive_v1.md", "Written; not re-read line by line", "E2", 1, "Research: partly", "Ethiopia 2027/28 call and Type A agency number NOT found; ask MoE. Phase 6 adds university-level checks"),
 ("P4a", "CMU-Africa materials", "Package v3 (archived in Phase 6; replaced by 07_CMU_AFRICA draft)",
  "14_ARCHIVE/phase5_superseded/CMU-Africa_Application_Package_v3.md", "Assembled by script from v2 + new sections; not re-read in full", "E2 facts; essays contain placeholders", 2, "No", "Superseded by APPLICATION_EXECUTION_CENTER/07_CMU_AFRICA (Phase 6)"),
 ("P4b", "AISS materials", "Package v2 (moved to the workspace and updated in Phase 6)",
  "APPLICATION_EXECUTION_CENTER/09_AISS/AISS_Application_Package.md", "Written; not re-read in full", "E1 rules", 2, "No", "Phase 6 adds the portfolio document, social impact statement"),
 ("P4c", "EDISS materials", "Package v2 (moved to the workspace in Phase 6)",
  "APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/EDISS_Application_Package.md", "Written; not re-read in full", "E1 rules", 2, "No (optional, label C)", "Only if you choose to apply"),
 ("P4d", "Sweden SI materials", "Package v2 (moved to the workspace and updated in Phase 6)",
  "APPLICATION_EXECUTION_CENTER/08_SWEDEN_SI/SI_Application_Package.md", "Assembled from v1 + new sections", "E1 rules", 2, "No", "SI templates publish 13 Nov; programme shortlist added in Phase 6"),
 ("P4e", "China CSC materials", "Package v1",
  "APPLICATION_EXECUTION_CENTER/11_CHINA/China_Study_Plan_and_Emails.md", "Written; not re-read in full", "E2", 2, "No", "Superseded by 11_CHINA (Phase 6)"),
 ("P5", "Technical project portfolio (PDF + editable)", "Portfolio v1",
  "APPLICATION_EXECUTION_CENTER/06_FINAL_CV/Technical_Project_Portfolio.md", "Rendered PDF page 1 viewed", "Structure only: most fields await your facts", 2, "No", "Project title, HR facts, incident/automation stories; demo repos not built"),
 ("P6", "Supervisor and university contact system with customised e-mails", "System + tracker",
  "APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Research_MSc_Canada/Supervisor_Contact_System.md", "Written; tracker recalc 0 errors", "E1 (umanitoba.ca, ucalgary.ca)", 5, "Drafts: yes, after you add one paper per professor", "Calgary e-mail addresses are on profile pages (not fetched)"),
 ("P7", "Portal submission guides with copy-paste answers", "Guides v1",
  "APPLICATION_EXECUTION_CENTER/14_SUBMISSION_GUIDES.md", "Written; not tested on live portals (portals not opened)", "E1/E2 mix", 5, "Yes as guidance", "Portal fields may differ from the guide; check when you open each portal"),
 ("P8", "Reusable profile-driven document generator", "generate.py + templates",
  "00_MASTER_PROFILE/generator/generate.py", "Tested: E5 exclusion, review protection, change detection, final refusal", "Tool test", 5, "Yes (tool)", "Generated texts are drafts with pending facts"),
 ("P9", "English strategy across programmes", "Strategy v3",
  "APPLICATION_EXECUTION_CENTER/13_INTERVIEW_PREPARATION/English_Test_Plan.md", "Written; rendered, not viewed", "E1 (ielts.org, programme pages)", 5, "Yes", "IELTS not booked (Phase 6 re-confirms necessity)"),
 ("P10", "Transparent readiness score (10 dimensions, A-D)", "Scorecard v1",
  "14_ARCHIVE/phase5_superseded/Readiness_Scorecard_v1.xlsx", "Recalc 0 errors; labels compared with documents", "Scores are judgements on stated rules", 5, "Yes", "Re-scored in Phase 6 (02_TOP_SCHOLARSHIPS.xlsx)"),
 ("P11", "One-page MY_SCHOLARSHIP_ACTION_CENTER.pdf", "Action Center",
  "14_ARCHIVE/phase5_superseded/MY_SCHOLARSHIP_ACTION_CENTER.pdf", "Full page viewed at 80 dpi; fits one page", "Mix E1/E2", 5, "Yes (as of 8 Nov)", "Replaced by APPLICATION_EXECUTION_CENTER dashboard"),
 ("P12", "Deliverables manifest + organised archive", "Manifest + zip",
  "DELIVERABLES_MANIFEST.md", "Generated by script; zip scanned for personal data (0 hits)", "File-system check", 5, "Yes", "The zip lived only in the temporary scratchpad and was sent to you"),
]
wb = Workbook(); ws = wb.active; ws.title = "Phase 5 audit"
H = ["ID", "Phase 5 requirement", "Deliverable produced", "Exact filename", "Exact location", "File exists?", "Opened / reviewed by Claude",
     "Reviewed by you", "Factual verification", "Editable?", "Highest honest status reached", "Ready for application use?", "Remaining work"]
W = [6, 44, 26, 40, 40, 9, 34, 10, 30, 9, 30, 22, 46]
HDR = PatternFill("solid", fgColor="1F3864"); th = Side(style="thin", color="BFBFBF"); BD = Border(left=th, right=th, top=th, bottom=th)
for j, (h, w) in enumerate(zip(H, W), 1):
    c = ws.cell(row=1, column=j, value=h); c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9); c.fill = HDR; c.border = BD
    c.alignment = Alignment(wrap_text=True, vertical="center"); ws.column_dimensions[c.column_letter].width = w
ws.row_dimensions[1].height = 32
for i, (rid, req, deliv, rel, how, ver, lvl, ready, rem) in enumerate(R, 2):
    p = os.path.join(ROOT, rel)
    exists = "Yes" if os.path.exists(p) else "NO"
    ext = os.path.splitext(rel)[1]
    editable = "Yes" if ext in (".md", ".xlsx", ".docx", ".py") else "PDF (source editable)"
    vals = [rid, req, deliv, os.path.basename(rel), os.path.dirname(rel) or "(project root)", exists, how, "No", ver, editable, LADDER[lvl - 1], ready, rem]
    for j, v in enumerate(vals, 1):
        c = ws.cell(row=i, column=j, value=v); c.font = Font(name="Arial", size=9); c.border = BD; c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[i].height = 48
last = len(R) + 1
ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['F2="NO"'], fill=PatternFill("solid", fgColor="FFC7CE")))
ws.freeze_panes = "C2"
lg = wb.create_sheet("Status ladder")
for i, t in enumerate(["Status ladder used (a document only reaches a level when the level below is true):"] + LADDER +
                      ["", "No Phase 5 item is at level 6. No application has been submitted.",
                       "'Reviewed by you' is No everywhere until you confirm you have read a document."], 1):
    lg.cell(row=i, column=1, value=t).font = Font(name="Arial", size=10, bold=(i == 1))
lg.column_dimensions["A"].width = 100
os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)
print("rows", len(R), "missing", sum(1 for r in R if not os.path.exists(os.path.join(ROOT, r[3]))))
