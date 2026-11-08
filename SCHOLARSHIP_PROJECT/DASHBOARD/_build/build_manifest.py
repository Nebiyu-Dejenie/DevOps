"""Builds DELIVERABLES_MANIFEST.md + .xlsx at the project root from the files actually present."""
import os, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
# base path -> (purpose, completion status)
P = {
 "START_HERE": ("Entry point: where everything is", "Complete"),
 "APPLICATION_EXECUTION_CENTER/00_DASHBOARD": ("Main dashboard: 5 tasks, top 5 opportunities, folder links", "Complete (8 Nov 2026)"),
 "APPLICATION_EXECUTION_CENTER/00_PHASE5_COMPLETION_AUDIT": ("Phase 5 completion audit (honest status ladder)", "Complete"),
 "APPLICATION_EXECUTION_CENTER/01_MY_PROFILE": ("Verified facts every application uses", "Complete; work facts pending HR letters"),
 "APPLICATION_EXECUTION_CENTER/02_TOP_SCHOLARSHIPS": ("Priority strategy, readiness scorecard (A-D), financial gap", "Complete (re-score as evidence arrives)"),
 "APPLICATION_EXECUTION_CENTER/03_APPLICATION_DEADLINES": ("All deadlines with days left", "Complete"),
 "APPLICATION_EXECUTION_CENTER/04_MY_MISSING_DOCUMENTS": ("Evidence checklist linked to the prepared request letters", "Complete"),
 "APPLICATION_EXECUTION_CENTER/05_READY_TO_SEND_EMAILS": ("E-mails to MoE, DUT, HUST, SMACCs, CMU-Africa", "Ready for your approval"),
 "APPLICATION_EXECUTION_CENTER/06_FINAL_CV/CV_CMU-Africa_2pages": ("CMU-Africa CV (2 pages), generated from the master profile", "Draft: needs HR facts"),
 "APPLICATION_EXECUTION_CENTER/06_FINAL_CV/CV_Europe_Academic": ("CV for AISS, SMACCs, EDISS", "Draft: needs HR facts"),
 "APPLICATION_EXECUTION_CENTER/06_FINAL_CV/CV_content_for_SI_template": ("Blocks to paste into SI's official CV template", "Draft: template due 13 Nov"),
 "APPLICATION_EXECUTION_CENTER/06_FINAL_CV/Technical_Project_Portfolio": ("Master list of real projects + evidence gaps", "Draft: needs your facts"),
 "APPLICATION_EXECUTION_CENTER/06_FINAL_CV/Profile_Strengthening_Plan": ("Genuine improvements before deadlines vs long term", "Complete"),
 "APPLICATION_EXECUTION_CENTER/00_VERIFICATION_REGISTER": ("Every key fact: verified (E1) vs search-indexed (E2) vs not found", "Complete (8 Nov 2026)"),
 "APPLICATION_EXECUTION_CENTER/07_CMU_AFRICA/CMU-Africa_SUBMISSION_PACK": ("CMU-Africa: URLs, field map, 4 essays, video, need statement, guide, audit", "Draft ready for your review"),
 "APPLICATION_EXECUTION_CENTER/08_SWEDEN_SI/Sweden_SI_SUBMISSION_PACK": ("Sweden SI: 3,000-h rule (E1), programmes, motivation, SU letter, forms, guide, audit", "Draft ready (templates due 13 Nov)"),
 "APPLICATION_EXECUTION_CENTER/09_AISS/AISS_SUBMISSION_PACK": ("AISS (SITeS-JM): identity, rules, video, research interests, self-funding, guide, audit", "Draft ready for your review"),
 "APPLICATION_EXECUTION_CENTER/10_SMACCs/SMACCs_SUBMISSION_PACK": ("SMACCs: facts (E2), letter, documents, guide, interview, audit", "Draft ready for your review"),
 "APPLICATION_EXECUTION_CENTER/11_CHINA/China_DUT_SUBMISSION_PACK": ("DUT Type B: rules, study plan, e-mails, guide, audit", "Draft ready for your review"),
 "APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Italy_IYT_SUBMISSION_PACK": ("Italy IYT: rules incl. age, video draft, documents, audit", "Conditional (private age check)"),
 "APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Germany_USA_Canada_UK_Assessment": ("Fully funded options in DE/US/CA/UK for 2027 and next cycle", "Complete"),
 "DASHBOARD/Privacy_Audit_Git_History": ("Full Git history privacy audit (masked findings) + protections", "Complete"),
 "APPLICATION_EXECUTION_CENTER/07_CMU_AFRICA/CMU-Africa_Application_Draft": ("Complete CMU-Africa draft: essays, need, leadership, video, test, interview", "Draft ready for your review"),
 "APPLICATION_EXECUTION_CENTER/08_SWEDEN_SI/SI_Application_Package": ("SI + universityadmissions.se package with the 13 Nov procedure", "Draft: templates due 13 Nov"),
 "APPLICATION_EXECUTION_CENTER/08_SWEDEN_SI/SI_Programme_Shortlist": ("100 open programmes (E1) + 5 checked candidates", "Provisional until the 13 Nov list"),
 "APPLICATION_EXECUTION_CENTER/09_AISS/AISS_Application_Package": ("AISS video, research interests, social impact, interview, self-funding", "Draft ready for your review"),
 "APPLICATION_EXECUTION_CENTER/09_AISS/AISS_Portfolio_Summaries": ("The AISS portfolio document (5 executed projects)", "Draft: needs your facts + demo repos"),
 "APPLICATION_EXECUTION_CENTER/10_SMACCs/SMACCs_Application_Strategy": ("SMACCs facts, fit, motivation letter, interview strategy", "Draft ready for your review (facts E2)"),
 "APPLICATION_EXECUTION_CENTER/11_CHINA/China_Application_Operation": ("DUT/CSU/NUIST/UESTC/HUST/NPU + Type A comparison", "Complete (E2)"),
 "APPLICATION_EXECUTION_CENTER/11_CHINA/China_Study_Plan_and_Emails": ("DUT study plan, e-mails, step-by-step", "Draft ready for your review"),
 "APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Backups_and_Financial_Gap": ("Backups and real funding gaps", "Complete"),
 "APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/EDISS_Application_Package": ("EDISS (optional, label C)", "Draft"),
 "APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Stipendium_Hungaricum_Package": ("Hungary package", "Draft: call pending"),
 "APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Research_MSc_Canada/Supervisor_Contact_System": ("Manitoba/Calgary supervisor drafts", "Ready (approve before sending)"),
 "APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Research_MSc_Canada/Supervisor_Contact_Tracker": ("Supervisor contact log", "Ready to use"),
 "APPLICATION_EXECUTION_CENTER/13_INTERVIEW_PREPARATION/Interview_Bank": ("6 core stories + hard questions + practice plan", "Template: needs your stories"),
 "APPLICATION_EXECUTION_CENTER/13_INTERVIEW_PREPARATION/English_Test_Plan": ("One-test plan; necessity check; study plan", "Complete"),
 "APPLICATION_EXECUTION_CENTER/14_SUBMISSION_CHECKLIST": ("Final checks before each submission", "Complete"),
 "APPLICATION_EXECUTION_CENTER/14_SUBMISSION_GUIDES": ("Portal-by-portal steps with copy-paste answers", "Complete (E1/E2)"),
 "DASHBOARD/Scholarship_Command_Center": ("Full database: 55 opportunities, 14 sheets, formulas", "Complete (update as evidence arrives)"),
 "00_MASTER_PROFILE/Master_Profile_and_Discrepancy_Register_v3": ("Profile + discrepancy register DR-01-DR-17", "Complete; DRs open until HR letters"),
 "00_MASTER_PROFILE/profile/master_profile": ("Single source of facts for the generator (no private data)", "Draft: pending facts marked"),
 "00_MASTER_PROFILE/generator/README": ("How to use the document generator", "Complete"),
 "00_MASTER_PROFILE/generator/generate": ("Generator: status/build/review/final/diff/export", "Complete (tested)"),
 "03_SCHOLARSHIP_RESEARCH/P1_Research_Corrections_v1": ("AISS, EDISS, CMU-Africa, SI hours, age handling", "Complete"),
 "03_SCHOLARSHIP_RESEARCH/New_Opportunities_2027_v1": ("16 new options + rejected list", "Complete (E2)"),
 "03_SCHOLARSHIP_RESEARCH/China_CSC_2027-28_Deep_Dive_v1": ("CSC routes and rules", "Complete (E2); university detail now in 11_CHINA"),
 "03_SCHOLARSHIP_RESEARCH/Verified_Opportunity_Register_v3": ("Phase 4 register (with correction notice)", "Superseded in part"),
 "APPLICATION_EXECUTION_CENTER/15_EMPLOYMENT_EVIDENCE/SantimPay_Experience_Letter_Request": ("Official request: dates per title, weekly hours, leave, duties, leadership, verifier", "Ready for your signature"),
 "APPLICATION_EXECUTION_CENTER/15_EMPLOYMENT_EVIDENCE/HR_Certificate_Draft_SantimPay": ("Suggested certificate wording for HR (letterhead, signature, stamp)", "Ready to attach"),
 "APPLICATION_EXECUTION_CENTER/15_EMPLOYMENT_EVIDENCE/Kifiya_Experience_Letter_Request": ("Official request: dates, hours, duties, verifier", "Ready for your signature"),
 "APPLICATION_EXECUTION_CENTER/15_EMPLOYMENT_EVIDENCE/HR_Certificate_Draft_Kifiya": ("Suggested certificate wording for HR", "Ready to attach"),
 "APPLICATION_EXECUTION_CENTER/16_RECOMMENDATIONS/00_Referee_Requirements_by_Programme": ("Who writes which letter, format, route, deadline", "Complete"),
 "APPLICATION_EXECUTION_CENTER/16_RECOMMENDATIONS/Academic_Draft_1_Project_Supervisor": ("Draft for the referee to edit and sign", "Ready for referee"),
 "APPLICATION_EXECUTION_CENTER/16_RECOMMENDATIONS/Academic_Draft_2_Course_Lecturer": ("Draft for the referee to edit and sign", "Ready for referee"),
 "APPLICATION_EXECUTION_CENTER/16_RECOMMENDATIONS/Professional_Draft_1_SantimPay_Manager": ("Draft for the referee (also content for SI template)", "Ready for referee"),
 "APPLICATION_EXECUTION_CENTER/16_RECOMMENDATIONS/Professional_Draft_2_Second_Referee": ("Draft for the second work referee", "Ready for referee"),
 "09_APPLICATION_DOCUMENTS/00_Document_Collection_Kit/University_Registrar_Request": ("Transcripts, UHR delivery, GPA/grading/MOI letter, class rank", "Ready to use"),
 "09_APPLICATION_DOCUMENTS/00_Document_Collection_Kit/Lecturer_Recommendation_Request": ("Academic referee request", "Ready to use"),
 "09_APPLICATION_DOCUMENTS/00_Document_Collection_Kit/Manager_Recommendation_Request": ("Work referee request", "Ready to use"),
 "09_APPLICATION_DOCUMENTS/00_Document_Collection_Kit/Personal_Information_Confirmation_Form": ("Your answers in one place (no passport number)", "Waiting for your answers"),
 "09_APPLICATION_DOCUMENTS/00_Document_Collection_Kit/Employment_Evidence_Collection_Form_v2": ("Job history + SI hours (no double counting)", "Ready to use (tested)"),
 "10_RECOMMENDATIONS/Recommender_Briefing_Pack_v1": ("Facts for referees (they write their own letters)", "Ready to use"),
 "11_INTERVIEW_PREPARATION/English_Test_and_Interview_Plan_v1": ("Earlier question bank and 5-week plan", "Reference"),
}
EDIT = {".md", ".docx", ".xlsx", ".yaml", ".py", ".json", ".html"}
rows = []
for base, (purpose, status) in P.items():
    exts = [e for e in (".md", ".docx", ".pdf", ".xlsx", ".yaml", ".py", ".json") if os.path.exists(os.path.join(ROOT, base + e))]
    if not exts:
        continue
    name = os.path.basename(base)
    m = re.search(r"_v(\d+(?:\.\d+)?)", name)
    ver = ("v" + m.group(1)) if m else "–"
    folder = os.path.dirname(base) or "(project root)"
    rows.append((name, folder, purpose, ver, " ".join(e[1:] for e in exts), status, "Yes" if any(e in EDIT for e in exts) else "PDF only", "No", "No"))
extra = [("Official evidence (46 snapshots + 3 data files)", "03_SCHOLARSHIP_RESEARCH/Official_Evidence_2026-11-07", "Saved copies of official pages (proof)", "7 Nov 2026", "txt json", "Complete", "n/a", "n/a", "n/a"),
         ("Generated drafts (12 documents)", "00_MASTER_PROFILE/generator/output", "Drafts built from the master profile", "v1", "md docx", "Draft: pending facts listed in each", "Yes", "No", "No"),
         ("Build scripts", "DASHBOARD/_build, 09_APPLICATION_DOCUMENTS/_build, 00_Document_Collection_Kit/_build", "Rebuild every workbook and the Action Center", "–", "py js html", "Complete", "Yes", "n/a", "n/a"),
         ("Archive of superseded versions", "14_ARCHIVE", "Older versions (kept, never deleted; NOT in the download)", "–", "mixed", "Historical", "–", "–", "–")]
rows += extra
H = ["File", "Location", "Purpose", "Version", "Formats", "Completion status", "Editable", "Reviewed by you", "Final"]
md = ["# Deliverables Manifest (8 Nov 2026, after Phase 6)", "",
      "Every Phase 1–5 deliverable: where it is, what it is for, and its status.",
      "",
      "- **Reviewed by you** stays \"No\" until you confirm that you have read the document.",
      "- **Final** stays \"No\" until the facts are complete and you approve it.",
      "",
      "| " + " | ".join(H) + " |", "|" + "---|" * len(H)]
for r in rows:
    md.append("| " + " | ".join(str(x).replace("|", "/") for x in r) + " |")
md += ["", "**Storage:** everything is committed locally in git (not pushed, per your instruction).",
       "",
       "A downloadable archive (`SCHOLARSHIP_PROJECT_2026-11-08.zip`) excludes:",
       "",
       "- `14_ARCHIVE`",
       "- private files",
       "- generator outputs built with private data"]
open(os.path.join(ROOT, "DELIVERABLES_MANIFEST.md"), "w").write("\n".join(md) + "\n")
wb = Workbook(); ws = wb.active; ws.title = "Manifest"
HDR = PatternFill("solid", fgColor="1F3864"); th = Side(style="thin", color="BFBFBF"); BD = Border(left=th, right=th, top=th, bottom=th)
for j, (h, w) in enumerate(zip(H, [42, 44, 56, 9, 16, 30, 10, 12, 8]), 1):
    c = ws.cell(row=1, column=j, value=h); c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10); c.fill = HDR; c.border = BD
    ws.column_dimensions[c.column_letter].width = w
for i, r in enumerate(rows, 2):
    for j, v in enumerate(r, 1):
        c = ws.cell(row=i, column=j, value=v); c.font = Font(name="Arial", size=9, color="0000FF" if j in (8, 9) else "000000"); c.border = BD
        c.alignment = Alignment(wrap_text=True, vertical="top")
ws.freeze_panes = "B2"; ws.auto_filter.ref = f"A1:I{len(rows) + 1}"
wb.save(os.path.join(ROOT, "DELIVERABLES_MANIFEST.xlsx"))
print("rows", len(rows))
