"""Builds 02_TOP_SCHOLARSHIPS.xlsx (dashboard, priority strategy, readiness scorecard, financial gap)
and 03_APPLICATION_DEADLINES.xlsx. Re-run after any change; then recalc with LibreOffice."""
import os
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

W = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
F = "Arial"; HDR = PatternFill("solid", fgColor="1F3864"); th = Side(style="thin", color="BFBFBF"); BD = Border(left=th, right=th, top=th, bottom=th)
B = Font(name=F, size=9); BB = Font(name=F, size=9, bold=True); INP = Font(name=F, size=9, color="0000FF"); LINK = Font(name=F, size=9, color="0563C1", underline="single")
WR = Alignment(wrap_text=True, vertical="top")


def header(ws, row, cols, widths):
    for j, (h, w) in enumerate(zip(cols, widths), 1):
        c = ws.cell(row=row, column=j, value=h); c.font = Font(name=F, bold=True, color="FFFFFF", size=9); c.fill = HDR; c.border = BD
        c.alignment = Alignment(wrap_text=True, vertical="center"); ws.column_dimensions[L(j)].width = w
    ws.row_dimensions[row].height = 32


def put(ws, r, vals, fonts=None, h=None):
    for j, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=j, value=v); c.font = (fonts or {}).get(j, B); c.border = BD; c.alignment = WR
    if h: ws.row_dimensions[r].height = h


# ---------------- 02_TOP_SCHOLARSHIPS.xlsx ----------------
wb = Workbook()
d = wb.active; d.title = "Dashboard"
d["A1"] = "TOP SCHOLARSHIPS: what to work on"; d["A1"].font = Font(name=F, bold=True, size=14, color="1F3864")
d["A2"] = "Labels: A strong · B possible, needs evidence · C low probability · D not eligible. Not admission probabilities. Nothing is submitted without you."; d["A2"].font = Font(name=F, size=9, italic=True)
d["A4"] = "Your 5 immediate tasks"; d["A4"].font = Font(name=F, bold=True, size=11, color="1F3864")
TASKS = [("Sign and hand in the SantimPay + Kifiya experience-letter requests (with the draft certificates) and the DTU Registrar request", "12 Nov", "15_EMPLOYMENT_EVIDENCE"),
         ("Ask 4 referees; give each the matching draft from 16_RECOMMENDATIONS", "12 Nov", "16_RECOMMENDATIONS"),
         ("Open the CMU-Africa portal; paste the exact essay questions into the pack", "15 Nov", "07_CMU_AFRICA"),
         ("Approve and send e-mails 1-4 (Ministry of Education, DUT, HUST, SMACCs)", "15 Nov", "05_READY_TO_SEND_EMAILS.docx"),
         ("Check privately: does your passport birth date meet Italy IYT's rule (born on/after 1 Jan 2000)?", "15 Nov", "12_OTHER_SCHOLARSHIPS")]
header(d, 5, ["#", "Task", "By", "Open"], [5, 80, 10, 50])
for i, (t, by, link) in enumerate(TASKS, 6):
    put(d, i, [i - 5, t, by, None])
    c = d.cell(row=i, column=4, value=link); c.hyperlink = link; c.font = LINK; c.border = BD
d["A12"] = "Top 5 opportunities (time order of work)"; d["A12"].font = Font(name=F, bold=True, size=11, color="1F3864")
header(d, 13, ["#", "Opportunity", "Label", "Folder", "Next deadline"], [5, 80, 10, 50, 22])
TOP = [(1, "CMU-Africa MSIT + Mastercard Foundation Scholars (Kigali, Rwanda)", "B", "07_CMU_AFRICA", "15 Dec 2026 (submit 12 Dec)"),
       (2, "Sweden: SI Scholarship for Global Professionals", "B", "08_SWEDEN_SI", "15 Jan / 1 Feb / 24 Feb 2027"),
       (3, "SMACCs (Erasmus Mundus): ICT specialisation", "B", "10_SMACCs", "1 Feb 2027"),
       (4, "Artificial Intelligence for Sustainable Societies (SITeS-JM), Erasmus Mundus", "B", "09_AISS", "2 Feb 2027"),
       (5, "China CSC via Dalian University of Technology (Type B)", "B", "11_CHINA", "15 Feb 2027")]
for i, row in enumerate(TOP, 14):
    put(d, i, [row[0], row[1], row[2], None, row[4]], {3: BB})
    c = d.cell(row=i, column=4, value=row[3]); c.hyperlink = row[3]; c.font = LINK; c.border = BD
d["A20"] = "Also prepared: Italy IYT (conditional on its age rule), Stipendium Hungaricum, ICCR India, EDISS (optional), Canada research MSc, Germany/USA/UK next-cycle assessment → 12_OTHER_SCHOLARSHIPS"; d["A20"].font = Font(name=F, size=9, italic=True)

# Priority strategy
ps = wb.create_sheet("Priority strategy")
cols = ["Your priority", "Opportunity", "Admission eligibility", "Scholarship eligibility", "Academic competitiveness", "Funding coverage",
        "Application cost", "Deadline", "Evidence status", "Time to give it now (and why)"]
header(ps, 1, cols, [8, 26, 34, 34, 30, 30, 16, 22, 18, 46])
PS = [
 ("1", "CMU-Africa MSIT + Mastercard Foundation", "Eligible: bachelor's degree; no GPA cut-off", "Mastercard Foundation: African citizens with need; academic talent, giving back, leadership (selected only after admission)",
  "Weak point: most admits ≥ 3.0; must show talent through the final-year trend, A− courses, a passed Kubernetes exam (expired) and code", "Full, only if selected; otherwise US$16,000 tuition + living",
  "None", "Early 15 Dec 2026; final 15 Jan 2027", "E2 (site blocked here); Mastercard Foundation E1",
  "HIGH NOW: earliest deadline, free, draft complete. About 12 hours of your time before 12 Dec"),
 ("2", "Sweden SI (Global Professionals)", "4 candidate programmes where you meet the formal requirements (E2)", "Ethiopia eligible; needs ≥ 3,000 work hours before 8 Feb 2027 + leadership, proven on SI forms",
  "Moderate: depends on programme selection (grades/motivation); SI ranks professionals", "Full: tuition + SEK 12,000/month + travel grant",
  "SEK 900", "Universities 15 Jan; fee + documents 1 Feb; SI 24 Feb (14:59 CET)", "E1",
  "HIGH from 13 Nov: best value; blocked only by the HR letters and the 13 Nov templates"),
 ("3", "SMACCs (Erasmus Mundus)", "Eligible: science/engineering bachelor's; English by test or prior English instruction", "Ethiopia: regular EMJM + NDICI third-country scholarships; ≤ 10% per nationality",
  "Moderate-low: grades and R&D are weak; work experience counts", "Full: tuition waived + €1,400/month, insurance by the consortium",
  "None found", "1 Feb 2027 (scholarship)", "E2 (smaccs.eu via search)",
  "MEDIUM: reuses the AISS/CV material; letter drafted; about 6 hours"),
 ("4", "Artificial Intelligence for Sustainable Societies (SITeS-JM)", "Eligible", "Top-10 ranking only; at most 1 Ethiopian; self-funding plan mandatory",
  "Moderate-low: transcript scored inside documents (50%)", "Full: tuition waived + €1,400/month",
  "IELTS ETB 32,800", "2 Feb 2027", "E1",
  "MEDIUM-HIGH: portfolio, video and IELTS take time; start the demo repos now"),
 ("5", "China CSC via DUT (Type B)", "No published GPA floor; ask which CS master's are English-taught", "CSC Type B; age rule (check privately)",
  "Unknown (selection by university)", "Nearly full: tuition, housing, CNY 3,000/month, insurance; airfare not covered",
  "CNY 400", "15 Feb 2027", "E2",
  "MEDIUM: send the DUT e-mail now; study plan drafted"),
 ("6", "Stipendium Hungaricum", "University entrance exam (e.g. ELTE)", "MoE nomination; 2027/28 call pending",
  "Moderate", "Nearly full: tuition + dorm + about €110/month", "None", "About mid-Jan (expected)", "E1 (2026/27 terms)",
  "LOW until the call opens; package ready"),
 ("7", "ICCR India", "Indian university admission", "Ethiopia quota (130 places in 2026-27)", "Moderate",
  "Nearly full: tuition + INR 20,000/month + housing allowance; airfare depends on the scheme", "None", "Portal about 27 Feb – 15 Apr 2027 (expected)", "E2",
  "LOW now; act in February"),
 ("8", "Funded research MSc (Manitoba / Calgary)", "Supervisor needed first", "Department funding (Manitoba min. $20,000/yr)", "Weak (last two years 2.98)",
  "Partial", "Calgary CAD 145", "Manitoba 1 Dec; Calgary 15 Jan", "E1",
  "ONLY if a professor agrees this week"),
 ("cond.", "EDISS (Erasmus Mundus)", "Eligible", "Erasmus Mundus scholarship", "Low: 70% of the score from grades", "Full", "IELTS", "12 Jan 2027", "E1",
  "OPTIONAL: at most one day of effort"),
]
for i, r in enumerate(PS, 2):
    put(ps, i, list(r), {2: BB, 10: BB}, 70)
ps.freeze_panes = "C2"
r = len(PS) + 3
for t in ["Why this order: value (full funding) × feasibility (eligibility proven or likely) ÷ cost, adjusted for deadline urgency and reuse of materials.",
          "CMU-Africa first because it is earliest, free and already drafted. SI second because it is the best-value full scholarship and E1-verified. SMACCs and AISS share material (CV, portfolio, references), so both cost little extra once one is done.",
          "China DUT is the cheapest 'nearly full' option with no published GPA floor. Hungary and ICCR open later, so they get little time now. EDISS is optional because grades decide 70% of its score."]:
    ps.cell(row=r, column=1, value=t).font = Font(name=F, size=9, italic=True); r += 1

# Readiness scorecard (Phase 6 re-score)
sc = wb.create_sheet("Readiness scorecard")
DIMS = ["Academic fit", "Scholarship eligibility", "Documents", "Professional profile", "English", "Leadership", "Financial evidence",
        "Recommendations", "Deadline feasibility", "Official verification"]
header(sc, 1, ["Opportunity"] + DIMS + ["Total /40", "Readiness %", "Label", "What would raise it"], [34] + [9] * 10 + [8, 10, 7, 46])
SC = [
 ("CMU-Africa + Mastercard Foundation", [2, 3, 2, 3, 3, 2, 1, 2, 3, 2], "Prompts copied; HR letters; financial figures; demo repo"),
 ("Sweden SI", [3, 2, 1, 3, 4, 2, 4, 2, 4, 4], "HR letters proving 3,000 h; leadership forms; 13 Nov templates"),
 ("SMACCs", [2, 3, 2, 3, 3, 3, 4, 2, 4, 2], "Confirm facts on smaccs.eu; 2 reference letters; letter final"),
 ("AISS (SITeS-JM)", [2, 3, 2, 3, 2, 3, 2, 4, 4, 4], "Portfolio repos; IELTS; self-funding figures; video"),
 ("China CSC via DUT (Type B)", [2, 3, 2, 3, 3, 4, 4, 2, 4, 2], "DUT confirms an English-taught CS master's; 2 professor letters"),
 ("China CSC Type A (MoE)", [2, 1, 1, 3, 3, 4, 4, 2, 2, 1], "MoE publishes the call and agency number"),
 ("Stipendium Hungaricum", [2, 2, 1, 2, 3, 4, 4, 2, 3, 3], "2027/28 call; MoE nomination"),
 ("ICCR India", [2, 3, 1, 2, 3, 4, 4, 3, 4, 2], "Portal opens; choose 5 universities"),
 ("Invest Your Talent Italy", [2, 3, 1, 2, 2, 4, 4, 3, 4, 2], "Spring call; participating programme admits you"),
 ("EDISS", [1, 3, 2, 2, 2, 4, 4, 2, 4, 4], "Grades decide 70%: cannot be raised now"),
 ("HUST (China)", [1, 3, 1, 3, 3, 4, 4, 2, 4, 2], "Written confirmation that 2.62/4.00 is acceptable"),
 ("NPU (China)", [0, 3, 1, 3, 3, 4, 4, 2, 4, 2], "Not eligible: 80% average required for CSC applicants"),
 ("Manitoba MSc (thesis)", [1, 2, 1, 3, 1, 4, 4, 2, 1, 4], "A supervisor this week"),
 ("Calgary MSc (thesis)", [1, 2, 1, 3, 2, 4, 4, 2, 3, 4], "A professor agrees to supervise"),
 ("VLIR-UOS Belgium", [2, 1, 1, 2, 2, 4, 4, 2, 4, 1], "2027-28 list with an ICT/data programme"),
]
for i, (name, s, nxt) in enumerate(SC, 2):
    sc.cell(row=i, column=1, value=name).font = BB
    for k, v in enumerate(s):
        c = sc.cell(row=i, column=2 + k, value=v); c.font = INP; c.alignment = Alignment(horizontal="center")
    sc.cell(row=i, column=12, value=f"=SUM(B{i}:K{i})").font = B
    p = sc.cell(row=i, column=13, value=f"=L{i}/40"); p.number_format = "0%"; p.font = B
    sc.cell(row=i, column=14, value=f'=IF(OR(B{i}=0,C{i}=0,J{i}=0),"D",IF(OR(B{i}<=1,C{i}<=1),"C",IF(AND(B{i}>=3,C{i}>=3,M{i}>=0.7),"A","B")))').font = BB
    c = sc.cell(row=i, column=15, value=nxt); c.font = B; c.alignment = WR
    for col in range(1, 16): sc.cell(row=i, column=col).border = BD
last = len(SC) + 1
dv = DataValidation(type="whole", operator="between", formula1="0", formula2="4"); sc.add_data_validation(dv); dv.add(f"B2:K{last}")
for lab, fill in (("A", "C6EFCE"), ("B", "DDEBF7"), ("C", "FFEB9C"), ("D", "FFC7CE")):
    sc.conditional_formatting.add(f"N2:N{last}", CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=fill)))
r = last + 2
for t in ["Each dimension 0-4 (rubric: 4 = met/ready … 0 = impossible). Label rules: D if Academic fit, Scholarship eligibility or Deadline = 0; C if Academic fit or Scholarship eligibility <= 1; A if both >= 3 and readiness >= 70%; otherwise B.",
          "Re-scored on 8 Nov 2026 after the Phase 6 drafts. Change the blue scores as evidence arrives; the label updates itself."]:
    sc.cell(row=r, column=1, value=t).font = Font(name=F, size=9, italic=True); r += 1
sc.freeze_panes = "B2"

# Financial gap summary
fg = wb.create_sheet("Financial gap")
header(fg, 1, ["Opportunity", "Class", "Covered", "Your gap", "Application cost"], [30, 18, 50, 50, 18])
FG = [("CMU-Africa + Mastercard Foundation", "Full (if selected)", "Tuition, living, travel, insurance", "If not selected: US$16,000 tuition + living", "None"),
      ("Sweden SI", "Full", "Tuition, SEK 12,000/month, SEK 15,000 travel", "Application fee; residence-permit fee; insurance check", "SEK 900"),
      ("SMACCs", "Full", "Tuition waived, €1,400/month, insurance", "Mobility, visas, Summer/Winter Schools from the stipend", "None found"),
      ("AISS (SITeS-JM)", "Full", "Tuition waived, €1,400/month", "Budget mobility from the stipend", "IELTS ETB 32,800"),
      ("China DUT (Type B)", "Nearly full", "Tuition, housing, CNY 3,000/month, insurance", "International airfare; visa; medical; notarisation", "CNY 400"),
      ("Stipendium Hungaricum", "Nearly full", "Tuition, dorm, HUF 43,700/month (about €110), insurance", "Living costs above about €110/month; travel", "None"),
      ("ICCR India", "Nearly full", "Tuition, INR 20,000/month, house-rent allowance", "Airfare (scheme-dependent)", "None"),
      ("Manitoba / Calgary", "Partial", "Manitoba min. $20,000/yr", "Tuition + living above the package; travel", "Calgary CAD 145")]
for i, r_ in enumerate(FG, 2):
    put(fg, i, list(r_), {1: BB}, 30)
wb.save(os.path.join(W, "02_TOP_SCHOLARSHIPS.xlsx"))

# ---------------- 03_APPLICATION_DEADLINES.xlsx ----------------
wb2 = Workbook(); ws = wb2.active; ws.title = "Deadlines"
ws["A1"] = "Application deadlines (days left update automatically)"; ws["A1"].font = Font(name=F, bold=True, size=13, color="1F3864")
header(ws, 3, ["Date", "Days left", "What", "Opportunity", "Type", "Evidence", "Folder", "Done? (Y)"], [12, 9, 60, 26, 12, 10, 26, 9])
DL = [
 (date(2026, 11, 12), "Hand in HR + Registrar requests; ask 4 referees; Manitoba go/no-go", "All", "Your target", "–", "09_APPLICATION_DOCUMENTS/00_Document_Collection_Kit"),
 (date(2026, 11, 13), "SI templates + eligible programme list published", "Sweden SI", "Official", "E1", "08_SWEDEN_SI"),
 (date(2026, 11, 16), "EDISS round opens (08:00 Helsinki)", "EDISS", "Official", "E1", "12_OTHER_SCHOLARSHIPS"),
 (date(2026, 11, 20), "Book IELTS (only if you confirm AISS)", "AISS/EDISS", "Your target", "–", "13_INTERVIEW_PREPARATION"),
 (date(2026, 11, 30), "Demo repos built; DUT reply received", "AISS/CMU/China", "Your target", "–", "09_AISS"),
 (date(2026, 12, 1), "University of Manitoba MSc CS (only if go)", "Manitoba", "Official", "E1", "12_OTHER_SCHOLARSHIPS"),
 (date(2026, 12, 2), "Artificial Intelligence for Sustainable Societies (SITeS-JM) window opens", "AISS (SITeS-JM)", "Official", "E1", "09_AISS"),
 (date(2026, 12, 12), "Submit CMU-Africa (your target)", "CMU-Africa", "Your target", "–", "07_CMU_AFRICA"),
 (date(2026, 12, 15), "CMU-Africa early deadline (Duolingo covered)", "CMU-Africa", "Official", "E2", "07_CMU_AFRICA"),
 (date(2027, 1, 12), "EDISS deadline 15:00 Helsinki (optional)", "EDISS", "Official", "E1", "12_OTHER_SCHOLARSHIPS"),
 (date(2027, 1, 15), "universityadmissions.se deadline (4 programmes); CMU-Africa final; Calgary early review", "Sweden SI / CMU / Calgary", "Official", "E1/E2/E1", "08_SWEDEN_SI"),
 (date(2027, 1, 15), "Stipendium Hungaricum (expected, 2026 pattern)", "Hungary", "Expected", "E3", "12_OTHER_SCHOLARSHIPS"),
 (date(2027, 1, 25), "Submit SMACCs (your target)", "SMACCs", "Your target", "–", "10_SMACCs"),
 (date(2027, 1, 30), "Submit AISS (SITeS-JM) (your target)", "AISS (SITeS-JM)", "Your target", "–", "09_AISS"),
 (date(2027, 2, 1), "SMACCs scholarship deadline; Sweden SEK 900 fee + documents", "SMACCs / Sweden", "Official", "E2/E1", "10_SMACCs"),
 (date(2027, 2, 2), "Artificial Intelligence for Sustainable Societies (SITeS-JM) deadline", "AISS (SITeS-JM)", "Official", "E1", "09_AISS"),
 (date(2027, 2, 8), "SI portal opens", "Sweden SI", "Official", "E1", "08_SWEDEN_SI"),
 (date(2027, 2, 15), "DUT CSC Type B deadline", "China DUT", "Official", "E2", "11_CHINA"),
 (date(2027, 2, 24), "SI portal closes 14:59 CET", "Sweden SI", "Official", "E1", "08_SWEDEN_SI"),
 (date(2027, 2, 27), "ICCR portal expected to open (2026 pattern)", "ICCR", "Expected", "E2", "12_OTHER_SCHOLARSHIPS"),
 (date(2027, 4, 1), "Sweden admission results", "Sweden SI", "Official", "E1", "08_SWEDEN_SI"),
 (date(2027, 4, 7), "SMACCs interviews 7–18 April (if invited)", "SMACCs", "Official", "E2", "10_SMACCs"),
 (date(2027, 4, 22), "SI scholarship results", "Sweden SI", "Official", "E1", "08_SWEDEN_SI"),
 (date(2027, 5, 9), "SMACCs scholarship results", "SMACCs", "Official", "E2", "10_SMACCs"),
 (date(2027, 4, 20), "Italy IYT 2027/28 call expected (2026 call ran 20 Apr - 11 May)", "Italy IYT", "Expected", "E2", "12_OTHER_SCHOLARSHIPS"),
 (date(2027, 6, 1), "ASU Mastercard Foundation (Phase III, MS IT): watch for the 2028/29 call (2026 deadline was 27 Sep)", "USA ASU (Mastercard Foundation)", "Expected", "E2 pattern", "12_OTHER_SCHOLARSHIPS"),
 (date(2027, 8, 1), "DAAD EPOS course deadlines for 2028 entry start (only if the DTU rank letter shows the upper third)", "Germany DAAD EPOS", "Expected", "E2/E3", "12_OTHER_SCHOLARSHIPS"),
 (date(2027, 8, 4), "Chevening 2028/29 window expected to open (2026 pattern: 4 Aug - 6 Oct); needs 2,800 h after graduation (E1)", "UK Chevening", "Expected", "E1 pattern", "12_OTHER_SCHOLARSHIPS"),
]
for i, (dt, what, opp, typ, ev, folder) in enumerate(DL, 4):
    c = ws.cell(row=i, column=1, value=dt); c.number_format = "dd-mmm-yyyy"; c.font = BB; c.border = BD
    c = ws.cell(row=i, column=2, value=f'=IF(H{i}="Y","done",A{i}-TODAY())'); c.font = B; c.border = BD
    for j, v in enumerate([what, opp, typ, ev], 3):
        c = ws.cell(row=i, column=j, value=v); c.font = B; c.border = BD; c.alignment = WR
    c = ws.cell(row=i, column=7, value=folder); c.hyperlink = folder; c.font = LINK; c.border = BD
    c = ws.cell(row=i, column=8, value=None); c.font = INP; c.border = BD
n = len(DL) + 3
ws.conditional_formatting.add(f"B4:B{n}", FormulaRule(formula=[f'AND(ISNUMBER(B4),B4<0)'], fill=PatternFill("solid", fgColor="D9D9D9")))
ws.conditional_formatting.add(f"B4:B{n}", FormulaRule(formula=[f'AND(ISNUMBER(B4),B4>=0,B4<=14)'], fill=PatternFill("solid", fgColor="FFC7CE")))
ws.conditional_formatting.add(f"B4:B{n}", FormulaRule(formula=[f'AND(ISNUMBER(B4),B4>14,B4<=45)'], fill=PatternFill("solid", fgColor="FFEB9C")))
ws.freeze_panes = "C4"
wb2.save(os.path.join(W, "03_APPLICATION_DEADLINES.xlsx"))
print("saved 02 and 03")
