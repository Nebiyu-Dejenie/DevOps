"""Builds DASHBOARD/Readiness_Scorecard_v1.xlsx: 10-dimension readiness (0-4 each) + A-D label by explicit rules.
NOT an admission probability. Scores reflect evidence on the date below; edit the blue cells as evidence arrives."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter as L
F = "Arial"; HDR = PatternFill("solid", fgColor="1F3864"); HF = Font(name=F, bold=True, color="FFFFFF", size=9)
B = Font(name=F, size=9); INP = Font(name=F, size=9, color="0000FF"); BOLD = Font(name=F, size=9, bold=True)
th = Side(style="thin", color="BFBFBF"); BD = Border(left=th, right=th, top=th, bottom=th); WR = Alignment(wrap_text=True, vertical="top")
DATE = "8 Nov 2026"
DIMS = ["Academic eligibility / fit", "Scholarship eligibility", "Document readiness", "Professional profile", "English readiness",
        "Leadership evidence", "Financial evidence", "Recommendation readiness", "Deadline feasibility", "Official verification"]
RUB = [
 ("Academic eligibility / fit", "4 = no grade-based selection, or you exceed the requirement | 3 = grades considered but not dominant | 2 = grades important and you are below typical admits | 1 = grades dominant or a published minimum is borderline | 0 = below a published minimum"),
 ("Scholarship eligibility", "4 = every published condition met and documented | 3 = met, routine documents pending | 2 = depends on key evidence still missing (e.g. HR hours) or a pending call | 1 = uncertain (nationality list, programme or route unconfirmed) | 0 = ineligible"),
 ("Document readiness", "4 = all documents in hand | 3 = finishing touches | 2 = drafts ready, key documents requested | 1 = templates and plans only | 0 = not started"),
 ("Professional profile", "4 = verified experience strongly matching the criteria | 3 = strong match, verification pending | 2 = partly relevant | 1 = weakly relevant | 0 = none"),
 ("English readiness", "4 = requirement met or not required | 3 = test booked | 2 = test planned and feasible | 1 = timing very tight | 0 = cannot be met before the deadline"),
 ("Leadership evidence", "4 = documented and signed, or not assessed | 3 = examples confirmed, documents pending | 2 = candidate examples unconfirmed | 1 = assessed and none yet | 0 = required and impossible"),
 ("Financial evidence", "4 = not needed, or ready | 3 = figures gathered | 2 = plan drafted | 1 = needed, not started | 0 = cannot be provided"),
 ("Recommendation readiness", "4 = letters secured, or none required | 3 = referees agreed | 2 = requests drafted, not sent | 1 = referees not identified | 0 = required type unavailable"),
 ("Deadline feasibility", "4 = more than 8 weeks and on track | 3 = 4-8 weeks | 2 = under 4 weeks, feasible | 1 = very tight | 0 = passed or impossible"),
 ("Official verification", "4 = all key facts E1 | 3 = mostly E1 | 2 = E1/E2 mix | 1 = mostly E2/E3 | 0 = unverified"),
]
ROWS = [  # id, name, 10 scores, why, next step
 ("S02", "Sweden SI (Global Professionals)", [3, 2, 1, 3, 4, 2, 4, 2, 4, 4], "No SI GPA rule (universities may weigh grades). 3,000 h unproven until HR letters. Leadership unconfirmed. Templates out 13 Nov.", "HR letters; 2 leadership examples; pick 4 programmes on 13-20 Nov"),
 ("S30", "AISS (Erasmus Mundus)", [2, 3, 1, 3, 2, 3, 2, 4, 4, 4], "Transcript scored inside documents (50%). 10 scholarships, at most 1 per nationality. Portfolio not built; IELTS not taken; self-funding plan needed.", "Build demo repos P5/P6; IELTS; self-funding figures"),
 ("S01", "CMU-Africa MSIT + Mastercard Foundation", [2, 3, 2, 3, 3, 2, 1, 2, 3, 1], "No cut-off but most admits >= 3.0. Tests optional by 15 Dec. Financial-need evidence not gathered. Site only E2.", "Open the site; essays; 3 referees; financial figures; submit by 12 Dec"),
 ("S31", "EDISS (Erasmus Mundus)", [1, 3, 2, 2, 2, 4, 4, 2, 4, 4], "70 of 100 points from grades, weighted to your weakest areas; experience does not compensate.", "Apply only if IELTS + academic letters are ready anyway"),
 ("S39", "SMACCs (Erasmus Mundus)", [1, 3, 1, 3, 2, 3, 4, 2, 4, 1], "Merit score includes academic record and R&D experience (both weak) plus work experience (strong). Weights and deadline not verified on the official page.", "Read the official SMACCs pages in a browser"),
 ("S40", "CDE Copernicus Digital Earth (Erasmus Mundus)", [1, 3, 1, 2, 2, 4, 4, 2, 4, 2], "Needs Earth-observation/GIS evidence you do not have.", "Apply only with real geodata work"),
 ("S09", "Stipendium Hungaricum (ELTE / Szeged)", [2, 2, 1, 2, 3, 4, 4, 2, 3, 3], "Ethiopia is a sending partner (2026/27). 2027/28 call pending; ELTE entrance exam; MoE nomination.", "Ask the MoE (message 1); watch the call"),
 ("S05", "China CSC Type A (bilateral, via MoE)", [2, 2, 1, 3, 3, 4, 4, 2, 2, 1], "Ethiopia call not yet published; agency number unknown; pre-admission letter required.", "Ask the MoE; request pre-admission letters"),
 ("S43", "China CSC Type B (via universities)", [2, 3, 1, 3, 3, 4, 4, 2, 4, 1], "University windows open (DUT 15 Feb; HUST 25 Feb; NPU 28 Feb). E2 only.", "Apply to 3 universities; study plan per university"),
 ("S35", "ICCR (India)", [2, 3, 1, 2, 3, 4, 4, 3, 4, 1], "130 places for Ethiopians in 2026-27. Admission decided by Indian universities. Portal expected about late Feb 2027.", "Shortlist 5 Indian universities (MCA / M.Tech CS)"),
 ("S38", "Invest Your Talent in Italy", [2, 3, 1, 2, 2, 4, 4, 3, 4, 2], "Ethiopia eligible (official FAQ). Needs admission to a participating master's. 2027-28 call expected in spring.", "Shortlist participating ICT master's"),
 ("S37", "VLIR-UOS ICP Connect (Belgium)", [2, 1, 1, 2, 2, 4, 4, 2, 4, 1], "2027-28 country and programme lists not yet published; no ICT/data programme confirmed.", "Check the lists in mid-Nov, then re-score"),
 ("S10", "University of Manitoba MSc CS (thesis)", [1, 2, 1, 3, 1, 4, 4, 2, 1, 4], "Years 4-5 average 2.98 (borderline). Supervisor support needed before applying; 1 Dec deadline. Min. $20,000/yr if admitted (E1).", "Decide go/no-go by 12 Nov"),
 ("S32", "University of Calgary MSc CS (thesis)", [1, 2, 1, 3, 2, 4, 4, 2, 3, 4], "GPA competitiveness assessed; a professor must agree to supervise; minimum to verify on the graduate-studies page.", "Write to 3 matching professors (C1-C3)"),
 ("S49", "Mauritius-Africa Scholarship Scheme", [2, 1, 1, 2, 3, 4, 3, 2, 3, 2], "MPhil/PhD routes; Ethiopian nominating agency unconfirmed.", "Ask the MoE (message 1)"),
 ("S44", "Taiwan (MOE / university scholarships)", [2, 1, 1, 2, 3, 4, 3, 2, 4, 2], "Route for Ethiopians unclear (no Taiwan office in Ethiopia); university awards are partial.", "Ask NTUT (message 6)"),
]
wb = Workbook(); r0 = wb.active; r0.title = "How to read"
r0["A1"] = "Readiness Scorecard v1 (" + DATE + ")"; r0["A1"].font = Font(name=F, bold=True, size=14, color="1F3864")
lines = ["This is NOT an admission probability. It measures how ready and how well-matched each application is, using 10 dimensions scored 0-4.",
         "Readiness % = total / 40. Labels follow fixed rules (no judgement inside the formula):",
         "  D (not eligible): Academic = 0, Scholarship eligibility = 0 or Deadline = 0",
         "  C (low probability): Academic <= 1 or Scholarship eligibility <= 1",
         "  A (strong and realistic): Academic >= 3, Scholarship eligibility >= 3 and readiness >= 70%",
         "  B (possible, needs evidence): everything else",
         "Blue cells are scores: change them as evidence arrives (e.g. HR letters raise SI's Scholarship eligibility from 2 to 4 if they prove 3,000 hours).",
         "Only 2027-intake options are scored. Next-cycle options are listed in the Command Center."]
for i, t in enumerate(lines, 3):
    c = r0.cell(row=i, column=1, value=t); c.font = Font(name=F, size=10); c.alignment = WR
r0.column_dimensions["A"].width = 130
ru = wb.create_sheet("Rubric")
for j, (t, w) in enumerate(zip(["Dimension", "What each score means"], [28, 150]), 1):
    x = ru.cell(row=1, column=j, value=t); x.font, x.fill, x.border = HF, HDR, BD; ru.column_dimensions[L(j)].width = w
for i, (d, t) in enumerate(RUB, 2):
    ru.cell(row=i, column=1, value=d).font = BOLD; c = ru.cell(row=i, column=2, value=t); c.font = B; c.alignment = WR
    ru.row_dimensions[i].height = 30
sc = wb.create_sheet("Scorecard", 1)
heads = ["ID", "Opportunity"] + DIMS + ["Total /40", "Readiness %", "Label", "Why (evidence)", "Next step to raise it"]
widths = [6, 30] + [9] * 10 + [8, 10, 7, 60, 40]
for j, (t, w) in enumerate(zip(heads, widths), 1):
    x = sc.cell(row=1, column=j, value=t); x.font, x.fill, x.border = HF, HDR, BD
    x.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center"); sc.column_dimensions[L(j)].width = w
sc.row_dimensions[1].height = 54
for i, (oid, name, s, why, nxt) in enumerate(ROWS, 2):
    sc.cell(row=i, column=1, value=oid).font = B; sc.cell(row=i, column=2, value=name).font = BOLD
    for k, v in enumerate(s):
        c = sc.cell(row=i, column=3 + k, value=v); c.font = INP; c.alignment = Alignment(horizontal="center")
    sc.cell(row=i, column=13, value=f"=SUM(C{i}:L{i})").font = B
    p = sc.cell(row=i, column=14, value=f"=M{i}/40"); p.number_format = "0%"; p.font = B
    sc.cell(row=i, column=15, value=f'=IF(OR(C{i}=0,D{i}=0,K{i}=0),"D",IF(OR(C{i}<=1,D{i}<=1),"C",IF(AND(C{i}>=3,D{i}>=3,N{i}>=0.7),"A","B")))').font = BOLD
    for col, t in ((16, why), (17, nxt)):
        c = sc.cell(row=i, column=col, value=t); c.font = B; c.alignment = WR
    for col in range(1, 18):
        sc.cell(row=i, column=col).border = BD
    sc.row_dimensions[i].height = 42
last = 1 + len(ROWS)
dv = DataValidation(type="whole", operator="between", formula1="0", formula2="4", allow_blank=False); sc.add_data_validation(dv); dv.add(f"C2:L{last}")
for lab, fill in (("A", "C6EFCE"), ("B", "DDEBF7"), ("C", "FFEB9C"), ("D", "FFC7CE")):
    sc.conditional_formatting.add(f"O2:O{last}", CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=fill)))
sc.conditional_formatting.add(f"C2:L{last}", CellIsRule(operator="lessThanOrEqual", formula=["1"], fill=PatternFill("solid", fgColor="FCE4D6")))
sc.freeze_panes = "C2"
wb.save("../Readiness_Scorecard_v1.xlsx"); print("saved", len(ROWS))
