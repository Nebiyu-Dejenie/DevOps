"""Builds 15_EMPLOYMENT_EVIDENCE/Experience_Calculation.xlsx (Phase 8).

One workbook for all employment evidence:
  - employer-by-employer, title-by-title periods (data from the HR letters only),
  - three SEPARATE eligibility tests, each month counted once (no double counting):
      Swedish Institute  >= 3,000 h before 8 Feb 2027, max 3 employers, no internships        (E1 si.se)
      Chevening          >= 2,800 h after the bachelor's degree, internships/voluntary count   (E1 chevening.org)
      DAAD EPOS          >= 24 months of related professional experience after the degree      (E2/E3, daad.de blocked)
  - an illustrative scenario that is clearly NOT evidence (your statement + documented assumptions),
  - the evidence-collection sheets moved from Employment_Evidence_Collection_Form_v2.xlsx
    (that form is archived in 14_ARCHIVE/phase8_superseded; this workbook replaces it).

Run:  cd APPLICATION_EXECUTION_CENTER/_build && python3 build_experience.py
      then recalculate with LibreOffice (xlsx skill recalc.py) so cached values exist.
"""
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L

OUT = "../15_EMPLOYMENT_EVIDENCE/Experience_Calculation.xlsx"
F = "Arial"
NAVY = "1F3864"
HDR = PatternFill("solid", fgColor=NAVY); HF = Font(name=F, bold=True, color="FFFFFF", size=10)
BODY = Font(name=F, size=10); INP = Font(name=F, size=10, color="0000FF"); BOLD = Font(name=F, size=10, bold=True)
SMALL = Font(name=F, size=9, italic=True)
YEL = PatternFill("solid", fgColor="FFF2CC"); GREY = PatternFill("solid", fgColor="EDEDED")
GREEN = PatternFill("solid", fgColor="C6EFCE"); RED = PatternFill("solid", fgColor="FFC7CE"); AMB = PatternFill("solid", fgColor="FFEB9C")
th = Side(style="thin", color="BFBFBF"); BD = Border(left=th, right=th, top=th, bottom=th)
WR = Alignment(wrap_text=True, vertical="top")
EV = ["Pending – awaiting evidence", "E0 – own document (HR letter or contract)", "E1 – verified with employer or official source",
      "E4 – my confirmation only", "E5 – unsupported or contradictory"]
TYPES = ["Full-time employee", "Part-time employee", "Fixed-term contract", "Consultant or freelance", "Entrepreneurship",
         "Internship", "Voluntary work"]
DOCS = ["HR letter received", "HR letter requested", "Contract or payslips", "SI or Chevening form signed", "Not requested", "None"]


def title(ws, text, sub=None):
    ws["A1"] = text; ws["A1"].font = Font(name=F, bold=True, size=14, color=NAVY)
    if sub:
        ws["A2"] = sub; ws["A2"].font = SMALL


def hdr(ws, row, cols, widths=None, height=42):
    for i, c in enumerate(cols, 1):
        x = ws.cell(row=row, column=i, value=c); x.font, x.fill, x.border = HF, HDR, BD
        x.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        if widths:
            ws.column_dimensions[x.column_letter].width = widths[i - 1]
    ws.row_dimensions[row].height = height


def dv(ws, opts, rng):
    d = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=True)
    ws.add_data_validation(d); d.add(rng)


def cell(ws, r, c, v, font=BODY, fmt=None, fill=None, wrap=True):
    x = ws.cell(row=r, column=c, value=v); x.font = font; x.border = BD
    if wrap: x.alignment = WR
    if fmt: x.number_format = fmt
    if fill: x.fill = fill
    return x


wb = Workbook()

# =====================================================================================================
# 1. Read me
# =====================================================================================================
rm = wb.active; rm.title = "Read me"
title(rm, "Experience calculation: employer by employer, three separate tests",
      "Built 8 Nov 2026 (Phase 8). Replaces Employment_Evidence_Collection_Form_v2.xlsx. Nothing here is evidence until an HR letter supports it.")
lines = [
    ("WHAT THIS DOES", True),
    ("1. You enter each job title from the HR letters on 'Employment periods' (one row per title and period).", False),
    ("2. 'Timeline' spreads every documented period over calendar months. In a month with two jobs, only the job with more hours counts (no double counting).", False),
    ("3. 'Results' shows, per employer, the documented hours before and after graduation, and three separate tests: Swedish Institute, Chevening and DAAD EPOS.", False),
    ("4. 'Scenario (not evidence)' shows what your statement of 3 years 6 months would give under stated assumptions. Never copy it into an application.", False),
    ("", False),
    ("RULES USED (see 'Inputs & rules' for sources and dates read)", True),
    ("Swedish Institute (E1, si.se, read 8 Nov 2026): at least 3,000 hours before 8 Feb 2027; from at most 3 employers; full-time, part-time, freelancing or entrepreneurship count; internships do not. Formula: hours/week × weeks/year × years.", False),
    ("Chevening (E1, chevening.org, read 8 Nov 2026): at least 2,800 hours acquired after the undergraduate degree, over at least two years; full-time, part-time, voluntary work and paid or unpaid internships count; up to 15 periods; weeks worked × hours per week.", False),
    ("DAAD EPOS (E2/E3; daad.de blocks automated access): at least two years of professional experience, normally after the bachelor's and related to the course; degree not older than six years; upper third of the class (not tested here). Confirm on the course page.", False),
    ("", False),
    ("HOW TO FILL IT", True),
    ("Fill the BLUE cells only. Black cells are formulas. Yellow cells on 'Inputs & rules' are dates and thresholds you may update when official pages change.", False),
    ("Copy dates, titles and weekly hours EXACTLY from the HR letter (Gregorian calendar). Use one row per job title.", False),
    ("End date is required. For a job you still hold, enter the date of the HR letter ('employed to date'). Update it when SI's form is signed in February 2027.", False),
    ("Weeks worked per year: 52 minus annual leave and public holidays, as stated or confirmed by HR. If the letter says only 'full-time', ask HR for this figure.", False),
    ("A row counts in the tests only if its evidence level is E0 or E1 AND start, end, hours/week and weeks/year are filled. E4 (your confirmation) and E5 never count.", False),
    ("Related to IT/CS (column I) decides the EPOS test. Answer honestly from the duties on the HR letter.", False),
    ("SI counts at most three employers: column O lets you exclude rows (N) if more than three employers are documented.", False),
    ("", False),
    ("WHY 'EACH MONTH COUNTED ONCE'", True),
    ("SI's FAQ adds separate experiences together but does not say whether jobs held at the same time may be added. Chevening's portal adds each entry. The safe figure counts the larger job in a month only; the 'added' figure is shown beside it. Decisions use the safe figure.", False),
    ("Months where one job ends and another starts are slightly under-counted by this method (it keeps the larger part). The error is small and always on the safe side.", False),
    ("", False),
    ("EVIDENCE SCALE", True),
    ("E0 your own official document | E1 official source or employer verified | E2 official page seen only through search | E3 secondary | E4 your confirmation only | E5 unsupported or contradictory (never used).", False),
]
for i, (t, b) in enumerate(lines, 4):
    c = rm.cell(row=i, column=1, value=t); c.font = Font(name=F, size=10, bold=b, color=NAVY if b else "000000"); c.alignment = WR
rm.column_dimensions["A"].width = 150

# =====================================================================================================
# 2. Inputs & rules
# =====================================================================================================
ip = wb.create_sheet("Inputs & rules")
title(ip, "Inputs and rules", "Yellow cells feed every formula. Change them only when an official source changes; record the source and date.")
hdr(ip, 3, ["Item", "Value", "Evidence", "Source and note"], [52, 16, 14, 110], height=24)
INPUTS = [  # (label, value, fmt, evidence, note)   rows 4..
    ("Bachelor's degree awarded (graduation date)", date(2022, 7, 14), "dd-mmm-yyyy", "E0", "DTU official transcript (read 7 Nov 2026)."),
    ("SI: hours count only BEFORE this date", date(2027, 2, 8), "dd-mmm-yyyy", "E1", "si.se SGP page and FAQ (read 8 Nov 2026): 'at least 3,000 hours of work experience prior to 8 February 2027'."),
    ("SI: minimum hours", 3000, "#,##0", "E1", "Ethiopia is on SI's 3,000-hour list (si.se, 8 Nov 2026). Internships do not count."),
    ("SI: maximum number of employers", 3, "0", "E1", "'The 3,000 hours must have been accumulated from a maximum of three employers.'"),
    ("Chevening: count hours up to (application deadline)", date(2027, 10, 6), "dd-mmm-yyyy", "Assumption",
     "The 2027/28 window closed 6 Oct 2026 (E1 GOV.UK). The 2028/29 deadline is not announced yet: replace this date when Chevening publishes it (expected Aug–Oct 2027)."),
    ("Chevening: minimum hours after graduation", 2800, "#,##0", "E1",
     "chevening.org work-experience page (read 8 Nov 2026): 2,800 hours after the undergraduate degree; full-time, part-time, voluntary work, paid or unpaid internships; up to 15 periods."),
    ("Chevening: degree completed at least this many months before the deadline", 24, "0", "E1",
     "'You must have acquired this experience over a minimum of two years' (current cycle: degree completed by Oct 2024)."),
    ("DAAD EPOS: count related experience up to (course deadline)", date(2027, 8, 31), "dd-mmm-yyyy", "Assumption",
     "EPOS course deadlines differ (most fall Aug–Oct for the following year's intake). Set this to your course's deadline."),
    ("DAAD EPOS: minimum months of related experience after graduation", 24, "0", "E2/E3",
     "DAAD EPOS: at least two years of professional experience (daad.de blocked this session; confirm on the course page and DAAD's EPOS FAQ)."),
    ("Full-time reference for EPOS full-time-equivalent months (hours/week)", 35, "0", "Assumption",
     "EPOS gives no figure. 35 h is Chevening's own definition of a full-time week (E1). Used only to flag part-time work."),
    ("DAAD EPOS: maximum age of the degree at the course deadline (years)", 6, "0", "E2/E3",
     "Commonly stated EPOS rule: the last degree should not be older than six years. Confirm per course."),
]
INP_ROW = {}
for i, (lab, val, fmt, ev, note) in enumerate(INPUTS, 4):
    cell(ip, i, 1, lab, BOLD)
    cell(ip, i, 2, val, INP, fmt, YEL, wrap=False)
    cell(ip, i, 3, ev)
    cell(ip, i, 4, note)
    INP_ROW[lab] = i
IN = "'Inputs & rules'!"
GRAD = f"{IN}$B$4"; SICUT = f"{IN}$B$5"; SIMIN = f"{IN}$B$6"; SIMAXEMP = f"{IN}$B$7"
CHDATE = f"{IN}$B$8"; CHMIN = f"{IN}$B$9"; CHSPAN = f"{IN}$B$10"
EPDATE = f"{IN}$B$11"; EPMIN = f"{IN}$B$12"; FTREF = f"{IN}$B$13"; EPAGE = f"{IN}$B$14"
ip.cell(row=16, column=1, value="Not tested in this workbook: SI leadership experience (required, no minimum hours; record it on 'Leadership evidence'); "
        "DAAD EPOS upper-third academic condition (needs a DTU class-rank letter).").font = SMALL

# =====================================================================================================
# 3. Employment periods (data entry)
# =====================================================================================================
ws = wb.create_sheet("Employment periods")
title(ws, "Employment periods: one row per job title and period (copy from the HR letters)",
      "Blue = you fill in. Black = formulas. Row 5 is an example and is ignored by every formula.")
cols = ["#", "Employer", "Job title (exactly as on the HR letter)", "Start date", "End date (required; current job = HR letter date)",
        "Employment type", "Hours per week (HR letter)", "Weeks worked per year (52 minus leave)", "Related to an IT/CS master's? (Y/N)",
        "Leadership in this role? (Y/N)", "Leadership: what you did, who can confirm", "Evidence document", "Evidence level",
        "Counted in the tests?", "SI: include this row? (Y/N)", "Hours before graduation", "Hours after graduation (to SI cut-off)",
        "Total hours (to SI cut-off)", "Notes / discrepancy ID"]
widths = [5, 28, 30, 12, 15, 17, 10, 11, 11, 11, 30, 18, 24, 20, 10, 11, 12, 11, 34]
hdr(ws, 4, cols, widths, height=58)
P_ROWS = [  # (#, employer, title, start, end, type, h/w, w/y, related, leader, lead_note, doc, ev, si_incl, note)
    ("EXAMPLE", "Example Bank S.C.", "IT Officer", date(2021, 3, 1), date(2023, 2, 28), "Full-time employee", 40, 48, "Y", "N", "",
     "HR letter received", EV[1], "Y", "Example only: shows the expected format"),
    (1, "SantimPay Financial Solution S.C.", "[title 1 from HR letter]", None, None, "", None, None, "", "", "", "HR letter requested", EV[0], "Y",
     "DR-06: four different titles/dates in your CVs (E5). Use the HR letter only."),
    (2, "SantimPay Financial Solution S.C.", "[title 2 from HR letter, if any]", None, None, "", None, None, "", "", "", "HR letter requested", EV[0], "Y", "DR-06"),
    (3, "SantimPay Financial Solution S.C.", "[title 3 from HR letter, if any]", None, None, "", None, None, "", "", "", "HR letter requested", EV[0], "Y", "DR-06"),
    (4, "Kifiya Financial Technology PLC", "[title from HR letter]", None, None, "", None, None, "", "", "", "HR letter requested", EV[0], "Y",
     "DR-07: three date versions in your CVs, all overlapping SantimPay (E5)."),
    (5, "Kifiya Financial Technology PLC", "[second title, if any]", None, None, "", None, None, "", "", "", "HR letter requested", EV[0], "Y", "DR-07"),
    (6, "Infonans Solutions PLC", "[title – confirm]", None, None, "", None, None, "", "", "", "Not requested", EV[0], "Y", "DR-10: optional; two date versions"),
    (7, "Gebeya Inc.", "[freelance work – confirm]", None, None, "Consultant or freelance", None, None, "", "", "", "Not requested", EV[0], "Y", "DR-10: needs contracts or invoices"),
    (8, "Dereja", "[co-trainer – confirm]", None, None, "", None, None, "", "", "", "Not requested", EV[0], "Y", "DR-10: check whether it was a programme or an internship"),
    (9, "GAO Tek Inc.", "Unpaid intern", None, None, "Internship", None, None, "", "", "", "Not requested", EV[0], "N",
     "DR-10: internship (E0 acceptance terms). Not for SI; counts for Chevening only if documented."),
    (10, "Plan International Ethiopia", "(excluded)", None, None, "", None, None, "", "", "", "None", EV[4], "N", "DR-08: duplicates the SantimPay text – EXCLUDED"),
    (11, "Yandex", "(excluded)", None, None, "", None, None, "", "", "", "None", EV[4], "N", "DR-09: no evidence – EXCLUDED"),
    (12, "[other employer]", "", None, None, "", None, None, "", "", "", "", EV[0], "Y", ""),
    (13, "[other employer]", "", None, None, "", None, None, "", "", "", "", EV[0], "Y", ""),
    (14, "[other employer]", "", None, None, "", None, None, "", "", "", "", EV[0], "Y", ""),
]
R0 = 6                       # first real data row (row 5 = example)
N = len(P_ROWS) - 1          # number of real rows
R1 = R0 + N - 1              # last real data row
for i, row in enumerate(P_ROWS):
    r = 5 + i
    n, emp, ttl, s, e, typ, hpw, wpy, rel, lead, ln, doc, ev, si_in, note = row
    vals = {1: n, 2: emp, 3: ttl, 4: s, 5: e, 6: typ, 7: hpw, 8: wpy, 9: rel, 10: lead, 11: ln, 12: doc, 13: ev, 15: si_in, 19: note}
    for c in range(1, 20):
        v = vals.get(c)
        x = cell(ws, r, c, v, INP if c in (2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 19) else BODY)
        if c in (4, 5): x.number_format = "dd-mmm-yyyy"
    # N: counted?
    ws.cell(row=r, column=14, value=(
        f'=IF(OR(B{r}="",LEFT(B{r},1)="["),"",IF(LEFT(M{r},2)="E5","NO – excluded",IF(OR(LEFT(M{r},2)="E0",LEFT(M{r},2)="E1"),'
        f'IF(AND(D{r}<>"",E{r}<>"",G{r}<>"",H{r}<>"",E{r}>=D{r}),"YES","INCOMPLETE – add dates, hours, weeks"),'
        f'IF(LEFT(M{r},2)="E4","NO – your confirmation only","NOT YET – evidence pending"))))'))
    full = f'AND(D{r}<>"",E{r}<>"",G{r}<>"",H{r}<>"")'
    # P: hours before graduation (days up to and including the graduation date, and before the SI cut-off)
    ws.cell(row=r, column=16, value=f'=IF({full},ROUND(G{r}*H{r}/52*MAX(0,MIN(E{r},{SICUT}-1,{GRAD})-D{r}+1)/7,0),"")')
    # Q: hours after graduation, up to the day before the SI cut-off
    ws.cell(row=r, column=17, value=f'=IF({full},ROUND(G{r}*H{r}/52*MAX(0,MIN(E{r},{SICUT}-1)-MAX(D{r},{GRAD}+1)+1)/7,0),"")')
    ws.cell(row=r, column=18, value=f'=IF({full},P{r}+Q{r},"")')
    ws.cell(row=r, column=20, value=f'=IF(N{r}="YES",D{r},"")').number_format = "dd-mmm-yyyy"   # helper: start if counted
    ws.cell(row=r, column=21, value=f'=IF(N{r}="YES",E{r},"")').number_format = "dd-mmm-yyyy"   # helper: end if counted
    for c in (14, 16, 17, 18):
        x = ws.cell(row=r, column=c); x.font = BODY; x.border = BD; x.alignment = WR
        if c >= 16: x.number_format = "#,##0"
    if n == "EXAMPLE":
        for c in range(1, 20): ws.cell(row=r, column=c).fill = GREY
dv(ws, TYPES, f"F5:F{R1 + 10}"); dv(ws, DOCS, f"L5:L{R1 + 10}"); dv(ws, EV, f"M5:M{R1 + 10}")
dv(ws, ["Y", "N"], f"I5:J{R1 + 10}"); dv(ws, ["Y", "N"], f"O5:O{R1 + 10}")
ws.conditional_formatting.add(f"N{R0}:N{R1}", FormulaRule(formula=[f'N{R0}="YES"'], fill=GREEN))
ws.conditional_formatting.add(f"N{R0}:N{R1}", FormulaRule(formula=[f'LEFT(N{R0},2)="NO"'], fill=RED))
ws.conditional_formatting.add(f"N{R0}:N{R1}", FormulaRule(formula=[f'OR(LEFT(N{R0},3)="NOT",LEFT(N{R0},3)="INC")'], fill=AMB))
ws.column_dimensions["T"].hidden = True; ws.column_dimensions["U"].hidden = True
ws.freeze_panes = "D5"
ws.cell(row=R1 + 2, column=2, value="Per-row hours may overlap between rows. The tests on 'Results' use 'Timeline', which counts each month once.").font = SMALL
EP = "'Employment periods'!"

# =====================================================================================================
# 4. Timeline (formulas only)
# =====================================================================================================
tl = wb.create_sheet("Timeline")
title(tl, "Monthly timeline (formulas only – do not edit)",
      "Four blocks, one column per 'Employment periods' row: SI hours | Chevening hours | EPOS months covered | EPOS full-time-equivalent months.")
months = []
y, m = 2017, 1
while (y, m) <= (2027, 12):
    months.append(date(y, m, 1)); m += 1
    if m == 13: y, m = y + 1, 1
T0 = 7                                  # first month row (header row 6, block labels row 5)
TN = T0 + len(months) - 1
BLK = {"SI": 3, "CH": 3 + N, "EC": 3 + 2 * N, "EF": 3 + 3 * N}  # first column of each block
S0 = 3 + 4 * N                          # first summary column
SUMCOLS = ["SI: counted (larger job)", "SI: jobs added", "SI: overlap?", "Chevening: counted (larger job)", "Chevening: jobs added",
           "Chevening: overlap?", "EPOS: months covered", "EPOS: full-time-equivalent"]
labels = {"SI": "SI hours (documented, no internships/voluntary, column O = Y, before the SI cut-off)",
          "CH": "Chevening hours (documented, all types, after graduation, up to the Chevening date)",
          "EC": "EPOS: share of month covered (documented, related = Y, no internships/voluntary, after graduation)",
          "EF": "EPOS: covered share × min(1, hours/week ÷ full-time reference)"}
for k, c0 in BLK.items():
    x = tl.cell(row=5, column=c0, value=labels[k]); x.font = Font(name=F, bold=True, size=9, color=NAVY)
heads = ["Month", "Month end"] + [f"{k} r{j + 1}" for k in ("SI", "CH", "EC", "EF") for j in range(N)] + SUMCOLS
hdr(tl, 6, heads, None, height=48)
tl.column_dimensions["A"].width = 10; tl.column_dimensions["B"].width = 11
for c in range(3, S0): tl.column_dimensions[L(c)].width = 6.5
for c in range(S0, S0 + len(SUMCOLS)): tl.column_dimensions[L(c)].width = 12
for k, d0 in enumerate(months):
    r = T0 + k
    tl.cell(row=r, column=1, value=d0).number_format = "mmm-yyyy"
    tl.cell(row=r, column=2, value=f"=EOMONTH(A{r},0)").number_format = "dd-mmm-yy"
    for j in range(N):
        pr = R0 + j
        ok = f'{EP}$N${pr}="YES"'
        not_int = f'{EP}$F${pr}<>"Internship",{EP}$F${pr}<>"Voluntary work"'
        hw = f"{EP}$G${pr}*{EP}$H${pr}/52"
        si_days = f"MAX(0,MIN({EP}$E${pr},{SICUT}-1,$B{r})-MAX({EP}$D${pr},$A{r})+1)"
        ch_days = f"MAX(0,MIN({EP}$E${pr},{CHDATE},$B{r})-MAX({EP}$D${pr},{GRAD}+1,$A{r})+1)"
        ep_days = f"MAX(0,MIN({EP}$E${pr},{EPDATE},$B{r})-MAX({EP}$D${pr},{GRAD}+1,$A{r})+1)"
        tl.cell(row=r, column=BLK["SI"] + j, value=f'=IF(AND({ok},{not_int},{EP}$O${pr}="Y"),ROUND({hw}*{si_days}/7,2),0)')
        tl.cell(row=r, column=BLK["CH"] + j, value=f'=IF({ok},ROUND({hw}*{ch_days}/7,2),0)')
        tl.cell(row=r, column=BLK["EC"] + j, value=f'=IF(AND({ok},{EP}$I${pr}="Y",{not_int}),ROUND({ep_days}/DAY($B{r}),4),0)')
        ec = f"{L(BLK['EC'] + j)}{r}"
        tl.cell(row=r, column=BLK["EF"] + j, value=f'=IF({ec}=0,0,ROUND({ec}*MIN(1,{EP}$G${pr}/{FTREF}),4))')
    rng = lambda key: f"{L(BLK[key])}{r}:{L(BLK[key] + N - 1)}{r}"
    tl.cell(row=r, column=S0 + 0, value=f"=MAX({rng('SI')})")
    tl.cell(row=r, column=S0 + 1, value=f"=SUM({rng('SI')})")
    tl.cell(row=r, column=S0 + 2, value=f'=IF(COUNTIF({rng("SI")},">0")>1,"OVERLAP","")')
    tl.cell(row=r, column=S0 + 3, value=f"=MAX({rng('CH')})")
    tl.cell(row=r, column=S0 + 4, value=f"=SUM({rng('CH')})")
    tl.cell(row=r, column=S0 + 5, value=f'=IF(COUNTIF({rng("CH")},">0")>1,"OVERLAP","")')
    tl.cell(row=r, column=S0 + 6, value=f"=MAX({rng('EC')})")
    tl.cell(row=r, column=S0 + 7, value=f"=MAX({rng('EF')})")
    for c in range(1, S0 + len(SUMCOLS)):
        x = tl.cell(row=r, column=c); x.font = Font(name=F, size=8); x.border = BD
        if 3 <= c < BLK["EC"] or c in (S0, S0 + 1, S0 + 3, S0 + 4): x.number_format = "0.0"
        elif BLK["EC"] <= c < S0 or c in (S0 + 6, S0 + 7): x.number_format = "0.00"
tl.freeze_panes = f"C{T0}"
TLc = lambda i: f"Timeline!${L(S0 + i)}${T0}:${L(S0 + i)}${TN}"

# =====================================================================================================
# 5. Results
# =====================================================================================================
rs = wb.create_sheet("Results", 1)
title(rs, "Results: documented experience, employer by employer, and the three tests",
      "Only rows counted on 'Employment periods' (evidence E0/E1, complete data) are used. Today: no HR letter has arrived, so every documented figure is 0.")
hdr(rs, 4, ["Employer", "Job titles entered", "Rows counted (documented)", "First start (documented)", "Last end (documented)",
            "Hours before graduation", "Hours after graduation (to SI cut-off)", "Total documented hours (to SI cut-off)",
            "Leadership recorded? (any row)", "Status"], [32, 10, 12, 13, 13, 13, 15, 15, 13, 46], height=48)
BLOCKS = [(R0, R0 + 2), (R0 + 3, R0 + 4), (R0 + 5, R0 + 5), (R0 + 6, R0 + 6), (R0 + 7, R0 + 7), (R0 + 8, R0 + 8),
          (R0 + 9, R0 + 9), (R0 + 10, R0 + 10), (R0 + 11, R0 + 11), (R0 + 12, R0 + 12), (R0 + 13, R0 + 13)]
assert BLOCKS[-1][1] == R1
for k, (a, b) in enumerate(BLOCKS):
    r = 5 + k
    rg = lambda col: f"{EP}${col}${a}:${col}${b}"
    cnt = f'COUNTIF({rg("N")},"YES")'
    cell(rs, r, 1, f"={EP}B{a}")
    cell(rs, r, 2, f'=COUNTIFS({rg("C")},"<>",{rg("C")},"<>[*")', fmt="0")
    cell(rs, r, 3, f"={cnt}", fmt="0")
    cell(rs, r, 4, f'=IF({cnt}=0,"",MIN({rg("T")}))', fmt="dd-mmm-yyyy")
    cell(rs, r, 5, f'=IF({cnt}=0,"",MAX({rg("U")}))', fmt="dd-mmm-yyyy")
    cell(rs, r, 6, f'=SUMIFS({rg("P")},{rg("N")},"YES")', fmt="#,##0")
    cell(rs, r, 7, f'=SUMIFS({rg("Q")},{rg("N")},"YES")', fmt="#,##0")
    cell(rs, r, 8, f'=SUMIFS({rg("R")},{rg("N")},"YES")', fmt="#,##0")
    cell(rs, r, 9, f'=IF(COUNTIFS({rg("J")},"Y",{rg("N")},"YES")>0,"YES",IF(COUNTIF({rg("J")},"Y")>0,"Claimed, not documented","No"))')
    cell(rs, r, 10, f'=IF(COUNTIF({rg("M")},"E5*")=COUNTA({rg("B")}),"Excluded (E5)",IF({cnt}>0,"Documented: "&{cnt}&" row(s)",'
                    f'IF(COUNTIF({rg("N")},"INCOMPLETE*")>0,"Evidence entered but data incomplete","Pending: no HR evidence yet")))')
RT = 5 + len(BLOCKS)
cell(rs, RT, 1, "All employers (rows may overlap; tests below count each month once)", BOLD)
for c in (6, 7, 8):
    cell(rs, RT, c, f"=SUM({L(c)}5:{L(c)}{RT - 1})", BOLD, "#,##0")

# ---- the three tests
t0 = RT + 3
rs.cell(row=t0 - 1, column=1, value="THE THREE TESTS (each month counted once; decisions use the 'counted' figure)").font = Font(name=F, bold=True, size=12, color=NAVY)
hdr(rs, t0, ["Test", "Requirement", "Documented result", "If overlapping jobs were added", "Other condition", "Other condition met?",
             "RESULT", "", "", "Source"], None, height=36)
rs.merge_cells(start_row=t0, start_column=7, end_row=t0, end_column=9)
# SI employers counted (distinct employers among rows used for SI)
cond = (f'({EP}$N${R0}:$N${R1}="YES")*({EP}$O${R0}:$O${R1}="Y")*({EP}$F${R0}:$F${R1}<>"Internship")*({EP}$F${R0}:$F${R1}<>"Voluntary work")')
emp_cnt = (f'SUMPRODUCT({cond}/(COUNTIFS({EP}$B${R0}:$B${R1},{EP}$B${R0}:$B${R1},{EP}$N${R0}:$N${R1},"YES",{EP}$O${R0}:$O${R1},"Y",'
           f'{EP}$F${R0}:$F${R1},"<>Internship",{EP}$F${R0}:$F${R1},"<>Voluntary work")+(1-{cond})))')
r = t0 + 1  # SI
cell(rs, r, 1, "Swedish Institute (SI Scholarship for Global Professionals)", BOLD)
cell(rs, r, 2, f'="≥ "&TEXT({SIMIN},"#,##0")&" h before "&TEXT({SICUT},"d mmm yyyy")&"; ≤ "&{SIMAXEMP}&" employers; no internships"')
cell(rs, r, 3, f"=ROUND(SUM({TLc(0)}),0)", BOLD, "#,##0")
cell(rs, r, 4, f"=ROUND(SUM({TLc(1)}),0)", fmt="#,##0")
cell(rs, r, 5, "Employers used (max 3)")
cell(rs, r, 6, f"=ROUND({emp_cnt},0)", fmt="0")
cell(rs, r, 7, f'=IF(F{r}>{SIMAXEMP},"CHECK: more than "&{SIMAXEMP}&" employers – set column O to N for the smallest",IF(C{r}>={SIMIN},"MEETS (documented)",'
                f'IF(D{r}>={SIMIN},"MEETS only if SI accepts jobs held at the same time – ask SI","NOT YET PROVEN: "&TEXT({SIMIN}-C{r},"#,##0")&" documented hours missing")))', BOLD)
cell(rs, r, 10, "E1 si.se (SGP page + FAQ), read 8 Nov 2026")
r = t0 + 2  # Chevening
cell(rs, r, 1, "Chevening (next cycle: 2028/29)", BOLD)
cell(rs, r, 2, f'="≥ "&TEXT({CHMIN},"#,##0")&" h after graduation, up to "&TEXT({CHDATE},"d mmm yyyy")')
cell(rs, r, 3, f"=ROUND(SUM({TLc(3)}),0)", BOLD, "#,##0")
cell(rs, r, 4, f"=ROUND(SUM({TLc(4)}),0)", fmt="#,##0")
cell(rs, r, 5, "Months from graduation to the deadline (≥ 24)")
cell(rs, r, 6, f'=DATEDIF({GRAD},{CHDATE},"m")', fmt="0")
cell(rs, r, 7, f'=IF(F{r}<{CHSPAN},"NOT ELIGIBLE: degree completed less than two years before the deadline",IF(C{r}>={CHMIN},"MEETS (documented)",'
                f'IF(D{r}>={CHMIN},"MEETS only if overlapping entries are added (the portal adds entries) – check","NOT YET PROVEN: "&TEXT({CHMIN}-C{r},"#,##0")&" documented hours missing")))', BOLD)
cell(rs, r, 10, "E1 chevening.org work-experience page, read 8 Nov 2026")
r = t0 + 3  # EPOS
cell(rs, r, 1, "DAAD EPOS (development-related master's)", BOLD)
cell(rs, r, 2, f'="≥ "&{EPMIN}&" months related experience after graduation, up to "&TEXT({EPDATE},"d mmm yyyy")')
cell(rs, r, 3, f"=ROUND(SUM({TLc(6)}),1)", BOLD, "0.0")
cell(rs, r, 4, f'="Full-time equivalent: "&TEXT(SUM({TLc(7)}),"0.0")')
cell(rs, r, 5, "Degree age at the course deadline (≤ 6 years)")
cell(rs, r, 6, f'=ROUND(DATEDIF({GRAD},{EPDATE},"m")/12,1)', fmt="0.0")
cell(rs, r, 7, f'=IF(F{r}>{EPAGE},"NOT ELIGIBLE: degree older than "&{EPAGE}&" years at this deadline",IF(C{r}>={EPMIN},IF(SUM({TLc(7)})>={EPMIN},"MEETS (documented)",'
                f'"CHECK: 24 months reached only with part-time work – ask the course"),"NOT YET PROVEN: "&TEXT({EPMIN}-C{r},"0.0")&" documented months missing"))', BOLD)
cell(rs, r, 10, "E2/E3 (daad.de blocked); confirm on the course page. Upper third not tested here")
for rr in range(t0 + 1, t0 + 4):
    rs.merge_cells(start_row=rr, start_column=7, end_row=rr, end_column=9)
    rs.row_dimensions[rr].height = 44
    g = f"G{rr}"
    rs.conditional_formatting.add(g, FormulaRule(formula=[f'LEFT({g},5)="MEETS"'], fill=GREEN))
    rs.conditional_formatting.add(g, FormulaRule(formula=[f'OR(LEFT({g},3)="NOT",LEFT({g},3)="CHE")'], fill=AMB))
    rs.conditional_formatting.add(g, FormulaRule(formula=[f'LEFT({g},12)="NOT ELIGIBLE"'], fill=RED))
nb = t0 + 5
notes = [
    "How to read this: 'NOT YET PROVEN' means the HR letters have not arrived. It is not a judgement that you lack the experience.",
    "The months with overlapping jobs are shown on 'Timeline' (columns marked OVERLAP). Your CVs show SantimPay and Kifiya overlapping (DR-07): only the HR letters can settle this.",
    "SI also requires leadership experience (no minimum hours): record it on 'Leadership evidence'. EPOS also requires the upper third of your class: needs a DTU class-rank letter.",
    "Chevening 2,800 h ≈ two years at Chevening's own 35-hour week. SI's formula is hours/week × weeks/year × years; this workbook applies the same weekly rate to the days in each month.",
]
for i, t in enumerate(notes):
    c = rs.cell(row=nb + i, column=1, value=t); c.font = SMALL
rs.freeze_panes = "B5"

# =====================================================================================================
# 6. Scenario (not evidence)
# =====================================================================================================
sc = wb.create_sheet("Scenario (not evidence)", 2)
title(sc, "Scenario: what your statement would give IF the HR letters confirm it (ILLUSTRATION – NOT EVIDENCE)",
      "Uses your Phase 8 statement (3 years 6 months, E4) and the assumptions below. Never type these numbers into an application.")
hdr(sc, 4, ["Assumption", "Value", "Basis"], [62, 14, 100], height=24)
SC = [
    ("Total experience you state (months)", 42, "0", "E4: your Phase 8 statement '3 years and 6 months'. Your CVs say '4 years' and 'over two years' (DR-11, E5): not used."),
    ("Hours per week", 40, "0", "Assumption: a standard full-time week. Ethiopian law allows up to 48 h; the HR letter governs."),
    ("Weeks worked per year", 48, "0", "Assumption: 52 weeks minus annual leave and public holidays (the earlier project estimate; SI's own example uses 50)."),
    ("Months of the 42 that fall AFTER 14 Jul 2022 (graduation)", 42, "0",
     "Assumption: all of them, which holds only if you started work after graduation. CVs suggest Aug 2022 (E5)."),
    ("Months of the 42 that are related to an IT/CS master's", 42, "0", "Assumption: infrastructure, DevOps and security roles are related. HR duties must show it."),
    ("Low case: hours per week", 35, "0", "Assumption: Chevening's reference week (E1)."),
    ("Low case: weeks worked per year", 44, "0", "Assumption: generous leave and holidays."),
]
for i, (lab, v, fmt, basis) in enumerate(SC, 5):
    cell(sc, i, 1, lab, BOLD); cell(sc, i, 2, v, INP, fmt, YEL, wrap=False); cell(sc, i, 3, basis)
hdr(sc, 13, ["Illustrative result", "Value", "Compared with the requirement"], None, height=24)
OUTS = [
    ("SI hours (months ÷ 12 × hours/week × weeks/year)", "=ROUND(B5/12*B6*B7,0)", "#,##0",
     f'=IF(B14>={SIMIN},"Would meet ≥ "&TEXT({SIMIN},"#,##0")&" ("&TEXT(B14/{SIMIN},"0.0")&"× the minimum)","Would NOT meet")'),
    ("SI hours, low case", "=ROUND(B5/12*B10*B11,0)", "#,##0", f'=IF(B15>={SIMIN},"Would still meet","Would NOT meet")'),
    ("Chevening hours after graduation", "=ROUND(B8/12*B6*B7,0)", "#,##0", f'=IF(B16>={CHMIN},"Would meet ≥ "&TEXT({CHMIN},"#,##0"),"Would NOT meet")'),
    ("Chevening hours after graduation, low case", "=ROUND(B8/12*B10*B11,0)", "#,##0", f'=IF(B17>={CHMIN},"Would still meet","Would NOT meet")'),
    ("EPOS related months after graduation", "=MIN(B8,B9)", "0", f'=IF(B18>={EPMIN},"Would meet ≥ "&{EPMIN}&" months","Would NOT meet")'),
]
for i, (lab, f, fmt, cmp_) in enumerate(OUTS, 14):
    cell(sc, i, 1, lab, BOLD); cell(sc, i, 2, f, BOLD, fmt); cell(sc, i, 3, cmp_)
hdr(sc, 20, ["What the HR letters must show at minimum (break-even)", "Value", "Meaning"], None, height=24)
BE = [
    ("SI: full-time months needed before 8 Feb 2027 (at the hours and weeks above)", f"=ROUND({SIMIN}/(B6*B7/12),1)", "0.0",
     "Documented full-time months needed. Fewer if your weekly hours are higher."),
    ("Chevening: full-time months needed after graduation", f"=ROUND({CHMIN}/(B6*B7/12),1)", "0.0",
     "Also needs the degree completed at least two years before the deadline (yours: July 2022)."),
    ("EPOS: related months needed after graduation", f"={EPMIN}", "0", "Calendar months of related professional work (full-time)."),
    ("SI: minimum average hours/week if all 42 months are documented", f"=ROUND({SIMIN}/(B5/12*B7),1)", "0.0", "Below this weekly average, 42 months would not reach 3,000 h."),
]
for i, (lab, f, fmt, mean) in enumerate(BE, 21):
    cell(sc, i, 1, lab, BOLD); cell(sc, i, 2, f, BOLD, fmt); cell(sc, i, 3, mean)
sc.cell(row=26, column=1, value="Conclusion: if the HR letters confirm roughly what you state, all three experience tests are met with a wide margin. "
        "The risk is not the amount of experience but proof: titles, dates, weekly hours and the SantimPay/Kifiya overlap (DR-06, DR-07).").font = Font(name=F, size=10, bold=True, color=NAVY)
sc.cell(row=26, column=1).alignment = WR; sc.merge_cells("A26:C26"); sc.row_dimensions[26].height = 32

# =====================================================================================================
# 7–9. Evidence-collection sheets (moved from Employment_Evidence_Collection_Form_v2.xlsx)
# =====================================================================================================
ra = wb.create_sheet("Responsibilities & results")
title(ra, "Responsibilities and measurable results – only claims you can prove and defend in an interview")
hdr(ra, 3, ["#", "Employer / period", "Responsibility or result (plain words)", "Number (only if real)", "Evidence (document / person)",
            "Evidence level", "Can you explain it in an interview? (Y/N)", "Use in CV?"], [6, 26, 52, 16, 30, 24, 16, 14])
seed = [("EXAMPLE", "Example Bank, 2022", "Automated server patching with Ansible", "about 30 servers", "Manager [name] / change records", EV[0], "Y")]
seed += [(i, "", "", "", "", EV[0], "") for i in range(1, 16)]
for i, rw in enumerate(seed):
    rr = 4 + i
    for c, v in enumerate(rw, 1):
        cell(ra, rr, c, v, INP if c > 1 else BODY)
    cell(ra, rr, 8, f'=IF(OR(B{rr}="",C{rr}=""),"",IF(AND(OR(LEFT(F{rr},2)="E0",LEFT(F{rr},2)="E1"),G{rr}="Y"),"YES",IF(LEFT(F{rr},2)="E5","NO","Not yet")))')
    if rw[0] == "EXAMPLE":
        for c in range(1, 9): ra.cell(row=rr, column=c).fill = GREY
dv(ra, EV, "F4:F40"); dv(ra, ["Y", "N"], "G4:G40")

le = wb.create_sheet("Leadership evidence")
title(le, "Leadership evidence (SI: lead colleagues, allocate tasks, influence strategy or decisions) – SI accepts up to 2 leadership forms")
hdr(le, 3, ["#", "Situation", "Your role / mandate", "People you led or coordinated", "Decision you made", "Result",
            "Who can confirm (name, title)", "Evidence level"], [5, 34, 26, 14, 30, 30, 26, 24])
for i in range(1, 8):
    for c in range(1, 9):
        cell(le, 3 + i, c, i if c == 1 else (EV[0] if c == 8 else None), INP if c > 1 else BODY)
dv(le, EV, "H4:H20")

ce = wb.create_sheet("Certifications")
title(ce, "Certifications – list only those with a verifiable ID or link")
hdr(ce, 3, ["Certification", "Issuer", "Issued", "Expires", "Verification ID / link", "Status", "Evidence level"], [36, 20, 12, 12, 40, 24, 24])
certs = [("Certified Kubernetes Administrator (CKA)", "Linux Foundation", "Jul 2022", "Jul 2025"), ("Azure Fundamentals (AZ-900)", "Microsoft", "May 2021", "–"),
         ("Fortinet NSE 1", "Fortinet", "Jun 2023", "Jun 2025"), ("Fortinet NSE 2", "Fortinet", "Jun 2023", "Jun 2025"), ("Fortinet NSE 3", "Fortinet", "Jun 2023", "Jun 2025")]
for i, (nm, iss, a, e) in enumerate(certs):
    rr = 4 + i
    for c, v in enumerate([nm, iss, a, e, None, None, EV[0]], 1):
        cell(ce, rr, c, v, INP if c >= 5 else BODY)
    cell(ce, rr, 6, f'=IF(E{rr}="","Needs ID – do not list yet",IF(D{rr}="–","Valid","Expired – list as expired"))')
dv(ce, EV, "G4:G20")

ORDER = ["Read me", "Employment periods", "Results", "Scenario (not evidence)", "Inputs & rules", "Timeline",
         "Responsibilities & results", "Leadership evidence", "Certifications"]
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
for s in wb.worksheets:
    s.sheet_view.zoomScale = 90
    s.sheet_view.tabSelected = (s.title == "Read me")
wb.save(OUT)
print("saved", OUT, "| period rows", N, "| months", len(months), "| timeline cols", S0 + len(SUMCOLS) - 1)
