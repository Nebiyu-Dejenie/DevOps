"""Builds Employment_Evidence_Collection_Form_v2.xlsx (v2 adds the SI no-double-count sheet). Run from this folder, then recalc with LibreOffice."""
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment

OUT = "../Employment_Evidence_Collection_Form_v2.xlsx"
F = "Arial"
HDR = PatternFill("solid", fgColor="1F3864"); HF = Font(name=F, bold=True, color="FFFFFF", size=10)
BODY = Font(name=F, size=10); INP = Font(name=F, size=10, color="0000FF"); BOLD = Font(name=F, size=10, bold=True)
YEL = PatternFill("solid", fgColor="FFF2CC"); GREY = PatternFill("solid", fgColor="EDEDED")
GREEN = PatternFill("solid", fgColor="C6EFCE"); RED = PatternFill("solid", fgColor="FFC7CE"); AMB = PatternFill("solid", fgColor="FFEB9C")
th = Side(style="thin", color="BFBFBF"); BD = Border(left=th, right=th, top=th, bottom=th)
WR = Alignment(wrap_text=True, vertical="top")
SI_CUTOFF = date(2027, 2, 8)  # SI: hours must be accumulated prior to 8 Feb 2027 (E1 si.se FAQ)
EV = ["Pending – awaiting evidence", "E0 – own document", "E1 – verified with employer/official", "E4 – my confirmation only", "E5 – unsupported/contradictory"]


def hdr(ws, row, cols, widths):
    for i, (c, w) in enumerate(zip(cols, widths), 1):
        x = ws.cell(row=row, column=i, value=c); x.font, x.fill, x.border = HF, HDR, BD
        x.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[x.column_letter].width = w
    ws.row_dimensions[row].height = 42


def dv(ws, opts, rng):
    d = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=True); ws.add_data_validation(d); d.add(rng)


wb = Workbook()
# ---------------- How to use ----------------
h = wb.active; h.title = "How to use"
h["A1"] = "Employment Evidence Collection Form"; h["A1"].font = Font(name=F, bold=True, size=14, color="1F3864")
lines = [
    "Purpose: build ONE verified employment history from official evidence (HR letters), and check the Swedish Institute 3,000-hour rule.",
    "1. Fill the BLUE cells only. Black cells are formulas. Yellow cells are key inputs.",
    "2. Enter dates exactly as written on the HR certificate (Gregorian calendar, dd/mm/yyyy).",
    "3. Leave 'End date' empty if you still work there. Hours are counted only up to 8 Feb 2027 (SI rule).",
    "4. Weekly hours and weeks per year must match the HR letter (e.g. full-time 40 h x 48 weeks).",
    "5. 'Evidence level' decides whether a row may appear in applications: E0/E1 = yes; E4 = only with care; E5 = never.",
    "6. SI counts at most THREE employers and does NOT count internships (si.se FAQ, verified 7 Nov 2026).",
    "7. Row 5 of 'Employment periods' is an EXAMPLE showing the expected format. Leave it: every formula skips row 5.",
    "8. Overlapping jobs: the sheet 'SI hours - no double count' counts each month only once (the job with the most hours that month). Use that figure. SI's FAQ does not say overlapping hours may be added.",
    "9. Only rows with evidence level E0 or E1 count in the SI check. SI needs its own signed and stamped form for each employer, so a personal confirmation (E4) is not proof.",
    "Evidence scale: E0 your document | E1 verified with the employer/official source | E2 official search result | E3 secondary | E4 your confirmation | E5 unsupported/contradictory.",
]
for i, t in enumerate(lines, 3):
    c = h.cell(row=i, column=1, value=t); c.font = BODY; c.alignment = WR
h.column_dimensions["A"].width = 120

# ---------------- Employment periods ----------------
ws = wb.create_sheet("Employment periods")
ws["A1"] = "Employment periods (one row per job title / period)"; ws["A1"].font = Font(name=F, bold=True, size=13, color="1F3864")
ws["A2"] = f"SI hours cut-off date:"; ws["A2"].font = BOLD
ws["C2"] = SI_CUTOFF; ws["C2"].number_format = "dd-mmm-yyyy"; ws["C2"].font = INP; ws["C2"].fill = YEL
ws["D2"] = "(SI: 3,000 h prior to 8 Feb 2027, max 3 employers; internships excluded)"; ws["D2"].font = Font(name=F, italic=True, size=9)
cols = ["#", "Employer", "Job title (exactly as HR states)", "Start date", "End date (blank = current)", "Type",
        "Hours / week", "Weeks / year", "Years counted (to cut-off)", "SI hours", "Evidence document", "Evidence level",
        "Use in applications?", "Notes / discrepancy ID"]
hdr(ws, 4, cols, [4, 26, 30, 13, 15, 16, 10, 10, 12, 11, 24, 24, 16, 34])
rows = [
    ("EXAMPLE – delete", "Example Bank S.C.", "IT Officer", date(2021, 3, 1), date(2023, 2, 28), "Full-time employee", 40, 48, "HR letter received", EV[2], "Example only"),
    (1, "SantimPay Financial Solution S.C.", "[title 1 from HR letter]", None, None, "", None, 48, "HR letter requested", EV[0], "DR-06: 4 versions in CVs; use HR letter only"),
    (2, "SantimPay Financial Solution S.C.", "[title 2 from HR letter]", None, None, "", None, 48, "HR letter requested", EV[0], "DR-06"),
    (3, "SantimPay Financial Solution S.C.", "[title 3 from HR letter]", None, None, "", None, 48, "HR letter requested", EV[0], "DR-06"),
    (4, "Kifiya Financial Technology PLC", "[title from HR letter]", None, None, "", None, 48, "HR letter requested", EV[0], "DR-07: 3 date versions + overlap"),
    (5, "Infonans Solutions PLC", "Network Engineer [confirm]", None, None, "", None, 48, "Not requested", EV[0], "DR-10: short role; optional"),
    (6, "Gebeya Inc.", "[freelance web development – confirm]", None, None, "Freelance", None, None, "Not requested", EV[0], "DR-10: evidence needed"),
    (7, "Dereja", "Co-Trainer [confirm]", None, None, "", None, None, "Not requested", EV[0], "DR-10: check if programme/internship"),
    (8, "GAO Tek Inc.", "Unpaid intern", None, None, "Internship", None, None, "Not requested", EV[0], "Internships do NOT count for SI"),
    (9, "Plan International Ethiopia", "(excluded)", None, None, "", None, None, "None", EV[4], "DR-08: duplicates SantimPay text – EXCLUDED"),
    (10, "Yandex", "(excluded)", None, None, "", None, None, "None", EV[4], "DR-09: no evidence – EXCLUDED"),
]
for i, r in enumerate(rows):
    rr = 5 + i
    n, emp, title, s, e, typ, hpw, wpy, evdoc, evl, note = r
    vals = [n, emp, title, s, e, typ, hpw, wpy, None, None, evdoc, evl, None, note]
    for c, v in enumerate(vals, 1):
        x = ws.cell(row=rr, column=c, value=v); x.border = BD; x.alignment = WR; x.font = INP if c in (2, 3, 4, 5, 6, 7, 8, 11, 12, 14) else BODY
        if c in (4, 5): x.number_format = "dd-mmm-yyyy"
    # years counted up to SI cut-off
    ws.cell(row=rr, column=9, value=f'=IF(D{rr}="","",MAX(0,(MIN(IF(E{rr}="",$C$2,E{rr}),$C$2)-D{rr})/365.25))').number_format = "0.00"
    ws.cell(row=rr, column=10, value=f'=IF(OR(I{rr}="",G{rr}="",H{rr}="",F{rr}="Internship"),0,ROUND(G{rr}*H{rr}*I{rr},0))').number_format = "#,##0"
    ws.cell(row=rr, column=13, value=f'=IF(LEFT(L{rr},2)="E5","NO – excluded",IF(OR(LEFT(L{rr},2)="E0",LEFT(L{rr},2)="E1"),"YES",IF(LEFT(L{rr},2)="E4","Only with care","PENDING")))')
    for c in (9, 10, 13):
        ws.cell(row=rr, column=c).font = BODY; ws.cell(row=rr, column=c).border = BD
    if n == "EXAMPLE – delete":
        for c in range(1, 15): ws.cell(row=rr, column=c).fill = GREY
LAST = 4 + len(rows)
dv(ws, ["Full-time employee", "Part-time employee", "Fixed-term contract", "Consultant", "Freelance", "Internship", "Entrepreneurship"], f"F5:F{LAST + 20}")
dv(ws, ["HR letter received", "HR letter requested", "Contract/payslip", "Not requested", "None"], f"K5:K{LAST + 20}")
dv(ws, EV, f"L5:L{LAST + 20}")
ws.conditional_formatting.add(f"M5:M{LAST}", FormulaRule(formula=['LEFT(M5,3)="YES"'], fill=GREEN))
ws.conditional_formatting.add(f"M5:M{LAST}", FormulaRule(formula=['LEFT(M5,2)="NO"'], fill=RED))
ws.conditional_formatting.add(f"M5:M{LAST}", FormulaRule(formula=['OR(LEFT(M5,4)="Only",LEFT(M5,7)="PENDING")'], fill=AMB))
ws.freeze_panes = "C5"

# SI check block
b = LAST + 3
ws.cell(row=b, column=2, value="Simple sum per employer (may double-count jobs that overlap). The figure to use is on the sheet 'SI hours - no double count'.").font = Font(name=F, bold=True, size=11, color="1F3864")
hdr_row = b + 1
for i, t in enumerate(["Employer", "SI hours (rows marked YES or 'Only with care')"], 2):
    x = ws.cell(row=hdr_row, column=i, value=t); x.font, x.fill, x.border = HF, HDR, BD
emps = ["SantimPay Financial Solution S.C.", "Kifiya Financial Technology PLC", "Infonans Solutions PLC", "Gebeya Inc.", "Dereja"]
first = hdr_row + 1
for k, e in enumerate(emps):
    r = first + k
    ws.cell(row=r, column=2, value=e).border = BD
    c = ws.cell(row=r, column=3, value=f'=SUMIFS($J$6:$J${LAST},$B$6:$B${LAST},B{r},$M$6:$M${LAST},"YES")+SUMIFS($J$6:$J${LAST},$B$6:$B${LAST},B{r},$M$6:$M${LAST},"Only with care")')
    c.number_format = "#,##0"; c.border = BD
lastE = first + len(emps) - 1
r = lastE + 1
ws.cell(row=r, column=2, value="Best 3 employers combined").font = BOLD
ws.cell(row=r, column=3, value=f"=LARGE(C{first}:C{lastE},1)+LARGE(C{first}:C{lastE},2)+LARGE(C{first}:C{lastE},3)").number_format = "#,##0"
ws.cell(row=r + 1, column=2, value="SI requirement (≥ 3,000 h) met?").font = BOLD
ws.cell(row=r + 1, column=3, value="='SI hours - no double count'!C8")
ws.conditional_formatting.add(f"C{r + 1}", FormulaRule(formula=[f'C{r + 1}="YES"'], fill=GREEN))
ws.conditional_formatting.add(f"C{r + 1}", FormulaRule(formula=[f'LEFT(C{r + 1},3)="NOT"'], fill=RED))
ws.cell(row=r + 2, column=2, value="Source: si.se FAQ (verified 7 Nov 2026). Formula = hours/week × weeks/year × years, up to 8 Feb 2027, max 3 employers, no internships.").font = Font(name=F, italic=True, size=9)

# ---------------- SI hours - no double count (v2) ----------------
si = wb.create_sheet("SI hours - no double count", 2)
si["A1"] = "Swedish Institute 3,000-hour check, counting each month only once"; si["A1"].font = Font(name=F, bold=True, size=13, color="1F3864")
si["A2"] = ("SI's FAQ calculates hours per experience (hours/week x weeks/year x years) and adds experiences together. It does not say whether jobs held at the "
            "same time may be added. This sheet takes the safe route: in any month with two jobs, only the job with more hours counts. Use the safe figure unless SI confirms in writing that overlapping hours can be added.")
si["A2"].font = BODY; si["A2"].alignment = WR; si.merge_cells("A2:N2"); si.row_dimensions[2].height = 44
EP = "'Employment periods'"
NJ = LAST - 5  # employment rows 6..LAST (row 5 is the example)
# summary block rows 4-10
labels = [("Safe total: each month counted once (E0/E1 rows you selected)", "=SUM(M{a}:M{b})"),
          ("Total if overlapping jobs could be added (ask SI before relying on it)", "=SUM(N{a}:N{b})"),
          ("Months with two or more jobs at the same time", '=COUNTIF(O{a}:O{b},"OVERLAP")'),
          ("Different employers selected (SI maximum: 3)", None),
          ("RESULT", None)]
TL0 = 30  # first timeline row (header at TL0-1)
months = []
y, m = 2015, 1
while (y, m) <= (2027, 2):
    months.append(date(y, m, 1)); m += 1
    if m == 13: y, m = y + 1, 1
a, b = TL0, TL0 + len(months) - 1
SEL0 = 15  # selection table header at row 14, rows 15..
for i, (lab, f) in enumerate(labels):
    r = 4 + i
    x = si.cell(row=r, column=1, value=lab); x.font = BOLD; x.border = BD; si.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    if f: si.cell(row=r, column=3, value=f.format(a=a, b=b))
si.cell(row=7, column=3, value=f'=SUMPRODUCT(--(E{SEL0}:E{SEL0 + NJ - 1}="Y"),1/(COUNTIFS(B{SEL0}:B{SEL0 + NJ - 1},B{SEL0}:B{SEL0 + NJ - 1},E{SEL0}:E{SEL0 + NJ - 1},"Y")+(E{SEL0}:E{SEL0 + NJ - 1}<>"Y")))')
si.cell(row=8, column=3, value='=IF(C7>3,"Too many employers selected (SI counts at most 3)",IF(C4>=3000,"MEETS 3,000 h (safe count)",IF(C5>=3000,"Meets 3,000 h only if SI allows overlapping hours: ask SI","NOT YET: need more evidenced hours")))')
for r in range(4, 9):
    c = si.cell(row=r, column=3); c.border = BD; c.font = BOLD; c.number_format = "#,##0"
si.merge_cells("C8:H8")
si.conditional_formatting.add("C8", FormulaRule(formula=['LEFT(C8,5)="MEETS"'], fill=GREEN))
si.conditional_formatting.add("C8", FormulaRule(formula=['OR(LEFT(C8,3)="NOT",LEFT(C8,3)="Too")'], fill=RED))
si.conditional_formatting.add("C8", FormulaRule(formula=['LEFT(C8,5)="Meets"'], fill=AMB))
si.cell(row=10, column=1, value="Cut-off (from 'Employment periods' C2): hours count up to the day before this date.").font = Font(name=F, italic=True, size=9)
si.cell(row=10, column=6, value=f"={EP}!C2").number_format = "dd-mmm-yyyy"
# selection table
si.cell(row=12, column=1, value="Step 1. Choose the rows to count (Y/N). Several job titles at the same employer count as ONE employer.").font = Font(name=F, bold=True, size=11, color="1F3864")
for i, (t, w) in enumerate(zip(["Row", "Employer", "Job title", "Evidence OK? (from periods sheet)", "Count for SI? (Y/N)", "Start", "End", "Hours/week", "Weeks/year", "Type"], [7, 30, 30, 18, 12, 12, 12, 10, 10, 16]), 1):
    x = si.cell(row=SEL0 - 1, column=i, value=t); x.font, x.fill, x.border = HF, HDR, BD; x.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
si.row_dimensions[SEL0 - 1].height = 32
for j in range(NJ):
    r = SEL0 + j; src = 6 + j
    vals = [j + 1, f"={EP}!B{src}", f"={EP}!C{src}", f"={EP}!M{src}", "Y" if j < 5 else "N", f'=IF({EP}!D{src}="","",{EP}!D{src})',
            f'=IF({EP}!E{src}="","",{EP}!E{src})', f'=IF({EP}!G{src}="","",{EP}!G{src})', f'=IF({EP}!H{src}="","",{EP}!H{src})', f"={EP}!F{src}"]
    for c, v in enumerate(vals, 1):
        x = si.cell(row=r, column=c, value=v); x.border = BD; x.font = INP if c == 5 else BODY
        if c in (6, 7): x.number_format = "dd-mmm-yyyy"
dv(si, ["Y", "N"], f"E{SEL0}:E{SEL0 + NJ - 1}")
# timeline
si.cell(row=TL0 - 2, column=1, value="Step 2. Monthly timeline (formulas only). Columns C-L = hours in that month for selection rows 1-10.").font = Font(name=F, bold=True, size=11, color="1F3864")
heads = ["Month", "Counted to"] + [f"Row {j + 1}" for j in range(NJ)] + ["Counted (largest)", "Added together", "Overlap?"]
for i, t in enumerate(heads, 1):
    x = si.cell(row=TL0 - 1, column=i, value=t); x.font, x.fill, x.border = HF, HDR, BD; x.alignment = Alignment(wrap_text=True, horizontal="center")
from openpyxl.utils import get_column_letter as L
for k, d0 in enumerate(months):
    r = TL0 + k
    si.cell(row=r, column=1, value=d0).number_format = "mmm-yyyy"
    si.cell(row=r, column=2, value=f"=MIN(EOMONTH(A{r},0),$F$10-1)").number_format = "dd-mmm-yyyy"
    for j in range(NJ):
        sr = SEL0 + j; col = 3 + j
        f = (f'=IF(AND($E${sr}="Y",LEFT($D${sr},3)="YES",$J${sr}<>"Internship",$F${sr}<>"",$H${sr}<>"",$I${sr}<>""),'
             f'ROUND($H${sr}*$I${sr}/12*MAX(0,MIN(IF($G${sr}="",$F$10-1,$G${sr}),$B{r})-MAX($F${sr},$A{r})+1)/DAY(EOMONTH($A{r},0)),1),0)')
        si.cell(row=r, column=col, value=f).number_format = "0.0"
    si.cell(row=r, column=3 + NJ, value=f"=MAX(C{r}:{L(2 + NJ)}{r})").number_format = "0.0"
    si.cell(row=r, column=4 + NJ, value=f"=SUM(C{r}:{L(2 + NJ)}{r})").number_format = "0.0"
    si.cell(row=r, column=5 + NJ, value=f'=IF(COUNTIF(C{r}:{L(2 + NJ)}{r},">0")>1,"OVERLAP","")')
    for c in range(1, 6 + NJ):
        si.cell(row=r, column=c).border = BD; si.cell(row=r, column=c).font = BODY
si.conditional_formatting.add(f"O{a}:O{b}", FormulaRule(formula=[f'O{a}="OVERLAP"'], fill=AMB))
for col, w in zip("ABCDEFGHIJKLMNO", [12, 30, 30, 18, 12, 12, 12, 10, 10, 16, 9, 9, 12, 12, 10]):
    si.column_dimensions[col].width = w
si.freeze_panes = f"C{TL0}"

# ---------------- Responsibilities & achievements ----------------
ra = wb.create_sheet("Responsibilities & results")
ra["A1"] = "Responsibilities and measurable results – only claims you can prove and defend in an interview"; ra["A1"].font = Font(name=F, bold=True, size=12, color="1F3864")
hdr(ra, 3, ["#", "Employer / period", "Responsibility or result (plain words)", "Number (only if real)", "Evidence (document / person)", "Evidence level", "Can you explain it in an interview? (Y/N)", "Use in CV?"],
    [4, 26, 52, 16, 30, 22, 16, 14])
ex = [("EXAMPLE", "Example Bank, 2022", "Automated server patching with Ansible", "about 30 servers", "Manager [name] / change records", EV[0], "Y")]
seed = ex + [(i, "", "", "", "", EV[0], "") for i in range(1, 16)]
for i, r in enumerate(seed):
    rr = 4 + i
    for c, v in enumerate(r, 1):
        x = ra.cell(row=rr, column=c, value=v); x.border = BD; x.alignment = WR; x.font = INP if c > 1 else BODY
    ra.cell(row=rr, column=8, value=f'=IF(OR(B{rr}="",C{rr}=""),"",IF(AND(OR(LEFT(F{rr},2)="E0",LEFT(F{rr},2)="E1"),G{rr}="Y"),"YES",IF(LEFT(F{rr},2)="E5","NO","Not yet")))').border = BD
    if r[0] == "EXAMPLE":
        for c in range(1, 9): ra.cell(row=rr, column=c).fill = GREY
dv(ra, EV, "F4:F40"); dv(ra, ["Y", "N"], "G4:G40")

# ---------------- Leadership evidence ----------------
le = wb.create_sheet("Leadership evidence")
le["A1"] = "Leadership evidence (SI: lead colleagues, allocate tasks, influence strategy or decisions)"; le["A1"].font = Font(name=F, bold=True, size=12, color="1F3864")
hdr(le, 3, ["#", "Situation", "Your role / mandate", "People you led or coordinated", "Decision you made", "Result", "Who can confirm (name, title)", "Evidence level"],
    [4, 34, 26, 14, 30, 30, 26, 22])
for i in range(1, 8):
    for c in range(1, 9):
        x = le.cell(row=3 + i, column=c, value=i if c == 1 else (EV[0] if c == 8 else None)); x.border = BD; x.alignment = WR; x.font = INP if c > 1 else BODY
dv(le, EV, "H4:H20")

# ---------------- Certifications ----------------
ce = wb.create_sheet("Certifications")
ce["A1"] = "Certifications – list only those with a verifiable ID or link"; ce["A1"].font = Font(name=F, bold=True, size=12, color="1F3864")
hdr(ce, 3, ["Certification", "Issuer", "Issued", "Expires", "Verification ID / link", "Status", "Evidence level"], [36, 20, 12, 12, 40, 22, 22])
certs = [("Certified Kubernetes Administrator (CKA)", "Linux Foundation", "Jul 2022", "Jul 2025"), ("Azure Fundamentals (AZ-900)", "Microsoft", "May 2021", "–"),
         ("Fortinet NSE 1", "Fortinet", "Jun 2023", "Jun 2025"), ("Fortinet NSE 2", "Fortinet", "Jun 2023", "Jun 2025"), ("Fortinet NSE 3", "Fortinet", "Jun 2023", "Jun 2025")]
for i, (n, iss, a, e) in enumerate(certs):
    rr = 4 + i
    for c, v in enumerate([n, iss, a, e, None, None, EV[0]], 1):
        x = ce.cell(row=rr, column=c, value=v); x.border = BD; x.alignment = WR; x.font = INP if c >= 5 else BODY
    ce.cell(row=rr, column=6, value=f'=IF(E{rr}="","Needs ID – do not list yet",IF(D{rr}="–","Valid","Expired – list as expired"))').font = BODY
dv(ce, EV, "G4:G20")

for s in wb.worksheets:
    s.sheet_view.zoomScale = 90
wb.save(OUT)
print("saved", OUT)
