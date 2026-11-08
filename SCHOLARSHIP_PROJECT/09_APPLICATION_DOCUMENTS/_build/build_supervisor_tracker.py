"""Builds APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Research_MSc_Canada/Supervisor_Contact_Tracker.xlsx (run from 09_APPLICATION_DOCUMENTS/_build, then recalc)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
F = "Arial"; HDR = PatternFill("solid", fgColor="1F3864"); HF = Font(name=F, bold=True, color="FFFFFF", size=10)
BODY = Font(name=F, size=10); INP = Font(name=F, size=10, color="0000FF")
th = Side(style="thin", color="BFBFBF"); BD = Border(left=th, right=th, top=th, bottom=th); WR = Alignment(wrap_text=True, vertical="top")
GREEN = PatternFill("solid", fgColor="C6EFCE"); RED = PatternFill("solid", fgColor="FFC7CE"); AMB = PatternFill("solid", fgColor="FFEB9C")
wb = Workbook(); h = wb.active; h.title = "How to use"
h["A1"] = "Supervisor Contact Tracker"; h["A1"].font = Font(name=F, bold=True, size=14, color="1F3864")
for i, t in enumerate([
    "1. Blue cells are yours to fill. Black cells are formulas.",
    "2. Fill 'Paper read' before drafting: a specific, recent paper or project by the professor.",
    "3. 'Approved to send' = Y only after you have reviewed the e-mail. Nothing is sent automatically; you send from your own address.",
    "4. 'Status' tells you what to do: Draft / Ready to send / Waiting / FOLLOW UP NOW / the outcome.",
    "5. Follow up once: 7 days after sending (Manitoba, deadline 1 Dec) or 10 days (others). Then stop.",
    "6. Research descriptions are from official university pages (E1, 7-8 Nov 2026). Funding: Manitoba states a minimum of $20,000/yr for admitted MSc students (E1); do not assume any other funding."], 3):
    c = h.cell(row=i, column=1, value=t); c.font = BODY; c.alignment = WR
h.column_dimensions["A"].width = 120
ws = wb.create_sheet("Contacts")
cols = ["ID", "University", "Professor", "Research (official page, E1)", "Fit", "Official profile / directory", "Official e-mail",
        "Supervises MSc / taking students?", "Paper read (title, year)", "Approved to send (Y/N)", "Date sent", "Follow-up due", "Reply date", "Outcome", "Status"]
widths = [6, 18, 22, 46, 10, 34, 30, 16, 30, 12, 12, 13, 12, 18, 18]
for i, (c, w) in enumerate(zip(cols, widths), 1):
    x = ws.cell(row=1, column=i, value=c); x.font, x.fill, x.border = HF, HDR, BD; x.alignment = Alignment(wrap_text=True, vertical="center")
    ws.column_dimensions[x.column_letter].width = w
ws.row_dimensions[1].height = 34
UM, UC = "University of Manitoba", "University of Calgary"
rows = [
 ("M1", UM, "Sogand Sadrhaghighi", "HiPe-Net Lab: communication and security technologies for next-generation networks; data centres to IoT", "Strong", "home.cs.umanitoba.ca/~sadrhags/", "Sogand.Sadrhaghighi@umanitoba.ca"),
 ("M2", UM, "Ruppa (Tulsi) Thulasiram", "Computational Financial Derivatives Lab: finance and computing", "Moderate", "www.cs.umanitoba.ca/~tulsi/", "Tulsi.Thulasiram@umanitoba.ca"),
 ("M3", UM, "Parimala Thulasiraman", "IDEAS Lab: parallel/distributed algorithms on heterogeneous multicore architectures", "Moderate", "www.cs.umanitoba.ca/~thulasir/", "Parimala.Thulasiraman@umanitoba.ca"),
 ("C1", UC, "Mahmoud Alfadel", "Software ecosystems, releases, build systems, security vulnerabilities, mining software repositories", "Strong", "profiles.ucalgary.ca/mahmoud-alfadel", "[from profile]"),
 ("C2", UC, "Kaisong Huang", "Database systems, transaction processing, data storage management", "Strong", "profiles.ucalgary.ca/kaisong-huang", "[from profile]"),
 ("C3", UC, "Sara Rouhani", "Distributed systems and cybersecurity; decentralised architectures (also adjunct, Manitoba)", "Strong", "profiles.ucalgary.ca/sara-rouhani", "[from profile]"),
 ("C4", UC, "Majid Ghaderi", "Design, optimisation and implementation of network systems and algorithms", "Good", "profiles.ucalgary.ca/majid-ghaderi", "[from profile]"),
 ("C5", UC, "Mea Wang", "Networking systems and algorithms; QoS; overlay and P2P protocols", "Good", "profiles.ucalgary.ca/mea-wang", "[from profile]"),
 ("C6", UC, "Philip Fong", "Access control models, policy languages, enforcement", "Possible", "profiles.ucalgary.ca/philip-fong", "[from profile]"),
]
for k in range(len(rows), 20):
    rows.append((f"X{k - len(rows) + 1}" if False else "", "", "", "", "", "", ""))
for i, r in enumerate(rows):
    rr = 2 + i
    vals = list(r) + ["[check profile]" if r[0] else "", "", "", None, None, None, "", None]
    for c, v in enumerate(vals, 1):
        x = ws.cell(row=rr, column=c, value=v); x.border = BD; x.alignment = WR
        x.font = INP if c in (2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14) else BODY
        if c in (11, 12, 13): x.number_format = "dd-mmm-yyyy"
    ws.cell(row=rr, column=12, value=f'=IF(K{rr}="","",K{rr}+IF(B{rr}="{UM}",7,10))')
    ws.cell(row=rr, column=15, value=f'=IF(C{rr}="","",IF(M{rr}<>"",IF(N{rr}="","Reply received",N{rr}),IF(K{rr}="",IF(J{rr}="Y","Ready to send","Draft"),IF(TODAY()>L{rr},"FOLLOW UP NOW","Waiting"))))')
LAST = 1 + len(rows)
d1 = DataValidation(type="list", formula1='"Y,N"', allow_blank=True); ws.add_data_validation(d1); d1.add(f"J2:J{LAST}")
d2 = DataValidation(type="list", formula1='"Positive - will support,Asked for more info,Interview scheduled,Not taking students,No reply after follow-up"', allow_blank=True); ws.add_data_validation(d2); d2.add(f"N2:N{LAST}")
ws.conditional_formatting.add(f"O2:O{LAST}", FormulaRule(formula=['O2="FOLLOW UP NOW"'], fill=RED))
ws.conditional_formatting.add(f"O2:O{LAST}", FormulaRule(formula=['LEFT(O2,8)="Positive"'], fill=GREEN))
ws.conditional_formatting.add(f"O2:O{LAST}", FormulaRule(formula=['O2="Ready to send"'], fill=AMB))
ws.freeze_panes = "D2"
wb.save("../../APPLICATION_EXECUTION_CENTER/12_OTHER_SCHOLARSHIPS/Research_MSc_Canada/Supervisor_Contact_Tracker.xlsx"); print("saved")
