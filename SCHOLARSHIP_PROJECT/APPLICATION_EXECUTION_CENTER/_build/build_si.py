"""08_SWEDEN_SI/SI_Programme_Shortlist.xlsx
Sheet 1: provisional top candidates (entry requirements checked, E2).
Sheet 2: every open English, full-time master's programme for Autumn 2027 in computing/security/networks/data, pulled from the
universityadmissions.se search API on 8 Nov 2026 (E1). The SI-eligible list is published on 13 Nov: tick 'On SI list?' then."""
import json, os, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

SP = os.path.join(os.path.dirname(__file__), "..", "08_SWEDEN_SI", "_ua_programmes_2026-11-08.json")
OUT = os.path.join(os.path.dirname(__file__), "..", "08_SWEDEN_SI", "SI_Programme_Shortlist.xlsx")
HDR = PatternFill("solid", fgColor="1F3864"); th = Side(style="thin", color="BFBFBF"); BD = Border(left=th, right=th, top=th, bottom=th)

def header(ws, cols, widths):
    for j, (h, w) in enumerate(zip(cols, widths), 1):
        c = ws.cell(row=1, column=j, value=h); c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9); c.fill = HDR; c.border = BD
        c.alignment = Alignment(wrap_text=True, vertical="center"); ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[1].height = 34

wb = Workbook()
top = wb.active; top.title = "Provisional top candidates"
cols = ["Rank (provisional)", "Programme", "University", "Application code", "Tuition total (SEK, E1)", "Entry requirements (E2)",
        "Your fit with the requirements", "Selection basis (E2)", "Career / alignment", "On SI list? (13 Nov)", "Decision"]
header(top, cols, [9, 34, 20, 26, 12, 46, 40, 32, 34, 12, 14])
CANDS = [
 (1, "Master Programme in Information Security", "Luleå University of Technology", "MASTERHT27-LTU-97403 / -97401", 160000,
  "Degree of at least 180 ECTS in information systems, systems science, computer science or equivalent; English 6. Work experience cannot replace the formal requirements",
  "Meets: BSc Computer Science (240 credits); English met through the Ethiopian degree", "Academic qualifications: quality and quantity",
  "Security governance and management; fits fintech security roles"),
 (2, "Master's Programme in Information Security", "Stockholm University", "MASTERHT27-SU-43004", 270000,
  "Bachelor's degree of 180 ECTS in any field; English 6. Autumn 2027 asks for new supporting documentation (see the programme page)",
  "Meets the formal requirement", "Grades, relevance of previous studies, **motivation letter** (one A4 page on an older page)",
  "Security in organisations; motivation letter lets you explain your work"),
 (3, "Master's Programme in Cybersecurity", "Karlstad University", "MASTERHT27-KAU-80917", 310000,
  "Bachelor 180 credits + 90 credits in computer science including programming, software development methodology, operating systems, computer networking I and mathematical statistics",
  "Likely meets: Programming (6 courses), Fundamentals of SE, Operating Systems, Data Communication and Networks, Probability and Statistics. Confirm the course matching",
  "Not stated on the page found", "Technical cybersecurity; closest to your infrastructure work"),
 (4, "Master's Programme in Intelligent Communication Systems", "Karlstad University", "MASTERHT27-KAU-80919", 310000,
  "Same as Karlstad Cybersecurity", "Likely meets (same course matching)", "Not stated on the page found", "Networks and communication systems; strong match with infrastructure operations"),
 (5, "Master Programme in Cyber Security", "Luleå University of Technology", "MASTERHT27-LTU-97409 / -97410", 280000,
  "At least 7.5 ECTS programming, 5 ECTS data communication/computer networks, 5 ECTS machine learning/AI (from the programme page)",
  "Meets by subject: programming 34 credits, networks 21, Introduction to AI 6", "Not stated on the page found", "Technical cybersecurity"),
 ("Avoid", "Cybersecurity, Master's Programme", "Linköping University", "MASTERHT27-LIU-91140", 332000,
  "Computing major + at least 22 ECTS mathematics including discrete mathematics, linear algebra **and calculus**",
  "Does NOT meet: no calculus course on the transcript", "Merit rating (grades)", "–"),
]
for i, r in enumerate(CANDS, 2):
    for j, v in enumerate(list(r) + ["", ""], 1):
        c = top.cell(row=i, column=j, value=v); c.font = Font(name="Arial", size=9, bold=(j == 2)); c.border = BD
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if j == 5: c.number_format = "#,##0"
    top.row_dimensions[i].height = 80
dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True); top.add_data_validation(dv); dv.add(f"J2:J{len(CANDS) + 1}")
dv2 = DataValidation(type="list", formula1='"Choose,Reserve,Drop"', allow_blank=True); top.add_data_validation(dv2); dv2.add(f"K2:K{len(CANDS) + 1}")
r = len(CANDS) + 3
for t in ["Ranking logic: (1) you meet the formal entry requirements; (2) selection uses more than grades; (3) alignment with infrastructure/security work; (4) career value.",
          "On 13 Nov: mark 'On SI list?'. Only programmes on SI's 2027/28 list can be funded. Then choose 4, ranked by real preference, on universityadmissions.se.",
          "Pay the SEK 900 fee only after the 4 programmes are fixed and their entry requirements confirmed (fee deadline 1 Feb 2027; review starts only after payment).",
          "Evidence: tuition and application codes are E1 (universityadmissions.se API, 8 Nov 2026). Entry requirements are E2 (university pages via search): confirm each on the university page."]:
    top.cell(row=r, column=1, value=t).font = Font(name="Arial", size=9, italic=True); r += 1

al = wb.create_sheet("All open programmes (E1)")
cols2 = ["University", "Programme", "Credits", "Tuition total (SEK)", "City", "Application code", "Programme page", "Relevant to you?", "On SI list? (13 Nov)"]
header(al, cols2, [30, 60, 8, 14, 14, 28, 50, 12, 12])
P = json.load(open(SP))
kw = re.compile(r"comput|software|secur|cyber|network|distributed|cloud|communication|information technology|embedded|systems|data", re.I)
rows = [p for p in P if p["status"] == "anmalan.oppen" and (p["lang"] or "").lower().startswith("english") and p["pace"] == 100 and kw.search(p["title"])]
for i, p in enumerate(sorted(rows, key=lambda x: (x["uni"], x["title"])), 2):
    vals = [p["uni"], p["title"], p["credits"], p["tuition_total_sek"], p["city"], p["code"], p["url"], "", ""]
    for j, v in enumerate(vals, 1):
        c = al.cell(row=i, column=j, value=v); c.font = Font(name="Arial", size=9); c.border = BD
        if j == 4: c.number_format = "#,##0"
al.freeze_panes = "C2"; al.auto_filter.ref = f"A1:I{len(rows) + 1}"
wb.save(OUT); print("saved shortlist; all-programmes rows:", len(rows))
