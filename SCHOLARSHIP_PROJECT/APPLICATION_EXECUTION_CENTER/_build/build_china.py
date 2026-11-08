"""11_CHINA/China_Application_Operation.xlsx: one row per route/university, with the evidence level for every claim."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule

OUT = os.path.join(os.path.dirname(__file__), "..", "11_CHINA", "China_Application_Operation.xlsx")
H = ["Priority", "Route / university", "Type", "English-taught master's (CS-related)", "Curriculum verified?", "Academic rule (published)",
     "Your status vs academic rule", "Age rule (check privately)", "English rule / exemption", "Scholarship coverage", "Application fee",
     "Deadline", "Agency number", "Portal", "Pre-admission letter?", "Recommendation", "Evidence"]
W = [8, 26, 9, 34, 14, 30, 28, 18, 30, 34, 16, 26, 22, 30, 22, 40, 22]
ROWS = [
 ("1", "Dalian University of Technology (DUT)", "B (Silk Road / High-Level Postgraduate)",
  "CS school lists Computer Science and Technology (academic) and Computer Technology (professional); which are English-taught must be checked on the 2027 English-taught list",
  "No (list not captured)", "No minimum GPA published in any page found", "No published bar; competitiveness unknown",
  "Under 35", "TOEFL 80 / IELTS 5.5 or an English-medium certificate from your university (two page versions differ)",
  "Tuition, free campus housing, CNY 3,000/month; international travel self-funded; master's scholarship lasts 3 years",
  "CNY 400 (CSC applicants)", "1 Nov 2026 – 15 Feb 2027", "Type B: DUT's code (likely 10141: verify in the CSC system)",
  "http://iso.dlut.edu.cn + CSC online system", "Optional supervisor acceptance letter; DUT application number issued after submission",
  "APPLY. Ask DUT which CS master's are English-taught; send the admissions e-mail now", "E2 (sie.dlut.edu.cn pages via search)"),
 ("2", "Central South University (CSU)", "B (Silk Road / University Program)",
  "Computer Science and Technology (English-taught, in the 2026 Silk Road list); International Master Program in CS&T",
  "No", "'Strong academic background'; no GPA number found", "Unclear",
  "Under 35", "IELTS 6.0 / TOEFL 85", "CSC standard (as DUT)", "Check", "2026 pattern; 2027 page not found", "10533 (from a 2026 page; verify)",
  "CSU international admissions portal", "Check", "SECOND CHOICE, once the 2027 guide is published", "E2"),
 ("3", "Nanjing Univ. of Information Science & Technology (NUIST)", "B (Silk Road)",
  "Computer Science and Technology (English-medium)", "No", "Not found", "Unclear", "Under 35", "TOEFL 80 / IELTS 6.0 (2022 page)",
  "CSC standard", "Check", "2026 Silk Road deadline was 20 May; 2027 not found", "10300 (from older pages; verify)", "NUIST portal",
  "Check", "THIRD CHOICE (potentially more accessible)", "E2 (2022/2026 pages)"),
 ("4", "UESTC", "B", "CS, software engineering, cybersecurity (English) per earlier pages", "No", "Not found for 2027", "Unclear", "Under 35",
  "Check", "CSC standard", "Check", "2027 page not found", "10614 (older page; verify)", "UESTC portal", "Check",
  "Only if its 2027 guide shows no GPA bar you fail", "E2 (older page)"),
 ("–", "HUST", "B", "Computer Science and Technology (English) listed under the Silk Road track, which is open to only 10 named countries",
  "No", "2024 guide: minimum 3.0/4, 3.75/5 or 75%; 2027 figure not confirmed", "2.62/4.00 is BELOW 3.0/4 and 65.5% < 75% if that rule still applies",
  "Under 35 (born on/after 1 Sep 1992)", "TOEFL 80 / IELTS 6.0 / DET 120", "CSC standard", "CNY 600 (2026 guide) vs CNY 800 (scholarship page): conflict",
  "15 Oct 2026 – 25 Feb 2027 (some pages show odd dates)", "Verify", "admission.hust.edu.cn", "Check",
  "DO NOT PAY the fee unless HUST confirms in writing that 2.62/4.00 is acceptable and Ethiopia qualifies for its CS track", "E2"),
 ("–", "Northwestern Polytechnical University (NPU)", "B", "Various", "No",
  "Minimum average 75%; CSC applicants 80%", "NOT ELIGIBLE: 2.62/4.00 = 65.5% on a straight conversion, below 80%",
  "Under 30 on 1 Sep 2027 (not needed for this decision)", "IELTS 6.5 (no band < 5.5) / TOEFL 85 / DET 120 / PTE 58", "CSC standard",
  "CNY 600 (2026)", "Batch 1: 1 Nov–31 Dec 2026; batch 2: 1 Jan–28 Feb 2027", "–", "NPU online portal", "–",
  "DO NOT APPLY (academic rule not met)", "E2 (2027 NPU guide)"),
 ("A", "Type A: Ethiopian government nomination (Bilateral Program)", "A",
  "Any CSC university with English-taught master's", "–", "Set by the MoE nomination and the university", "Unknown",
  "Under 35", "University rule", "CSC full scholarship; some Type A notices include airfare", "None known",
  "Ethiopian 2027/28 call NOT yet published (other countries: 20 Dec 2026 – 14 Feb 2027)",
  "Given by the dispatching authority (MoE / Chinese Embassy Addis). NOT found; never use another country's number",
  "CSC online system after MoE nomination", "DUT: after nomination, apply to DUT for a pre-admission letter",
  "Ask the MoE now (e-mail 1 in 05_READY_TO_SEND_EMAILS); apply if the call opens", "E2/E3"),
]
wb = Workbook(); ws = wb.active; ws.title = "China routes"
HDR = PatternFill("solid", fgColor="1F3864"); th = Side(style="thin", color="BFBFBF"); BD = Border(left=th, right=th, top=th, bottom=th)
for j, (h, w) in enumerate(zip(H, W), 1):
    c = ws.cell(row=1, column=j, value=h); c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9); c.fill = HDR; c.border = BD
    c.alignment = Alignment(wrap_text=True, vertical="center"); ws.column_dimensions[c.column_letter].width = w
ws.row_dimensions[1].height = 34
for i, r in enumerate(ROWS, 2):
    for j, v in enumerate(r, 1):
        c = ws.cell(row=i, column=j, value=v); c.font = Font(name="Arial", size=9, bold=(j == 16)); c.border = BD
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[i].height = 92
n = len(ROWS) + 1
ws.conditional_formatting.add(f"P2:P{n}", FormulaRule(formula=['LEFT(P2,3)="DO "'], fill=PatternFill("solid", fgColor="FFC7CE")))
ws.conditional_formatting.add(f"P2:P{n}", FormulaRule(formula=['LEFT(P2,5)="APPLY"'], fill=PatternFill("solid", fgColor="C6EFCE")))
ws.freeze_panes = "C2"
nt = wb.create_sheet("Rules")
for i, t in enumerate([
    "Type B (Chinese University Program) = you apply to the university; the agency number is that university's code.",
    "Type A (Bilateral Program) = the Ethiopian dispatching authority nominates you; its agency number comes from the MoE/embassy notice.",
    "Never use another country's embassy agency number.",
    "Pay no application fee until the university's published rules show you meet its academic and age requirements.",
    "International airfare is NOT covered on Type B (DUT states travel is self-funded).",
    "Evidence: E2 = official university pages found by search (not opened from this machine). Re-check each page in your browser before paying."], 1):
    nt.cell(row=i, column=1, value=t).font = Font(name="Arial", size=10)
nt.column_dimensions["A"].width = 130
wb.save(OUT); print("saved", len(ROWS))
