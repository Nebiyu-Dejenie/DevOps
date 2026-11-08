"""Builds Scholarship_Tracker_v1.xlsx (12 sheets).

Run:  python3 build_tracker.py   then recalc with LibreOffice.
All facts mirror 03_SCHOLARSHIP_RESEARCH/Phase1-2_Scholarship_Intelligence_Report_v1.md
(research date 2026-11-05). Edit data here, not by hand, to keep versions consistent.
"""
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

VERIFIED_ON = date(2026, 11, 5)
OUT = "Scholarship_Tracker_v1.xlsx"

F = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(name=F, bold=True, color="FFFFFF", size=10)
BODY = Font(name=F, size=10)
INPUT = Font(name=F, size=10, color="0000FF")
TITLE = Font(name=F, bold=True, size=14, color="1F3864")
SUB = Font(name=F, italic=True, size=9, color="555555")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")

# --------------------------------------------------------------------------
# Master data. Scores: AE, FC, PR, PE, AF, FE (0-10 each).
# deadline: date or None (TBC). status: OPEN NOW / OPENING SOON / UPCOMING /
# CLOSED / DISCONTINUED / ELIGIBILITY UNCERTAIN
# --------------------------------------------------------------------------
OPPS = [
    dict(id="S01", pri="A", country="Rwanda (US degree)", host="Carnegie Mellon University Africa (Kigali)",
         prog="MS Information Technology (MSIT)", sch="Mastercard Foundation Scholars Program at CMU-Africa",
         fund="FULLY FUNDED", cover="Tuition & fees, health insurance, travel, basic living, project/entrepreneurship support",
         gpa="No GPA cut-off; most admits >=3.0 (OFFICIAL)", elig="Eligible",
         eng="IELTS 6.5 / TOEFL 81 / DET 120 (recommended minimums); DET fee covered if applied by 15 Dec",
         work="Not required (valued)", opens=date(2026, 10, 1), deadline=date(2027, 1, 15), tz="Not stated (assume Kigali, CAT)",
         status="OPEN NOW", intake="Aug 2027", url="https://www.africa.engineering.cmu.edu/admissions/index.html",
         ver="OFFICIAL (year inferred)", next="Book DET now; draft 3x300-word essays; submit by EARLY deadline 15 Dec 2026",
         scores=(8, 10, 10, 9, 8, 7), notes="Early deadline 15 Dec 2026. Scholarship considered after admission offer."),
    dict(id="S02", pri="A", country="Sweden", host="Eligible Swedish university (choose from SI list on 13 Nov 2026)",
         prog="MSc Software Engineering / Cybersecurity / Computer Systems (TBC from list)",
         sch="SI Scholarship for Global Professionals (SISGP)", fund="FULLY FUNDED",
         cover="Tuition paid to university; monthly living allowance; SEK 15,000 travel grant. Insurance: SI FAQ says not included - confirm university cover",
         gpa="No national minimum; university merit ranking", elig="Eligible",
         eng="English 6 - met by Ethiopian bachelor's (OFFICIAL universityadmissions.se)",
         work=">=3,000 hours by Feb 2027 + documented leadership experience", opens=date(2026, 10, 15),
         deadline=date(2027, 1, 15), tz="CET (universityadmissions.se)", status="OPEN NOW", intake="Aug/Sep 2027",
         url="https://si.se/en/apply/scholarships/swedish-institute-scholarships-for-global-professionals/",
         ver="OFFICIAL (via domain search)", next="13 Nov: pick up to 4 eligible programmes; apply on universityadmissions.se by 15 Jan; SI form 8-24 Feb",
         scores=(8, 9, 9, 10, 6, 9), notes="Opening date of university round assumed mid-Oct: UNVERIFIED. SI window 8-24 Feb 2027."),
    dict(id="S03", pri="A", country="Norway / Finland / Sweden / Denmark / Estonia", host="NTNU (coordinator) + Aalto and partners",
         prog="CYBERSURE - Cybersecurity and Assurance (EMJM)", sch="Erasmus Mundus Joint Master scholarship",
         fund="FULLY FUNDED", cover="EM scholarship (EUR 1,400/month + participation costs) - confirm on programme site",
         gpa="Not published; CS/IT bachelor required", elig="Potentially eligible", eng="UNVERIFIED (check call)",
         work="Not required", opens=date(2026, 11, 16), deadline=date(2027, 1, 4), tz="Finnish time (EET) - verify",
         status="OPENING SOON", intake="Aug 2027", url="https://www.ntnu.edu/studies/mscybsure",
         ver="OFFICIAL (Aalto/NTNU; 4 vs 5 Jan conflict)", next="Read 2027 call when it opens 16 Nov; draft motivation letter",
         scores=(7, 10, 9, 7, 4, 7), notes="Aalto pages show 4 Jan and 5 Jan 2027 - use 4 Jan."),
    dict(id="S04", pri="A", country="Finland / Italy / Netherlands", host="LUT University, Univ. of L'Aquila, VU Amsterdam",
         prog="SE4GD - Software Engineers for Green Deal (EMJM)", sch="Erasmus Mundus Joint Master scholarship",
         fund="FULLY FUNDED", cover="EM scholarship if selected", gpa="UNVERIFIED", elig="Potentially eligible",
         eng="UNVERIFIED", work="Not required", opens=None, deadline=None, tz="TBC", status="UPCOMING", intake="Sep 2027",
         url="https://se4gd.eu", ver="UNVERIFIED dates", next="Check se4gd.eu weekly for 2027 call",
         scores=(7, 10, 9, 7, 5, 6), notes="Deadline expected Jan-Feb 2027 (secondary)."),
    dict(id="S05", pri="B", country="China", host="Chinese universities (Type B) / Ethiopian MoE quota (Type A)",
         prog="MSc Computer Science / Software Engineering (English-taught)", sch="Chinese Government Scholarship (CSC)",
         fund="FULLY FUNDED", cover="Tuition, accommodation, monthly stipend, insurance (amounts to verify)",
         gpa="University-specific", elig="Potentially eligible", eng="University-specific (often IELTS/TOEFL or MOI)",
         work="Not required", opens=None, deadline=None, tz="TBC", status="OPEN NOW", intake="Sep 2027",
         url="https://www.campuschina.org", ver="OFFICIAL cycle open (Indonesia notice); Ethiopia date UNVERIFIED",
         next="Confirm Ethiopia Type A route with MoE / Chinese Embassy Addis; shortlist 3 Type B universities",
         scores=(7, 9, 8, 6, 7, 6), notes="Indonesian embassy deadline 14 Feb 2027 - not applicable to Ethiopia."),
    dict(id="S06", pri="B", country="France / Belgium / Luxembourg", host="Univ. Bretagne Sud, ULB, Univ. of Luxembourg",
         prog="CYBERUS - Cybersecurity (EMJM)", sch="Erasmus Mundus Joint Master scholarship", fund="FULLY FUNDED",
         cover="EM scholarship if selected", gpa="UNVERIFIED", elig="Eligibility requires university confirmation",
         eng="UNVERIFIED", work="Not required", opens=None, deadline=None, tz="TBC", status="ELIGIBILITY UNCERTAIN",
         intake="Sep 2027", url="https://master-cyberus.eu", ver="UNVERIFIED (2027 scholarship call not confirmed)",
         next="Check master-cyberus.eu; use as 3rd EMJM only if scholarships offered", scores=(7, 10, 9, 7, 4, 3),
         notes="Max 3 EMJM applications per cycle."),
    dict(id="S07", pri="B", country="Belgium / Spain / others", host="ULB, UPC BarcelonaTech + partners",
         prog="DEAI - Data Engineering & Artificial Intelligence (EMJM)", sch="Erasmus Mundus Joint Master scholarship",
         fund="FULLY FUNDED", cover="EUR 1,400/month up to 24 months, tuition, insurance (secondary)",
         gpa="Not published", elig="Potentially eligible", eng="UNVERIFIED", work="Not required", opens=date(2026, 10, 20),
         deadline=date(2027, 1, 11), tz="12:00 GMT", status="OPEN NOW", intake="Sep 2027", url="https://deai.ulb.be/",
         ver="SECONDARY (multiple agree)", next="Alternative 3rd EMJM; maths-heavy - weigh transcript risk",
         scores=(6, 10, 8, 6, 4, 7), notes="References due 18 Jan 2027 (secondary)."),
    dict(id="S08", pri="B", country="South Korea", host="Korean universities via Embassy (Addis Ababa) or University track",
         prog="MSc Computer Science / IT", sch="Global Korea Scholarship - Graduate (GKS-G) 2027", fund="FULLY FUNDED",
         cover="Airfare, tuition, monthly allowance, settlement, Korean language year, insurance (per guideline)",
         gpa="CGPA >=2.64/4.0 OR top 20% of class (OFFICIAL 2026)", elig="Eligibility requires university confirmation",
         eng="Not mandatory; scores add points", work="Not required", opens=None, deadline=None, tz="KST",
         status="UPCOMING", intake="Sep 2027 (after language year: Sep 2028 degree)",
         url="https://www.studyinkorea.go.kr", ver="OFFICIAL (2026 rules)",
         next="DECISIVE: confirm CGPA 2.62 vs 2.67 or obtain top-20% rank letter", scores=(4, 10, 8, 6, 3, 7),
         notes="2.62 fails by 0.02; 2.67 passes."),
    dict(id="S09", pri="B", country="Hungary", host="Univ. of Szeged / ELTE / Univ. of Pannonia (TBC)",
         prog="MSc Computer Science", sch="Stipendium Hungaricum", fund="NEARLY FULLY FUNDED",
         cover="Tuition-free; HUF 43,700/month stipend; dorm OR HUF 40,000/month housing; medical insurance",
         gpa="Entrance test/interview; Szeged scores core-course grades (50/100)", elig="Potentially eligible",
         eng="B2 (e.g. IELTS ~6.0) per programme", work="Not required", opens=None, deadline=None,
         tz="CET 14:00 (2026/27 pattern)", status="OPENING SOON", intake="Sep 2027",
         url="https://stipendiumhungaricum.hu", ver="OFFICIAL (2026/27 terms); 2027/28 dates UNVERIFIED",
         next="Watch for 2027/28 call (~mid-Nov); confirm Ethiopian MoE nomination steps",
         scores=(8, 6, 8, 5, 7, 8), notes="Stipend officially a partial contribution - funding gap exists."),
    dict(id="S10", pri="B", country="Canada", host="University of Manitoba (Winnipeg)", prog="MSc Computer Science (thesis)",
         sch="Departmental minimum funding (CAD 20,000/yr x 2)", fund="PARTIALLY FUNDED",
         cover="CAD 20,000/yr minimum for admitted MSc; tuition & living gap likely", gpa="3.0/4.5 over last two years (FGPS minimum)",
         elig="Eligibility requires university confirmation", eng="IELTS Academic 6.5", work="Not required",
         opens=None, deadline=None, tz="Central (CST)", status="OPEN NOW", intake="Sep 2027",
         url="https://umanitoba.ca/graduate-studies/admissions/programs-of-study/computer-science-msc",
         ver="OFFICIAL", next="Email cs.gradprogram@umanitoba.ca for intl deadline; contact 3 potential supervisors",
         scores=(6, 6, 8, 6, 5, 6), notes="Proof of funds CAD 22,895 (IRCC, from 1 Sep 2025) + tuition."),
    dict(id="S11", pri="B", country="Italy", host="Public university (e.g. Padova, Politecnico di Torino)",
         prog="MSc Computer Science / Cybersecurity", sch="Regional DSU grant (ESU / EDISU)", fund="NEARLY FULLY FUNDED",
         cover="~EUR 8,000+/yr + fee exemption (need-based; secondary)", gpa="Admission per university; DSU need-based at entry",
         elig="Potentially eligible", eng="Programme-specific", work="Not required", opens=None, deadline=None, tz="CET",
         status="UPCOMING", intake="Sep/Oct 2027", url="https://www.esu.pd.it", ver="SECONDARY",
         next="Phase 2b: shortlist 3 Italian programmes and pre-enrolment windows", scores=(6, 7, 8, 5, 6, 6),
         notes="Requires family income documentation (ISEE parificato)."),
    dict(id="S12", pri="B", country="Turkiye", host="Turkish universities", prog="MSc Computer Engineering / CS",
         sch="Turkiye Burslari", fund="FULLY FUNDED", cover="Tuition, stipend, accommodation, insurance, airfare, language year",
         gpa=">=75% for master's (secondary); 4.0-scale conversion UNVERIFIED", elig="Eligibility requires university confirmation",
         eng="Programme-specific", work="Not required", opens=date(2027, 1, 10), deadline=date(2027, 2, 20), tz="TRT",
         status="UPCOMING", intake="Sep 2027", url="https://www.turkiyeburslari.gov.tr", ver="SECONDARY",
         next="Confirm how 2.62/4.0 is assessed against 75%", scores=(4, 9, 7, 5, 4, 7), notes=""),
    dict(id="S13", pri="B", country="United Kingdom", host="UK university of choice (x3)", prog="MSc Cloud / Cyber / Software Eng.",
         sch="Chevening Scholarship 2028/29", fund="FULLY FUNDED",
         cover="Tuition, monthly stipend, return airfare, visa, arrival/departure allowances", gpa="Unconditional UK offer required",
         elig="Eligible", eng="Per university offer", work=">=2,800 hours post-degree (met)", opens=None, deadline=None,
         tz="UTC", status="UPCOMING", intake="Sep 2028",
         url="https://www.chevening.org/scholarship/ethiopia/", ver="OFFICIAL (2027/28 closed 6 Oct 2026 11:00 UTC)",
         next="Build leadership/networking evidence now; cycle opens ~Aug 2027", scores=(7, 10, 8, 10, 6, 2),
         notes="2027/28 CLOSED. Requires 2-year return to Ethiopia."),
    dict(id="S14", pri="B", country="United States", host="US universities (placed by Fulbright)", prog="MS Computer Science / IT",
         sch="Fulbright Foreign Student Program 2028-29", fund="FULLY FUNDED",
         cover="Tuition, living stipend, airfare, health cover (per Fulbright)", gpa="UNVERIFIED for Ethiopia",
         elig="Potentially eligible", eng="TOEFL/IELTS per commission", work="Valued", opens=None, deadline=None, tz="EAT",
         status="UPCOMING", intake="Aug 2028", url="https://foreign.fulbrightonline.org",
         ver="SECONDARY (2027-28 deadline 2 May 2026)", next="Monitor US Embassy Addis Ababa from early 2027",
         scores=(7, 10, 8, 8, 6, 3), notes="J-1: two-year home residency requirement."),
    dict(id="S15", pri="B", country="EU (multi)", host="UPF Barcelona + partners", prog="EMAI - Artificial Intelligence (EMJM)",
         sch="Erasmus Mundus Joint Master scholarship", fund="FULLY FUNDED", cover="If scholarship round exists - dates UNVERIFIED",
         gpa="UNVERIFIED", elig="Eligibility requires university confirmation", eng="UNVERIFIED", work="Not required",
         opens=date(2026, 11, 15), deadline=None, tz="CEST", status="ELIGIBILITY UNCERTAIN", intake="Sep 2027",
         url="https://www.upf.edu/web/emai/access-admission", ver="OFFICIAL (self-funded windows only)",
         next="Low priority: AI/maths focus", scores=(6, 8, 8, 6, 4, 5), notes="Self-funded window to 5 Jun 2027."),
    dict(id="S16", pri="B", country="EU (multi)", host="ULB + partners", prog="BDMA - Big Data Management & Analytics (EMJM)",
         sch="Erasmus Mundus Joint Master scholarship", fund="FULLY FUNDED", cover="If selected",
         gpa="UNVERIFIED", elig="Eligibility requires university confirmation", eng="UNVERIFIED", work="Not required",
         opens=None, deadline=None, tz="CET", status="ELIGIBILITY UNCERTAIN", intake="Sep 2027", url="https://bdma.ulb.ac.be",
         ver="SECONDARY", next="Low priority", scores=(6, 8, 8, 6, 4, 5), notes=""),
    dict(id="S17", pri="C", country="Canada", host="University of British Columbia", prog="Eligible master's",
         sch="Mastercard Foundation Scholars Program at UBC", fund="FULLY FUNDED", cover="Full", gpa="UBC grad minimum",
         elig="Potentially eligible", eng="Per programme", work="Valued", opens=date(2026, 8, 31),
         deadline=date(2026, 9, 6), tz="PT", status="CLOSED", intake="Sep 2027",
         url="https://mastercardfdn.scholars.ubc.ca/graduate-admissions/", ver="OFFICIAL",
         next="Next cycle: watch Aug 2027", scores=(5, 10, 6, 7, 4, 1), notes="EOI closed 6 Sep 2026."),
    dict(id="S18", pri="C", country="Canada", host="McGill University", prog="Eligible master's",
         sch="Mastercard Foundation Scholars Program at McGill", fund="FULLY FUNDED", cover="Full", gpa="Per programme",
         elig="Potentially eligible", eng="Per programme", work="Valued", opens=None, deadline=date(2026, 8, 21),
         tz="EST", status="CLOSED", intake="Sep 2027", url="https://www.mcgill.ca", ver="SECONDARY",
         next="Next cycle", scores=(5, 10, 6, 7, 4, 1), notes=""),
    dict(id="S19", pri="C", country="Canada", host="University of Toronto", prog="Eligible master's",
         sch="Mastercard Foundation Scholars Program at U of T", fund="FULLY FUNDED", cover="Full",
         gpa="Per programme", elig="Ineligible", eng="Per programme", work="Valued", opens=None, deadline=None,
         tz="ET", status="ELIGIBILITY UNCERTAIN", intake="2027-28 / 2028-29",
         url="https://internationalexperience.utoronto.ca/global-experiences/global-scholarships/mastercard-foundation-scholars",
         ver="OFFICIAL", next="None unless partner list changes", scores=(1, 10, 6, 6, 3, 2),
         notes="2026-27: Africa Health Collaborative partner-institution graduates only."),
    dict(id="S20", pri="C", country="United Kingdom", host="University of Edinburgh", prog="Listed sustainability MSc (no CS)",
         sch="Mastercard Foundation Scholars Programme", fund="FULLY FUNDED", cover="Full", gpa="Per programme",
         elig="Ineligible", eng="Per programme", work="Valued", opens=None, deadline=date(2026, 11, 19), tz="UK",
         status="CLOSED", intake="Sep 2027",
         url="https://edinburgh-global.ed.ac.uk/mastercard-foundation-scholars-program/apply-for-a-scholarship/application-timelines",
         ver="OFFICIAL", next="None", scores=(6, 10, 2, 6, 5, 0),
         notes="Mandatory info session 15-24 Sep 2026 missed; programme mismatch."),
    dict(id="S21", pri="C", country="United Kingdom", host="University of Manchester", prog="Eligible MSc (development-focused list)",
         sch="Equity and Merit Scholarship", fund="FULLY FUNDED", cover="Tuition + living (on-campus)",
         gpa="Per offer", elig="Potentially eligible", eng="Per offer", work="Valued", opens=date(2026, 8, 1),
         deadline=date(2026, 8, 31), tz="UK", status="CLOSED", intake="Sep 2027",
         url="https://www.manchester.ac.uk/study/masters/fees-and-funding/masters-student-funding/equity-merit-scholarships/",
         ver="OFFICIAL/SECONDARY (date conflict: one snippet 6 Dec 2026)", next="Verify date conflict; else next cycle",
         scores=(6, 9, 4, 7, 5, 1), notes="Ethiopia is an eligible country."),
    dict(id="S22", pri="C", country="Germany", host="EPOS-listed German universities", prog="Development-related master's (no CS confirmed)",
         sch="DAAD EPOS", fund="FULLY FUNDED", cover="Monthly stipend, insurance, travel, (tuition where applicable)",
         gpa="Far above average (upper third)", elig="Ineligible", eng="Per course", work=">=2 years (met)",
         opens=None, deadline=None, tz="CET", status="ELIGIBILITY UNCERTAIN", intake="Oct 2027",
         url="https://www2.daad.de/deutschland/stipendium/datenbank/en/21148-scholarship-database/?detail=50076777",
         ver="OFFICIAL", next="Only if registrar rank letter shows upper third", scores=(2, 10, 3, 9, 2, 3),
         notes="Degree must normally be <=6 years old (2022: OK)."),
    dict(id="S23", pri="C", country="Germany", host="Public universities (e.g. TH Koln)", prog="MSc Communication Systems & Networks / CS",
         sch="None (tuition-free; self-funded)", fund="SELF-FUNDED", cover="Tuition-free at most public universities",
         gpa="Programme-specific (uni-assist evaluation)", elig="Eligibility requires university confirmation",
         eng="Programme-specific", work="Not required", opens=None, deadline=None, tz="CET", status="UPCOMING",
         intake="Oct 2027", url="https://www.th-koeln.de/en/academics/communication-systems-and-networks-master--how-to-apply_6834.php",
         ver="OFFICIAL", next="Backup only", scores=(5, 2, 9, 6, 5, 6), notes="Blocked account EUR 11,904/yr."),
    dict(id="S24", pri="C", country="United States", host="Stanford", prog="Any graduate programme", sch="Knight-Hennessy Scholars",
         fund="FULLY FUNDED", cover="Full", gpa="Highly selective", elig="Ineligible", eng="Per programme", work="Valued",
         opens=None, deadline=None, tz="PT", status="CLOSED", intake="Sep 2027", url="https://knight-hennessy.stanford.edu",
         ver="UNVERIFIED this session", next="None", scores=(2, 10, 7, 7, 2, 0),
         notes="Not competitive at 2.62; 2027 cycle presumed closed (not re-verified)."),
    dict(id="S25", pri="C", country="UK (Commonwealth)", host="UK universities", prog="Any", sch="Commonwealth Shared / Master's Scholarships",
         fund="FULLY FUNDED", cover="Full", gpa="n/a", elig="Ineligible", eng="n/a", work="n/a", opens=None, deadline=None,
         tz="n/a", status="CLOSED", intake="n/a", url="https://cscuk.fcdo.gov.uk", ver="Fact: Ethiopia is not a Commonwealth member",
         next="None", scores=(0, 10, 7, 7, 0, 0), notes="Nationality restriction."),
]

PRIORITIES = ["A", "B", "C"]
STATUSES = ["OPEN NOW", "OPENING SOON", "UPCOMING", "CLOSED", "DISCONTINUED", "ELIGIBILITY UNCERTAIN"]
ELIG = ["Eligible", "Potentially eligible", "Ineligible", "Eligibility requires university confirmation"]
FUNDS = ["FULLY FUNDED", "NEARLY FULLY FUNDED", "PARTIALLY FUNDED", "TUITION ONLY", "SELF-FUNDED"]
PROG = ["Not started", "In progress", "Done", "N/A"]


def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font, cell.fill, cell.border = HDR_FONT, HDR_FILL, BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 32


def body(ws, r1, r2, c1, c2):
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            cell.border = BORDER
            cell.alignment = WRAP
            if cell.font is None or cell.font.color is None or cell.font.color.rgb in (None, "FF000000"):
                cell.font = BODY


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = TITLE
    ws["A2"] = sub
    ws["A2"].font = SUB


def add_dv(ws, options, rng):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


def priority_cf(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"A"'], fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"B"'], fill=PatternFill("solid", fgColor="FFEB9C")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"C"'], fill=PatternFill("solid", fgColor="FFC7CE")))


wb = Workbook()

# ---------------- 2. Master Scholarship Database -------------------------
db = wb.active
db.title = "Master Scholarship Database"
title(db, "Master Scholarship Database", f"Verified {VERIFIED_ON:%d %b %Y}. Never add fictional entries. Blue = editable status fields.")
cols = ["ID", "Priority", "Country", "University / Host", "Programme", "Scholarship", "Funding class", "Coverage summary",
        "GPA / academic rule", "Eligibility status", "English requirement", "Work-experience requirement", "Opening date",
        "Deadline", "Deadline timezone", "Deadline status", "Intake", "Official URL", "Verification status",
        "Verification date", "Admission status", "Scholarship status", "Next action", "Evidence & notes"]
header(db, 4, cols, [6, 8, 16, 26, 28, 26, 16, 38, 26, 20, 28, 22, 12, 12, 16, 16, 12, 40, 24, 12, 14, 14, 40, 40])
R0 = 5
for i, o in enumerate(OPPS):
    r = R0 + i
    vals = [o["id"], o["pri"], o["country"], o["host"], o["prog"], o["sch"], o["fund"], o["cover"], o["gpa"], o["elig"],
            o["eng"], o["work"], o["opens"] or "TBC", o["deadline"] or "TBC", o["tz"], o["status"], o["intake"], o["url"],
            o["ver"], VERIFIED_ON, "Not started", "Not started", o["next"], o["notes"]]
    for c, v in enumerate(vals, 1):
        cell = db.cell(row=r, column=c, value=v)
        if isinstance(v, date):
            cell.number_format = "dd-mmm-yyyy"
    db.cell(row=r, column=18).hyperlink = o["url"]
    for c in (2, 10, 16, 21, 22):
        db.cell(row=r, column=c).font = INPUT
RL = R0 + len(OPPS) - 1
body(db, R0, RL, 1, len(cols))
for c in (2, 10, 16, 21, 22):
    for r in range(R0, RL + 1):
        db.cell(row=r, column=c).font = INPUT
add_dv(db, PRIORITIES, f"B{R0}:B{RL + 50}")
add_dv(db, ELIG, f"J{R0}:J{RL + 50}")
add_dv(db, STATUSES, f"P{R0}:P{RL + 50}")
add_dv(db, FUNDS, f"G{R0}:G{RL + 50}")
add_dv(db, ["Not started", "Submitted", "Under review", "Interview", "Offer", "Conditional offer", "Rejected", "Withdrawn"], f"U{R0}:V{RL + 50}")
priority_cf(db, f"B{R0}:B{RL}")
db.conditional_formatting.add(f"P{R0}:P{RL}", CellIsRule(operator="equal", formula=['"CLOSED"'], font=Font(name=F, color="9C0006", bold=True)))
db.conditional_formatting.add(f"P{R0}:P{RL}", CellIsRule(operator="equal", formula=['"OPEN NOW"'], font=Font(name=F, color="006100", bold=True)))
db.freeze_panes = "C5"
db.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{RL}"

# ---------------- 4. Eligibility Assessment -------------------------------
el = wb.create_sheet("Eligibility Assessment")
title(el, "Eligibility Assessment & Programme Matching Score",
      "Strategic score (0-100), NOT an admission probability. Edit blue scores (0-10) or the weights in row 4.")
el["A4"] = "Weights ->"
el["A4"].font = Font(name=F, bold=True)
weights = [0.25, 0.25, 0.15, 0.15, 0.10, 0.10]
for j, w in enumerate(weights):
    c = el.cell(row=4, column=5 + j, value=w)
    c.number_format = "0%"
    c.font, c.fill = INPUT, YELLOW
el["K4"] = "=SUM(E4:J4)"
el["K4"].number_format = "0%"
el["L4"] = "<- must equal 100%"
el["L4"].font = SUB
ecols = ["ID", "Opportunity", "Priority", "Eligibility status", "Academic eligibility (25%)", "Funding coverage (25%)",
         "Programme relevance (15%)", "Professional exp. alignment (15%)", "Admission flexibility (10%)",
         "Application feasibility (10%)", "Matching score /100", "Rank", "Rationale"]
header(el, 5, ecols, [6, 42, 9, 22, 13, 13, 13, 13, 13, 13, 12, 8, 60])
E0 = 6
for i, o in enumerate(OPPS):
    r = E0 + i
    dbr = R0 + i
    el.cell(row=r, column=1, value=f"='Master Scholarship Database'!A{dbr}").font = Font(name=F, color="008000")
    el.cell(row=r, column=2, value=f"='Master Scholarship Database'!F{dbr}&\" - \"&'Master Scholarship Database'!E{dbr}").font = Font(name=F, color="008000")
    el.cell(row=r, column=3, value=f"='Master Scholarship Database'!B{dbr}").font = Font(name=F, color="008000")
    el.cell(row=r, column=4, value=f"='Master Scholarship Database'!J{dbr}").font = Font(name=F, color="008000")
    for j, s in enumerate(o["scores"]):
        el.cell(row=r, column=5 + j, value=s).font = INPUT
    el.cell(row=r, column=11, value=f"=ROUND(SUMPRODUCT(E{r}:J{r},$E$4:$J$4)*10,1)")
    el.cell(row=r, column=12, value=f"=RANK(K{r},$K${E0}:$K${E0 + len(OPPS) - 1})+COUNTIF($K${E0 - 1}:K{r - 1},K{r})")
    el.cell(row=r, column=13, value=o["notes"] or o["next"])
EL = E0 + len(OPPS) - 1
body(el, E0, EL, 1, len(ecols))
for i in range(len(OPPS)):
    for j in range(6):
        el.cell(row=E0 + i, column=5 + j).font = INPUT
for r in range(E0, EL + 1):
    for c in (1, 2, 3, 4):
        el.cell(row=r, column=c).font = Font(name=F, size=10, color="008000")
dv = DataValidation(type="whole", operator="between", formula1="0", formula2="10")
el.add_data_validation(dv)
dv.add(f"E{E0}:J{EL}")
el.conditional_formatting.add(f"K{E0}:K{EL}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color="5B9BD5"))
priority_cf(el, f"C{E0}:C{EL}")
el.freeze_panes = "C6"
el.auto_filter.ref = f"A5:M{EL}"

# ---------------- 5. Application Deadlines --------------------------------
dl = wb.create_sheet("Application Deadlines")
title(dl, "Application Deadline Calendar", "Days remaining recalculates daily (TODAY()). Re-verify each date on the official page before submitting.")
MIL = [
    ("ALL", "Resolve CV discrepancies; confirm CGPA; request registrar & employer letters", date(2026, 11, 12), "Internal", "Internal target"),
    ("S02", "SI publishes 2027/28 eligible programme list", date(2026, 11, 13), "CET", "OFFICIAL"),
    ("S03", "CYBERSURE application opens", date(2026, 11, 16), "EET", "OFFICIAL"),
    ("S01", "Target: Duolingo English Test taken", date(2026, 12, 1), "Internal", "Internal target"),
    ("S01", "CMU-Africa EARLY deadline (DET fee covered)", date(2026, 12, 15), "CAT (verify)", "OFFICIAL (year inferred)"),
    ("S03", "CYBERSURE deadline (earlier of 4/5 Jan)", date(2027, 1, 4), "EET (verify)", "OFFICIAL (conflict)"),
    ("S07", "DEAI scholarship deadline", date(2027, 1, 11), "12:00 GMT", "SECONDARY"),
    ("S02", "universityadmissions.se master's deadline", date(2027, 1, 15), "CET", "OFFICIAL/SECONDARY"),
    ("S01", "CMU-Africa FINAL deadline", date(2027, 1, 15), "CAT (verify)", "OFFICIAL (year inferred)"),
    ("S07", "DEAI reference letters due", date(2027, 1, 18), "GMT", "SECONDARY"),
    ("S12", "Turkiye Burslari window opens", date(2027, 1, 10), "TRT", "SECONDARY"),
    ("S02", "SI scholarship application opens", date(2027, 2, 8), "CET", "OFFICIAL"),
    ("S12", "Turkiye Burslari window closes", date(2027, 2, 20), "TRT", "SECONDARY"),
    ("S02", "SI scholarship application closes", date(2027, 2, 24), "CET", "OFFICIAL"),
    ("S02", "SI results announced", date(2027, 4, 22), "CET", "OFFICIAL"),
    ("S13", "Chevening 2027/28 deadline (CLOSED - reference)", date(2026, 10, 6), "11:00 UTC", "OFFICIAL"),
    ("S17", "Mastercard Foundation UBC EOI (CLOSED - reference)", date(2026, 9, 6), "PT", "OFFICIAL"),
]
MIL.sort(key=lambda x: x[2])
header(dl, 4, ["ID", "Milestone", "Date", "Timezone", "Days remaining", "Urgency", "Source status", "Done?"],
       [7, 60, 13, 16, 13, 14, 24, 10])
for i, (oid, m, d, tz, src) in enumerate(MIL):
    r = 5 + i
    dl.cell(row=r, column=1, value=oid)
    dl.cell(row=r, column=2, value=m)
    dl.cell(row=r, column=3, value=d).number_format = "dd-mmm-yyyy"
    dl.cell(row=r, column=4, value=tz)
    dl.cell(row=r, column=5, value=f"=C{r}-TODAY()")
    dl.cell(row=r, column=6, value=f'=IF(H{r}="Yes","DONE",IF(E{r}<0,"PASSED",IF(E{r}<=14,"URGENT",IF(E{r}<=45,"SOON","PLANNED"))))')
    dl.cell(row=r, column=7, value=src)
    dl.cell(row=r, column=8, value="No").font = INPUT
DLL = 4 + len(MIL)
body(dl, 5, DLL, 1, 8)
for r in range(5, DLL + 1):
    dl.cell(row=r, column=8).font = INPUT
add_dv(dl, ["Yes", "No"], f"H5:H{DLL + 30}")
dl.conditional_formatting.add(f"F5:F{DLL}", CellIsRule(operator="equal", formula=['"URGENT"'], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=F, bold=True, color="9C0006")))
dl.conditional_formatting.add(f"F5:F{DLL}", CellIsRule(operator="equal", formula=['"SOON"'], fill=PatternFill("solid", fgColor="FFEB9C")))
dl.conditional_formatting.add(f"F5:F{DLL}", CellIsRule(operator="equal", formula=['"PASSED"'], font=Font(name=F, color="808080", italic=True)))
dl.conditional_formatting.add(f"F5:F{DLL}", CellIsRule(operator="equal", formula=['"DONE"'], fill=PatternFill("solid", fgColor="C6EFCE")))
dl.freeze_panes = "A5"
dl.auto_filter.ref = f"A4:H{DLL}"

# ---------------- 1. Executive Dashboard (inserted first) -----------------
dash = wb.create_sheet("Executive Dashboard", 0)
title(dash, "Executive Dashboard - Fully Funded Master's 2027",
      f"Applicant: Nebyu Dejenie Tilahun | BSc CS, Debre Tabor Univ. (2022) | CGPA 2.62 on transcript (2.67 reported - TO CONFIRM) | Verified {VERIFIED_ON:%d %b %Y}")
dash.column_dimensions["A"].width = 38
for col in "BCDEFG":
    dash.column_dimensions[col].width = 16
DBR = f"'Master Scholarship Database'!$B${R0}:$B${RL}"
DBS = f"'Master Scholarship Database'!$P${R0}:$P${RL}"
DBF = f"'Master Scholarship Database'!$G${R0}:$G${RL}"
dash["A4"] = "Opportunities by priority"
dash["A4"].font = Font(name=F, bold=True, size=11)
header_row = 5
for j, h in enumerate(["Priority", "Count", "Meaning"]):
    c = dash.cell(row=header_row, column=1 + j, value=h)
    c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
meanings = {"A": "High potential - apply in 2027 cycle", "B": "Possible - needs clarification / backup / 2028",
            "C": "Low potential or ineligible - reason recorded"}
for i, p in enumerate(PRIORITIES):
    r = 6 + i
    dash.cell(row=r, column=1, value=p)
    dash.cell(row=r, column=2, value=f'=COUNTIF({DBR},A{r})')
    dash.cell(row=r, column=3, value=meanings[p])
dash.cell(row=9, column=1, value="Total tracked")
dash.cell(row=9, column=2, value="=SUM(B6:B8)")
body(dash, 6, 9, 1, 3)
priority_cf(dash, "A6:A8")
dash.merge_cells("C6:G6"); dash.merge_cells("C7:G7"); dash.merge_cells("C8:G8")

dash["A11"] = "Opportunities by deadline status"
dash["A11"].font = Font(name=F, bold=True, size=11)
for j, h in enumerate(["Status", "Count"]):
    c = dash.cell(row=12, column=1 + j, value=h)
    c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
for i, s in enumerate(STATUSES):
    r = 13 + i
    dash.cell(row=r, column=1, value=s)
    dash.cell(row=r, column=2, value=f'=COUNTIF({DBS},A{r})')
body(dash, 13, 18, 1, 2)

dash["D11"] = "By funding class"
dash["D11"].font = Font(name=F, bold=True, size=11)
for j, h in enumerate(["Funding class", "Count"]):
    c = dash.cell(row=12, column=4 + j, value=h)
    c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
for i, s in enumerate(FUNDS):
    r = 13 + i
    dash.cell(row=r, column=4, value=s)
    dash.cell(row=r, column=5, value=f'=COUNTIF({DBF},D{r})')
body(dash, 13, 17, 4, 5)
dash.column_dimensions["D"].width = 24

dash["A21"] = "Next 8 deadlines / milestones (from 'Application Deadlines')"
dash["A21"].font = Font(name=F, bold=True, size=11)
for j, h in enumerate(["Milestone", "Date", "Days left", "Urgency"]):
    c = dash.cell(row=22, column=1 + j, value=h)
    c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
future = [i for i, m in enumerate(MIL) if m[2] >= VERIFIED_ON][:8]
for k, idx in enumerate(future):
    r, src = 23 + k, 5 + idx
    dash.cell(row=r, column=1, value=f"='Application Deadlines'!B{src}")
    dash.cell(row=r, column=2, value=f"='Application Deadlines'!C{src}").number_format = "dd-mmm-yyyy"
    dash.cell(row=r, column=3, value=f"='Application Deadlines'!E{src}")
    dash.cell(row=r, column=4, value=f"='Application Deadlines'!F{src}")
body(dash, 23, 22 + len(future), 1, 4)
for r in range(23, 23 + len(future)):
    for c in range(1, 5):
        dash.cell(row=r, column=c).font = Font(name=F, size=10, color="008000")
dash.conditional_formatting.add(f"D23:D{22 + len(future)}", CellIsRule(operator="equal", formula=['"URGENT"'], fill=PatternFill("solid", fgColor="FFC7CE")))

dash["A33"] = "Top-ranked opportunities (matching score)"
dash["A33"].font = Font(name=F, bold=True, size=11)
for j, h in enumerate(["Opportunity", "Score", "Priority"]):
    c = dash.cell(row=34, column=1 + j, value=h)
    c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
for k in range(1, 9):
    r = 34 + k
    rng_rank = f"'Eligibility Assessment'!$L${E0}:$L${EL}"
    dash.cell(row=r, column=1, value=f"=IFERROR(INDEX('Eligibility Assessment'!$B${E0}:$B${EL},MATCH({k},{rng_rank},0)),\"\")")
    dash.cell(row=r, column=2, value=f"=IFERROR(INDEX('Eligibility Assessment'!$K${E0}:$K${EL},MATCH({k},{rng_rank},0)),\"\")")
    dash.cell(row=r, column=3, value=f"=IFERROR(INDEX('Eligibility Assessment'!$C${E0}:$C${EL},MATCH({k},{rng_rank},0)),\"\")")
body(dash, 35, 42, 1, 3)
priority_cf(dash, "C35:C42")

dash["A45"] = "Overall application progress"
dash["A45"].font = Font(name=F, bold=True, size=11)
dash["A46"] = "Average completion across tracked Priority A/B applications"
dash["B46"] = "=IFERROR(AVERAGE('Application Progress'!L5:L14),0)"
dash["B46"].number_format = "0%"
body(dash, 46, 46, 1, 2)
dash.conditional_formatting.add("B46", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"))
dash["A48"] = "Critical risks"
dash["A48"].font = Font(name=F, bold=True, size=11)
risks = ["1. CV versions conflict (titles, dates, duplicate bullets across employers) - fix before ANY submission.",
         "2. CGPA 2.62 (transcript) vs 2.67 (reported) - decides GKS eligibility (cut-off 2.64).",
         "3. No English test score yet - DET needed by early Dec for CMU-Africa.",
         "4. Five major deadlines between 15 Dec 2026 and 24 Feb 2027.",
         "5. Some official sites not directly reachable in research session - re-verify before submitting."]
for i, t in enumerate(risks):
    dash.cell(row=49 + i, column=1, value=t).font = BODY
    dash.merge_cells(start_row=49 + i, start_column=1, end_row=49 + i, end_column=7)
dash["A55"] = "Legend: blue text = editable input; green text = linked from another sheet; black = formula/fixed fact."
dash["A55"].font = SUB

# ---------------- 3. University Comparison --------------------------------
uc = wb.create_sheet("University Comparison", 2)
title(uc, "University / Programme Comparison (priority programmes)", "UNVERIFIED = not yet confirmed from an official source. Completed progressively in Phase 4.")
ucols = ["University", "Country", "City", "Programme", "Degree type", "Programme website", "Admission website",
         "Scholarship website", "Tuition fee", "Scholarship coverage", "Minimum GPA", "Ethiopian qualification acceptance",
         "English requirement", "Work-experience requirement", "Application opening", "Application deadline", "Intake",
         "Application fee", "Fee waiver", "Required documents", "Funding gap", "Eligibility assessment",
         "Estimated competitiveness", "Next action"]
header(uc, 4, ucols, [26, 12, 12, 30, 10, 34, 34, 34, 22, 30, 22, 26, 26, 20, 14, 16, 10, 12, 18, 36, 22, 20, 16, 36])
UROWS = [
    ["Carnegie Mellon University Africa", "Rwanda", "Kigali", "MS Information Technology", "MSc",
     "https://www.africa.engineering.cmu.edu/academics/programs/msit.html", "https://www.africa.engineering.cmu.edu/admissions/index.html",
     "https://www.africa.engineering.cmu.edu/impact/mastercard-foundation-scholars.html", "USD 16,000/yr (African students, 2026-27)",
     "Full (Mastercard Foundation)", "None (most >=3.0)", "Accepted (African applicants)", "IELTS 6.5 / TOEFL 81 / DET 120",
     "Not required", "01-Oct-2026", "15-Dec-2026 early / 15-Jan-2027 final", "Aug 2027", "None", "DET fee covered if early",
     "Transcripts (>=3/4 of degree), CV, 3 essays <=300 words, 3 recommenders, English score, competency tests",
     "None if scholarship awarded", "Eligible", "Moderate", "Book DET; draft essays"],
    ["Swedish university (from SI list)", "Sweden", "TBC", "MSc Software Eng. / Cybersecurity / Computer Systems", "MSc",
     "https://www.universityadmissions.se", "https://www.universityadmissions.se", "https://si.se/en/apply/scholarships/swedish-institute-scholarships-for-global-professionals/",
     "Programme-specific (UNVERIFIED)", "Tuition + monthly allowance + SEK 15,000 travel", "No national minimum",
     "Accepted if ETA-accredited programme", "English 6 - met by Ethiopian bachelor's", ">=3,000 h + leadership (SI)",
     "~mid-Oct 2026 (UNVERIFIED)", "15-Jan-2027 (univ) / 24-Feb-2027 (SI)", "Aug 2027", "SEK 900 (UNVERIFIED)", "UNVERIFIED",
     "Diploma, transcripts, passport, CV, motivation; SI: CV + motivation + reference letters (UNVERIFIED detail)",
     "Insurance to confirm", "Eligible", "Moderate", "Pick programmes on 13 Nov"],
    ["NTNU (coordinator) / Aalto", "Norway / Finland", "Trondheim / Espoo", "CYBERSURE Cybersecurity and Assurance", "MSc (double degree)",
     "https://www.ntnu.edu/studies/mscybsure", "https://www.aalto.fi/en/study-options/cybersecurity-and-assurance-cybersure-erasmus-mundus-master-of-science-technology",
     "https://erasmus-plus.ec.europa.eu/opportunities/individuals/students/erasmus-mundus-joint-masters", "Covered if EM scholar",
     "EM scholarship", "Not published", "UNVERIFIED (check call)", "UNVERIFIED", "Not required", "16-Nov-2026", "04-Jan-2027",
     "Aug 2027", "UNVERIFIED", "UNVERIFIED", "Per 2027 call", "None if EM scholar", "Potentially eligible", "Low-Moderate", "Read call 16 Nov"],
    ["LUT / Univ. L'Aquila / VU Amsterdam", "Finland / Italy / Netherlands", "Multiple", "SE4GD Software Engineers for Green Deal",
     "MSc (triple degree)", "https://se4gd.eu", "https://se4gd.eu", "https://se4gd.eu", "Covered if EM scholar", "EM scholarship",
     "UNVERIFIED", "UNVERIFIED", "UNVERIFIED", "Not required", "UNVERIFIED", "UNVERIFIED (Jan-Feb 2027 expected)", "Sep 2027",
     "UNVERIFIED", "UNVERIFIED", "Per 2027 call", "None if EM scholar", "Potentially eligible", "Low-Moderate", "Monitor site"],
    ["University of Szeged", "Hungary", "Szeged", "MSc Computer Science", "MSc",
     "https://apply.stipendiumhungaricum.hu/courses/course/1530-msc-computer-science", "https://apply.stipendiumhungaricum.hu",
     "https://stipendiumhungaricum.hu", "Covered by SH", "Tuition + HUF 43,700/mo + housing + insurance",
     "Score >=50/100 (grades 50 pts, interview 20, diploma 20, extras 10)", "Via Ethiopian MoE nomination", "B2",
     "Not required", "~Nov 2026 (UNVERIFIED)", "~15-Jan-2027 (UNVERIFIED)", "Sep 2027", "UNVERIFIED", "UNVERIFIED",
     "Diploma, transcript, CV, motivation, medical cert., English proof, MoE nomination", "Living-cost gap (stipend partial)",
     "Potentially eligible", "Moderate", "Await call"],
    ["University of Manitoba", "Canada", "Winnipeg", "MSc Computer Science (thesis)", "MSc",
     "https://umanitoba.ca/graduate-studies/admissions/programs-of-study/computer-science-msc",
     "https://umanitoba.ca/graduate-studies/admissions/requirements/international",
     "https://umanitoba.ca/science/student-experience/financial-aid-awards/computer-science", "UNVERIFIED",
     "Min CAD 20,000/yr x 2 (dept)", "3.0/4.5 last two years", "Assessed on last 2 years", "IELTS 6.5", "Not required",
     "Up to 16 months before term", "Confirm by email (15 Jan / 1 Dec shown)", "Sep 2027", "UNVERIFIED", "UNVERIFIED",
     "Transcripts, CV, statement, references, IELTS, supervisor contact", "See Financial Coverage sheet",
     "Requires university confirmation", "Low-Moderate", "Email department + supervisors"],
]
for i, row in enumerate(UROWS):
    for c, v in enumerate(row, 1):
        cell = uc.cell(row=5 + i, column=c, value=v)
        if isinstance(v, str) and v.startswith("https://"):
            cell.hyperlink = v
body(uc, 5, 4 + len(UROWS), 1, len(ucols))
uc.freeze_panes = "B5"
uc.auto_filter.ref = f"A4:{get_column_letter(len(ucols))}{4 + len(UROWS)}"

# ---------------- 6. Required Documents -----------------------------------
rd = wb.create_sheet("Required Documents")
title(rd, "Required Documents Register", "Status dropdown in blue. Do NOT store passport numbers or ID numbers in this file.")
header(rd, 4, ["#", "Document", "Needed for", "Status", "Location / source", "Verification needed", "Notes"], [5, 46, 26, 18, 30, 30, 46])
DOCS = [
    ("Temporary graduation certificate + student academic record (transcript)", "All", "Have (needs official copy)", "Google Drive", "Official/sealed or authenticated copy", "Generated 15 Jul 2022; 'official only when signed and sealed'"),
    ("Permanent printed BSc diploma", "All", "Not started", "Debre Tabor University registrar", "Check if 2.67 appears", "Certificate says diploma printing pending"),
    ("Class rank / percentile letter", "S08 GKS, S22 EPOS", "Not started", "Registrar", "Only if genuinely top 20% / upper third", ""),
    ("Last-two-years GPA statement", "S10 Manitoba", "Not started", "Registrar", "Unofficial estimate ~3.08", ""),
    ("MoE / MoFA authentication of degree & transcript", "S05, S09, visas (programme-specific)", "Not started", "Ethiopian MoE, MoFA", "Confirm per programme", ""),
    ("Employer experience letter - SantimPay", "S02, S13, all", "Not started", "SantimPay HR", "Title + dates must match CV", "Resolve DevOps Engineer vs Technical POS Manager"),
    ("Employer experience letter - Kifiya", "S02, S13, all", "Not started", "Kifiya HR", "Explain overlap with SantimPay", "Dates differ across CVs"),
    ("Experience letter - Infonans (short role)", "Optional", "Not started", "Infonans", "", "Jul-Aug 2022"),
    ("Clarification: Plan International Ethiopia / Yandex roles", "All", "Needs verification", "Applicant", "Bullets duplicate SantimPay text", "Do not list unless documented"),
    ("Master CV (international, consistent)", "All", "Not started", "09/02 folders", "After truth table", ""),
    ("English test - Duolingo English Test", "S01 (by 15 Dec)", "Not started", "englishtest.duolingo.com", "", "Target >=120"),
    ("English test - IELTS Academic", "S10, S03?, S09", "Not started", "IELTS centre Addis Ababa", "", "Target 6.5-7.0, no band <6.0"),
    ("Passport bio page (redacted copy for screening)", "All", "Have (not reviewed)", "Google Drive 'Passport' folder", "Name spelling & expiry", "Do not upload number until needed"),
    ("Recommendation - academic 1 (final-project supervisor)", "All", "Not started", "DTU", "", "Final Project I/II: A-"),
    ("Recommendation - academic 2", "All", "Not started", "DTU", "", "Prefer lecturer of A-/B grade course"),
    ("Recommendation - professional (current manager)", "All", "Not started", "SantimPay", "", ""),
    ("Recommendation - professional 2 (Kifiya supervisor)", "S02, S13", "Not started", "Kifiya", "", ""),
    ("Existing letter - [former lecturer, name withheld] (Nov 2024)", "None (replace)", "Needs verification", "Google Drive", "Overstates grades; template error", "Request a new, specific letter"),
    ("Certification evidence (CKA, AZ-900, Fortinet NSE)", "All", "Needs verification", "Credly / vendor portals", "IDs; mark expired as expired", "CKA expired Jul 2025"),
    ("Portfolio (GitHub, website)", "S01, S02, EMJM", "Not started", "GitHub / nebiyu.dev.et", "Remove employer-confidential data", ""),
    ("Statement of Purpose / Motivation letters", "All", "Not started", "09_STATEMENTS_AND_LETTERS", "", "One per programme"),
    ("Medical certificate", "S09 Hungary, S05 CSC", "Not started", "Approved clinic", "", "Programme form"),
]
for i, d in enumerate(DOCS):
    r = 5 + i
    rd.cell(row=r, column=1, value=i + 1)
    for c, v in enumerate(d, 2):
        rd.cell(row=r, column=c, value=v)
    rd.cell(row=r, column=4).font = INPUT
body(rd, 5, 4 + len(DOCS), 1, 7)
for r in range(5, 5 + len(DOCS)):
    rd.cell(row=r, column=4).font = INPUT
add_dv(rd, ["Have", "Have (needs official copy)", "Have (not reviewed)", "Requested", "Not started", "Needs verification", "N/A"], f"D5:D{60}")
rd.conditional_formatting.add(f"D5:D{4 + len(DOCS)}", FormulaRule(formula=[f'LEFT(D5,4)="Have"'], fill=PatternFill("solid", fgColor="C6EFCE")))
rd.conditional_formatting.add(f"D5:D{4 + len(DOCS)}", CellIsRule(operator="equal", formula=['"Needs verification"'], fill=PatternFill("solid", fgColor="FFEB9C")))
rd.freeze_panes = "C5"
rd.auto_filter.ref = f"A4:G{4 + len(DOCS)}"

# ---------------- 7. Recommendation Tracking ------------------------------
rt = wb.create_sheet("Recommendation Tracking")
title(rt, "Recommendation Tracking", "Placeholders in [brackets] - replace with real referees. Never draft or sign on a referee's behalf without their review and approval.")
header(rt, 4, ["Referee (name, title)", "Type", "Relationship / course or role", "Email on file?", "For opportunities", "Request date",
               "Briefing pack sent", "Portal link sent", "Status", "Deadline", "Notes"], [34, 12, 34, 12, 22, 12, 14, 14, 14, 12, 40])
REFS = [
    ["[Final-project supervisor, DTU]", "Academic", "Supervised Final Project I & II (A-)", "No", "S01, S02, S03, S04", None, "No", "No", "Not requested", date(2026, 12, 10), "Ask this week"],
    ["[Lecturer, DTU - Compiler Design / Advanced Programming]", "Academic", "Course lecturer (A- grade)", "No", "S01, S03, S04", None, "No", "No", "Not requested", date(2026, 12, 10), ""],
    ["[Line manager, SantimPay]", "Professional", "Direct supervisor - DevOps/IT", "No", "All", None, "No", "No", "Not requested", date(2026, 12, 10), "Also request experience letter"],
    ["[Supervisor, Kifiya]", "Professional", "IT Operations supervisor", "No", "S02, S13", None, "No", "No", "Not requested", date(2027, 1, 10), ""],
]
for i, row in enumerate(REFS):
    for c, v in enumerate(row, 1):
        cell = rt.cell(row=5 + i, column=c, value=v)
        if isinstance(v, date):
            cell.number_format = "dd-mmm-yyyy"
body(rt, 5, 8, 1, 11)
for r in range(5, 9):
    for c in (1, 4, 6, 7, 8, 9):
        rt.cell(row=r, column=c).font = INPUT
add_dv(rt, ["Not requested", "Requested", "Agreed", "Submitted", "Declined"], "I5:I40")
add_dv(rt, ["Yes", "No"], "D5:D40")
add_dv(rt, ["Yes", "No"], "G5:H40")
rt.conditional_formatting.add("I5:I40", CellIsRule(operator="equal", formula=['"Submitted"'], fill=PatternFill("solid", fgColor="C6EFCE")))

# ---------------- 8. Application Progress --------------------------------
ap = wb.create_sheet("Application Progress")
title(ap, "Application Progress", "Set each step to Done / In progress / Not started / N/A. Completion % = Done / (steps not N/A).")
steps = ["Portal account", "Documents uploaded", "SOP / motivation", "References", "English test", "Fee / waiver", "Submitted"]
header(ap, 4, ["ID", "Opportunity", "Priority"] + steps + ["Completion %", "Admission status", "Scholarship status", "Interview status", "Visa status"],
       [6, 46, 9] + [13] * 7 + [12, 14, 14, 14, 12])
track = [o for o in OPPS if o["pri"] in ("A", "B")][:10]
for i, o in enumerate(track):
    r = 5 + i
    ap.cell(row=r, column=1, value=o["id"])
    ap.cell(row=r, column=2, value=f'{o["sch"]} - {o["prog"]}')
    ap.cell(row=r, column=3, value=o["pri"])
    for j in range(7):
        ap.cell(row=r, column=4 + j, value="Not started").font = INPUT
    ap.cell(row=r, column=11, value=f'=IFERROR(COUNTIF(D{r}:J{r},"Done")/(7-COUNTIF(D{r}:J{r},"N/A")),0)').number_format = "0%"
    for c in range(12, 16):
        ap.cell(row=r, column=c, value="Not started").font = INPUT
APL = 4 + len(track)
body(ap, 5, APL, 1, 15)
for r in range(5, APL + 1):
    for c in list(range(4, 11)) + list(range(12, 16)):
        ap.cell(row=r, column=c).font = INPUT
add_dv(ap, PROG, f"D5:J{APL + 20}")
add_dv(ap, ["Not started", "Submitted", "Under review", "Interview", "Offer", "Conditional offer", "Rejected", "Withdrawn"], f"L5:N{APL + 20}")
add_dv(ap, ["Not started", "Preparing", "Submitted", "Biometrics done", "Approved", "Refused"], f"O5:O{APL + 20}")
ap.conditional_formatting.add(f"K5:K{APL}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"))
ap.conditional_formatting.add(f"D5:J{APL}", CellIsRule(operator="equal", formula=['"Done"'], fill=PatternFill("solid", fgColor="C6EFCE")))
ap.conditional_formatting.add(f"D5:J{APL}", CellIsRule(operator="equal", formula=['"In progress"'], fill=PatternFill("solid", fgColor="FFEB9C")))
priority_cf(ap, f"C5:C{APL}")
# The dashboard average references L5:L14 originally; point it at completion column K
dash["B46"] = f"=IFERROR(AVERAGE('Application Progress'!K5:K{APL}),0)"
ap.freeze_panes = "D5"

# ---------------- 9. Financial Coverage -----------------------------------
fc = wb.create_sheet("Financial Coverage")
title(fc, "Financial Coverage Matrix & Funding-Gap Calculator", "Yes / Partial / No / UNVERIFIED. Never assume airfare, insurance or visa costs are covered without evidence.")
items = ["Tuition", "Monthly living", "Accommodation", "Return airfare", "Health insurance", "Visa costs", "Settlement", "Books", "Research"]
header(fc, 4, ["ID", "Opportunity", "Funding class"] + items + ["Covered items (Yes)", "Notes"], [6, 40, 18] + [12] * 9 + [12, 40])
COV = {
    "S01": ["Yes", "Yes", "UNVERIFIED", "Yes", "Yes", "UNVERIFIED", "UNVERIFIED", "UNVERIFIED", "Yes"],
    "S02": ["Yes", "Yes", "Partial", "Partial", "No", "UNVERIFIED", "No", "No", "No"],
    "S03": ["Yes", "Yes", "Partial", "Partial", "Yes", "Partial", "UNVERIFIED", "No", "No"],
    "S04": ["Yes", "Yes", "Partial", "Partial", "Yes", "Partial", "UNVERIFIED", "No", "No"],
    "S05": ["Yes", "Yes", "Yes", "UNVERIFIED", "Yes", "UNVERIFIED", "UNVERIFIED", "No", "No"],
    "S08": ["Yes", "Yes", "Partial", "Yes", "Yes", "UNVERIFIED", "Yes", "UNVERIFIED", "UNVERIFIED"],
    "S09": ["Yes", "Partial", "Yes", "No", "Yes", "No", "No", "No", "No"],
    "S10": ["Partial", "Partial", "No", "No", "No", "No", "No", "No", "Partial"],
    "S13": ["Yes", "Yes", "Partial", "Yes", "UNVERIFIED", "Yes", "Yes", "UNVERIFIED", "UNVERIFIED"],
    "S23": ["Yes", "No", "No", "No", "No", "No", "No", "No", "No"],
}
r = 5
for o in OPPS:
    if o["id"] not in COV:
        continue
    fc.cell(row=r, column=1, value=o["id"])
    fc.cell(row=r, column=2, value=f'{o["sch"]} - {o["prog"]}')
    fc.cell(row=r, column=3, value=o["fund"])
    for j, v in enumerate(COV[o["id"]]):
        fc.cell(row=r, column=4 + j, value=v).font = INPUT
    fc.cell(row=r, column=13, value=f'=COUNTIF(D{r}:L{r},"Yes")')
    fc.cell(row=r, column=14, value="EM scholarship: 'Partial' = paid via participation-cost contribution; confirm in call" if o["id"] in ("S03", "S04") else o["notes"])
    r += 1
FCL = r - 1
body(fc, 5, FCL, 1, 14)
for rr in range(5, FCL + 1):
    for c in range(4, 13):
        fc.cell(row=rr, column=c).font = INPUT
add_dv(fc, ["Yes", "Partial", "No", "UNVERIFIED"], f"D5:L{FCL + 20}")
fc.conditional_formatting.add(f"D5:L{FCL}", CellIsRule(operator="equal", formula=['"Yes"'], fill=PatternFill("solid", fgColor="C6EFCE")))
fc.conditional_formatting.add(f"D5:L{FCL}", CellIsRule(operator="equal", formula=['"No"'], fill=PatternFill("solid", fgColor="FFC7CE")))
fc.conditional_formatting.add(f"D5:L{FCL}", CellIsRule(operator="equal", formula=['"UNVERIFIED"'], fill=PatternFill("solid", fgColor="FFEB9C")))

g = FCL + 3
fc.cell(row=g, column=1, value="Funding-gap calculator (official figures only; yellow = inputs to confirm)").font = Font(name=F, bold=True, size=11)
header(fc, g + 1, ["Case", "Item", "Amount", "Currency", "Source"], None)
GAP = [
    ("Germany self-funded (S23)", "Blocked account requirement per year", 11904, "EUR", "DAAD - since Jan 2025 (2026 update UNVERIFIED)"),
    ("Germany self-funded (S23)", "Scholarship income per year", 0, "EUR", "No scholarship"),
    ("Germany self-funded (S23)", "Annual gap (living only; excludes semester fees)", "=C{a}-C{b}", "EUR", "Formula"),
    ("Canada - Manitoba (S10)", "IRCC proof-of-funds (living, 1 applicant)", 22895, "CAD", "IRCC - from 1 Sep 2025; check 1 Sep 2026 figure"),
    ("Canada - Manitoba (S10)", "International tuition per year (INPUT)", None, "CAD", "UNVERIFIED - enter from UManitoba fee page"),
    ("Canada - Manitoba (S10)", "Department minimum funding per year", 20000, "CAD", "UManitoba CS funding page"),
    ("Canada - Manitoba (S10)", "Year-1 gap vs IRCC test (living + tuition - funding)", "=C{d}+C{e}-C{f}", "CAD", "Formula; tuition blank = understated"),
    ("Hungary (S09)", "Monthly stipend", 43700, "HUF", "SH Call 2026/27"),
    ("Hungary (S09)", "Housing contribution if no dorm", 40000, "HUF", "SH Call 2026/27"),
    ("Hungary (S09)", "Total monthly support (no dorm)", "=C{h}+C{i}", "HUF", "Formula; compare with real local costs"),
]
start = g + 2
refs = {}
keys = "abcdefghij"
for k, row in enumerate(GAP):
    refs[keys[k]] = start + k
for k, (case, item, amt, cur, src) in enumerate(GAP):
    rr = start + k
    fc.cell(row=rr, column=1, value=case)
    fc.cell(row=rr, column=2, value=item)
    if isinstance(amt, str):
        amt = amt.format(**refs)
        fc.cell(row=rr, column=3, value=amt)
    else:
        cc = fc.cell(row=rr, column=3, value=amt)
        cc.font, cc.fill = INPUT, YELLOW
    fc.cell(row=rr, column=3).number_format = "#,##0;(#,##0);-"
    fc.cell(row=rr, column=4, value=cur)
    fc.cell(row=rr, column=5, value=src)
body(fc, start, start + len(GAP) - 1, 1, 5)
for k, (case, item, amt, cur, src) in enumerate(GAP):
    if not isinstance(amt, str):
        c = fc.cell(row=start + k, column=3)
        c.font, c.fill = INPUT, YELLOW

# ---------------- 10. Visa Preparation ------------------------------------
vp = wb.create_sheet("Visa Preparation")
title(vp, "Visa Preparation (summary - full country guides in 13_VISA_PREPARATION, Phase 10)",
      "Visa rules are separate from admission rules. Never use false documents or borrowed funds presented as your own.")
header(vp, 4, ["Country", "Visa / permit", "Official link", "Financial requirement (verified)", "Key documents", "Ethiopia-specific notes", "Status"],
       [16, 26, 46, 38, 46, 46, 14])
VISA = [
    ["Germany", "National (D) student visa", "https://www.germany.info", "Blocked account EUR 11,904/yr (EUR 992/month) since Jan 2025 - DAAD; or scholarship letter",
     "Admission letter, proof of funds, health insurance, passport, certificates", "Apply via German Embassy Addis Ababa; appointment wait times UNVERIFIED", "Not started"],
    ["United States", "F-1 (or J-1 for Fulbright)", "https://travel.state.gov", "I-20 shows funding; SEVIS fee; DS-160",
     "I-20/DS-2019, SEVIS receipt, DS-160, financial evidence, interview", "Ethiopia not on Proclamation 10998 lists (eff. 1 Jan 2026) per White House text - re-check before applying", "Not started"],
    ["Canada", "Study permit", "https://www.canada.ca/en/immigration-refugees-citizenship/services/study-canada/graduate-student.html",
     "CAD 22,895 living (from 1 Sep 2025) + tuition; check 2026 figure", "Letter of acceptance, proof of funds, biometrics, medical exam if required",
     "Master's at public DLIs exempt from PAL/TAL from 1 Jan 2026 (except Quebec rules)", "Not started"],
    ["United Kingdom", "Student visa", "https://www.gov.uk/student-visa", "GBP 1,136/month outside London, 1,483 London (financial-evidence guidance); 'Money you need' page shows 1,171 - CONFLICT",
     "CAS, financial evidence (28 days) or official sponsor letter, TB test certificate, passport", "Chevening covers visa fee & IHS (Chevening terms)", "Not started"],
    ["Sweden", "Residence permit for studies", "https://www.migrationsverket.se", "UNVERIFIED (check Migrationsverket monthly maintenance figure)",
     "Admission letter, SI scholarship letter, insurance, passport", "Apply online; biometrics at Swedish Embassy Addis Ababa (UNVERIFIED)", "Not started"],
    ["Rwanda", "Student visa / permit", "https://www.migration.gov.rw", "UNVERIFIED", "CMU-Africa admission letter", "Ethiopians - entry rules UNVERIFIED", "Not started"],
    ["Hungary", "Residence permit for study", "https://enterhungary.gov.hu", "SH award letter", "Award letter, accommodation, insurance", "Hungarian visa services for Ethiopia UNVERIFIED", "Not started"],
]
for i, row in enumerate(VISA):
    for c, v in enumerate(row, 1):
        cell = vp.cell(row=5 + i, column=c, value=v)
        if c == 3:
            cell.hyperlink = v
    vp.cell(row=5 + i, column=7).font = INPUT
body(vp, 5, 4 + len(VISA), 1, 7)
for r in range(5, 5 + len(VISA)):
    vp.cell(row=r, column=7).font = INPUT
add_dv(vp, ["Not started", "Preparing", "Submitted", "Biometrics done", "Approved", "Refused", "N/A"], "G5:G30")

# ---------------- 11. Interview Preparation -------------------------------
ip = wb.create_sheet("Interview Preparation")
title(ip, "Interview Preparation & English Test Plan", "Model answers are drafted ONLY from verified profile facts (Phase 8). No invented experiences.")
header(ip, 4, ["Question bank", "Target #", "Drafted", "Practised (mock)", "% ready", "Notes"], [40, 10, 10, 14, 10, 50])
QB = [("Common scholarship questions", 50), ("Technical & academic", 30), ("Leadership", 20), ("Professional experience", 20),
      ("GPA questions", 5), ("Motivation", 5), ("Returning value to Ethiopia", 5), ("Future objectives", 5)]
for i, (q, n) in enumerate(QB):
    r = 5 + i
    ip.cell(row=r, column=1, value=q)
    ip.cell(row=r, column=2, value=n)
    ip.cell(row=r, column=3, value=0).font = INPUT
    ip.cell(row=r, column=4, value=0).font = INPUT
    ip.cell(row=r, column=5, value=f"=IFERROR(D{r}/B{r},0)").number_format = "0%"
r = 5 + len(QB)
ip.cell(row=r, column=1, value="TOTAL").font = Font(name=F, bold=True)
ip.cell(row=r, column=2, value=f"=SUM(B5:B{r - 1})")
ip.cell(row=r, column=3, value=f"=SUM(C5:C{r - 1})")
ip.cell(row=r, column=4, value=f"=SUM(D5:D{r - 1})")
ip.cell(row=r, column=5, value=f"=IFERROR(D{r}/B{r},0)").number_format = "0%"
body(ip, 5, r, 1, 6)
for rr in range(5, r):
    ip.cell(row=rr, column=3).font = INPUT
    ip.cell(row=rr, column=4).font = INPUT
ip.conditional_formatting.add(f"E5:E{r}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5"))

e0 = r + 3
ip.cell(row=e0, column=1, value="English test plan - 4 weeks to DET (target >=120) then IELTS Academic (target 7.0, no band <6.5)").font = Font(name=F, bold=True, size=11)
header(ip, e0 + 1, ["Week", "Daily routine (~90 min)", "Weekly target", "Mock test", "Score", "Notes"], None)
PLAN = [
    ("Week 1 (from 9 Nov)", "20 min listening (podcasts/lectures, note-taking); 20 min reading (tech articles, timed); 20 min writing (1 task, self-check); 30 min speaking (record 3 x 2-min answers on work topics)",
     "Learn DET question types; baseline", "Official DET practice test", None, "Speaking is the priority (self-rated B1)"),
    ("Week 2", "Same + vocabulary log (20 words/day, academic & tech)", "Fluency: 2-min answers without long pauses", "DET practice", None, ""),
    ("Week 3", "Timed full sections; writing samples reviewed", "Writing: clear structure, 150+ words in time", "DET practice x2", None, ""),
    ("Week 4 (DET by 1 Dec)", "Full mocks; fix weakest subscore", "Score >=120", "Real DET", None, "Retake allowed if time permits before 15 Dec"),
]
for i, row in enumerate(PLAN):
    for c, v in enumerate(row, 1):
        cell = ip.cell(row=e0 + 2 + i, column=c, value=v)
    ip.cell(row=e0 + 2 + i, column=5).font = INPUT
body(ip, e0 + 2, e0 + 1 + len(PLAN), 1, 6)
ip.column_dimensions["B"].width = 60
ip.column_dimensions["C"].width = 30

# ---------------- 12. Research Sources ------------------------------------
rs = wb.create_sheet("Research Sources")
title(rs, "Research Sources & Evidence Log", "Source hierarchy: government > university > scholarship body > immigration > foundations > secondary.")
header(rs, 4, ["Claim", "Source title", "URL", "Date accessed", "Publication / update date", "Verification status", "Conditions / notes"],
       [48, 34, 60, 13, 16, 20, 40])
SRC = [
    ("Chevening 2027/28 Ethiopia deadline 6 Oct 2026 11:00 UTC", "Chevening - Ethiopia", "https://www.chevening.org/scholarship/ethiopia/", "OFFICIAL (domain search)", "2,800 h; 2-year return"),
    ("SISGP eligible countries incl. Ethiopia; 3,000 h; leadership", "SI Scholarship for Global Professionals", "https://si.se/en/apply/scholarships/swedish-institute-scholarships-for-global-professionals/", "OFFICIAL (domain search)", "34 countries"),
    ("SI insurance not included", "SI FAQ master's scholarships", "https://si.se/en/apply/faq-scholarships-on-masters-level/", "OFFICIAL (domain search)", "Confirm university cover"),
    ("Ethiopian bachelor's meets English 6; ETA accreditation", "UniversityAdmissions.se - Ethiopia", "https://www.universityadmissions.se/en/apply-to-masters/provide-application-documents-masters/country-instructions/ethiopia/", "OFFICIAL", ""),
    ("CMU-Africa no GPA cut-off; English minimums; no fee", "CMU-Africa FAQ", "https://www.africa.engineering.cmu.edu/admissions/faq.html", "OFFICIAL", "French FAQ says TOEFL 93"),
    ("CMU-Africa dates 1 Oct / 15 Dec / 15 Jan", "CMU-Africa Admissions", "https://www.africa.engineering.cmu.edu/admissions/index.html", "OFFICIAL", "Year inferred"),
    ("Mastercard Foundation at CMU-Africa coverage", "CMU-Africa MCF Scholars", "https://www.africa.engineering.cmu.edu/impact/mastercard-foundation-scholars.html", "OFFICIAL/SECONDARY", ""),
    ("CYBERSURE window 16 Nov 2026 - 5 Jan 2027", "Aalto - CYBERSURE", "https://www.aalto.fi/en/study-options/cybersecurity-and-assurance-cybersure-erasmus-mundus-master-of-science-technology", "OFFICIAL", "Other Aalto page: 4 Jan"),
    ("CYBERSURE entry: CS/IT bachelor; 25 places", "NTNU - CYBERSURE", "https://www.ntnu.edu/studies/mscybsure", "OFFICIAL", ""),
    ("DEAI scholarship deadline 11 Jan 2027", "Gizpulse / ErasmusScholarship (DEAI)", "https://gizpulse.com/opportunities/erasmus-mundus-deai-scholarship-2027", "SECONDARY", "Official: deai.ulb.be"),
    ("EMJM application timing; catalogue", "Erasmus+ EMJM students", "https://erasmus-plus.ec.europa.eu/opportunities/individuals/students/erasmus-mundus-joint-masters", "OFFICIAL", ""),
    ("GKS-G CGPA 2.64/4.0 or top 20%", "Korean Embassy notice 2026 GKS-G", "https://www.mofa.go.kr/sg-en/brd/m_2435/view.do?seq=761574", "OFFICIAL (2026)", "Ethiopia notice: et.mofa.go.kr"),
    ("Stipendium Hungaricum: Ethiopia partner; HUF 43,700; housing; insurance", "SH Call 2026/27", "https://stipendiumhungaricum.hu/wp-content/uploads/2025/10/BA_MA_OTM_Call_for_Applications_2026_27.pdf", "OFFICIAL", "Deadline 15 Jan 2026 (previous cycle)"),
    ("Szeged MSc CS scoring", "SH course page", "https://apply.stipendiumhungaricum.hu/courses/course/1530-msc-computer-science", "OFFICIAL", ""),
    ("CSC 2027/28 open (Indonesia, 14 Feb 2027)", "Chinese Embassy Indonesia", "https://id.china-embassy.gov.cn/eng/sgdt/202610/t20261027_12047753.htm", "OFFICIAL (Indonesia only)", ""),
    ("DAAD EPOS upper third; 2 yrs exp.", "DAAD scholarship database - EPOS", "https://www2.daad.de/deutschland/stipendium/datenbank/en/21148-scholarship-database/?detail=50076777", "OFFICIAL", ""),
    ("EPOS 2027/28 course list", "DAAD EPOS deadlines PDF", "https://static.daad.de/media/daad_de/pdfs_nicht_barrierefrei/in-deutschland-studieren-forschen-lehren/epos-course_overview__2027_28.pdf", "OFFICIAL", "No CS course seen in snippets"),
    ("Blocked account EUR 11,904/yr", "DAAD - finances", "https://www.daad.de/en/studying-in-germany/living-in-germany/finances/", "OFFICIAL", "2026 update UNVERIFIED"),
    ("Manitoba 3.0/4.5 last two years", "UManitoba FGPS international requirements", "https://umanitoba.ca/graduate-studies/admissions/requirements/international", "OFFICIAL", ""),
    ("Manitoba CS min funding CAD 20,000/yr", "UManitoba CS funding", "https://umanitoba.ca/science/student-experience/financial-aid-awards/computer-science", "OFFICIAL", "Contingent on progress"),
    ("Canada PAL exemption for master's from 1 Jan 2026", "IRCC graduate student", "https://www.canada.ca/en/immigration-refugees-citizenship/services/study-canada/graduate-student.html", "OFFICIAL", ""),
    ("IRCC proof of funds CAD 22,895", "IRCC DM transition binder 2026", "https://www.canada.ca/en/immigration-refugees-citizenship/corporate/transparency/transition-binders/deputy-minister-2026/international-student-program.html", "OFFICIAL", "From 1 Sep 2025"),
    ("UK maintenance GBP 1,136 / 1,483", "GOV.UK financial evidence", "https://www.gov.uk/guidance/financial-evidence-for-student-and-child-student-route-applicants", "OFFICIAL (conflict 1,171)", ""),
    ("US Proclamation 10998 country lists", "White House proclamation Dec 2025", "https://www.whitehouse.gov/presidential-actions/2025/12/restricting-and-limiting-the-entry-of-foreign-nationals-to-protect-the-security-of-the-united-states/", "OFFICIAL", "Ethiopia not listed"),
    ("MCF UBC EOI 6 Sep 2026", "UBC MCF Graduate Admissions", "https://mastercardfdn.scholars.ubc.ca/graduate-admissions/graduate-admissions-2/", "OFFICIAL", ""),
    ("MCF Edinburgh 2027/28 process & programmes", "Edinburgh Global - timelines", "https://edinburgh-global.ed.ac.uk/mastercard-foundation-scholars-program/apply-for-a-scholarship/application-timelines", "OFFICIAL", ""),
    ("MCF Toronto partner restriction", "U of T MCF announcement 2026-27", "https://internationalexperience.utoronto.ca/wp-content/uploads/Graduate-MCF-Announcement-2026-27-1.pdf", "OFFICIAL", ""),
    ("Manchester E&M: Ethiopia eligible; Aug 2026", "UoM Equity & Merit", "https://www.manchester.ac.uk/study/masters/fees-and-funding/masters-student-funding/equity-merit-scholarships/", "OFFICIAL/SECONDARY", "Date conflict 6 Dec"),
    ("Fulbright Ethiopia 2027-28 deadline 2 May 2026", "etcareers guide", "https://etcareers.com/blog/fully-funded-scholarships-ethiopians-guide", "SECONDARY", ""),
    ("Turkiye Burslari 75%; Jan-Feb 2027", "opportunitiesnexus", "https://opportunitiesnexus.com/turkiye-scholarships-2027/", "SECONDARY", ""),
    ("Applicant GPA 2.62 (627.75/240)", "Applicant transcript (Google Drive)", "Nebyu_Dejenie_BSc_Computer_Science_Certificate_Transcript.pdf", "PRIMARY DOCUMENT", "Reported 2.67 not found"),
]
for i, (cl, st, url, vs, nt) in enumerate(SRC):
    r = 5 + i
    rs.cell(row=r, column=1, value=cl)
    rs.cell(row=r, column=2, value=st)
    c = rs.cell(row=r, column=3, value=url)
    if url.startswith("http"):
        c.hyperlink = url
    rs.cell(row=r, column=4, value=VERIFIED_ON).number_format = "dd-mmm-yyyy"
    rs.cell(row=r, column=5, value="Not stated")
    rs.cell(row=r, column=6, value=vs)
    rs.cell(row=r, column=7, value=nt)
body(rs, 5, 4 + len(SRC), 1, 7)
rs.freeze_panes = "A5"
rs.auto_filter.ref = f"A4:G{4 + len(SRC)}"

# Sheet order: exactly as specified
order = ["Executive Dashboard", "Master Scholarship Database", "University Comparison", "Eligibility Assessment",
         "Application Deadlines", "Required Documents", "Recommendation Tracking", "Application Progress",
         "Financial Coverage", "Visa Preparation", "Interview Preparation", "Research Sources"]
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
wb.active = 0
wb.save(OUT)
print("saved", OUT)
