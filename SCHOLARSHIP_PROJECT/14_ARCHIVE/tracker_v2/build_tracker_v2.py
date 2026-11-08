"""Builds Scholarship_Tracker_v2.xlsx (12 original sheets + Discrepancy Log, Action Items, Risk Register).

Run:  python3 build_tracker_v2.py   then recalc with LibreOffice (recalc.py).
Facts mirror 03_SCHOLARSHIP_RESEARCH/Scholarship_Eligibility_Report_v2.md (verification pass 2026-11-05).
Evidence classes: E0 own document | E1 official page read directly | E2 official-domain extract | E3 secondary.
"""
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

VERIFIED_ON = date(2026, 11, 5)
OUT = "Scholarship_Tracker_v2.xlsx"
F = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(name=F, bold=True, color="FFFFFF", size=10)
BODY = Font(name=F, size=10)
INPUT = Font(name=F, size=10, color="0000FF")
LINK = Font(name=F, size=10, color="008000")
TITLE = Font(name=F, bold=True, size=14, color="1F3864")
SUB = Font(name=F, italic=True, size=9, color="555555")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GREEN = PatternFill("solid", fgColor="C6EFCE")
AMBER = PatternFill("solid", fgColor="FFEB9C")
RED = PatternFill("solid", fgColor="FFC7CE")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")

# --------------------------------------------------------------------------------------------
# Opportunity data. scores = (Academic, Funding, Relevance, Prof-exp, Flexibility, Feasibility)
# gpa = GPA eligibility at 2.62 ; disc = depends on blocking discrepancies? ; eng = English status
# --------------------------------------------------------------------------------------------
OPPS = [
    dict(id="S02", pri="A", country="Sweden", host="Eligible Swedish university (choose from SI list, 13 Nov 2026)",
         prog="MSc Software Eng. / Cybersecurity / Computer Systems (TBC)", sch="SI Scholarship for Global Professionals (SISGP)",
         fund="FULLY FUNDED", cover="Tuition paid to university; monthly allowance; SEK 15,000 travel grant. SI excludes insurance - check university cover",
         gpa_rule="No SI minimum; university merit ranking", gpa="Eligible",
         eng_req="English 6 met by Ethiopian bachelor's", eng="Not needed",
         work=">=3,000 h before Feb 2027 from <=3 employers + leadership; SI templates signed & stamped",
         opens=date(2026, 10, 16), deadline=date(2027, 1, 15), tz="CET (univ.); SI closes 14:59 CET 24 Feb",
         status="OPEN NOW", intake="Aug 2027", url="https://si.se/en/apply/scholarships/swedish-institute-scholarships-for-global-professionals/",
         ev="E2", sep="Yes - universityadmissions.se (15 Jan) + SI portal (8-24 Feb)", disc="Yes (SI work/leadership forms)",
         my="Request HR letters; after 13 Nov pick 2-4 programmes; apply by 15 Jan; pay SEK 900 by 1 Feb",
         cl="Map programme selection criteria after 13 Nov; prepare SI-template CV + motivation once facts confirmed",
         risk="Medium", risk_note="Employer must sign SI forms - CV dates must match; university merit ranking at 2.62",
         scores=(8, 9, 9, 10, 6, 8)),
    dict(id="S01", pri="A", country="Rwanda (US degree)", host="Carnegie Mellon University Africa (Kigali)",
         prog="MS Information Technology (MSIT)", sch="Mastercard Foundation Scholars Program at CMU-Africa",
         fund="FULLY FUNDED", cover="IF selected: tuition & fees, insurance, travel, basic living, project support. Need-based; admission does not guarantee award",
         gpa_rule="No GPA cut-off (most admits >=3.0)", gpa="Eligible",
         eng_req="IELTS 6.5 / TOEFL 81 (FR page 87) / DET 120; PTE not listed", eng="Not started",
         work="Not required; MCF: leadership, giving back, financial need, age <=35",
         opens=date(2026, 10, 1), deadline=date(2026, 12, 15), tz="Not stated - submit 72 h early",
         status="OPEN NOW", intake="Aug 2027", url="https://www.africa.engineering.cmu.edu/admissions/index.html",
         ev="E2", sep="Admission first; scholarship considered after admission", disc="Yes (CV + essays)",
         my="Create portal account; paste essay prompts; confirm 3 recommenders; fill Evidence Worksheet",
         cl="Adapt essays to real prompts; finalise CMU CV after HR letters; video-essay mocks",
         risk="Medium", risk_note="MCF need criterion; GPA below typical admit; tight early deadline",
         scores=(8, 7, 10, 8, 8, 7)),
    dict(id="S04", pri="B", country="Finland / Italy / Netherlands", host="LUT, Univ. L'Aquila, VU Amsterdam",
         prog="SE4GD - Software Engineers for Green Deal (EMJM)", sch="Erasmus Mundus Joint Master scholarship",
         fund="FULLY FUNDED", cover="IF selected; self-funded fee EUR 9,000/yr (partner country)",
         gpa_rule="UNVERIFIED", gpa="Eligibility requires university confirmation", eng_req="UNVERIFIED", eng="Not started",
         work="Not required", opens=None, deadline=None, tz="TBC", status="ELIGIBILITY UNCERTAIN", intake="Sep 2027",
         url="https://se4gd.eu", ev="E2 (partial)", sep="No - single application", disc="Yes (CV)",
         my="Send eligibility email (template in Erasmus package)", cl="Monitor se4gd.eu; draft motivation once call published",
         risk="Medium", risk_note="Requirements and dates unknown", scores=(6, 9, 9, 7, 5, 5)),
    dict(id="S05", pri="B", country="China", host="Chinese universities (Type B) / Ethiopian MoE (Type A)",
         prog="MSc Computer Science / Software Engineering (English-taught)", sch="Chinese Government Scholarship (CSC)",
         fund="FULLY FUNDED", cover="Tuition, accommodation, stipend, insurance (amounts to verify)",
         gpa_rule="'Academically excellent'; university-specific", gpa="Eligibility requires university confirmation",
         eng_req="IELTS/TOEFL per university", eng="Not started", work="Not required", opens=date(2027, 1, 5), deadline=None,
         tz="Set by Ethiopian MoE", status="UPCOMING", intake="Sep 2027", url="https://www.campuschina.org/content/details3_74775.html",
         ev="E2", sep="Pre-admission letter from university + CSC application", disc="Yes (CV)",
         my="Ask MoE scholarship office about 2027 CSC quota & deadline", cl="Shortlist 3 English-taught CS programmes (Type B)",
         risk="Medium", risk_note="Ethiopia-specific dates unknown; 'academically excellent' wording", scores=(6, 9, 8, 6, 6, 6)),
    dict(id="S09", pri="B", country="Hungary", host="Univ. of Szeged / ELTE / Univ. of Pannonia (TBC)", prog="MSc Computer Science",
         sch="Stipendium Hungaricum", fund="NEARLY FULLY FUNDED",
         cover="Tuition-free; HUF 43,700/mo (partial); dorm or HUF 40,000/mo; insurance",
         gpa_rule="Entrance test/interview; 56/100 scholarship threshold", gpa="Potentially eligible",
         eng_req="B2 per programme", eng="Not started", work="Not required", opens=None, deadline=None, tz="CET 14:00 (prior pattern)",
         status="OPENING SOON", intake="Sep 2027", url="https://stipendiumhungaricum.hu", ev="E2 (2026/27 terms)",
         sep="DreamApply application + Ethiopian sending-partner nomination", disc="Yes (CV)",
         my="Watch for 2027/28 call (~mid-Nov); ask MoE about nomination steps", cl="Compare 3 MSc CS programmes' entrance tests",
         risk="Low", risk_note="Stipend is partial - living-cost gap", scores=(8, 6, 8, 5, 8, 7)),
    dict(id="S03", pri="B", country="Norway + 6 EU partners", host="NTNU (coordinator), Aalto, DTU + partners",
         prog="CYBERSURE - Cybersecurity and Assurance (EMJM)", sch="Erasmus Mundus Joint Master scholarship",
         fund="FULLY FUNDED", cover="IF selected: tuition, travel, monthly allowance (full & partial awards)",
         gpa_rule="Minimum GPA exists (number unpublished); academic screen FIRST", gpa="Eligibility requires university confirmation",
         eng_req="IELTS 6.5 (W>=6.0) / TOEFL / Cambridge; no DET; Ethiopia not exempt", eng="Not started",
         work="Not required", opens=date(2026, 11, 16), deadline=date(2027, 1, 4), tz="Finnish time (verify)",
         status="OPENING SOON", intake="Aug 2027", url="https://cybersure-master.eu/admission", ev="E2",
         sep="Single online application (scholarship within)", disc="Yes (CV)",
         my="Send GPA eligibility email by 12 Nov; book IELTS only if eligible", cl="Draft motivation letter after facts confirmed",
         risk="High", risk_note="2.62 may fail first academic screen; IELTS needed by early Jan", scores=(5, 9, 9, 6, 3, 6)),
    dict(id="S10", pri="B", country="Canada", host="University of Manitoba", prog="MSc Computer Science (thesis)",
         sch="Departmental funding (automatic consideration)", fund="PARTIALLY FUNDED",
         cover="Min. funding figure to re-verify (CAD 20,000/yr found in Phase 1); tuition+living gap likely",
         gpa_rule="3.0/4.5 over last two years (60 credit hours)", gpa="Eligibility requires university confirmation",
         eng_req="IELTS 6.5", eng="Not started", work="Not required", opens=None, deadline=date(2026, 12, 1), tz="Central (verify)",
         status="OPEN NOW", intake="Sep 2027", url="https://umanitoba.ca/graduate-studies/admissions/programs-of-study/computer-science-msc",
         ev="E2 (deadline unclear)", sep="Admission = funding consideration", disc="Yes (CV)",
         my="Decide by 12 Nov: pursue or defer (needs IELTS + supervisor by 1 Dec)", cl="Draft supervisor-contact email",
         risk="High", risk_note="Deadline ~1 Dec; IELTS timing; funding gap", scores=(6, 6, 8, 6, 5, 5)),
    dict(id="S08", pri="B", country="South Korea", host="Korean universities via Embassy (Addis) or University track",
         prog="MSc Computer Science / IT", sch="Global Korea Scholarship - Graduate (GKS-G) 2027", fund="FULLY FUNDED",
         cover="Airfare, tuition, allowance, language year, insurance (per guideline)",
         gpa_rule="CGPA >=2.64/4.0 OR top 20% / 80th percentile", gpa="Ineligible",
         eng_req="Optional (adds points)", eng="Not needed", work="Not required", opens=None, deadline=None, tz="KST",
         status="UPCOMING", intake="Sep 2027 (+ language year)", url="https://et.mofa.go.kr/et-en/brd/m_23675/view.do?seq=758548",
         ev="E2 (2026 cycle)", sep="Embassy track or University track", disc="No",
         my="Ask registrar for class-rank/percentile letter (use only if top 20%)", cl="Re-assess when 2027 notice appears (~Feb 2027)",
         risk="High", risk_note="2.62 < 2.64 - only rank-letter route", scores=(2, 10, 8, 6, 2, 6)),
    dict(id="S13", pri="B", country="United Kingdom", host="UK universities (x3 choices)", prog="MSc Cloud / Cyber / Software Eng.",
         sch="Chevening Scholarship 2028/29", fund="FULLY FUNDED", cover="Tuition, stipend, return airfare, visa, allowances",
         gpa_rule="Unconditional UK offer required", gpa="Eligible", eng_req="Per university", eng="Not started",
         work=">=2,800 h post-degree", opens=None, deadline=None, tz="UTC", status="UPCOMING", intake="Sep 2028",
         url="https://www.chevening.org/scholarship/ethiopia/", ev="E2", sep="Chevening + separate university applications",
         disc="Yes (CV)", my="Build leadership/network evidence over 2027", cl="Prepare Chevening plan by Jul 2027",
         risk="Low", risk_note="2027/28 closed 6 Oct 2026", scores=(7, 10, 8, 10, 6, 2)),
    dict(id="S14", pri="B", country="United States", host="US universities via Fulbright", prog="MS Computer Science / IT",
         sch="Fulbright Foreign Student Program 2028-29", fund="FULLY FUNDED", cover="Tuition, stipend, airfare, health cover",
         gpa_rule="UNVERIFIED (Ethiopia)", gpa="Eligibility requires university confirmation", eng_req="TOEFL/IELTS", eng="Not started",
         work="Valued", opens=None, deadline=None, tz="EAT", status="UPCOMING", intake="Aug 2028", url="https://foreign.fulbrightonline.org",
         ev="E3", sep="Fulbright handles placement", disc="Yes (CV)", my="Monitor US Embassy Addis from early 2027",
         cl="Verify Ethiopia cycle when published", risk="Medium", risk_note="J-1 two-year home residency; US policy changes",
         scores=(7, 10, 8, 8, 6, 3)),
    dict(id="S06", pri="B", country="France / Belgium / Luxembourg", host="UBS, ULB, Univ. Luxembourg", prog="CYBERUS - Cybersecurity (EMJM)",
         sch="Erasmus Mundus Joint Master scholarship", fund="FULLY FUNDED", cover="IF selected",
         gpa_rule="UNVERIFIED", gpa="Eligibility requires university confirmation", eng_req="UNVERIFIED", eng="Not started",
         work="Not required", opens=None, deadline=None, tz="TBC", status="ELIGIBILITY UNCERTAIN", intake="Sep 2027",
         url="https://master-cyberus.eu", ev="E3", sep="TBC", disc="Yes (CV)", my="None until call confirmed",
         cl="Check for 2027 scholarship call", risk="Medium", risk_note="2027 call unconfirmed", scores=(6, 9, 9, 7, 4, 3)),
    dict(id="S07", pri="B", country="Belgium / Spain / partners", host="ULB, UPC + partners", prog="DEAI - Data Engineering & AI (EMJM)",
         sch="Erasmus Mundus Joint Master scholarship", fund="FULLY FUNDED", cover="EUR 1,400/mo up to 24 mo, tuition, insurance (E3)",
         gpa_rule="Not published", gpa="Eligibility requires university confirmation", eng_req="UNVERIFIED", eng="Not started",
         work="Not required", opens=date(2026, 10, 20), deadline=date(2027, 1, 11), tz="12:00 GMT", status="OPEN NOW", intake="Sep 2027",
         url="https://deai.ulb.be/", ev="E3", sep="TBC", disc="Yes (CV)", my="Only if consortium confirms 2.62 acceptable",
         cl="None unless chosen as 3rd EMJM", risk="High", risk_note="Maths-heavy; transcript weakness", scores=(5, 9, 7, 5, 4, 6)),
    dict(id="S11", pri="B", country="Italy", host="Public university (e.g. Padova, PoliTo)", prog="MSc Computer Science / Cybersecurity",
         sch="Regional DSU grant (ESU / EDISU)", fund="NEARLY FULLY FUNDED", cover="~EUR 8,000+/yr + fee exemption (need-based, E3)",
         gpa_rule="Per university", gpa="Potentially eligible", eng_req="Per programme", eng="Not started", work="Not required",
         opens=None, deadline=None, tz="CET", status="UPCOMING", intake="Sep/Oct 2027", url="https://www.esu.pd.it", ev="E3",
         sep="University admission + regional DSU call", disc="Yes (CV)", my="Tell me if the Dec 2024 Italy application was submitted",
         cl="Phase 2b: verify 3 Italian programmes + DSU calls", risk="Medium", risk_note="Need-based paperwork (ISEE)", scores=(6, 7, 8, 5, 6, 6)),
    dict(id="S12", pri="B", country="Turkiye", host="Turkish universities", prog="MSc Computer Engineering / CS", sch="Turkiye Burslari",
         fund="FULLY FUNDED", cover="Tuition, stipend, accommodation, insurance, airfare, language year",
         gpa_rule=">=75% (E3); 4.0-scale conversion UNVERIFIED", gpa="Eligibility requires university confirmation",
         eng_req="Per programme", eng="Not started", work="Not required", opens=date(2027, 1, 10), deadline=date(2027, 2, 20), tz="TRT",
         status="UPCOMING", intake="Sep 2027", url="https://www.turkiyeburslari.gov.tr", ev="E3", sep="Single portal",
         disc="Yes (CV)", my="-", cl="Verify GPA conversion method", risk="High", risk_note="75% threshold vs 2.62", scores=(4, 9, 7, 5, 4, 7)),
    dict(id="S23", pri="C", country="Germany", host="Public universities (e.g. TH Koln)", prog="MSc Communication Systems & Networks / CS",
         sch="None (tuition-free; self-funded)", fund="SELF-FUNDED", cover="Tuition-free; blocked account EUR 11,904/yr",
         gpa_rule="Programme-specific (uni-assist)", gpa="Eligibility requires university confirmation", eng_req="Per programme",
         eng="Not started", work="Not required", opens=None, deadline=None, tz="CET", status="UPCOMING", intake="Oct 2027",
         url="https://www.th-koeln.de/en/academics/communication-systems-and-networks-master--how-to-apply_6834.php", ev="E2",
         sep="n/a", disc="Yes (CV)", my="-", cl="Backup only", risk="Medium", risk_note="Not funded", scores=(5, 2, 9, 6, 5, 6)),
    dict(id="S22", pri="C", country="Germany", host="EPOS-listed universities", prog="Development-related master's (no CS confirmed)",
         sch="DAAD EPOS", fund="FULLY FUNDED", cover="Stipend, insurance, travel", gpa_rule="Upper third ('far above average')",
         gpa="Ineligible", eng_req="Per course", eng="Not needed", work=">=2 years", opens=None, deadline=None, tz="CET",
         status="ELIGIBILITY UNCERTAIN", intake="Oct 2027", url="https://www2.daad.de/deutschland/stipendium/datenbank/en/21148-scholarship-database/?detail=50076777",
         ev="E2", sep="Apply to course", disc="No", my="-", cl="Reconsider only with upper-third rank letter", risk="High",
         risk_note="Academic bar", scores=(2, 10, 3, 9, 2, 3)),
    dict(id="S26", pri="C", country="Saudi Arabia", host="KAUST", prog="MS Computer Science", sch="KAUST Fellowship (automatic if admitted)",
         fund="FULLY FUNDED", cover="Fellowship per offer", gpa_rule="Minimum GPA 3.0/4.0", gpa="Ineligible", eng_req="Required",
         eng="Not needed", work="Not required", opens=None, deadline=date(2027, 1, 3), tz="AST", status="OPEN NOW", intake="Aug 2027",
         url="https://admissions.kaust.edu.sa/how-to-apply/entry-requirements", ev="E2", sep="No", disc="No", my="-",
         cl="None", risk="High", risk_note="Mandatory GPA minimum", scores=(0, 10, 9, 6, 1, 6)),
    dict(id="S27", pri="C", country="UAE", host="MBZUAI", prog="MSc AI / ML / Applied AI", sch="MBZUAI funding (unconfirmed)",
         fund="PARTIALLY FUNDED", cover="Not confirmed this pass; Applied AI: no scholarships", gpa_rule="Minimum CGPA 3.0 (some 3.5)",
         gpa="Ineligible", eng_req="IELTS 6.5 / TOEFL 90", eng="Not needed", work="Not required", opens=date(2026, 9, 7),
         deadline=date(2026, 12, 15), tz="17:00 GST", status="OPEN NOW", intake="Aug 2027", url="https://mbzuai.ac.ae/admissions/graduate-masters-admissions",
         ev="E2", sep="-", disc="No", my="-", cl="None", risk="High", risk_note="Mandatory GPA minimum", scores=(0, 5, 7, 5, 1, 6)),
    dict(id="S28", pri="C", country="Africa (AU)", host="Pan African University", prog="Master's (various)", sch="AU full scholarship",
         fund="FULLY FUNDED", cover="Full", gpa_rule="Age <=25 (male) for master's", gpa="Ineligible", eng_req="-", eng="Not needed",
         work="-", opens=None, deadline=date(2025, 12, 15), tz="-", status="CLOSED", intake="2026/27",
         url="https://au.int/en/announcements/20251115/pan-african-university-call-scholarships-applications-academic-year-2026-2027",
         ev="E2", sep="-", disc="No", my="-", cl="None", risk="High", risk_note="Age limit", scores=(0, 9, 6, 5, 1, 0)),
    dict(id="S29", pri="C", country="Rwanda / Senegal", host="AIMS", prog="African Master's in Machine Intelligence (AMMI)",
         sch="Full scholarship for all admitted", fund="FULLY FUNDED", cover="Full", gpa_rule="~B / 80% average", gpa="Ineligible",
         eng_req="-", eng="Not needed", work="-", opens=None, deadline=None, tz="-", status="ELIGIBILITY UNCERTAIN", intake="TBC",
         url="https://aimsammi.org/admission-2/", ev="E2", sep="-", disc="No", my="-", cl="None", risk="High",
         risk_note="Academic bar; 2027 cycle unknown", scores=(2, 10, 6, 4, 2, 2)),
    dict(id="S17", pri="C", country="Canada", host="UBC", prog="Eligible master's", sch="Mastercard Foundation Scholars at UBC",
         fund="FULLY FUNDED", cover="Full", gpa_rule="UBC minimum", gpa="Potentially eligible", eng_req="Per programme", eng="Not needed",
         work="Valued", opens=date(2026, 8, 31), deadline=date(2026, 9, 6), tz="PT", status="CLOSED", intake="Sep 2027",
         url="https://mastercardfdn.scholars.ubc.ca/graduate-admissions/", ev="E2", sep="EOI then invitation", disc="No",
         my="-", cl="Watch Aug 2027", risk="Low", risk_note="Closed", scores=(5, 10, 6, 7, 4, 1)),
    dict(id="S21", pri="C", country="United Kingdom", host="University of Manchester", prog="Eligible MSc list",
         sch="Equity and Merit Scholarship", fund="FULLY FUNDED", cover="Tuition + living", gpa_rule="Per offer", gpa="Potentially eligible",
         eng_req="Per offer", eng="Not needed", work="Valued", opens=date(2026, 8, 1), deadline=date(2026, 8, 31), tz="UK",
         status="CLOSED", intake="Sep 2027", url="https://www.manchester.ac.uk/study/masters/fees-and-funding/masters-student-funding/equity-merit-scholarships/",
         ev="E2/E3 (date conflict)", sep="Offer first", disc="No", my="-", cl="Verify date conflict", risk="Low", risk_note="Closed",
         scores=(6, 9, 4, 7, 5, 1)),
    dict(id="S20", pri="C", country="United Kingdom", host="University of Edinburgh", prog="Listed sustainability MSc (no CS)",
         sch="Mastercard Foundation Scholars Programme", fund="FULLY FUNDED", cover="Full", gpa_rule="Per programme", gpa="Ineligible",
         eng_req="-", eng="Not needed", work="-", opens=None, deadline=date(2026, 11, 19), tz="UK", status="CLOSED", intake="Sep 2027",
         url="https://edinburgh-global.ed.ac.uk/mastercard-foundation-scholars-program/apply-for-a-scholarship/application-timelines",
         ev="E2", sep="-", disc="No", my="-", cl="None", risk="High", risk_note="Missed mandatory session; no CS", scores=(6, 10, 2, 6, 5, 0)),
    dict(id="S19", pri="C", country="Canada", host="University of Toronto", prog="Eligible master's", sch="Mastercard Foundation Scholars at U of T",
         fund="FULLY FUNDED", cover="Full", gpa_rule="Partner-institution graduates only (2026-27)", gpa="Ineligible", eng_req="-",
         eng="Not needed", work="-", opens=None, deadline=None, tz="ET", status="ELIGIBILITY UNCERTAIN", intake="2027-28",
         url="https://internationalexperience.utoronto.ca/global-experiences/global-scholarships/mastercard-foundation-scholars",
         ev="E2", sep="-", disc="No", my="-", cl="None", risk="High", risk_note="Institutional restriction", scores=(1, 10, 6, 6, 3, 2)),
    dict(id="S24", pri="C", country="United States", host="Stanford", prog="Any", sch="Knight-Hennessy Scholars", fund="FULLY FUNDED",
         cover="Full", gpa_rule="Highly selective", gpa="Ineligible", eng_req="-", eng="Not needed", work="-", opens=None, deadline=None,
         tz="PT", status="CLOSED", intake="Sep 2027", url="https://knight-hennessy.stanford.edu", ev="E3 (not re-verified)", sep="-",
         disc="No", my="-", cl="None", risk="High", risk_note="Not competitive", scores=(2, 10, 7, 7, 2, 0)),
    dict(id="S25", pri="C", country="UK (Commonwealth)", host="UK universities", prog="Any", sch="Commonwealth Shared / Master's",
         fund="FULLY FUNDED", cover="Full", gpa_rule="n/a", gpa="Ineligible", eng_req="-", eng="Not needed", work="-", opens=None,
         deadline=None, tz="-", status="CLOSED", intake="-", url="https://cscuk.fcdo.gov.uk", ev="Fact", sep="-", disc="No", my="-",
         cl="None", risk="High", risk_note="Ethiopia not a Commonwealth member", scores=(0, 10, 7, 7, 0, 0)),
]

PRIORITIES = ["A", "B", "C"]
STATUSES = ["OPEN NOW", "OPENING SOON", "UPCOMING", "CLOSED", "DISCONTINUED", "ELIGIBILITY UNCERTAIN"]
ELIG = ["Eligible", "Potentially eligible", "Ineligible", "Eligibility requires university confirmation"]
FUNDS = ["FULLY FUNDED", "NEARLY FULLY FUNDED", "PARTIALLY FUNDED", "TUITION ONLY", "SELF-FUNDED"]
ENG = ["Not needed", "Not started", "Booked", "Done - meets", "Done - below"]
EVC = ["E0", "E1", "E2", "E2 (partial)", "E3", "Fact"]
RISK = ["Low", "Medium", "High"]
PROG = ["Not started", "In progress", "Done", "N/A"]


def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font, cell.fill, cell.border = HDR_FONT, HDR_FILL, BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 36


def body(ws, r1, r2, c1, c2, font_cols=None):
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            cell.border = BORDER
            cell.alignment = WRAP
            f = (font_cols or {}).get(cell.column)
            cell.font = f if f is not None else BODY


def title(ws, text, sub):
    ws["A1"], ws["A2"] = text, sub
    ws["A1"].font, ws["A2"].font = TITLE, SUB


def add_dv(ws, options, rng):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


def priority_cf(ws, rng):
    for v, fill in (("A", GREEN), ("B", AMBER), ("C", RED)):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{v}"'], fill=fill))


wb = Workbook()
N = len(OPPS)

# =============================== Discrepancy Log (built early: dashboard refers to it) ==========
DISC = [
    ("DR-01", "CGPA 2.62 vs 2.67", "CRITICAL", "Resolved", "Transcript (E0)", "2.62 used everywhere", "You: say where 2.67 came from"),
    ("DR-02", "Study start date (CV 14/07/2018 vs AcYear 2017/18)", "MODERATE", "Partly resolved", "Registrar record", "AcYear 2017/18; start month pending", "You: confirm first-semester month"),
    ("DR-03", "EQF level 5 / extension mode omitted", "LOW", "Resolved", "Transcript", "EQF 6; extension stated", "-"),
    ("DR-04", "Faculty/school naming", "LOW", "Resolved", "Transcript", "As printed on D1", "-"),
    ("DR-05", "Name spelling (Nebyu / Nebiyu)", "BLOCKING", "Open", "Passport", "Passport spelling", "You: give passport name line"),
    ("DR-06", "SantimPay titles & dates (4 versions)", "BLOCKING", "Open", "SantimPay HR letter", "Pending", "You: request HR experience letter"),
    ("DR-07", "Kifiya dates (3 versions) + overlap", "BLOCKING", "Open", "Kifiya HR letter", "Pending", "You: request HR letter; explain overlap"),
    ("DR-08", "Plan International role (duplicate bullets)", "CRITICAL", "Excluded", "Official letter if real", "Excluded", "You: confirm it was an error; fix LinkedIn"),
    ("DR-09", "Yandex Odoo developer", "HIGH", "Excluded", "Contract / payment record", "Excluded", "You: provide evidence or drop"),
    ("DR-10", "Infonans / Gebeya / GAO Tek / Dereja", "MODERATE", "Open", "Letters / certificates", "Pending", "You: confirm each"),
    ("DR-11", "Total experience 3y6m vs ~4y3m", "BLOCKING", "Open", "Derived from HR letters", "Pending", "Follows DR-06/07"),
    ("DR-12", "Certifications (CKA ID etc.)", "MODERATE", "Open", "Verification IDs", "Expired items marked", "You: send IDs/links"),
    ("DR-13", "Tools changed between CV versions", "BLOCKING", "Open", "Your confirmation + manager", "Pending", "You: fill Evidence Worksheet"),
    ("DR-14", "Lecturer letter overstates grades", "HIGH", "Resolved", "Transcript", "Do not reuse", "You: request new letters"),
    ("DR-15", "2024 Italy intent letter", "LOW", "Open", "Your confirmation", "Do not reuse", "You: was it submitted?"),
    ("DR-16", "Other person's documents in Drive", "ADMIN", "Resolved", "-", "Excluded", "You: move to separate folder"),
]
dl_ = wb.active
dl_.title = "Discrepancy Log"
title(dl_, "Discrepancy Log (see 00_MASTER_PROFILE/Discrepancy_Resolution_Report_v1)",
      "BLOCKING items must be Resolved before any CV/essay is final. Status dropdown in blue.")
header(dl_, 4, ["ID", "Issue", "Severity", "Status", "Authoritative evidence", "Current verified value", "Who acts / next step"],
       [8, 44, 12, 15, 30, 30, 42])
for i, d in enumerate(DISC):
    for c, v in enumerate(d, 1):
        dl_.cell(row=5 + i, column=c, value=v)
DLAST = 4 + len(DISC)
body(dl_, 5, DLAST, 1, 7, {4: INPUT})
add_dv(dl_, ["Open", "Partly resolved", "Resolved", "Excluded"], f"D5:D{DLAST + 10}")
for v, fill in (("Open", RED), ("Partly resolved", AMBER), ("Resolved", GREEN), ("Excluded", PatternFill("solid", fgColor="D9D9D9"))):
    dl_.conditional_formatting.add(f"D5:D{DLAST}", CellIsRule(operator="equal", formula=[f'"{v}"'], fill=fill))
dl_.conditional_formatting.add(f"C5:C{DLAST}", CellIsRule(operator="equal", formula=['"BLOCKING"'], font=Font(name=F, bold=True, color="9C0006")))
dl_.freeze_panes = "B5"
dl_.auto_filter.ref = f"A4:G{DLAST}"
BLOCK_FLAG = f"IF(COUNTIFS('Discrepancy Log'!$C$5:$C${DLAST},\"BLOCKING\",'Discrepancy Log'!$D$5:$D${DLAST},\"<>Resolved\")=0,1,0)"

# =============================== Financial Coverage ===========================================
fc = wb.create_sheet("Financial Coverage")
title(fc, "Financial Coverage Matrix & Funding-Gap Calculator",
      "Yes / Partial / No / UNVERIFIED. Completeness = (Yes + 0.5 x Partial) / 9. Never assume airfare, insurance or visa costs without evidence.")
items = ["Tuition", "Monthly living", "Accommodation", "Return airfare", "Health insurance", "Visa costs", "Settlement", "Books", "Research"]
header(fc, 4, ["ID", "Opportunity", "Funding class"] + items + ["Yes", "Partial", "Completeness", "Notes"], [7, 40, 18] + [11] * 9 + [7, 8, 12, 40])
COV = {
    "S01": ["Yes", "Yes", "UNVERIFIED", "Yes", "Yes", "UNVERIFIED", "UNVERIFIED", "UNVERIFIED", "Yes"],
    "S02": ["Yes", "Yes", "Partial", "Partial", "No", "UNVERIFIED", "No", "No", "No"],
    "S03": ["Yes", "Yes", "Partial", "Partial", "UNVERIFIED", "Partial", "UNVERIFIED", "No", "No"],
    "S04": ["Yes", "Yes", "Partial", "Partial", "UNVERIFIED", "Partial", "UNVERIFIED", "No", "No"],
    "S05": ["Yes", "Yes", "Yes", "UNVERIFIED", "Yes", "UNVERIFIED", "UNVERIFIED", "No", "No"],
    "S08": ["Yes", "Yes", "Partial", "Yes", "Yes", "UNVERIFIED", "Yes", "UNVERIFIED", "UNVERIFIED"],
    "S09": ["Yes", "Partial", "Yes", "No", "Yes", "No", "No", "No", "No"],
    "S10": ["Partial", "Partial", "No", "No", "No", "No", "No", "No", "Partial"],
    "S13": ["Yes", "Yes", "Partial", "Yes", "UNVERIFIED", "Yes", "Yes", "UNVERIFIED", "UNVERIFIED"],
    "S23": ["Yes", "No", "No", "No", "No", "No", "No", "No", "No"],
}
FCNOTES = {"S01": "Only if selected as Mastercard Foundation Scholar (need-based)", "S02": "SI FAQ: insurance not included",
           "S03": "EM 'Partial' = paid via participation-cost contribution", "S04": "EM 'Partial' = participation-cost contribution",
           "S09": "Stipend officially partial", "S10": "Funding minimum to re-verify"}
r = 5
for o in OPPS:
    if o["id"] not in COV:
        continue
    fc.cell(row=r, column=1, value=o["id"])
    fc.cell(row=r, column=2, value=f'{o["sch"]} - {o["prog"]}')
    fc.cell(row=r, column=3, value=o["fund"])
    for j, v in enumerate(COV[o["id"]]):
        fc.cell(row=r, column=4 + j, value=v)
    fc.cell(row=r, column=13, value=f'=COUNTIF(D{r}:L{r},"Yes")')
    fc.cell(row=r, column=14, value=f'=COUNTIF(D{r}:L{r},"Partial")')
    fc.cell(row=r, column=15, value=f"=(M{r}+0.5*N{r})/9").number_format = "0%"
    fc.cell(row=r, column=16, value=FCNOTES.get(o["id"], ""))
    r += 1
FCL = r - 1
body(fc, 5, FCL, 1, 16, {c: INPUT for c in range(4, 13)})
add_dv(fc, ["Yes", "Partial", "No", "UNVERIFIED"], f"D5:L{FCL + 20}")
for v, fill in (("Yes", GREEN), ("No", RED), ("UNVERIFIED", AMBER)):
    fc.conditional_formatting.add(f"D5:L{FCL}", CellIsRule(operator="equal", formula=[f'"{v}"'], fill=fill))
fc.conditional_formatting.add(f"O5:O{FCL}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"))
g = FCL + 3
fc.cell(row=g, column=1, value="Funding-gap calculator (official figures only; yellow = inputs to confirm)").font = Font(name=F, bold=True, size=11)
header(fc, g + 1, ["Case", "Item", "Amount", "Currency", "Source"])
GAP = [
    ("Germany self-funded (S23)", "Blocked account per year", 11904, "EUR", "DAAD, since Jan 2025 (2026 update unverified)"),
    ("Germany self-funded (S23)", "Scholarship income per year", 0, "EUR", "No scholarship"),
    ("Germany self-funded (S23)", "Annual gap (living only)", "=C{a}-C{b}", "EUR", "Formula"),
    ("Canada - Manitoba (S10)", "IRCC proof of funds (living)", 22895, "CAD", "IRCC from 1 Sep 2025; check 2026 figure"),
    ("Canada - Manitoba (S10)", "International tuition per year (INPUT)", None, "CAD", "UNVERIFIED - enter from UM fees page"),
    ("Canada - Manitoba (S10)", "Department minimum funding per year (re-verify)", 20000, "CAD", "UM CS page (Phase 1); truncated in Phase 3"),
    ("Canada - Manitoba (S10)", "Year-1 gap vs IRCC test", "=C{d}+C{e}-C{f}", "CAD", "Formula; blank tuition understates gap"),
    ("Hungary (S09)", "Monthly stipend", 43700, "HUF", "SH Call 2026/27"),
    ("Hungary (S09)", "Housing contribution if no dorm", 40000, "HUF", "SH Call 2026/27"),
    ("Hungary (S09)", "Total monthly support (no dorm)", "=C{h}+C{i}", "HUF", "Formula"),
    ("CMU-Africa (S01)", "Tuition per year, African students 2026-27", 16000, "USD", "CMU-Africa tuition page (E2) - covered only if MCF Scholar"),
]
start = g + 2
refs = {k: start + i for i, k in enumerate("abcdefghijk")}
for k, (case, item, amt, cur, src) in enumerate(GAP):
    rr = start + k
    fc.cell(row=rr, column=1, value=case)
    fc.cell(row=rr, column=2, value=item)
    fc.cell(row=rr, column=3, value=amt.format(**refs) if isinstance(amt, str) else amt)
    fc.cell(row=rr, column=4, value=cur)
    fc.cell(row=rr, column=5, value=src)
body(fc, start, start + len(GAP) - 1, 1, 5)
for k, row in enumerate(GAP):
    c = fc.cell(row=start + k, column=3)
    c.number_format = "#,##0;(#,##0);-"
    if not isinstance(row[2], str):
        c.font, c.fill = INPUT, YELLOW

# =============================== Application Progress =========================================
ap = wb.create_sheet("Application Progress")
title(ap, "Application Progress", "Done / In progress / Not started / N/A. Completion = Done / (steps not N/A). Only YOU create accounts, pay or submit.")
steps = ["Eligibility confirmed", "Portal account (you)", "Documents uploaded", "CV final", "SOP / essays", "References", "English test", "Fee / waiver", "Submitted (you)"]
header(ap, 4, ["ID", "Opportunity", "Priority"] + steps + ["Completion %", "Admission status", "Scholarship status", "Interview status", "Visa status"],
       [7, 44, 9] + [12] * 9 + [12, 14, 14, 14, 12])
track = [o for o in OPPS if o["pri"] in ("A", "B")][:10]
ns = len(steps)
for i, o in enumerate(track):
    r = 5 + i
    ap.cell(row=r, column=1, value=o["id"])
    ap.cell(row=r, column=2, value=f'{o["sch"]} - {o["prog"]}')
    ap.cell(row=r, column=3, value=o["pri"])
    vals = ["Done" if o["gpa"] == "Eligible" else "In progress"] + ["Not started"] * (ns - 1)
    if o["eng"] == "Not needed":
        vals[6] = "N/A"
    for j, v in enumerate(vals):
        ap.cell(row=r, column=4 + j, value=v)
    lc = get_column_letter(3 + ns)
    ap.cell(row=r, column=4 + ns, value=f'=IFERROR(COUNTIF(D{r}:{lc}{r},"Done")/({ns}-COUNTIF(D{r}:{lc}{r},"N/A")),0)').number_format = "0%"
    for c in range(5 + ns, 9 + ns):
        ap.cell(row=r, column=c, value="Not started")
APL = 4 + len(track)
CMP = get_column_letter(4 + ns)
body(ap, 5, APL, 1, 8 + ns, {c: INPUT for c in list(range(4, 4 + ns)) + list(range(5 + ns, 9 + ns))})
add_dv(ap, PROG, f"D5:{get_column_letter(3 + ns)}{APL + 20}")
add_dv(ap, ["Not started", "Submitted", "Under review", "Interview", "Offer", "Conditional offer", "Rejected", "Withdrawn"],
       f"{get_column_letter(5 + ns)}5:{get_column_letter(7 + ns)}{APL + 20}")
add_dv(ap, ["Not started", "Preparing", "Submitted", "Biometrics done", "Approved", "Refused"], f"{get_column_letter(8 + ns)}5:{get_column_letter(8 + ns)}{APL + 20}")
ap.conditional_formatting.add(f"{CMP}5:{CMP}{APL}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"))
ap.conditional_formatting.add(f"D5:{get_column_letter(3 + ns)}{APL}", CellIsRule(operator="equal", formula=['"Done"'], fill=GREEN))
ap.conditional_formatting.add(f"D5:{get_column_letter(3 + ns)}{APL}", CellIsRule(operator="equal", formula=['"In progress"'], fill=AMBER))
priority_cf(ap, f"C5:C{APL}")
ap.freeze_panes = "D5"

# =============================== Master Scholarship Database ==================================
db = wb.create_sheet("Master Scholarship Database", 0)
title(db, "Master Scholarship Database (v2)",
      f"Verification pass {VERIFIED_ON:%d %b %Y}. Evidence: E0 own doc | E1 official page read | E2 official-domain extract | E3 secondary. Blue = editable.")
cols = ["ID", "Priority", "Country", "University / Host", "Programme", "Scholarship", "Funding class", "Coverage summary",
        "GPA / academic rule", "GPA eligibility at 2.62", "English requirement", "English status", "Work / leadership requirement",
        "Opening date", "Deadline", "Deadline timezone", "Deadline status", "Application urgency", "Days to deadline",
        "Intake", "Official URL", "Evidence class", "Separate univ. & scholarship applications?", "Depends on blocking discrepancies?",
        "Funding completeness", "Readiness score /100", "Risk level", "Risk note", "YOUR next action", "CLAUDE next action",
        "Verification date", "Admission status", "Scholarship status"]
widths = [7, 8, 16, 26, 28, 26, 16, 36, 26, 20, 28, 13, 28, 12, 12, 16, 16, 13, 10, 12, 36, 11, 28, 18, 12, 12, 9, 30, 40, 40, 12, 14, 14]
header(db, 4, cols, widths)
R0 = 5
RL = R0 + N - 1
ci = {c: i + 1 for i, c in enumerate(cols)}
L = {c: get_column_letter(i) for c, i in ci.items()}
for i, o in enumerate(OPPS):
    r = R0 + i
    vals = {"ID": o["id"], "Priority": o["pri"], "Country": o["country"], "University / Host": o["host"], "Programme": o["prog"],
            "Scholarship": o["sch"], "Funding class": o["fund"], "Coverage summary": o["cover"], "GPA / academic rule": o["gpa_rule"],
            "GPA eligibility at 2.62": o["gpa"], "English requirement": o["eng_req"], "English status": o["eng"],
            "Work / leadership requirement": o["work"], "Opening date": o["opens"] or "TBC", "Deadline": o["deadline"] or "TBC",
            "Deadline timezone": o["tz"], "Deadline status": o["status"], "Intake": o["intake"], "Official URL": o["url"],
            "Evidence class": o["ev"], "Separate univ. & scholarship applications?": o["sep"],
            "Depends on blocking discrepancies?": o["disc"], "Risk level": o["risk"], "Risk note": o["risk_note"],
            "YOUR next action": o["my"], "CLAUDE next action": o["cl"], "Verification date": VERIFIED_ON,
            "Admission status": "Not started", "Scholarship status": "Not started"}
    for k, v in vals.items():
        c = db.cell(row=r, column=ci[k], value=v)
        if isinstance(v, date):
            c.number_format = "dd-mmm-yyyy"
    db.cell(row=r, column=ci["Official URL"]).hyperlink = o["url"]
    dcol = L["Deadline"]
    db.cell(row=r, column=ci["Days to deadline"], value=f'=IF(ISNUMBER({dcol}{r}),{dcol}{r}-TODAY(),"")')
    dd = f'{L["Days to deadline"]}{r}'
    db.cell(row=r, column=ci["Application urgency"],
            value=f'=IF({L["Deadline status"]}{r}="CLOSED","CLOSED",IF(NOT(ISNUMBER({dcol}{r})),"TBC",IF({dd}<0,"PASSED",IF({dd}<=21,"URGENT",IF({dd}<=60,"HIGH","NORMAL")))))')
    db.cell(row=r, column=ci["Funding completeness"],
            value=f"=IFERROR(INDEX('Financial Coverage'!$O$5:$O${FCL},MATCH(A{r},'Financial Coverage'!$A$5:$A${FCL},0)),\"n/a\")").number_format = "0%"
    gp = f'{L["GPA eligibility at 2.62"]}{r}'
    en = f'{L["English status"]}{r}'
    dp = f'{L["Depends on blocking discrepancies?"]}{r}'
    prog_ = f"IFERROR(INDEX('Application Progress'!${CMP}$5:${CMP}${APL},MATCH(A{r},'Application Progress'!$A$5:$A${APL},0)),0)"
    gscore = f'IF({gp}="Eligible",1,IF({gp}="Ineligible",0,0.5))'
    escore = f'IF(OR({en}="Not needed",{en}="Done - meets"),1,IF({en}="Booked",0.5,0))'
    dscore = f'IF(LEFT({dp},2)="No",1,{BLOCK_FLAG})'
    db.cell(row=r, column=ci["Readiness score /100"], value=f"=ROUND(100*(0.3*{gscore}+0.2*{escore}+0.2*{dscore}+0.3*{prog_}),0)")
RLc = len(cols)
inputs = {ci[k]: INPUT for k in ["Priority", "GPA eligibility at 2.62", "English status", "Deadline status", "Risk level", "Admission status", "Scholarship status"]}
body(db, R0, RL, 1, RLc, inputs)
add_dv(db, PRIORITIES, f'{L["Priority"]}{R0}:{L["Priority"]}{RL + 50}')
add_dv(db, ELIG, f'{L["GPA eligibility at 2.62"]}{R0}:{L["GPA eligibility at 2.62"]}{RL + 50}')
add_dv(db, ENG, f'{L["English status"]}{R0}:{L["English status"]}{RL + 50}')
add_dv(db, STATUSES, f'{L["Deadline status"]}{R0}:{L["Deadline status"]}{RL + 50}')
add_dv(db, FUNDS, f'{L["Funding class"]}{R0}:{L["Funding class"]}{RL + 50}')
add_dv(db, RISK, f'{L["Risk level"]}{R0}:{L["Risk level"]}{RL + 50}')
add_dv(db, ["Not started", "Submitted", "Under review", "Interview", "Offer", "Conditional offer", "Rejected", "Withdrawn"],
       f'{L["Admission status"]}{R0}:{L["Scholarship status"]}{RL + 50}')
priority_cf(db, f'B{R0}:B{RL}')
gcol = L["GPA eligibility at 2.62"]
db.conditional_formatting.add(f"{gcol}{R0}:{gcol}{RL}", CellIsRule(operator="equal", formula=['"Eligible"'], fill=GREEN))
db.conditional_formatting.add(f"{gcol}{R0}:{gcol}{RL}", CellIsRule(operator="equal", formula=['"Ineligible"'], fill=RED))
db.conditional_formatting.add(f"{gcol}{R0}:{gcol}{RL}", FormulaRule(formula=[f'LEFT({gcol}{R0},4)="Elig"'], fill=AMBER, stopIfTrue=False))
ucol = L["Application urgency"]
db.conditional_formatting.add(f"{ucol}{R0}:{ucol}{RL}", CellIsRule(operator="equal", formula=['"URGENT"'], fill=RED, font=Font(name=F, bold=True, color="9C0006")))
db.conditional_formatting.add(f"{ucol}{R0}:{ucol}{RL}", CellIsRule(operator="equal", formula=['"HIGH"'], fill=AMBER))
rk = L["Risk level"]
db.conditional_formatting.add(f"{rk}{R0}:{rk}{RL}", CellIsRule(operator="equal", formula=['"High"'], fill=RED))
db.conditional_formatting.add(f"{rk}{R0}:{rk}{RL}", CellIsRule(operator="equal", formula=['"Medium"'], fill=AMBER))
db.conditional_formatting.add(f"{rk}{R0}:{rk}{RL}", CellIsRule(operator="equal", formula=['"Low"'], fill=GREEN))
for k in ("Readiness score /100",):
    c = L[k]
    db.conditional_formatting.add(f"{c}{R0}:{c}{RL}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color="5B9BD5"))
fcm = L["Funding completeness"]
db.conditional_formatting.add(f"{fcm}{R0}:{fcm}{RL}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"))
db.freeze_panes = "C5"
db.auto_filter.ref = f"A4:{get_column_letter(RLc)}{RL}"

# =============================== Eligibility Assessment =======================================
el = wb.create_sheet("Eligibility Assessment", 1)
title(el, "Eligibility Assessment & Programme Matching Score (v2)",
      "Strategic score (0-100), NOT an admission probability. Edit blue scores (0-10) or weights in row 4.")
el["A4"] = "Weights ->"
el["A4"].font = Font(name=F, bold=True)
for j, w in enumerate([0.25, 0.25, 0.15, 0.15, 0.10, 0.10]):
    c = el.cell(row=4, column=6 + j, value=w)
    c.number_format, c.font, c.fill = "0%", INPUT, YELLOW
el["L4"] = "=SUM(F4:K4)"
el["L4"].number_format = "0%"
ecols = ["ID", "Opportunity", "Priority", "GPA eligibility at 2.62", "Evidence class", "Academic eligibility (25%)", "Funding coverage (25%)",
         "Programme relevance (15%)", "Professional exp. alignment (15%)", "Admission flexibility (10%)", "Application feasibility (10%)",
         "Matching score /100", "Rank", "Rationale"]
header(el, 5, ecols, [7, 46, 9, 22, 11, 12, 12, 12, 12, 12, 12, 11, 7, 56])
E0 = 6
EL = E0 + N - 1
for i, o in enumerate(OPPS):
    r, dbr = E0 + i, R0 + i
    el.cell(row=r, column=1, value=f"='Master Scholarship Database'!A{dbr}")
    el.cell(row=r, column=2, value=f"='Master Scholarship Database'!F{dbr}&\" - \"&'Master Scholarship Database'!E{dbr}")
    el.cell(row=r, column=3, value=f"='Master Scholarship Database'!B{dbr}")
    el.cell(row=r, column=4, value=f"='Master Scholarship Database'!{gcol}{dbr}")
    el.cell(row=r, column=5, value=f"='Master Scholarship Database'!{L['Evidence class']}{dbr}")
    for j, s in enumerate(o["scores"]):
        el.cell(row=r, column=6 + j, value=s)
    el.cell(row=r, column=12, value=f"=ROUND(SUMPRODUCT(F{r}:K{r},$F$4:$K$4)*10,1)")
    el.cell(row=r, column=13, value=f"=RANK(L{r},$L${E0}:$L${EL})+COUNTIF($L${E0 - 1}:L{r - 1},L{r})")
    el.cell(row=r, column=14, value=o["risk_note"])
body(el, E0, EL, 1, 14, {**{c: LINK for c in range(1, 6)}, **{c: INPUT for c in range(6, 12)}})
dv = DataValidation(type="whole", operator="between", formula1="0", formula2="10")
el.add_data_validation(dv)
dv.add(f"F{E0}:K{EL}")
el.conditional_formatting.add(f"L{E0}:L{EL}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color="5B9BD5"))
priority_cf(el, f"C{E0}:C{EL}")
el.freeze_panes = "C6"
el.auto_filter.ref = f"A5:N{EL}"

# =============================== Application Deadlines ========================================
dls = wb.create_sheet("Application Deadlines", 2)
title(dls, "Application Deadline & Milestone Calendar (v2)", "Days remaining recalculates daily. Re-verify each date on the official page before acting.")
MIL = [
    ("ALL", "Request HR letters (SantimPay, Kifiya); registrar rank letter; recommenders asked", date(2026, 11, 12), "Internal", "Internal"),
    ("S03/S04", "Send GPA-eligibility emails to CYBERSURE & SE4GD", date(2026, 11, 12), "Internal", "Internal"),
    ("S10", "Decide: pursue Manitoba (needs IELTS + supervisor by 1 Dec) or defer", date(2026, 11, 12), "Internal", "Internal"),
    ("S02", "SI templates + eligible programme list published", date(2026, 11, 13), "CET", "E2"),
    ("S03", "CYBERSURE application opens", date(2026, 11, 16), "EET", "E2"),
    ("ALL", "Book IELTS Academic (test early-mid Dec)", date(2026, 11, 20), "Internal", "Internal"),
    ("S10", "University of Manitoba MSc CS deadline (likely)", date(2026, 12, 1), "CST (verify)", "E2 (unclear)"),
    ("S01", "Recommenders entered in CMU-Africa portal", date(2026, 12, 1), "Internal", "Internal"),
    ("S01", "Target CMU-Africa submission (3-day buffer)", date(2026, 12, 12), "Internal", "Internal"),
    ("S01", "CMU-Africa EARLY deadline (DET & test fees covered)", date(2026, 12, 15), "Not stated", "E2 (year inferred)"),
    ("S03", "Target CYBERSURE submission", date(2027, 1, 2), "Internal", "Internal"),
    ("S03", "CYBERSURE deadline (4 Jan; another page says 5 Jan)", date(2027, 1, 4), "EET (verify)", "E2"),
    ("S07", "DEAI scholarship deadline (only if chosen)", date(2027, 1, 11), "12:00 GMT", "E3"),
    ("S02", "universityadmissions.se deadline (rank programmes)", date(2027, 1, 15), "CET", "E2"),
    ("S01", "CMU-Africa FINAL deadline", date(2027, 1, 15), "Not stated", "E2"),
    ("S02", "Swedish application fee SEK 900 due", date(2027, 2, 1), "CET", "E2"),
    ("S02", "SI scholarship portal opens", date(2027, 2, 8), "CET", "E2"),
    ("S02", "SI scholarship portal closes (14:59 CET)", date(2027, 2, 24), "14:59 CET", "E2"),
    ("S02", "SI results", date(2027, 4, 22), "CET", "E2"),
]
MIL.sort(key=lambda x: x[2])
header(dls, 4, ["ID", "Milestone", "Date", "Timezone", "Days remaining", "Urgency", "Evidence", "Done?"], [9, 62, 13, 16, 12, 12, 18, 9])
for i, (oid, m, d, tz, src) in enumerate(MIL):
    r = 5 + i
    for c, v in enumerate([oid, m, d, tz, f"=C{r}-TODAY()",
                           f'=IF(H{r}="Yes","DONE",IF(E{r}<0,"PASSED",IF(E{r}<=14,"URGENT",IF(E{r}<=45,"SOON","PLANNED"))))', src, "No"], 1):
        cell = dls.cell(row=r, column=c, value=v)
        if c == 3:
            cell.number_format = "dd-mmm-yyyy"
DLL = 4 + len(MIL)
body(dls, 5, DLL, 1, 8, {8: INPUT})
add_dv(dls, ["Yes", "No"], f"H5:H{DLL + 30}")
for v, fill in (("URGENT", RED), ("SOON", AMBER), ("DONE", GREEN)):
    dls.conditional_formatting.add(f"F5:F{DLL}", CellIsRule(operator="equal", formula=[f'"{v}"'], fill=fill))
dls.freeze_panes = "A5"
dls.auto_filter.ref = f"A4:H{DLL}"

# =============================== University Comparison ========================================
uc = wb.create_sheet("University Comparison", 1)
title(uc, "University / Programme Comparison (priority programmes, v2)", "UNVERIFIED = not confirmed from an official source. E2 = official-domain extract.")
ucols = ["University", "Country", "City", "Programme", "Degree", "Programme website", "Admission website", "Scholarship website",
         "Tuition fee", "Scholarship coverage", "Minimum GPA", "Ethiopian qualification", "English requirement", "Work experience",
         "Opening", "Deadline", "Intake", "Application fee", "Fee waiver", "Required documents", "Funding gap", "Eligibility (2.62)",
         "Competitiveness", "Next action"]
header(uc, 4, ucols, [24, 12, 11, 28, 9, 30, 30, 30, 20, 28, 22, 24, 28, 20, 13, 20, 10, 12, 18, 40, 20, 20, 14, 34])
UROWS = [
    ["Swedish university (from SI list)", "Sweden", "TBC", "MSc Software Eng. / Cybersecurity / Computer Systems", "MSc",
     "https://www.universityadmissions.se", "https://www.universityadmissions.se",
     "https://si.se/en/apply/scholarships/swedish-institute-scholarships-for-global-professionals/", "Programme-specific",
     "Tuition + monthly allowance + SEK 15,000 travel (no insurance)", "No national minimum", "ETA-accredited programme required",
     "English 6 met by Ethiopian bachelor's", ">=3,000 h (<=3 employers) + leadership", "16-Oct-2026",
     "15-Jan-2027 (univ) / 24-Feb-2027 14:59 CET (SI)", "Aug 2027", "SEK 900", "No (SI scholars still pay)",
     "Univ: diploma, transcripts, passport. SI: SI-template CV (<=3 p), work+leadership proof (signed/stamped), 2 refs (SI template), motivation, ID",
     "Insurance", "Eligible", "Moderate", "Select programmes 13 Nov"],
    ["Carnegie Mellon University Africa", "Rwanda", "Kigali", "MS Information Technology", "MSc",
     "https://www.africa.engineering.cmu.edu/academics/programs/msit.html", "https://www.africa.engineering.cmu.edu/admissions/index.html",
     "https://www.africa.engineering.cmu.edu/impact/mastercard-foundation-scholars.html", "USD 16,000/yr (African, 2026-27)",
     "Full IF MCF Scholar (need-based)", "None (most >=3.0)", "Accepted", "IELTS 6.5 / TOEFL 81 / DET 120 (no PTE)",
     "Not required", "01-Oct-2026", "15-Dec-2026 early / 15-Jan-2027 final", "Aug 2027", "None", "DET + test fees if early",
     "CV, transcript, 3 essays <=300 words (+optional), video essay, 3 recommenders, English score, 1-h technical test",
     "Full if MCF; else tuition + living", "Eligible (admission)", "Moderate", "Portal + essays + referees"],
    ["NTNU (coord.) / Aalto / DTU + partners", "Norway etc.", "Trondheim + 2nd yr", "CYBERSURE Cybersecurity and Assurance", "MSc (double)",
     "https://www.ntnu.edu/studies/mscybsure", "https://cybersure-master.eu/admission", "https://cybersure-master.eu/admission",
     "Covered if EM scholar", "EM full/partial", "Minimum exists (unpublished); academic screen first", ">=180 ECTS / 3 yrs (met)",
     "IELTS 6.5 (W6.0) / TOEFL / Cambridge; no DET", "Not required", "16-Nov-2026", "04-Jan-2027", "Aug 2027", "UNVERIFIED",
     "UNVERIFIED", "Online application; transcripts, CV, motivation, English, references, residence proof", "None if EM scholar",
     "Requires confirmation", "Low", "Eligibility email"],
    ["LUT / Univ. L'Aquila / VU Amsterdam", "FI / IT / NL", "Multiple", "SE4GD", "MSc (triple)", "https://se4gd.eu", "https://se4gd.eu",
     "https://se4gd.eu", "EUR 9,000/yr partner-country (self-funded)", "EM if selected", "UNVERIFIED", "UNVERIFIED", "UNVERIFIED",
     "Not required", "UNVERIFIED", "UNVERIFIED", "Sep 2027", "UNVERIFIED", "UNVERIFIED", "Single application for programme + scholarship",
     "None if EM scholar", "Requires confirmation", "Low-Moderate", "Eligibility email"],
    ["University of Szeged", "Hungary", "Szeged", "MSc Computer Science", "MSc", "https://apply.stipendiumhungaricum.hu/courses/course/1530-msc-computer-science",
     "https://apply.stipendiumhungaricum.hu", "https://stipendiumhungaricum.hu", "Covered", "Tuition + HUF 43,700/mo + housing + insurance",
     "Score >=50/100 (grades 50, interview 20, diploma 20, extras 10)", "Via Ethiopian sending-partner nomination", "B2", "Not required",
     "~Nov 2026 (UNVERIFIED)", "~15-Jan-2027 (UNVERIFIED)", "Sep 2027", "UNVERIFIED", "UNVERIFIED",
     "Diploma, transcript, CV, motivation, medical cert., English, nomination", "Living-cost gap", "Potentially eligible", "Moderate", "Await call"],
    ["University of Manitoba", "Canada", "Winnipeg", "MSc Computer Science (thesis)", "MSc",
     "https://umanitoba.ca/graduate-studies/admissions/programs-of-study/computer-science-msc",
     "https://umanitoba.ca/graduate-studies/admissions/requirements/international",
     "https://umanitoba.ca/science/student-experience/financial-aid-awards/computer-science", "UNVERIFIED",
     "Automatic consideration; minimum to re-verify", "3.0/4.5 last two years", "Assessed on last 2 years", "IELTS 6.5", "Not required",
     "Up to 16 months before", "~01-Dec-2026 (verify)", "Sep 2027", "CAD 130", "UNVERIFIED",
     "Transcripts, CV, statement, references, IELTS, supervisor contact", "See Financial Coverage", "Requires confirmation", "Low-Moderate",
     "Decide by 12 Nov"],
]
for i, row in enumerate(UROWS):
    for c, v in enumerate(row, 1):
        cell = uc.cell(row=5 + i, column=c, value=v)
        if isinstance(v, str) and v.startswith("https://"):
            cell.hyperlink = v
body(uc, 5, 4 + len(UROWS), 1, len(ucols))
uc.freeze_panes = "B5"
uc.auto_filter.ref = f"A4:{get_column_letter(len(ucols))}{4 + len(UROWS)}"

# =============================== Required Documents ===========================================
rd = wb.create_sheet("Required Documents")
title(rd, "Required Documents Register (v2)", "No passport/ID numbers in this file. Status dropdown in blue.")
header(rd, 4, ["#", "Document", "Needed for", "Status", "Location / source", "Verification needed", "Notes"], [5, 46, 24, 20, 28, 30, 44])
DOCS = [
    ("SantimPay HR experience letter (titles, dates, full-time, duties)", "ALL - BLOCKING", "Not started", "SantimPay HR", "Resolves DR-06/11", "Request this week"),
    ("Kifiya HR experience letter (dates, type)", "ALL - BLOCKING", "Not started", "Kifiya HR", "Resolves DR-07/11", "Request this week"),
    ("SI 'proof of work & leadership' form (2027/28 template)", "S02", "Not started", "si.se after 13 Nov", "Signed + stamped by employer", "Max 3 employers"),
    ("SI reference letters x2 (2027/28 template)", "S02", "Not started", "Referees", "Signed + stamped", "1 work-based; 1 civil society = merit"),
    ("Official/sealed transcript + temporary certificate", "All", "Have (needs official copy)", "Drive / registrar", "Certified copies", "CGPA 2.62"),
    ("Printed BSc diploma", "All", "Not started", "DTU registrar", "-", "Temporary certificate says pending"),
    ("Class-rank / percentile letter", "S08, S22", "Not started", "Registrar", "Use only if favourable", ""),
    ("Last-two-years GPA statement", "S10", "Not started", "Registrar", "-", "Unofficial ~3.1"),
    ("MoFA consular confirmation of certificates", "S08, S05 (programme-specific)", "Not started", "Ethiopian MoFA", "-", "GKS requires for docs 11-19"),
    ("Passport name & expiry (redacted copy for screening)", "All", "Have (not reviewed)", "Drive 'Passport' folder", "DR-05", "Number never stored here"),
    ("IELTS Academic result", "S03, S10, S09, S05", "Not started", "British Council Addis Ababa", "Target 7.0, no band <6.5", "Book by 20 Nov"),
    ("Duolingo English Test result", "S01", "Not started", "Online", "Target >=120", "Fee covered if CMU early applicant"),
    ("Recommendation - final-project supervisor", "All", "Not started", "DTU", "Briefing pack", ""),
    ("Recommendation - lecturer (A- course)", "All", "Not started", "DTU", "Briefing pack", ""),
    ("Recommendation - SantimPay manager", "All", "Not started", "SantimPay", "Briefing pack", ""),
    ("Recommendation - 2nd professional", "S02, S01", "Not started", "Kifiya / SantimPay", "Briefing pack", ""),
    ("Achievement Evidence Worksheet (completed by you)", "All", "Not started", "02_PROFESSIONAL_CV", "-", "Needed for CV and essays"),
    ("Final-project title/abstract", "All", "Not started", "Your files / DTU", "-", "Strongest academic evidence"),
    ("Certification IDs (CKA, AZ-900)", "All", "Needs verification", "Vendor portals", "DR-12", ""),
    ("[former lecturer, name withheld] letter (Nov 2024)", "None", "Needs verification", "Drive", "DR-14 - do not reuse", ""),
    ("Medical certificate", "S09, S05", "Not started", "Approved clinic", "-", "Programme form"),
]
for i, d in enumerate(DOCS):
    rd.cell(row=5 + i, column=1, value=i + 1)
    for c, v in enumerate(d, 2):
        rd.cell(row=5 + i, column=c, value=v)
RDL = 4 + len(DOCS)
body(rd, 5, RDL, 1, 7, {4: INPUT})
add_dv(rd, ["Have", "Have (needs official copy)", "Have (not reviewed)", "Requested", "Not started", "Needs verification", "N/A"], f"D5:D{RDL + 30}")
rd.conditional_formatting.add(f"D5:D{RDL}", FormulaRule(formula=['LEFT(D5,4)="Have"'], fill=GREEN))
rd.conditional_formatting.add(f"D5:D{RDL}", CellIsRule(operator="equal", formula=['"Needs verification"'], fill=AMBER))
rd.freeze_panes = "C5"
rd.auto_filter.ref = f"A4:G{RDL}"

# =============================== Recommendation Tracking ======================================
rt = wb.create_sheet("Recommendation Tracking")
title(rt, "Recommendation Tracking", "Referees write and submit their own letters. Placeholders in [brackets].")
header(rt, 4, ["Referee (name, title)", "Type", "Relationship", "Agreed?", "For", "Briefing pack sent", "Portal link sent", "Status", "Deadline", "Notes"],
       [34, 12, 32, 10, 18, 14, 14, 14, 12, 40])
REFS = [["[Final-project supervisor, DTU]", "Academic", "Supervised Final Project I & II (A-)", "No", "S01, S02, S03", "No", "No", "Not requested", date(2026, 12, 1), "Ask by 12 Nov"],
        ["[Lecturer, DTU - A- course]", "Academic", "Compiler Design / Advanced Programming", "No", "S01, S03", "No", "No", "Not requested", date(2026, 12, 1), ""],
        ["[Line manager, SantimPay]", "Professional", "Direct supervisor", "No", "All", "No", "No", "Not requested", date(2026, 12, 1), "Also HR letter"],
        ["[2nd professional referee]", "Professional", "Kifiya supervisor or senior SantimPay colleague", "No", "S02, S01", "No", "No", "Not requested", date(2027, 2, 1), "SI template"]]
for i, row in enumerate(REFS):
    for c, v in enumerate(row, 1):
        cell = rt.cell(row=5 + i, column=c, value=v)
        if isinstance(v, date):
            cell.number_format = "dd-mmm-yyyy"
body(rt, 5, 8, 1, 10, {c: INPUT for c in (1, 4, 6, 7, 8)})
add_dv(rt, ["Not requested", "Requested", "Agreed", "Submitted", "Declined"], "H5:H40")
add_dv(rt, ["Yes", "No"], "D5:D40")
add_dv(rt, ["Yes", "No"], "F5:G40")

# =============================== Visa Preparation =============================================
vp = wb.create_sheet("Visa Preparation")
title(vp, "Visa Preparation (summary)", "Visa rules are separate from admission. Never use false documents or borrowed funds presented as your own.")
header(vp, 4, ["Country", "Visa / permit", "Official link", "Financial requirement", "Key documents", "Notes", "Status"], [16, 24, 44, 38, 44, 44, 13])
VISA = [
    ["Sweden", "Residence permit for studies", "https://www.migrationsverket.se", "UNVERIFIED (Migrationsverket monthly figure); SI award letter", "Admission, SI letter, insurance, passport", "Priority A destination", "Not started"],
    ["Rwanda", "Student permit", "https://www.migration.gov.rw", "UNVERIFIED", "CMU-Africa admission letter", "Priority A destination", "Not started"],
    ["Hungary", "Residence permit (study)", "https://enterhungary.gov.hu", "SH award letter", "Award, accommodation, insurance", "", "Not started"],
    ["Germany", "National (D) student visa", "https://www.germany.info", "Blocked account EUR 11,904/yr (DAAD, Jan 2025)", "Admission, funds, insurance", "Backup only", "Not started"],
    ["Canada", "Study permit", "https://www.canada.ca/en/immigration-refugees-citizenship/services/study-canada/graduate-student.html",
     "CAD 22,895 + tuition (from 1 Sep 2025)", "LOA, funds, biometrics", "Master's at public DLIs exempt from PAL since 1 Jan 2026", "Not started"],
    ["United Kingdom", "Student visa", "https://www.gov.uk/student-visa", "GBP 1,136/1,483 per month (conflict: 1,171)", "CAS, funds or sponsor letter, TB test", "2028 via Chevening", "Not started"],
    ["United States", "F-1 / J-1", "https://travel.state.gov", "Per I-20 / DS-2019", "DS-160, SEVIS, interview", "Ethiopia not on Proclamation 10998 lists (re-check)", "Not started"],
]
for i, row in enumerate(VISA):
    for c, v in enumerate(row, 1):
        cell = vp.cell(row=5 + i, column=c, value=v)
        if c == 3:
            cell.hyperlink = v
body(vp, 5, 4 + len(VISA), 1, 7, {7: INPUT})
add_dv(vp, ["Not started", "Preparing", "Submitted", "Biometrics done", "Approved", "Refused", "N/A"], "G5:G30")

# =============================== Interview Preparation ========================================
ip = wb.create_sheet("Interview Preparation")
title(ip, "Interview Preparation & English Test Plan (v2)", "Model answers only from verified facts. Log mock scores in blue cells.")
header(ip, 4, ["Question bank", "Target #", "Drafted", "Practised", "% ready", "Notes"], [40, 10, 10, 11, 10, 52])
QB = [("GPA & academic weaknesses", 6, "Prepare first"), ("Professional experience", 9, "After HR letters"), ("Leadership", 5, "Real examples only"),
      ("Motivation & fit", 4, ""), ("Ethiopia/Africa & future", 6, ""), ("Technical & academic", 10, "CS fundamentals"),
      ("Full bank expansion (50/30/20/20)", 120, "After verification")]
for i, (q, n, note) in enumerate(QB):
    r = 5 + i
    for c, v in enumerate([q, n, 0, 0, f"=IFERROR(D{r}/B{r},0)", note], 1):
        ip.cell(row=r, column=c, value=v)
    ip.cell(row=r, column=5).number_format = "0%"
r = 5 + len(QB)
ip.cell(row=r, column=1, value="TOTAL")
for c, f_ in ((2, f"=SUM(B5:B{r - 1})"), (3, f"=SUM(C5:C{r - 1})"), (4, f"=SUM(D5:D{r - 1})"), (5, f"=IFERROR(D{r}/B{r},0)")):
    ip.cell(row=r, column=c, value=f_)
ip.cell(row=r, column=5).number_format = "0%"
body(ip, 5, r, 1, 6, {3: INPUT, 4: INPUT})
ip.cell(row=r, column=1).font = Font(name=F, bold=True)
e0 = r + 3
ip.cell(row=e0, column=1, value="English plan: IELTS Academic (target 7.0, no band <6.5) + DET for CMU-Africa (target >=120)").font = Font(name=F, bold=True, size=11)
header(ip, e0 + 1, ["Week", "Focus", "Saturday", "Target", "Mock score", "Notes"])
PLAN = [("W1 9-15 Nov", "Daily: listening 20, reading 20, writing 20, SPEAKING 30 min", "Diagnostic IELTS + DET practice", "Baseline per skill", None, "Speaking self-rated B1"),
        ("W2 16-22 Nov", "Add vocab log; Task 1/Task 2 timing", "Writing under time", "Clear 4-paragraph essays", None, "Book IELTS by 20 Nov"),
        ("W3 23-29 Nov", "Speaking Part 2 daily; T/F/NG reading", "Full mock #2", "+0.5 weakest skill", None, "Mock interviews start"),
        ("W4 30 Nov-6 Dec", "Error-log review; DET task types", "DET practice + IELTS mock #3", "DET practice >=120", None, ""),
        ("W5 7-13 Dec", "Light review, logistics", "Real tests", "IELTS 7.0 / DET 120", None, "")]
for i, row in enumerate(PLAN):
    for c, v in enumerate(row, 1):
        ip.cell(row=e0 + 2 + i, column=c, value=v)
body(ip, e0 + 2, e0 + 1 + len(PLAN), 1, 6, {5: INPUT})
ip.column_dimensions["B"].width = 52
ip.column_dimensions["C"].width = 30

# =============================== Research Sources =============================================
rs = wb.create_sheet("Research Sources")
title(rs, "Research Sources & Evidence Log (v2)", "Full register: 03_SCHOLARSHIP_RESEARCH/Evidence_and_Source_Register_v2. E1 impossible this session (network policy).")
header(rs, 4, ["ID", "Claim", "URL", "Accessed", "Evidence class", "Action before submission"], [8, 56, 62, 12, 12, 30])
SRC = [
    ("E0-1", "CGPA 2.62; 627.75/240; award 14 Jul 2022; Extension; English medium", "Drive: DTU transcript + certificate", "E0", "-"),
    ("S-01", "SISGP: Ethiopia eligible; 3,000 h <=3 employers; leadership; SI templates; Europass disqualified; 14:59 CET", "https://si.se/en/apply/scholarships/swedish-institute-scholarships-for-global-professionals/", "E2", "Open & confirm after 13 Nov"),
    ("S-02", "SI excludes insurance", "https://si.se/en/apply/faq-scholarships-on-masters-level/", "E2", "Confirm"),
    ("S-04", "Ethiopian bachelor's meets English 6; ETA accreditation", "https://www.universityadmissions.se/en/apply-to-masters/provide-application-documents-masters/country-instructions/ethiopia/", "E2", "Confirm"),
    ("S-05", "Round 16 Oct 2026 - 15 Jan 2027; fee due 1 Feb", "https://www.universityadmissions.se/en/key-dates-and-deadlines/autumn-semester-dates/", "E2", "Confirm (page may be old)"),
    ("C-01", "MCF criteria; admission first", "https://www.africa.engineering.cmu.edu/impact/mastercard-foundation-scholars.html", "E2", "Confirm"),
    ("C-02", "Financial need; disadvantaged; age <=35", "https://www.africa.engineering.cmu.edu/_files/documents/financial-aid/financial-aid.pdf", "E2", "Confirm current version"),
    ("C-03", "No GPA cut-off; English minimums; waiver rules", "https://www.africa.engineering.cmu.edu/admissions/faq.html", "E2", "Confirm TOEFL 81 vs 87"),
    ("C-04", "1 Oct / 15 Dec / 15 Jan; DET fees covered early", "https://www.africa.engineering.cmu.edu/admissions/index.html", "E2", "Confirm year"),
    ("C-05", "3 essays <=300 words + optional; video essay", "https://www.africa.engineering.cmu.edu/admissions/how-to-apply/step-1.html", "E2", "Read prompts in portal"),
    ("C-06", "1-hour online technical test", "https://www.africa.engineering.cmu.edu/admissions/how-to-apply/step-2.html", "E2", "-"),
    ("EM-01", "CYBERSURE: 180 ECTS; academic screen first; IELTS 6.5/W6.0; no DET", "https://cybersure-master.eu/admission", "E2", "Email for GPA minimum"),
    ("EM-02", "CYBERSURE 16 Nov 2026 - 5 Jan 2027 (4 Jan elsewhere)", "https://www.aalto.fi/en/study-options/cybersecurity-and-assurance-cybersure-erasmus-mundus-master-of-science-technology", "E2", "Confirm"),
    ("EM-04", "SE4GD single application; EUR 9,000/yr partner", "https://elut.lut.fi/en/completing-studies/tuition-fee/scholarship-masters-dd-studies", "E2", "Confirm at se4gd.eu"),
    ("H-01", "SH: Ethiopia partner; HUF 43,700; housing; insurance; 56/100", "https://stipendiumhungaricum.hu/wp-content/uploads/2025/10/BA_MA_OTM_Call_for_Applications_2026_27.pdf", "E2", "Await 2027/28 call"),
    ("K-01", "GKS Ethiopia 2026: 10+13 quota; 12-25 Feb KST; MoFA confirmation", "https://et.mofa.go.kr/et-en/brd/m_23675/view.do?seq=758548", "E2", "Await 2027 notice"),
    ("K-02", "GKS 2.64 / 80% / top 20%", "https://www.mofa.go.kr/sg-en/brd/m_2435/view.do?seq=761574", "E2", "-"),
    ("CN-01", "CSC bilateral window Jan-Apr; <35; pre-admission letter", "https://www.campuschina.org/content/details3_74775.html", "E2", "Ask Ethiopian MoE"),
    ("M-01", "Manitoba 3.0/4.5 last two years", "https://umanitoba.ca/graduate-studies/admissions/requirements/international", "E2", "-"),
    ("M-02", "Manitoba CS funding (minimum to re-verify)", "https://umanitoba.ca/science/student-experience/financial-aid-awards/computer-science", "E2 (partial)", "Confirm amount"),
    ("X-01", "KAUST min GPA 3.0", "https://admissions.kaust.edu.sa/how-to-apply/entry-requirements", "E2", "-"),
    ("X-02", "MBZUAI min CGPA 3.0/3.5", "https://mbzuai.ac.ae/admissions/graduate-masters-admissions", "E2", "-"),
    ("X-03", "PAU master's age <=25 (male); closed", "https://au.int/en/announcements/20251115/pan-african-university-call-scholarships-applications-academic-year-2026-2027", "E2", "-"),
    ("X-04", "AMMI ~B / 80%", "https://aimsammi.org/admission-2/", "E2", "-"),
    ("L-02", "Test prices TOEFL/PTE/DET", "Secondary sites", "E3", "Confirm with providers"),
]
for i, (sid, cl, url, ev, act) in enumerate(SRC):
    r = 5 + i
    rs.cell(row=r, column=1, value=sid)
    rs.cell(row=r, column=2, value=cl)
    c = rs.cell(row=r, column=3, value=url)
    if url.startswith("http"):
        c.hyperlink = url
    rs.cell(row=r, column=4, value=VERIFIED_ON).number_format = "dd-mmm-yyyy"
    rs.cell(row=r, column=5, value=ev)
    rs.cell(row=r, column=6, value=act)
body(rs, 5, 4 + len(SRC), 1, 6)
rs.freeze_panes = "A5"
rs.auto_filter.ref = f"A4:F{4 + len(SRC)}"

# =============================== Action Items =================================================
ai = wb.create_sheet("Action Items")
title(ai, "Action Items - YOU vs CLAUDE", "Claude never creates accounts, pays fees, signs or submits. Status dropdown in blue.")
header(ai, 4, ["#", "Owner", "Action", "Due", "Days left", "Status", "Linked to"], [5, 10, 70, 12, 10, 14, 16])
ACTS = [
    ("YOU", "Make GitHub repo private OR approve deletion of the public branch", date(2026, 11, 6), "Privacy"),
    ("YOU", "Answer the Phase 3 confirmation questions", date(2026, 11, 8), "DR-01..16"),
    ("YOU", "Request HR experience letters from SantimPay and Kifiya", date(2026, 11, 12), "DR-06/07"),
    ("YOU", "Ask registrar: class-rank letter, last-2-years GPA, sealed transcripts, printed diploma", date(2026, 11, 12), "DR-01/02"),
    ("YOU", "Ask 4 referees (send briefing pack)", date(2026, 11, 12), "Recommendations"),
    ("YOU", "Send GPA-eligibility emails to CYBERSURE and SE4GD", date(2026, 11, 12), "S03/S04"),
    ("YOU", "Decide on Manitoba (pursue/defer)", date(2026, 11, 12), "S10"),
    ("YOU", "Create CMU-Africa portal account; paste essay prompts into package", date(2026, 11, 10), "S01"),
    ("YOU", "Fill Achievement Evidence Worksheet", date(2026, 11, 15), "DR-13"),
    ("YOU", "Book IELTS Academic (early-mid Dec)", date(2026, 11, 20), "English"),
    ("YOU", "Allow official domains in network settings (or confirm pages yourself)", date(2026, 11, 10), "Verification"),
    ("CLAUDE", "Re-verify all E2 facts on live pages once domains allowed", date(2026, 11, 12), "Verification"),
    ("CLAUDE", "Map SI eligible programmes + selection criteria (after 13 Nov)", date(2026, 11, 15), "S02"),
    ("CLAUDE", "Finalise CV v1.0 + 6 variants after HR letters", date(2026, 11, 20), "CV"),
    ("CLAUDE", "Adapt CMU-Africa essays to real prompts; 2 review rounds", date(2026, 11, 28), "S01"),
    ("CLAUDE", "Weekly mock interview (from week of 23 Nov)", date(2026, 11, 23), "Interview"),
    ("CLAUDE", "CYBERSURE motivation letter (if eligible)", date(2026, 12, 10), "S03"),
    ("CLAUDE", "Stipendium Hungaricum 2027/28 call analysis when published", date(2026, 11, 20), "S09"),
]
for i, (own, act, due, link) in enumerate(ACTS):
    r = 5 + i
    for c, v in enumerate([i + 1, own, act, due, f"=D{r}-TODAY()", "Not started", link], 1):
        cell = ai.cell(row=r, column=c, value=v)
        if c == 4:
            cell.number_format = "dd-mmm-yyyy"
AIL = 4 + len(ACTS)
body(ai, 5, AIL, 1, 7, {6: INPUT})
add_dv(ai, ["Not started", "In progress", "Done", "Blocked"], f"F5:F{AIL + 30}")
ai.conditional_formatting.add(f"F5:F{AIL}", CellIsRule(operator="equal", formula=['"Done"'], fill=GREEN))
ai.conditional_formatting.add(f"F5:F{AIL}", CellIsRule(operator="equal", formula=['"Blocked"'], fill=RED))
ai.conditional_formatting.add(f"B5:B{AIL}", CellIsRule(operator="equal", formula=['"YOU"'], font=Font(name=F, bold=True, color="1F3864")))
ai.freeze_panes = "A5"
ai.auto_filter.ref = f"A4:G{AIL}"

# =============================== Risk Register ================================================
rr_ = wb.create_sheet("Risk Register")
title(rr_, "Risk Register", "Likelihood and impact 1-5; score = L x I. Blue = editable.")
header(rr_, 4, ["#", "Risk", "Likelihood (1-5)", "Impact (1-5)", "Score", "Mitigation", "Owner"], [5, 52, 12, 12, 9, 60, 10])
RISKS = [
    ("Public GitHub repo exposes personal academic/employment analysis", 5, 4, "Make repo private / delete branch; no further pushes without approval", "YOU"),
    ("CV dates contradict employer letters / SI forms", 4, 5, "HR letters first; single verified history", "YOU"),
    ("2.62 fails academic screens (CYBERSURE, GKS, KAUST, MBZUAI)", 4, 3, "Target holistic routes (SI, CMU-Africa, Hungary); eligibility emails", "BOTH"),
    ("MCF scholarship not awarded despite CMU admission (need-based)", 3, 4, "Truthful need statement; ask about other CMU aid; parallel applications", "YOU"),
    ("English score below thresholds (speaking B1)", 3, 4, "5-week plan; IELTS early Dec; DET for CMU", "YOU"),
    ("Official facts only E2 (pages not opened directly)", 3, 3, "Allow domains / you confirm key pages before submitting", "BOTH"),
    ("Referees unavailable or late", 3, 3, "Ask 4 by 12 Nov; briefing pack", "YOU"),
    ("Compressed deadlines (15 Dec - 24 Feb)", 4, 3, "Weekly action list; buffers of 3 days", "BOTH"),
    ("Funding gap on partial options (Hungary, Manitoba)", 3, 3, "Treat as backups; compute gaps", "BOTH"),
]
for i, (rk_, l, im, mit, own) in enumerate(RISKS):
    r = 5 + i
    for c, v in enumerate([i + 1, rk_, l, im, f"=C{r}*D{r}", mit, own], 1):
        rr_.cell(row=r, column=c, value=v)
RRL = 4 + len(RISKS)
body(rr_, 5, RRL, 1, 7, {3: INPUT, 4: INPUT})
rr_.conditional_formatting.add(f"E5:E{RRL}", CellIsRule(operator="greaterThanOrEqual", formula=["16"], fill=RED))
rr_.conditional_formatting.add(f"E5:E{RRL}", CellIsRule(operator="between", formula=["9", "15"], fill=AMBER))
rr_.conditional_formatting.add(f"E5:E{RRL}", CellIsRule(operator="lessThan", formula=["9"], fill=GREEN))

# =============================== Executive Dashboard ==========================================
d = wb.create_sheet("Executive Dashboard", 0)
title(d, "Executive Dashboard - Fully Funded Master's 2027 (v2)",
      f"Nebyu Dejenie Tilahun | BSc CS, Debre Tabor Univ. 2022 | CGPA 2.62 (verified, transcript) | Phase 3 verification {VERIFIED_ON:%d %b %Y}")
d.column_dimensions["A"].width = 44
for col in "BCDEFG":
    d.column_dimensions[col].width = 16
d.column_dimensions["D"].width = 30


def block_header(row, col, labels):
    for j, h in enumerate(labels):
        c = d.cell(row=row, column=col + j, value=h)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER


DBP = f"'Master Scholarship Database'!$B${R0}:$B${RL}"
DBG = f"'Master Scholarship Database'!${gcol}${R0}:${gcol}${RL}"
DBU = f"'Master Scholarship Database'!${ucol}${R0}:${ucol}${RL}"
d["A4"] = "Profile verification status"
d["A4"].font = Font(name=F, bold=True, size=11)
block_header(5, 1, ["Metric", "Value"])
KP = [("Discrepancies total", f"=COUNTA('Discrepancy Log'!A5:A{DLAST})"),
      ("BLOCKING discrepancies still open", f"=COUNTIFS('Discrepancy Log'!C5:C{DLAST},\"BLOCKING\",'Discrepancy Log'!D5:D{DLAST},\"<>Resolved\")"),
      ("Resolved or excluded", f"=COUNTIF('Discrepancy Log'!D5:D{DLAST},\"Resolved\")+COUNTIF('Discrepancy Log'!D5:D{DLAST},\"Excluded\")"),
      ("CV / essays may be finalised?", f'=IF({BLOCK_FLAG}=1,"YES","NO - resolve blocking items")')]
for i, (k, f_) in enumerate(KP):
    d.cell(row=6 + i, column=1, value=k)
    d.cell(row=6 + i, column=2, value=f_)
body(d, 6, 9, 1, 2)
d.conditional_formatting.add("B9", FormulaRule(formula=['LEFT(B9,2)="NO"'], fill=RED, font=Font(name=F, bold=True, color="9C0006")))

d["D4"] = "Opportunities by priority"
d["D4"].font = Font(name=F, bold=True, size=11)
block_header(5, 4, ["Priority", "Count"])
for i, p in enumerate(PRIORITIES):
    d.cell(row=6 + i, column=4, value=p)
    d.cell(row=6 + i, column=5, value=f"=COUNTIF({DBP},D{6 + i})")
d.cell(row=9, column=4, value="Total")
d.cell(row=9, column=5, value="=SUM(E6:E8)")
body(d, 6, 9, 4, 5)
priority_cf(d, "D6:D8")

d["A11"] = "GPA eligibility at 2.62"
d["A11"].font = Font(name=F, bold=True, size=11)
block_header(12, 1, ["Status", "Count"])
for i, s in enumerate(ELIG):
    d.cell(row=13 + i, column=1, value=s)
    d.cell(row=13 + i, column=2, value=f"=COUNTIF({DBG},A{13 + i})")
body(d, 13, 16, 1, 2)
d["D11"] = "Application urgency"
d["D11"].font = Font(name=F, bold=True, size=11)
block_header(12, 4, ["Urgency", "Count"])
for i, s in enumerate(["URGENT", "HIGH", "NORMAL", "TBC", "CLOSED"]):
    d.cell(row=13 + i, column=4, value=s)
    d.cell(row=13 + i, column=5, value=f"=COUNTIF({DBU},D{13 + i})")
body(d, 13, 17, 4, 5)
d.conditional_formatting.add("D13:E13", FormulaRule(formula=["$E$13>0"], fill=RED))

d["A20"] = "Top-ranked opportunities (matching score) with readiness"
d["A20"].font = Font(name=F, bold=True, size=11)
block_header(21, 1, ["Opportunity", "Score", "Priority", "GPA status", "Readiness /100"])
rank_rng = f"'Eligibility Assessment'!$M${E0}:$M${EL}"
for k in range(1, 9):
    r = 21 + k
    d.cell(row=r, column=1, value=f"=IFERROR(INDEX('Eligibility Assessment'!$B${E0}:$B${EL},MATCH({k},{rank_rng},0)),\"\")")
    d.cell(row=r, column=2, value=f"=IFERROR(INDEX('Eligibility Assessment'!$L${E0}:$L${EL},MATCH({k},{rank_rng},0)),\"\")")
    d.cell(row=r, column=3, value=f"=IFERROR(INDEX('Eligibility Assessment'!$C${E0}:$C${EL},MATCH({k},{rank_rng},0)),\"\")")
    d.cell(row=r, column=4, value=f"=IFERROR(INDEX('Eligibility Assessment'!$D${E0}:$D${EL},MATCH({k},{rank_rng},0)),\"\")")
    idr = f"INDEX('Eligibility Assessment'!$A${E0}:$A${EL},MATCH({k},{rank_rng},0))"
    d.cell(row=r, column=5, value=f"=IFERROR(INDEX('Master Scholarship Database'!${L['Readiness score /100']}${R0}:${L['Readiness score /100']}${RL},MATCH({idr},'Master Scholarship Database'!$A${R0}:$A${RL},0)),\"\")")
body(d, 22, 29, 1, 5, {c: LINK for c in range(1, 6)})
priority_cf(d, "C22:C29")
d.conditional_formatting.add("E22:E29", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color="5B9BD5"))

d["A32"] = "Next 8 milestones"
d["A32"].font = Font(name=F, bold=True, size=11)
block_header(33, 1, ["Milestone", "Date", "Days", "Urgency"])
fut = [i for i, m in enumerate(MIL) if m[2] >= VERIFIED_ON][:8]
for k, idx in enumerate(fut):
    r, src = 34 + k, 5 + idx
    d.cell(row=r, column=1, value=f"='Application Deadlines'!B{src}")
    d.cell(row=r, column=2, value=f"='Application Deadlines'!C{src}").number_format = "dd-mmm-yyyy"
    d.cell(row=r, column=3, value=f"='Application Deadlines'!E{src}")
    d.cell(row=r, column=4, value=f"='Application Deadlines'!F{src}")
body(d, 34, 33 + len(fut), 1, 4, {c: LINK for c in range(1, 5)})
d.conditional_formatting.add(f"D34:D{33 + len(fut)}", CellIsRule(operator="equal", formula=['"URGENT"'], fill=RED))

d["A44"] = "Open action items"
d["A44"].font = Font(name=F, bold=True, size=11)
block_header(45, 1, ["Owner", "Open items"])
for i, own in enumerate(["YOU", "CLAUDE"]):
    d.cell(row=46 + i, column=1, value=own)
    d.cell(row=46 + i, column=2, value=f"=COUNTIFS('Action Items'!B5:B{AIL},A{46 + i},'Action Items'!F5:F{AIL},\"<>Done\")")
body(d, 46, 47, 1, 2)
d["A50"] = "Legend: blue = editable input | green = linked from another sheet | black = formula/fixed fact. Evidence E2 = official-domain extract; re-check live pages before submitting."
d["A50"].font = SUB

# Final sheet order: original 12 (spec order) + 3 new
order = ["Executive Dashboard", "Master Scholarship Database", "University Comparison", "Eligibility Assessment", "Application Deadlines",
         "Required Documents", "Recommendation Tracking", "Application Progress", "Financial Coverage", "Visa Preparation",
         "Interview Preparation", "Research Sources", "Discrepancy Log", "Action Items", "Risk Register"]
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
wb.active = 0
wb.save(OUT)
print("saved", OUT, "opps", N)
