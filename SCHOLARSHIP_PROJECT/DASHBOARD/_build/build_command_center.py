"""Builds DASHBOARD/Scholarship_Command_Center.xlsx (v3, replaces Scholarship_Tracker_v2).
Run:  cd DASHBOARD/_build && python3 build_command_center.py   then recalc with LibreOffice (recalc.py).
Evidence scale: E0 own document | E1 verified directly (official) | E2 official search result | E3 secondary | E4 your confirmation | E5 unsupported.
"""
import os, re
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as CL

VERIFIED = date(2026, 11, 7)
OUT = "../Scholarship_Command_Center.xlsx"
EVIDENCE_INDEX = "../../03_SCHOLARSHIP_RESEARCH/Official_Evidence_2026-11-07/INDEX.md"
F = "Arial"
HFILL = PatternFill("solid", fgColor="1F3864"); HFONT = Font(name=F, bold=True, color="FFFFFF", size=10)
BODY = Font(name=F, size=10); INP = Font(name=F, size=10, color="0000FF"); LNK = Font(name=F, size=10, color="008000")
TITLE = Font(name=F, bold=True, size=15, color="1F3864"); SUB = Font(name=F, italic=True, size=9, color="555555")
QFONT = Font(name=F, bold=True, size=11, color="1F3864")
GREEN = PatternFill("solid", fgColor="C6EFCE"); AMBER = PatternFill("solid", fgColor="FFEB9C"); RED = PatternFill("solid", fgColor="FFC7CE")
GREY = PatternFill("solid", fgColor="D9D9D9"); BLUEF = PatternFill("solid", fgColor="EAF0F8"); YEL = PatternFill("solid", fgColor="FFF2CC")
th = Side(style="thin", color="BFBFBF"); BD = Border(left=th, right=th, top=th, bottom=th)
WR = Alignment(wrap_text=True, vertical="top")

PRI = ["A", "B", "C"]
STAT = ["OPEN NOW", "OPENING SOON", "UPCOMING", "CLOSED", "DISCONTINUED", "ELIGIBILITY UNCERTAIN"]
ELIG = ["Eligible", "Potentially eligible", "Ineligible", "Eligibility requires university confirmation"]
FUND = ["FULLY FUNDED", "NEARLY FULLY FUNDED", "PARTIALLY FUNDED", "TUITION ONLY", "SELF-FUNDED"]
ENG = ["Not needed", "Not started", "Booked", "Done - meets", "Done - below"]
PROG = ["Not started", "In progress", "Done", "N/A"]

# ---------------------------------------------------------------------------------------------
# Opportunities. scores=(Academic, Funding, Relevance, Prof-exp, Flexibility, Feasibility) 0-10
# cyc: "2027" (apply now) | "2028" (next cycle) | "monitor"
# ---------------------------------------------------------------------------------------------
O = []
def opp(**k): O.append(k)
opp(id="S02", pri="A", cyc="2027", country="Sweden", host="SI-eligible Swedish university (choose 4 after 13 Nov)", prog="MSc computer systems / cybersecurity / software (TBC)",
    sch="SI Scholarship for Global Professionals", fund="FULLY FUNDED", cover="Full tuition + SEK 12,000/month + SEK 15,000 travel. No insurance; UA fee not covered",
    gpa_rule="None at SI level; universities rank applicants", gpa="Eligible", eng_req="English 6 met by an Ethiopian bachelor's", eng="Not needed",
    work=">= 3,000 h before 8 Feb 2027 from <= 3 employers + leadership (SI forms signed & stamped)", opens=date(2026, 10, 16), deadline=date(2027, 1, 15),
    tz="UA 15 Jan; SI portal 8-24 Feb 2027, closes 14:59 CET", status="OPEN NOW", url="https://si.se/en/apply/scholarships/swedish-institute-scholarships-for-global-professionals/",
    ev="E1", disc="Yes", my="After 13 Nov: pick 4 programmes; apply on universityadmissions.se; pay SEK 900; registrar sends transcript to UHR",
    cl="Map eligible programmes & selection criteria (13-14 Nov); SI CV + motivation after HR letters", risk="Medium",
    note="Ethiopia 2026: 793 applications, 13 awards (1.6%)", scores=(8, 9, 9, 10, 6, 8))
opp(id="S30", pri="A", cyc="2027", country="Estonia / Portugal / Finland", host="Tallinn Univ. (coord.), Lusofona Univ., Tampere Univ.", prog="AISS - AI for Sustainable Societies (Erasmus Mundus)",
    sch="Erasmus Mundus scholarship", fund="FULLY FUNDED", cover="Tuition waived + EUR 1,400/month for 24 months (10 scholarships)",
    gpa_rule="No GPA minimum, but the transcript is scored inside the documents part (50%)", gpa="Eligible", eng_req="IELTS (incl. Online) 6.5, >= 5.5 each; PTE 62, >= 59 each; TOEFL",
    eng="Not started", work="Portfolio of executed projects (up to 5)", opens=date(2026, 12, 2), deadline=date(2027, 2, 2), tz="DreamApply (Tallinn) - confirm closing time",
    status="OPENING SOON", url="https://aissprogram.eu/application-procedure/", ev="E1", disc="Yes", my="Build 2-3 public demo repos; book IELTS",
    cl="Video script, portfolio texts, interview mocks (package ready)", risk="Medium", note="Label B. Scholarships go to the 10 highest-ranked; max 10% per nationality (about 1 Ethiopian); min 35/50 per part; self-funding plan mandatory", scores=(8, 9, 7, 8, 8, 8))
opp(id="S01", pri="A", cyc="2027", country="Rwanda (Kigali), Africa", host="Carnegie Mellon University Africa", prog="MS Information Technology",
    sch="Mastercard Foundation Scholars Program", fund="FULLY FUNDED", cover="Only if selected (need-based): tuition, fees, insurance, travel, living",
    gpa_rule="No cut-off, but most admits >= 3.0: 2.62 is a weakness to offset (report as printed)", gpa="Eligible", eng_req="Test optional if applying by 15 Dec (E2); after that DET/IELTS/TOEFL required", eng="Not started",
    work="Not required; leadership & giving back valued", opens=date(2026, 10, 1), deadline=date(2026, 12, 15), tz="Not stated - submit by 12 Dec",
    status="OPEN NOW", url="https://www.africa.engineering.cmu.edu/admissions/index.html", ev="E2 (site blocked here); E1 MCF target group", disc="Yes",
    my="Open the site in a browser; create the portal account; copy essay prompts; enter 3 referees by 1 Dec", cl="Essays v2 ready; adapt to real prompts",
    risk="Medium", note="Label B. Three separate decisions: admission, need-based aid (form after applying), Mastercard Foundation Scholar selection", scores=(8, 7, 10, 8, 8, 7))
opp(id="S31", pri="A", cyc="2027", country="Finland / Spain / Italy / Sweden", host="Abo Akademi (coord.) + partners", prog="EDISS - Data-intensive Intelligent Software Systems (Erasmus Mundus)",
    sch="Erasmus Mundus + EDISS scholarships", fund="FULLY FUNDED", cover="EM scholarship if selected; later tuition-waiver round 25 Feb - 11 Mar 2027",
    gpa_rule="Academic merit = 70 of 100 points from grades; IT experience does not compensate", gpa="Eligible", eng_req="IELTS 6.5, no band < 6.0 (online not accepted); TOEFL 93 / new 4.5",
    eng="Not started", work="Not required", opens=date(2026, 11, 16), deadline=date(2027, 1, 12), tz="15:00 Helsinki time", status="OPENING SOON",
    url="https://www.master-ediss.eu/upcoming-application-rounds-ediss-intake-7/", ev="E1", disc="Yes",
    my="Ask 2 lecturers (ACADEMIC letters only); book IELTS (in-centre)", cl="Motivation video script + grading statement ready", risk="High",
    note="Label C (low probability): EDISS-area grade average 2.72. Apply in round 1 only if IELTS and letters are ready anyway", scores=(4, 9, 9, 7, 5, 7))
opp(id="S09", pri="A", cyc="2027", country="Hungary", host="ELTE / Univ. of Szeged", prog="MSc Computer Science", sch="Stipendium Hungaricum",
    fund="NEARLY FULLY FUNDED", cover="Tuition-free; HUF 43,700/month; dorm or HUF 40,000/month; insurance", gpa_rule="ELTE: entrance exam; Szeged: 100-pt score (>= 50)",
    gpa="Potentially eligible", eng_req="B2 (ELTE: any certificate; Szeged: IELTS 5.5 / TOEFL 72)", eng="Not started", work="Not required", opens=None, deadline=None,
    tz="2026/27 was 15 Jan, 14:00 CET", status="OPENING SOON", url="https://stipendiumhungaricum.hu/", ev="E1 (2026/27 terms; 2027/28 call pending)", disc="Yes",
    my="Ask MoE about nomination steps", cl="Watch for the 2027/28 call; adapt motivation letter", risk="Low", note="Stipend officially a partial contribution", scores=(8, 6, 8, 5, 8, 7))
opp(id="S10", pri="B", cyc="2027", country="Canada", host="University of Manitoba", prog="MSc Computer Science (thesis)", sch="Department funding (min CAD 20,000/yr x 2)",
    fund="PARTIALLY FUNDED", cover="CAD 20,000/yr minimum; gap vs proof of funds + tuition", gpa_rule="3.0 (B) over the last two years", gpa="Eligibility requires university confirmation",
    eng_req="IELTS 6.5", eng="Not started", work="Supervisor support letter BEFORE applying", opens=None, deadline=date(2026, 12, 1), tz="International deadline",
    status="OPEN NOW", url="https://umanitoba.ca/graduate-studies/admissions/programs-of-study/computer-science-msc", ev="E1", disc="Yes",
    my="Go/no-go by 12 Nov; email 3-5 professors if go", cl="Outreach kit ready", risk="High", note="Your last-two-years GPA 2.98 (E0 calc) vs 3.0", scores=(5, 6, 8, 6, 4, 4))
opp(id="S32", pri="B", cyc="2027", country="Canada", host="University of Calgary", prog="MSc Computer Science (thesis)", sch="Departmental awards / assistantships",
    fund="PARTIALLY FUNDED", cover="Funding mostly awarded Feb-Mar; guarantee not confirmed on the page read", gpa_rule="'Competitiveness of academic record including GPA'",
    gpa="Eligibility requires university confirmation", eng_req="Per FGS rules (check)", eng="Not started", work="List >= 1 willing supervisor", opens=date(2026, 11, 1),
    deadline=date(2027, 1, 15), tz="Early review 15 Jan; final 1 Mar 2027", status="OPEN NOW",
    url="https://science.ucalgary.ca/computer-science/future-students/graduate/admission-requirements", ev="E1", disc="Yes", my="Optional: contact 2-3 professors",
    cl="Reuse the Manitoba outreach kit", risk="High", note="Fee CAD 145", scores=(5, 5, 8, 6, 5, 6))
opp(id="S05", pri="B", cyc="2027", country="China", host="Ethiopian MoE (Type A) or Chinese university (Type B)", prog="MSc CS / Software Engineering (English-taught)",
    sch="Chinese Government Scholarship (CSC) 2027/28", fund="FULLY FUNDED", cover="Tuition, housing, CNY 3,000/month, insurance; airfare only in some Type A notices", gpa_rule="University-specific",
    gpa="Eligibility requires university confirmation", eng_req="English-taught: per university (IELTS or MOI letter); Chinese-taught: HSK 4", eng="Not started", work="Not required", opens=None, deadline=None,
    tz="Ethiopia call not yet published; other countries 20 Dec 2026 - 14 Feb 2027", status="UPCOMING", url="https://www.campuschina.org/", ev="E2 (2027/28 embassy notices, other countries)",
    disc="Yes", my="Ask MoE for call date + Type A agency number; apply to 3 universities for pre-admission letters", cl="Deep-dive + shortlist ready (China_CSC_2027-28_Deep_Dive_v1)", risk="Medium",
    note="Master's: under 35 (check privately); pre-admission letter required", scores=(6, 9, 8, 6, 6, 6))
opp(id="S08", pri="B", cyc="2027", country="South Korea", host="Embassy or University track", prog="MSc CS / IT", sch="GKS Graduate 2027", fund="FULLY FUNDED",
    cover="Airfare, tuition, allowance, language year", gpa_rule="80% / top 20% / CGPA threshold (2026: 2.64)", gpa="Ineligible", eng_req="Optional", eng="Not needed",
    work="Not required", opens=None, deadline=None, tz="2026 Ethiopia round: 12-25 Feb", status="UPCOMING", url="https://www.studyinkorea.go.kr/ko/notice/scholarshipsList.do?boardSort=3",
    ev="E1 (2027 guideline not yet published)", disc="No", my="Registrar rank letter: only useful if top 20%", cl="Re-check when the 2027 GKS-G notice appears", risk="High",
    note="2.62 < 2.64", scores=(2, 10, 8, 6, 2, 6))
opp(id="S11", pri="B", cyc="2027", country="Italy", host="Public universities", prog="MSc Computer Science / Cybersecurity", sch="Regional DSU grants", fund="NEARLY FULLY FUNDED",
    cover="Need-based (~EUR 8,000+/yr, secondary source)", gpa_rule="Per university", gpa="Potentially eligible", eng_req="Per programme", eng="Not started", work="Not required",
    opens=None, deadline=None, tz="2027/28 calls about mid-2027", status="UPCOMING", url="https://www.esu.pd.it", ev="E3", disc="Yes", my="-", cl="Verify in a later pass",
    risk="Medium", note="Requires family income documents", scores=(6, 7, 8, 5, 6, 6))
opp(id="S35", pri="B", cyc="2027", country="India", host="ICCR (Indian Council for Cultural Relations)", prog="Master's in IT / Engineering (Indian universities)", sch="ICCR scholarships (A2A portal)",
    fund="FULLY FUNDED", cover="Per ICCR scheme (verify on portal)", gpa_rule="Indian university admission", gpa="Eligibility requires university confirmation", eng_req="Per university", eng="Not needed", work="-",
    opens=date(2027, 2, 27), deadline=date(2027, 4, 15), tz="2026-27 ran 27 Feb - 15 Apr 2026 (Ethiopia); 2027-28 dates expected similar", status="UPCOMING", url="https://a2ascholarships.iccr.gov.in/", ev="E2 (Indian Embassy Addis 2026-27 notice)",
    disc="No", my="Choose 5 Indian universities with MCA/M.Tech CS when the portal opens", cl="Shortlist universities", risk="Medium", note="130 places for Ethiopians in 2026-27. Dates are expected, not announced", scores=(6, 9, 7, 5, 7, 7))
opp(id="S13", pri="B", cyc="2028", country="United Kingdom", host="UK universities (3 choices)", prog="MSc Cloud / Cyber / Software", sch="Chevening 2028/29", fund="FULLY FUNDED",
    cover="Tuition, stipend, airfare, visa", gpa_rule="Unconditional UK offer required", gpa="Eligible", eng_req="Per university", eng="Not started", work=">= 2,800 h AFTER the bachelor's, over >= 2 years (E1 chevening.org, 8 Nov 2026). Documented today: 0 h (HR letters pending)",
    opens=None, deadline=None, tz="2027/28 window was 4 Aug - 6 Oct 2026; 2028/29 window expected Aug-Oct 2027", status="CLOSED", url="https://www.chevening.org/scholarships/who-can-apply/work-experience/",
    ev="E1 (cycle dates GOV.UK; work rule chevening.org)", disc="Yes", my="Prove 2,800 h with HR letters; plan for Aug 2027", cl="Chevening plan by Jul 2027", risk="Low",
    note="FUTURE-CYCLE TARGET (2028/29). Internships and voluntary work also count for Chevening (E1). 2-year return to Ethiopia after the award", scores=(7, 10, 8, 10, 6, 2), ver=date(2026, 11, 8))
opp(id="S14", pri="B", cyc="2028", country="United States", host="Fulbright via US Embassy Addis", prog="MS CS / IT", sch="Fulbright Foreign Student 2028-29", fund="FULLY FUNDED",
    cover="Tuition, stipend, airfare, health cover", gpa_rule="'Good academic record'", gpa="Eligibility requires university confirmation", eng_req="TOEFL iBT 79-80 / IELTS 6.5",
    eng="Not started", work="Valued", opens=None, deadline=None, tz="Ethiopia page blocked here", status="UPCOMING", url="https://foreign.fulbrightonline.org/apply",
    ev="E1 (general rules)", disc="Yes", my="Check the US Embassy page in a browser early 2027", cl="-", risk="Medium", note="J-1 two-year home residency", scores=(7, 10, 8, 8, 6, 3))
# ---- Priority C / eliminated ----
# ---- Phase 5 P2: new opportunities (7 Nov 2026) ----
def n(id, cyc, country, host, prog, sch, fund, cover, gpa_rule, gpa, deadline, tz, status, url, ev, my, note, scores, opens=None, eng_req="Per programme", risk="High"):
    opp(id=id, pri="B", cyc=cyc, country=country, host=host, prog=prog, sch=sch, fund=fund, cover=cover, gpa_rule=gpa_rule, gpa=gpa, eng_req=eng_req, eng="Not started",
        work="-", opens=opens, deadline=deadline, tz=tz, status=status, url=url, ev=ev, disc="No", my=my, cl="See New_Opportunities_2027_v1", risk=risk, note=note, scores=scores)
n("S37", "monitor", "Belgium", "Flemish universities", "ICP master's (12 programmes; ICT/data TBC)", "VLIR-UOS ICP Connect", "FULLY FUNDED", "Tuition, living, lodging, insurance, travel",
  "Programme-specific", "Potentially eligible", date(2027, 2, 28), "Programme deadlines may be earlier", "OPENING SOON", "https://www.vliruos.be/", "E2/E3", "Mid-Nov: check country + programme list", "Label B/C: needs an ICT/data programme", (6, 10, 6, 6, 6, 5))
n("S38", "2027", "Italy", "Participating Italian universities", "Engineering / high-tech master's", "Invest Your Talent in Italy (IYT)", "NEARLY FULLY FUNDED", "Tuition waiver + EUR 10,800 / 9 months",
  "Admission to participating master's", "Potentially eligible", None, "2027-28 call expected spring 2027", "UPCOMING", "https://investyourtalentapplication.esteri.it/", "E2 (official FAQ: Ethiopia eligible)", "Shortlist ICT master's that accept 2.62", "Label B/C", (6, 8, 8, 6, 6, 6))
n("S39", "2027", "Finland / Spain / Belgium or Greece", "Univ. of Vaasa, UPV/EHU Bilbao, UMONS or IHU Thessaloniki", "Smart Cities and Communities (Erasmus Mundus)", "Erasmus Mundus scholarship", "FULLY FUNDED", "EUR 1,400/month + fees (top-ranked)",
  "Global merit: academic record + motivation + WORK experience + R&D + 2 references", "Eligible", date(2027, 2, 1), "Scholarship round (E3 date)", "OPENING SOON", "https://www.smaccs.eu/", "E2 (FAQ) / E3 (date)", "Third Erasmus Mundus slot", "Label B: Ethiopia eligible for regular + NDICI scholarships; ICT track at IHU; work experience scored", (5, 9, 8, 7, 5, 6))
n("S40", "2027", "Austria / France / Czechia (consortium)", "CDE consortium", "Copernicus Master in Digital Earth (Erasmus Mundus)", "Erasmus Mundus scholarship", "FULLY FUNDED", "EUR 1,400/month + fee waiver (best-ranked)",
  "Holistic; needs Earth-observation/GIS evidence", "Potentially eligible", date(2027, 1, 22), "23:59 CET; call open 2 Nov 2026", "OPEN NOW", "https://master-cde.eu/", "E2 (official site)", "Only with real geodata work", "Label C", (4, 9, 4, 5, 4, 5), opens=date(2026, 11, 2))
n("S41", "monitor", "Turkiye / partners", "Kadir Has Univ. (coord.)", "CyberMACS Applied Cybersecurity (Erasmus Mundus)", "Erasmus Mundus scholarship", "FULLY FUNDED", "If selected", "Unclear (test + interview reported)",
  "Potentially eligible", None, "2027 intake not confirmed", "ELIGIBILITY UNCERTAIN", "http://cybermacs.eu/", "E3", "Check monthly in a browser", "Label C (monitor)", (4, 9, 9, 6, 4, 4))
n("S42", "2028", "Czech Republic", "Czech public universities", "English-taught master's", "Czech Government Scholarships (developing countries)", "FULLY FUNDED", "Tuition-free + stipend",
  "Completed bachelor's", "Potentially eligible", None, "2027/28 closed 30 Sep 2026; 2028/29 call expected Aug-Sep 2027", "CLOSED", "https://msmt.gov.cz/en/scholarships/government-scholarships-developing-countries", "E2 (official: Ethiopia on 2027/28 list)", "Reminder Aug 2027", "Label B (next cycle)", (6, 9, 7, 5, 6, 6))
n("S43", "2027", "China", "DUT (first) / CSU / NUIST / UESTC; HUST on hold; NPU excluded", "English-taught CS / SE / cybersecurity master's", "CSC Chinese University Program (Type B)", "FULLY FUNDED", "Tuition, housing, CNY 3,000/month, insurance; airfare usually not",
  "Set by each university", "Eligibility requires university confirmation", date(2027, 2, 15), "DUT 15 Feb; HUST 25 Feb; NPU 28 Feb 2027 17:00 UTC+8", "OPEN NOW", "https://sie.dlut.edu.cn/info/1180/2485.htm", "E2 (university pages)", "Apply to 3 universities (pre-admission letters)", "Label B/C by university; under 35 (check privately)", (5, 9, 8, 6, 6, 6), opens=date(2026, 11, 1), risk="Medium")
n("S44", "2027", "Taiwan", "NTUT and other universities", "Master's in CS / EE", "MOE Taiwan Scholarship / university scholarships", "PARTIALLY FUNDED", "University: tuition + NT$6,000/month; MOE: tuition + NT$20,000/month",
  "University-specific", "Eligibility requires university confirmation", None, "Usually Feb-Mar", "UPCOMING", "https://oia.ntut.edu.tw/p/412-1032-13793.php?Lang=en", "E2 (NTUT official)", "Ask NTUT which route applies to Ethiopians", "Label C; MOE normally via Taiwan missions (none in Ethiopia)", (5, 5, 7, 5, 6, 5))
n("S45", "monitor", "Brunei", "Universiti Brunei Darussalam", "Master's", "Brunei Darussalam Government Scholarship", "FULLY FUNDED", "Full (reported)", "English + academic record", "Eligibility requires university confirmation",
  None, "Call expected mid-Dec 2026", "ELIGIBILITY UNCERTAIN", "https://www.mofa.gov.bn/", "E3", "Check the call: is Ethiopia eligible?", "Label C (monitor)", (4, 9, 5, 5, 4, 4))
n("S46", "monitor", "Thailand", "Thai universities", "Master's (TIPP list)", "Thailand International Postgraduate Programme (TICA)", "FULLY FUNDED", "Full (reported)", "Government nomination",
  "Eligibility requires university confirmation", None, "Around Feb; via national focal point", "ELIGIBILITY UNCERTAIN", "https://tica-thaigov.mfa.go.th/", "E2", "Ask MoE whether Ethiopia nominates", "Label C (monitor)", (4, 9, 5, 5, 4, 4))
n("S47", "2028", "Japan", "Japanese universities", "Research student -> master's", "MEXT (Embassy recommendation)", "FULLY FUNDED", "Full", "Embassy screening; university route needs MEXT GPA >= 2.30 (yours about 1.5)",
  "Potentially eligible", None, "Next call expected spring 2027 (for 2028)", "CLOSED", "https://www.et.emb-japan.go.jp/", "E3", "Check Japanese Embassy Addis in Mar-Apr 2027", "Label C (next cycle); age limit: check privately", (3, 10, 6, 5, 4, 5))
n("S48", "monitor", "Rwanda", "University of Rwanda", "2-year master's (data/ICT TBC)", "Mastercard Foundation Scholars (UR)", "FULLY FUNDED", "Full", "UR admission criteria",
  "Potentially eligible", None, "2026-27 call closed; 2027-28 call expected about Jul 2027", "CLOSED", "https://ur.ac.rw/", "E2/E3", "Check ur.ac.rw in June 2027", "Label B/C; any African citizen", (5, 10, 7, 6, 6, 5))
n("S49", "2027", "Mauritius", "University of Mauritius (public)", "MPhil (research) in computing", "Mauritius-Africa Scholarship Scheme (MASS)", "NEARLY FULLY FUNDED", "Tuition + living contribution + return airfare",
  "Application to a Mauritian public institution", "Potentially eligible", date(2027, 2, 28), "Feb-Mar 2027 via national nominating agency (dates vary)", "OPEN NOW", "https://tertiaryeducation.govmu.org/te/index.php/mass/", "E2 (official site)", "Ask MoE if it is the nominating agency", "Label C (MPhil/PhD routes)", (4, 8, 5, 5, 4, 5))
n("S50", "monitor", "Africa (EU-funded)", "Project consortia", "Computing master's (if a project offers one)", "Intra-Africa Academic Mobility Scheme", "FULLY FUNDED", "Within each project", "Project-specific",
  "Potentially eligible", None, "Each project has its own call", "ELIGIBILITY UNCERTAIN", "https://www.eacea.ec.europa.eu/", "E2", "Ask DTU international office", "Label C (monitor)", (4, 9, 6, 5, 4, 4))
n("S51", "2027", "Canada", "University of Saskatchewan", "MSc Computer Science (thesis)", "Departmental funding (automatic consideration)", "PARTIALLY FUNDED", "Amount not confirmed (one site: about CAD 20,000/yr)",
  "No published minimum for the standard MSc", "Potentially eligible", None, "Check department page", "UPCOMING", "https://www.cs.usask.ca/students/graduate/graduate-programs/masters.php", "E2/E3", "Only if Manitoba goes ahead; email a professor", "Label C", (4, 6, 8, 6, 5, 4), eng_req="Required for international applicants")
n("S52", "2028", "Australia", "Australian universities", "Master's in a priority area", "Australia Awards Africa", "FULLY FUNDED", "Full", "Work experience; ICT not a priority area",
  "Potentially eligible", None, "Next round opens 1 Feb 2027 (for 2028)", "CLOSED", "https://australiaawardsafrica.org/", "E2 (DFAT/AAA)", "Only if a priority area fits truthfully", "Label C (next cycle)", (5, 10, 4, 6, 5, 4))


n("S61", "2028", "United States (Tempe, Arizona)", "Arizona State University (ASU)", "MS Information Technology / MS Global Technology and Development (E2 list)",
  "Mastercard Foundation Scholars Program at ASU - Phase III: Innovation and Technology", "FULLY FUNDED", "Full scholarship (confirm details on mcfscholars.asu.edu)",
  "ASU graduate admission: 3.00 in the last 60 credit hours (E3); your years 4-5 average 2.98, final year 3.20 (E0 calculation)", "Eligibility requires university confirmation", None,
  "2027/28 CLOSED: deadline was 27 Sep 2026 (E2); next cycle expected about Sep 2027 (assumption)", "CLOSED",
  "https://mcfscholars.asu.edu/phase-three-innovation-and-technology", "E1 (MCF page, 8 Nov 2026) / E2 (deadline, programmes)",
  "Monitor mcfscholars.asu.edu from Jun 2027", "Closed for 2027/28 (your information; E2 deadline). The MCF hub page still says 'open' (stale). IT-focused phase: the strongest US option next cycle. "
  "One third-party site reports an age limit (E3): check privately", (6, 10, 9, 8, 5, 4), eng_req="ASU English requirement (IELTS/TOEFL)", risk="Medium")
for _o in O:
    if _o["id"] in ("S22", "S61"): _o["ver"] = date(2026, 11, 8)
def c(id, country, host, prog, sch, reason, ev, status="ELIGIBILITY UNCERTAIN", gpa="Ineligible", fund="FULLY FUNDED", url="", scores=(1, 9, 6, 5, 1, 1), deadline=None):
    opp(id=id, pri="C", cyc="eliminated", country=country, host=host, prog=prog, sch=sch, fund=fund, cover="-", gpa_rule=reason, gpa=gpa, eng_req="-", eng="Not needed",
        work="-", opens=None, deadline=deadline, tz="-", status=status, url=url, ev=ev, disc="No", my="-", cl="None", risk="High", note=reason, scores=scores)
c("S04", "Finland / Italy / Netherlands", "LUT + partners", "SE4GD (Erasmus Mundus)", "EM scholarship", "Programme ended Aug 2026; no new intakes", "E1", status="DISCONTINUED", gpa="Potentially eligible", url="https://se4gd.lutsoftware.com/")
c("S03", "Norway + partners", "NTNU (coord.)", "CYBERSURE (Erasmus Mundus)", "EM scholarship", "First screen: 'very high weighted average grade or GPA'; deadline 4 Jan 2027", "E1", status="OPENING SOON", url="https://cybersure-master.eu/admission", deadline=date(2027, 1, 4), scores=(2, 9, 9, 6, 2, 6))
c("S36", "Finland + partners", "Aalto (coord.)", "SECCLO (Erasmus Mundus)", "EM scholarship", "Appears superseded by CYBERSURE; site not verifiable", "E3", url="https://secclo.eu")
c("S12", "Turkiye", "Turkish universities", "MSc CS", "Turkiye Burslari", "75% academic minimum (2.62/4.00 = 65.5% on a linear conversion); age rule checked privately, not used here", "E1", status="UPCOMING", url="https://www.turkiyeburslari.gov.tr/scholarshipsprograms", deadline=date(2027, 2, 20))
c("S33", "Poland", "NAWA", "Master's", "Poland My First Choice / Lukasiewicz", "Ethiopia not eligible; Lukasiewicz programme closed", "E1", status="CLOSED", url="https://nawa.gov.pl/en")
c("S34", "Japan", "JICA", "Master's + internship", "ABE Initiative", "Official pages removed; current schemes mostly for nominated officials", "E1 (404)", gpa="Eligibility requires university confirmation", url="https://www.jica.go.jp/english/")
n("S22", "2028", "Germany", "EPOS-listed German universities (ICT/IS courses to confirm)", "Development-related master's (prioritise IT / information systems / ICT4D)",
  "DAAD EPOS (Development-Related Postgraduate Courses)", "FULLY FUNDED", "Monthly stipend, travel, insurance (DAAD rates; confirm per course)",
  "Upper third of your class (E2/E3) + >= 2 years related work after the bachelor's + degree <= 6 years old", "Eligibility requires university confirmation", None,
  "Course deadlines mostly Aug-Oct 2027 for 2028 entry; your degree turns 6 on 14 Jul 2028, so later windows are likely closed", "UPCOMING",
  "https://www.daad.de/en/studying-in-germany/scholarships/daad-scholarships/development-related-postgraduate-courses-epos/", "E2/E3 (daad.de blocked here)",
  "Ask the DTU Registrar for a class-rank letter (request ready)", "CONDITIONAL: realistic only if the DTU rank letter shows the upper third. 2.62 alone does not show it", (4, 10, 7, 9, 4, 4),
  eng_req="Per course (often IELTS 6.0-6.5)", risk="High")
c("S26", "Saudi Arabia", "KAUST", "MS Computer Science", "KAUST Fellowship", "Minimum GPA 3.0", "E2", url="https://admissions.kaust.edu.sa")
c("S27", "UAE", "MBZUAI", "MSc AI", "MBZUAI funding", "Minimum CGPA 3.0 (some 3.5)", "E2", url="https://mbzuai.ac.ae")
c("S28", "Africa (AU)", "Pan African University", "Master's", "AU scholarship", "Master's age limit applies (E2, check privately); call closed", "E2", status="CLOSED", url="https://au.int")
c("S29", "Rwanda / Senegal", "AIMS", "AMMI", "Full scholarship", "Minimum about B / 80%", "E2", url="https://aimsammi.org")
c("S17", "Canada (Vancouver)", "University of British Columbia", "Forestry / Land and Food Systems master's only", "Mastercard Foundation Scholars Program at UBC (Phase 2)",
  "CLOSED for 2027: expression of interest closed 6 Sep 2026 (E2). Phase 2 is limited to forestry, agriculture, water and land management (E1 MCF page, 8 Nov 2026); "
  "young Africans with financial need who meet UBC admission. No computer-science route", "E1 (field focus) / E2 (dates)", status="CLOSED",
  gpa="Eligibility requires university confirmation", url="https://mastercardfdn.org/en/what-we-do/our-programs/mastercard-foundation-scholars-program/where-to-apply/university-of-british-columbia/")
c("S62", "Canada (Montreal)", "McGill University", "Master's in nutrition, public health, public policy, sustainable agriculture", "Mastercard Foundation Scholars Program at McGill",
  "CLOSED for Fall 2027: application period closed (McGill page via search, E2; deadline 21 Aug 2026, E2). Fields: nutrition, public health, public policy, "
  "sustainable agriculture (E1 MCF page, 8 Nov 2026). African citizens/residents; mandatory info session; MCF pre-screening before the McGill application. No computer-science route",
  "E1 (field focus) / E2 (dates)", status="CLOSED", gpa="Eligibility requires university confirmation",
  url="https://mastercardfdn.org/en/what-we-do/our-programs/mastercard-foundation-scholars-program/where-to-apply/mcgill-university/")
c("S20", "United Kingdom", "Edinburgh / Manchester", "Master's", "Mastercard Foundation / Equity & Merit", "Closed for 2027; Edinburgh has no CS programmes", "E2", status="CLOSED", gpa="Potentially eligible")
c("S23", "Germany", "Public universities", "MSc CS", "None (self-funded)", "Not funded (blocked account EUR 11,904/yr)", "E2", status="UPCOMING", gpa="Eligibility requires university confirmation", fund="SELF-FUNDED")
c("S24", "United States", "Stanford", "Any", "Knight-Hennessy", "Not competitive at 2.62", "E3", status="CLOSED")
c("S25", "United Kingdom", "UK universities", "Any", "Commonwealth scholarships", "Ethiopia is not a Commonwealth member", "Fact", status="CLOSED")
c("S06", "France / Belgium / Lux.", "UBS + partners", "CYBERUS (Erasmus Mundus)", "EM scholarship", "No 2027 scholarship call confirmed", "E3")
c("S07", "Belgium / Spain", "ULB + partners", "DEAI (Erasmus Mundus)", "EM scholarship", "Maths-heavy; competitive on grades", "E3", deadline=date(2027, 1, 11), gpa="Eligibility requires university confirmation")
c("S53", "Ireland", "Irish universities", "One-year master's", "Ireland Fellows Programme (Africa)", "Closed 26 Jul 2026; requires GPA 3.0", "E2", status="CLOSED", url="https://www.ireland.ie/en/ireland-fellows-programme/")
c("S54", "Saudi Arabia", "KFUPM", "MS Computer Science", "KFUPM graduate scholarship", "Scholarship applicants need GPA 3.00", "E2", url="https://cgis.kfupm.edu.sa/admission/admissionrequirement")
c("S55", "Canada", "Memorial University", "MSc Computer Science", "Departmental funding", "Minimum average 75% or upper second", "E2", fund="PARTIALLY FUNDED", url="https://www.mun.ca/become/graduate/programs-and-courses/computer-science/")
c("S56", "Canada", "University of Regina", "MSc Computer Science", "Departmental funding", "Overall >= 70% AND CS >= 75%; below B-level rejected", "E2", fund="PARTIALLY FUNDED", url="https://www.uregina.ca/science/computer-science/graduate-programs/msc-cs-thesis.html")
c("S57", "New Zealand", "NZ universities", "Master's", "Manaaki New Zealand Scholarships", "2027: Pacific and ASEAN countries only", "E2", status="CLOSED", url="https://www.nzscholarships.govt.nz/")
c("S58", "Japan", "Japanese universities", "Master's", "MEXT (University recommendation)", "MEXT-scale GPA about 1.5 < 2.30", "E0 calculation", url="https://www.studyinjapan.go.jp/")
c("S59", "China", "Northwestern Polytechnical University", "English-taught master's", "CSC Type B (NPU)", "CSC applicants need an 80% average (2.62/4.00 = 65.5%)", "E2", url="https://studyat.nwpu.edu.cn/info/1303/8791.htm")
c("S60", "China", "Huazhong Univ. of Science and Technology", "MSc Computer Science and Technology (English)", "CSC Type B (HUST)", "On hold: 2024 guide minimum 3.0/4 or 75%; 2027 rule unconfirmed; fee CNY 600/800", "E2", status="ELIGIBILITY UNCERTAIN", url="https://iso.hust.edu.cn/info/1194/5146.htm")

for _o in O:
    if _o["id"] in ("S17", "S62", "S22"): _o["ver"] = date(2026, 11, 8)
N = len(O)
wb = Workbook()

def header(ws, row, cols, widths=None, height=34):
    for i, c_ in enumerate(cols, 1):
        x = ws.cell(row=row, column=i, value=c_); x.font, x.fill, x.border = HFONT, HFILL, BD
        x.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    if widths:
        for i, w in enumerate(widths, 1): ws.column_dimensions[CL(i)].width = w
    ws.row_dimensions[row].height = height

def body(ws, r1, r2, c1, c2, fonts=None):
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for x in row:
            x.border = BD; x.alignment = WR; x.font = (fonts or {}).get(x.column, BODY)

def title(ws, t, s):
    ws["A1"], ws["A2"] = t, s; ws["A1"].font, ws["A2"].font = TITLE, SUB

def dv(ws, opts, rng):
    d = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=True); ws.add_data_validation(d); d.add(rng)

def pri_cf(ws, rng):
    for v, f in (("A", GREEN), ("B", AMBER), ("C", RED)):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{v}"'], fill=f))

# ============================ Discrepancy Log ============================
DISC = [
    ("DR-01", "CGPA: transcript 2.62 is final (earlier self-reported figure withdrawn)", "CRITICAL", "Resolved", "E0 transcript + your confirmation 8 Nov 2026", "2.62 / 4.00", "-"),
    ("DR-02", "Study start date / calendar", "MODERATE", "Partly resolved", "E0 transcript", "AcYear 2017/18-2021/22 (Gregorian); graduated 14 Jul 2022", "You: first-semester month"),
    ("DR-03", "EQF level / Extension mode", "LOW", "Resolved", "E0", "EQF 6; 'Extension programme, 5 academic years'", "-"),
    ("DR-04", "Faculty naming", "LOW", "Resolved", "E0", "As on transcript", "-"),
    ("DR-05", "Name spelling", "BLOCKING", "Open", "Passport (E4 form)", "Pending", "You: Personal Information form"),
    ("DR-06", "SantimPay titles & dates (4 versions)", "BLOCKING", "Open", "SantimPay HR certificate", "Pending", "You: send SantimPay_HR_Request"),
    ("DR-07", "Kifiya dates (3 versions) + overlap", "BLOCKING", "Open", "Kifiya HR certificate", "Pending", "You: send Kifiya_HR_Request"),
    ("DR-08", "Plan International (duplicate text)", "CRITICAL", "Excluded", "-", "Excluded (E5)", "You: remove from LinkedIn"),
    ("DR-09", "Yandex (no evidence)", "HIGH", "Excluded", "-", "Excluded (E5)", "-"),
    ("DR-10", "Infonans / Gebeya / Dereja / GAO Tek", "MODERATE", "Open", "Letters or E4", "Pending", "You: Personal Information form"),
    ("DR-11", "Total experience", "BLOCKING", "Open", "HR certificates", "Pending", "Follows DR-06/07"),
    ("DR-12", "Certifications without IDs", "MODERATE", "Open", "Verification links", "Pending", "You: send IDs"),
    ("DR-13", "Tools changed between CV versions", "BLOCKING", "Open", "E4 + manager", "Pending", "You: form section 5"),
    ("DR-14", "2024 lecturer letter overstates grades", "HIGH", "Resolved", "E0 transcript", "Do not use; request new letters", "You: Lecturer_Recommendation_Request"),
    ("DR-15", "2024 Italy statement of intent", "LOW", "Open", "E4", "Do not reuse", "You: was it submitted?"),
    ("DR-16", "Another person's documents in Drive", "ADMIN", "Resolved", "-", "Excluded", "-"),
    ("DR-17", "Final-year project title missing", "MODERATE", "Open", "Department letter + E4", "Grades A-/A- final", "You: Registrar pack (department letter)"),
]
dl = wb.active; dl.title = "Discrepancy Log"
title(dl, "Discrepancy Log (full detail: 00_MASTER_PROFILE/Master_Profile_and_Discrepancy_Register_v3.md)", "BLOCKING items must be Resolved before the CV/essays are final. Status = blue dropdown.")
header(dl, 4, ["ID", "Issue", "Severity", "Status", "Authoritative evidence", "Current value", "Who acts / next step", "Open YOU item #"], [8, 40, 12, 15, 26, 36, 36, 10])
for i, d in enumerate(DISC):
    r = 5 + i
    for c_, v in enumerate(d, 1): dl.cell(row=r, column=c_, value=v)
    dl.cell(row=r, column=8, value=f'=IF(AND(LEFT(G{r},3)="You",D{r}<>"Resolved",D{r}<>"Excluded"),COUNTIFS($G$5:G{r},"You*",$D$5:D{r},"<>Resolved",$D$5:D{r},"<>Excluded"),"")')
DL = 4 + len(DISC)
body(dl, 5, DL, 1, 8, {4: INP})
dv(dl, ["Open", "Partly resolved", "Resolved", "Excluded"], f"D5:D{DL + 20}")
for v, f in (("Open", RED), ("Partly resolved", AMBER), ("Resolved", GREEN), ("Excluded", GREY)):
    dl.conditional_formatting.add(f"D5:D{DL}", CellIsRule(operator="equal", formula=[f'"{v}"'], fill=f))
dl.freeze_panes = "B5"; dl.auto_filter.ref = f"A4:H{DL}"
BLOCKED = f"COUNTIFS('Discrepancy Log'!$C$5:$C${DL},\"BLOCKING\",'Discrepancy Log'!$D$5:$D${DL},\"<>Resolved\")"

# ============================ Financial Coverage ============================
fc = wb.create_sheet("Financial Coverage")
title(fc, "Financial Coverage (Yes / Partial / No / UNVERIFIED)", "Completeness = (Yes + 0.5 x Partial) / 9 cost items. Never assume airfare, insurance or visa costs without evidence.")
items = ["Tuition", "Monthly living", "Accommodation", "Return airfare", "Health insurance", "Visa costs", "Settlement", "Books", "Research"]
header(fc, 4, ["ID", "Opportunity", "Funding class"] + items + ["Yes", "Partial", "Completeness", "Notes (evidence)"], [7, 38, 18] + [11] * 9 + [6, 8, 12, 44])
COV = {"S02": ["Yes", "Yes", "Partial", "Partial", "No", "UNVERIFIED", "No", "No", "No", "E1: SEK 12,000/month + SEK 15,000 travel; no insurance"],
       "S30": ["Yes", "Yes", "Partial", "Partial", "Yes", "UNVERIFIED", "No", "No", "No", "E1: tuition waived + EUR 1,400/month (covers insurance, travel, housing, food)"],
       "S01": ["Yes", "Yes", "UNVERIFIED", "Yes", "Yes", "UNVERIFIED", "UNVERIFIED", "UNVERIFIED", "Yes", "E2: only if selected as Mastercard Foundation Scholar"],
       "S31": ["Yes", "Yes", "Partial", "Partial", "UNVERIFIED", "Partial", "UNVERIFIED", "No", "No", "E2: standard EM scholarship terms; confirm on EDISS page"],
       "S09": ["Yes", "Partial", "Yes", "No", "Yes", "No", "No", "No", "No", "E1: HUF 43,700/month + dorm or HUF 40,000"],
       "S10": ["Partial", "Partial", "No", "No", "No", "No", "No", "No", "Partial", "E1: min CAD 20,000/yr x 2"],
       "S05": ["Yes", "Yes", "Yes", "Partial", "Yes", "No", "UNVERIFIED", "No", "No", "E2: CNY 3,000/month; airfare only in some Type A notices"],
       "S13": ["Yes", "Yes", "Partial", "Yes", "UNVERIFIED", "Yes", "Yes", "UNVERIFIED", "UNVERIFIED", "E2 (next cycle)"]}
r = 5
for o in O:
    if o["id"] not in COV: continue
    v = COV[o["id"]]
    fc.cell(row=r, column=1, value=o["id"]); fc.cell(row=r, column=2, value=f'{o["sch"]} - {o["prog"]}'); fc.cell(row=r, column=3, value=o["fund"])
    for j in range(9): fc.cell(row=r, column=4 + j, value=v[j])
    fc.cell(row=r, column=13, value=f'=COUNTIF(D{r}:L{r},"Yes")'); fc.cell(row=r, column=14, value=f'=COUNTIF(D{r}:L{r},"Partial")')
    fc.cell(row=r, column=15, value=f"=(M{r}+0.5*N{r})/9").number_format = "0%"; fc.cell(row=r, column=16, value=v[9]); r += 1
FCL = r - 1
body(fc, 5, FCL, 1, 16, {c_: INP for c_ in range(4, 13)})
dv(fc, ["Yes", "Partial", "No", "UNVERIFIED"], f"D5:L{FCL + 10}")
for v, f in (("Yes", GREEN), ("No", RED), ("UNVERIFIED", AMBER)):
    fc.conditional_formatting.add(f"D5:L{FCL}", CellIsRule(operator="equal", formula=[f'"{v}"'], fill=f))
g = FCL + 3
fc.cell(row=g, column=1, value="Funding-gap calculator (official figures; yellow = inputs to confirm)").font = QFONT
header(fc, g + 1, ["Case", "Item", "Amount", "Currency", "Source"])
GAP = [("Canada - Manitoba", "IRCC proof of funds (living)", 22895, "CAD", "IRCC from 1 Sep 2025 (check 2026 figure)"),
       ("Canada - Manitoba", "International tuition per year (INPUT)", None, "CAD", "Enter from the UM fees page"),
       ("Canada - Manitoba", "Department minimum funding per year", 20000, "CAD", "E1 umanitoba.ca"),
       ("Canada - Manitoba", "Year-1 gap vs IRCC test", "=C{a}+C{b}-C{c}", "CAD", "Formula (blank tuition understates the gap)"),
       ("Hungary", "Monthly stipend", 43700, "HUF", "E1 SH call 2026/27"), ("Hungary", "Housing contribution if no dorm", 40000, "HUF", "E1"),
       ("Hungary", "Total monthly support (no dorm)", "=C{e}+C{f}", "HUF", "Formula"),
       ("UK (Chevening 2028)", "Student visa maintenance outside London / month", 1171, "GBP", "E1 gov.uk"),
       ("English tests", "IELTS Academic (BC Addis Ababa University)", 32800, "ETB", "E1 ielts.org"),
       ("English tests", "TOEFL iBT (Ethiopia)", 185, "USD", "E1 ets.org fee table")]
st = g + 2; refs = {k: st + i for i, k in enumerate("abcdefghij")}
for k, (cs, it, amt, cur, src) in enumerate(GAP):
    rr = st + k
    for c_, v in enumerate([cs, it, amt.format(**refs) if isinstance(amt, str) else amt, cur, src], 1): fc.cell(row=rr, column=c_, value=v)
    fc.cell(row=rr, column=3).number_format = "#,##0;(#,##0);-"
body(fc, st, st + len(GAP) - 1, 1, 5)
for k, row in enumerate(GAP):
    if not isinstance(row[2], str): fc.cell(row=st + k, column=3).font, fc.cell(row=st + k, column=3).fill = INP, YEL

# ============================ Application Progress ============================
ap = wb.create_sheet("Application Progress")
title(ap, "Application Progress (2027-cycle applications)", "Done / In progress / Not started / N/A. Completion = Done / (steps not N/A). Only YOU create accounts, pay or submit.")
steps = ["Eligibility confirmed", "Portal account (you)", "Documents ready", "CV final", "Essays / letters / video", "References", "English test", "Fee / waiver", "Submitted (you)"]
header(ap, 4, ["ID", "Opportunity", "Priority"] + steps + ["Completion %", "Admission status", "Scholarship status"], [7, 42, 8] + [12] * 9 + [12, 14, 14])
track = [o for o in O if o["cyc"] == "2027" and o["pri"] in ("A", "B")]
ns = len(steps)
for i, o in enumerate(track):
    r = 5 + i
    ap.cell(row=r, column=1, value=o["id"]); ap.cell(row=r, column=2, value=f'{o["sch"]} - {o["prog"]}'); ap.cell(row=r, column=3, value=o["pri"])
    vals = ["Done" if o["gpa"] == "Eligible" else "In progress"] + ["Not started"] * (ns - 1)
    if o["eng"] == "Not needed": vals[6] = "N/A"
    for j, v in enumerate(vals): ap.cell(row=r, column=4 + j, value=v)
    lc = CL(3 + ns)
    ap.cell(row=r, column=4 + ns, value=f'=IFERROR(COUNTIF(D{r}:{lc}{r},"Done")/({ns}-COUNTIF(D{r}:{lc}{r},"N/A")),0)').number_format = "0%"
    ap.cell(row=r, column=5 + ns, value="Not started"); ap.cell(row=r, column=6 + ns, value="Not started")
APL = 4 + len(track); CMP = CL(4 + ns)
body(ap, 5, APL, 1, 6 + ns, {c_: INP for c_ in list(range(4, 4 + ns)) + [5 + ns, 6 + ns]})
dv(ap, PROG, f"D5:{CL(3 + ns)}{APL + 10}")
dv(ap, ["Not started", "Submitted", "Under review", "Interview", "Offer", "Conditional offer", "Rejected", "Withdrawn"], f"{CL(5 + ns)}5:{CL(6 + ns)}{APL + 10}")
ap.conditional_formatting.add(f"{CMP}5:{CMP}{APL}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"))
ap.conditional_formatting.add(f"D5:{CL(3 + ns)}{APL}", CellIsRule(operator="equal", formula=['"Done"'], fill=GREEN))
ap.conditional_formatting.add(f"D5:{CL(3 + ns)}{APL}", CellIsRule(operator="equal", formula=['"In progress"'], fill=AMBER))
ap.freeze_panes = "D5"

# Phase 8: the funder and where it is based, kept separate from the country of study
FUNDER = {'S02': 'Swedish Institute (Swedish government agency, Stockholm)', 'S30': 'European Union (Erasmus+, EACEA, Brussels) via the SITeS-JM consortium', 'S01': 'Mastercard Foundation (Toronto, Canada) via CMU-Africa', 'S31': 'European Union (Erasmus+, EACEA, Brussels) via the EDISS consortium', 'S09': 'Hungarian Government (Tempus Public Foundation, Budapest)', 'S10': 'The Canadian university (department / supervisor grants)', 'S32': 'The Canadian university (department / supervisor grants)', 'S05': 'Chinese Government (China Scholarship Council, Beijing)', 'S08': 'Korean Government (NIIED, Seoul)', 'S11': 'Italian regional right-to-study agencies (DSU)', 'S35': 'Government of India (ICCR, New Delhi)', 'S13': 'UK Government (FCDO, London) and partner organisations', 'S14': 'US Government (Department of State) via the US Embassy', 'S37': 'Belgian (Flemish) development cooperation (VLIR-UOS, Brussels)', 'S38': 'Italian Ministry of Foreign Affairs (MAECI, Rome)', 'S39': 'European Union (Erasmus+, EACEA, Brussels) via the SMACCs consortium', 'S40': 'European Union (Erasmus+, EACEA, Brussels) via the CDE consortium', 'S41': 'European Union (Erasmus+, EACEA, Brussels) via the CyberMACS consortium', 'S42': 'Czech Government (Ministry of Education, Prague)', 'S43': 'Chinese Government (CSC, Beijing) via the Chinese university', 'S44': 'Taiwan Ministry of Education or the university (Taipei)', 'S45': 'Government of Brunei Darussalam', 'S46': 'Thai Government (TICA, Bangkok)', 'S47': 'Japanese Government (MEXT, Tokyo) via the embassy', 'S48': 'Mastercard Foundation (Toronto, Canada) via the University of Rwanda', 'S49': 'Government of Mauritius', 'S50': 'European Union (EACEA, Brussels) via African university consortia', 'S51': 'The Canadian university (department)', 'S52': 'Australian Government (DFAT, Canberra)', 'S04': 'European Union (Erasmus+, EACEA, Brussels) via the SE4GD consortium', 'S03': 'European Union (Erasmus+, EACEA, Brussels) via the CYBERSURE consortium', 'S36': 'European Union (Erasmus+, EACEA, Brussels) via the SECCLO consortium', 'S06': 'European Union (Erasmus+, EACEA, Brussels) via the CYBERUS consortium', 'S07': 'European Union (Erasmus+, EACEA, Brussels) via the DEAI consortium', 'S12': 'Turkish Government (YTB, Ankara)', 'S33': 'Polish Government (NAWA, Warsaw)', 'S34': 'Japanese Government (JICA, Tokyo)', 'S22': 'German Government (BMZ) funds via DAAD (Bonn)', 'S26': 'KAUST (the university, Saudi Arabia)', 'S27': 'MBZUAI (the university, Abu Dhabi)', 'S28': 'African Union Commission (Addis Ababa)', 'S29': 'AIMS (African Institute for Mathematical Sciences) and its sponsors', 'S17': 'Mastercard Foundation (Toronto, Canada) via UBC', 'S62': 'Mastercard Foundation (Toronto, Canada) via McGill', 'S20': 'Mastercard Foundation (Canada) via Edinburgh; University of Manchester (Equity & Merit)', 'S23': 'None (self-funded)', 'S24': 'Stanford University (Knight-Hennessy endowment, USA)', 'S25': 'UK Government (FCDO) via the Commonwealth Scholarship Commission', 'S53': 'Irish Government (Department of Foreign Affairs / Irish Aid, Dublin)', 'S54': 'KFUPM (the university, Saudi Arabia)', 'S55': 'The Canadian university (department)', 'S56': 'The Canadian university (department)', 'S57': 'New Zealand Government (MFAT, Wellington)', 'S58': 'Japanese Government (MEXT) via the university', 'S59': 'Chinese Government (CSC) via the university', 'S60': 'Chinese Government (CSC) via the university', 'S61': 'Mastercard Foundation (Toronto, Canada) via Arizona State University'}

# ============================ Master Scholarship Database ============================
db = wb.create_sheet("Master Scholarship Database", 0)
title(db, "Master Scholarship Database (v3, official-source verification 7 Nov 2026)",
      "Evidence: E0 own doc | E1 verified directly on official source | E2 official search result | E3 secondary | E4 your confirmation | E5 unsupported. Blue = editable.")
cols = ["ID", "Priority", "Cycle", "Country of study", "University / host", "Programme", "Scholarship", "Funder (and where it is based)", "Funding class", "Coverage (verified)", "GPA / academic rule",
        "GPA eligibility at 2.62", "English requirement", "English status", "Work / other requirement", "Opening date", "Deadline", "Deadline detail / timezone",
        "Deadline status", "Days to deadline", "Urgency", "Official URL", "Evidence level", "Blocked by open discrepancies?", "Funding completeness",
        "Readiness /100", "Category", "Risk", "Key note", "YOUR next action", "CLAUDE next action", "Verified on", "Cat seq"]
W_ = [6, 7, 9, 16, 26, 30, 26, 30, 15, 36, 28, 18, 26, 12, 28, 11, 11, 24, 15, 9, 11, 34, 16, 13, 11, 10, 15, 8, 32, 38, 34, 11, 7]
header(db, 4, cols, W_)
R0 = 5; RL = R0 + N - 1
ci = {c_: i + 1 for i, c_ in enumerate(cols)}; L = {c_: CL(i) for c_, i in ci.items()}
for i, o in enumerate(O):
    r = R0 + i
    vals = {"ID": o["id"], "Priority": o["pri"], "Cycle": o["cyc"], "Country of study": o["country"], "University / host": o["host"], "Programme": o["prog"], "Scholarship": o["sch"],
            "Funder (and where it is based)": FUNDER.get(o["id"], "See 'Scholarship'"),
            "Funding class": o["fund"], "Coverage (verified)": o["cover"], "GPA / academic rule": o["gpa_rule"], "GPA eligibility at 2.62": o["gpa"],
            "English requirement": o["eng_req"], "English status": o["eng"], "Work / other requirement": o["work"], "Opening date": o["opens"] or "TBC",
            "Deadline": o["deadline"] or "TBC", "Deadline detail / timezone": o["tz"], "Deadline status": o["status"], "Official URL": o["url"],
            "Evidence level": o["ev"], "Blocked by open discrepancies?": o["disc"], "Risk": o["risk"], "Key note": o["note"], "YOUR next action": o["my"],
            "CLAUDE next action": o["cl"], "Verified on": o.get("ver", VERIFIED)}
    for k, v in vals.items():
        x = db.cell(row=r, column=ci[k], value=v)
        if isinstance(v, date): x.number_format = "dd-mmm-yyyy"
    if o["url"].startswith("http"): db.cell(row=r, column=ci["Official URL"]).hyperlink = o["url"]
    d_ = f'{L["Deadline"]}{r}'; days = f'{L["Days to deadline"]}{r}'
    db.cell(row=r, column=ci["Days to deadline"], value=f'=IF(ISNUMBER({d_}),{d_}-TODAY(),"")')
    db.cell(row=r, column=ci["Urgency"], value=f'=IF(OR({L["Deadline status"]}{r}="CLOSED",{L["Deadline status"]}{r}="DISCONTINUED"),"CLOSED",IF(NOT(ISNUMBER({d_})),"TBC",IF({days}<0,"PASSED",IF({days}<=21,"URGENT",IF({days}<=60,"HIGH","NORMAL")))))')
    db.cell(row=r, column=ci["Funding completeness"], value=f"=IFERROR(INDEX('Financial Coverage'!$O$5:$O${FCL},MATCH(A{r},'Financial Coverage'!$A$5:$A${FCL},0)),\"n/a\")").number_format = "0%"
    gp = f'{L["GPA eligibility at 2.62"]}{r}'; en = f'{L["English status"]}{r}'; bl = f'{L["Blocked by open discrepancies?"]}{r}'
    prog_ = f"IFERROR(INDEX('Application Progress'!${CMP}$5:${CMP}${APL},MATCH(A{r},'Application Progress'!$A$5:$A${APL},0)),0)"
    db.cell(row=r, column=ci["Readiness /100"], value=f'=ROUND(100*(0.3*IF({gp}="Eligible",1,IF({gp}="Ineligible",0,0.5))+0.2*IF(OR({en}="Not needed",{en}="Done - meets"),1,IF({en}="Booked",0.5,0))+0.2*IF(OR({bl}="No",{BLOCKED}=0),1,0)+0.3*{prog_}),0)')
    rd = f'{L["Readiness /100"]}{r}'
    db.cell(row=r, column=ci["Category"], value=f'=IF(OR(B{r}="C",{L["Deadline status"]}{r}="DISCONTINUED"),"ELIMINATED",IF(C{r}="2028","NEXT CYCLE",IF(C{r}="monitor","MONITOR",IF({L["Urgency"]}{r}="PASSED","ELIMINATED",IF(AND({bl}="Yes",{BLOCKED}>0),"BLOCKED",IF({rd}>=80,"READY","IN PREPARATION"))))))')
    cat = f'{L["Category"]}{r}'
    db.cell(row=r, column=ci["Cat seq"], value=f'={cat}&"#"&COUNTIF(${L["Category"]}${R0}:{cat},{cat})')
ins = {ci[k]: INP for k in ["Priority", "GPA eligibility at 2.62", "English status", "Deadline status", "Risk", "Blocked by open discrepancies?"]}
body(db, R0, RL, 1, len(cols), ins)
dv(db, PRI, f'B{R0}:B{RL + 20}'); dv(db, ELIG, f'{L["GPA eligibility at 2.62"]}{R0}:{L["GPA eligibility at 2.62"]}{RL + 20}')
dv(db, ENG, f'{L["English status"]}{R0}:{L["English status"]}{RL + 20}'); dv(db, STAT, f'{L["Deadline status"]}{R0}:{L["Deadline status"]}{RL + 20}')
dv(db, FUND, f'{L["Funding class"]}{R0}:{L["Funding class"]}{RL + 20}'); dv(db, ["Low", "Medium", "High"], f'{L["Risk"]}{R0}:{L["Risk"]}{RL + 20}')
dv(db, ["Yes", "No"], f'{L["Blocked by open discrepancies?"]}{R0}:{L["Blocked by open discrepancies?"]}{RL + 20}')
pri_cf(db, f"B{R0}:B{RL}")
gc = L["GPA eligibility at 2.62"]
db.conditional_formatting.add(f"{gc}{R0}:{gc}{RL}", CellIsRule(operator="equal", formula=['"Eligible"'], fill=GREEN))
db.conditional_formatting.add(f"{gc}{R0}:{gc}{RL}", CellIsRule(operator="equal", formula=['"Ineligible"'], fill=RED))
uc = L["Urgency"]
db.conditional_formatting.add(f"{uc}{R0}:{uc}{RL}", CellIsRule(operator="equal", formula=['"URGENT"'], fill=RED))
db.conditional_formatting.add(f"{uc}{R0}:{uc}{RL}", CellIsRule(operator="equal", formula=['"HIGH"'], fill=AMBER))
cc = L["Category"]
for v, f in (("READY", GREEN), ("BLOCKED", RED), ("IN PREPARATION", AMBER), ("ELIMINATED", GREY)):
    db.conditional_formatting.add(f"{cc}{R0}:{cc}{RL}", CellIsRule(operator="equal", formula=[f'"{v}"'], fill=f))
db.conditional_formatting.add(f'{L["Readiness /100"]}{R0}:{L["Readiness /100"]}{RL}', DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color="5B9BD5"))
db.column_dimensions[L["Cat seq"]].hidden = True
db.freeze_panes = "C5"; db.auto_filter.ref = f"A4:{CL(len(cols) - 1)}{RL}"

# ============================ Eligibility Assessment ============================
el = wb.create_sheet("Eligibility Assessment", 1)
title(el, "Matching Score (0-100) - strategic, NOT an admission probability", "Edit blue scores (0-10) or the weights in row 4.")
el["A4"] = "Weights ->"; el["A4"].font = Font(name=F, bold=True)
for j, w in enumerate([0.25, 0.25, 0.15, 0.15, 0.10, 0.10]):
    x = el.cell(row=4, column=6 + j, value=w); x.number_format = "0%"; x.font, x.fill = INP, YEL
el["L4"] = "=SUM(F4:K4)"; el["L4"].number_format = "0%"
header(el, 5, ["ID", "Opportunity", "Priority", "GPA at 2.62", "Evidence", "Academic (25%)", "Funding (25%)", "Relevance (15%)", "Prof. exp. (15%)",
               "Flexibility (10%)", "Feasibility (10%)", "Score", "Rank", "Key note"], [6, 46, 8, 20, 14, 10, 10, 10, 10, 10, 10, 9, 7, 50])
E0 = 6; EL = E0 + N - 1
for i, o in enumerate(O):
    r, dr = E0 + i, R0 + i
    el.cell(row=r, column=1, value=f"='Master Scholarship Database'!A{dr}")
    el.cell(row=r, column=2, value=f"='Master Scholarship Database'!G{dr}&\" - \"&'Master Scholarship Database'!F{dr}")
    el.cell(row=r, column=3, value=f"='Master Scholarship Database'!B{dr}")
    el.cell(row=r, column=4, value=f"='Master Scholarship Database'!{gc}{dr}")
    el.cell(row=r, column=5, value=f"='Master Scholarship Database'!{L['Evidence level']}{dr}")
    for j, s in enumerate(o["scores"]): el.cell(row=r, column=6 + j, value=s)
    el.cell(row=r, column=12, value=f"=ROUND(SUMPRODUCT(F{r}:K{r},$F$4:$K$4)*10,1)")
    el.cell(row=r, column=13, value=f"=RANK(L{r},$L${E0}:$L${EL})+COUNTIF($L${E0 - 1}:L{r - 1},L{r})")
    el.cell(row=r, column=14, value=o["note"])
body(el, E0, EL, 1, 14, {**{c_: LNK for c_ in range(1, 6)}, **{c_: INP for c_ in range(6, 12)}})
d_ = DataValidation(type="whole", operator="between", formula1="0", formula2="10"); el.add_data_validation(d_); d_.add(f"F{E0}:K{EL}")
el.conditional_formatting.add(f"L{E0}:L{EL}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color="5B9BD5"))
pri_cf(el, f"C{E0}:C{EL}"); el.freeze_panes = "C6"

# ============================ Application Deadlines ============================
dls = wb.create_sheet("Application Deadlines", 2)
title(dls, "Deadline & Milestone Calendar", "Days remaining recalculates daily. Official deadlines carry an evidence code; 'Internal' = your target date.")
MIL = [("ALL", "Send HR, registrar and referee requests (kit ready)", date(2026, 11, 12), "Internal", "Internal"),
       ("S10", "Manitoba go / no-go decision", date(2026, 11, 12), "Internal", "Internal"),
       ("S02", "SI templates + eligible programme list published", date(2026, 11, 13), "CET", "E1"),
       ("S31", "EDISS scholarship round opens (08:00 Helsinki)", date(2026, 11, 16), "Helsinki", "E1"),
       ("ALL", "Book IELTS Academic (in-centre, ETB 32,800)", date(2026, 11, 20), "Internal", "Internal"),
       ("S10", "Supervisor confirmation needed for Manitoba", date(2026, 11, 25), "Internal", "Internal"),
       ("S10", "University of Manitoba MSc CS deadline (international)", date(2026, 12, 1), "Central", "E1"),
       ("S01", "CMU-Africa: referees entered in portal", date(2026, 12, 1), "Internal", "Internal"),
       ("S30", "AISS application window opens", date(2026, 12, 2), "Tallinn", "E1"),
       ("S01", "CMU-Africa target submission", date(2026, 12, 12), "Internal", "Internal"),
       ("S01", "CMU-Africa EARLY deadline (DET & test fees covered)", date(2026, 12, 15), "Not stated", "E2"),
       ("S31", "EDISS target submission", date(2027, 1, 9), "Internal", "Internal"),
       ("S31", "EDISS deadline (15:00 Helsinki)", date(2027, 1, 12), "15:00 Helsinki", "E1"),
       ("S02", "universityadmissions.se deadline (rank 4 programmes)", date(2027, 1, 15), "CET", "E1"),
       ("S01", "CMU-Africa FINAL deadline", date(2027, 1, 15), "Not stated", "E2"),
       ("S32", "University of Calgary early-review deadline", date(2027, 1, 15), "Mountain", "E1"),
       ("S09", "Stipendium Hungaricum (expected, based on 2026 pattern)", date(2027, 1, 15), "14:00 CET (2026)", "E3"),
       ("S30", "AISS target submission", date(2027, 1, 30), "Internal", "Internal"),
       ("S02", "UA Sweden: fee SEK 900 + documents (transcript from DTU) deadline", date(2027, 2, 1), "CET", "E1"),
       ("S30", "AISS deadline", date(2027, 2, 2), "Tallinn", "E1"),
       ("S02", "SI scholarship portal opens", date(2027, 2, 8), "CET", "E1"),
       ("S02", "SI scholarship portal CLOSES (14:59 CET)", date(2027, 2, 24), "14:59 CET", "E1"),
       ("S10", "Manitoba awards (UMGF/entrance) application deadline", date(2027, 2, 27), "Central", "E1"),
       ("S32", "University of Calgary final deadline", date(2027, 3, 1), "Mountain", "E1"),
       ("S02", "University Admissions results", date(2027, 4, 1), "CET", "E1"),
       ("S02", "SI scholarship results", date(2027, 4, 22), "CET", "E1"),
       ("S40", "CDE (Erasmus Mundus) deadline", date(2027, 1, 22), "23:59 CET", "E2"),
       ("S39", "SMACCs scholarship deadline", date(2027, 2, 1), "Check site", "E3"),
       ("S43", "CSC Type B: DUT deadline", date(2027, 2, 15), "China", "E2"),
       ("S43", "CSC Type B: HUST deadline", date(2027, 2, 25), "China", "E2"),
       ("S43", "CSC Type B: NPU deadline (17:00 UTC+8)", date(2027, 2, 28), "17:00 UTC+8", "E2"),
       ("S37", "VLIR-UOS ICP Connect (expected; programme deadlines may be earlier)", date(2027, 2, 28), "Expected", "E3"),
       ("S35", "ICCR portal (expected window 27 Feb - 15 Apr, based on 2026)", date(2027, 4, 15), "Expected", "E2")]
MIL.sort(key=lambda x: x[2])
header(dls, 4, ["ID", "Milestone", "Date", "Timezone", "Days left", "Urgency", "Evidence", "Done?"], [7, 60, 12, 16, 10, 11, 10, 8])
for i, (oid, m, dt, tz, ev) in enumerate(MIL):
    r = 5 + i
    for c_, v in enumerate([oid, m, dt, tz, f"=C{r}-TODAY()", f'=IF(H{r}="Yes","DONE",IF(E{r}<0,"PASSED",IF(E{r}<=14,"URGENT",IF(E{r}<=45,"SOON","PLANNED"))))', ev, "No"], 1):
        x = dls.cell(row=r, column=c_, value=v)
        if c_ == 3: x.number_format = "dd-mmm-yyyy"
MLL = 4 + len(MIL)
body(dls, 5, MLL, 1, 8, {8: INP}); dv(dls, ["Yes", "No"], f"H5:H{MLL + 20}")
for v, f in (("URGENT", RED), ("SOON", AMBER), ("DONE", GREEN)):
    dls.conditional_formatting.add(f"F5:F{MLL}", CellIsRule(operator="equal", formula=[f'"{v}"'], fill=f))
dls.freeze_panes = "A5"

# ============================ Required Documents ============================
rdoc = wb.create_sheet("Required Documents")
title(rdoc, "Required Documents Register", "Status = blue dropdown. No passport or ID numbers in this file.")
header(rdoc, 4, ["#", "Document", "Needed for", "Status", "How to get it (kit file)", "Notes", "Missing seq"], [5, 46, 22, 20, 36, 40, 8])
DOCS = [("SantimPay HR employment certificate", "ALL (blocking)", "Not started", "SantimPay_HR_Request.docx", "Titles, dates, FT/hours, duties, stamp"),
        ("Kifiya HR employment certificate", "ALL (blocking)", "Not started", "Kifiya_HR_Request.docx", "Dates, type, weekly hours"),
        ("Personal Information Confirmation Form (completed)", "ALL", "Not started", "Personal_Information_Confirmation_Form.docx", "E4 answers"),
        ("Experience calculation (HR-letter data entered)", "S02 + CV", "Not started", "15_EMPLOYMENT_EVIDENCE/Experience_Calculation.xlsx", "SI 3,000 h / Chevening 2,800 h / EPOS 24 months"),
        ("Transcript sent by DTU to UHR Sweden (e-mail or post)", "S02", "Not started", "University_Registrar_Request.docx", "Deadline 1 Feb 2027"),
        ("Sealed official transcripts (x3)", "S31, S09, others", "Not started", "University_Registrar_Request.docx", "University stamp across the flap"),
        ("Registrar letter: CGPA, grading scale, calendar, Extension, English medium", "S31, S09, S02", "Not started", "University_Registrar_Request.docx", ""),
        ("Printed (permanent) BSc diploma", "ALL", "Not started", "University_Registrar_Request.docx", "Temporary certificate held"),
        ("Department letter: final-project title + course syllabi", "S09, S31, CV", "Not started", "University_Registrar_Request.docx (letter B)", "DR-17"),
        ("IELTS Academic result (in-centre)", "S31, S30, S10", "Not started", "Book at BC Addis Ababa University", "EDISS: no band < 6.0"),
        ("Duolingo English Test result", "S01", "Not started", "Via CMU-Africa early-applicant route", "Target 120"),
        ("2 academic recommendation letters (signed)", "S31 (required), S02, S01", "Not started", "Lecturer_Recommendation_Request.docx", "EDISS accepts academic letters only"),
        ("2 professional references (SI template, signed & stamped)", "S02, S01", "Not started", "Manager_Recommendation_Request.docx", "Template published 13 Nov"),
        ("SI Proof of Work & Leadership forms (signed & stamped)", "S02", "Not started", "SantimPay/Kifiya HR (after 13 Nov)", "Max 5 pages total"),
        ("Portfolio: 2-3 public demo repositories + final project", "S30, S01", "Not started", "AISS package section 4", "No employer code"),
        ("Motivation videos (EDISS 2 min; AISS 90 s; CMU video essay)", "S31, S30, S01", "Not started", "Scripts in each package", ""),
        ("AISS self-funding plan", "S30", "Not started", "AISS package section 5", "Mandatory"),
        ("Passport copy (redacted for screening)", "ALL", "Have (not reviewed)", "Drive 'Passport' folder", "Number only in official portals"),
        ("Medical certificate (< 1 month old)", "S09", "Not started", "Approved clinic, when the call opens", "Szeged requirement"),
        ("Certification verification IDs (CKA, AZ-900)", "CV", "Needs verification", "Vendor portals", "DR-12")]
for i, d in enumerate(DOCS):
    r = 5 + i
    rdoc.cell(row=r, column=1, value=i + 1)
    for c_, v in enumerate(d, 2): rdoc.cell(row=r, column=c_, value=v)
    rdoc.cell(row=r, column=7, value=f'=IF(LEFT(D{r},4)="Have","",COUNTIF($D$5:D{r},"<>Have*"))')
RDL = 4 + len(DOCS)
body(rdoc, 5, RDL, 1, 7, {4: INP})
dv(rdoc, ["Have", "Have (not reviewed)", "Requested", "Not started", "Needs verification", "N/A"], f"D5:D{RDL + 20}")
rdoc.conditional_formatting.add(f"D5:D{RDL}", FormulaRule(formula=['LEFT(D5,4)="Have"'], fill=GREEN))
rdoc.conditional_formatting.add(f"D5:D{RDL}", CellIsRule(operator="equal", formula=['"Requested"'], fill=AMBER))
rdoc.freeze_panes = "C5"

# ============================ Recommendation Tracking ============================
rt = wb.create_sheet("Recommendation Tracking")
title(rt, "Recommendation Tracking", "Referees write and sign their own letters. [Brackets] = replace with real names.")
header(rt, 4, ["Referee", "Type", "Relationship (verified)", "For", "Request sent", "Agreed?", "Status", "Deadline", "Notes"], [34, 12, 34, 18, 12, 10, 14, 12, 40])
REFS = [["[Final-project supervisor, DTU]", "Academic", "Final Project I & II (A-, A-)", "S31, S01, S02", "No", "No", "Not requested", date(2026, 12, 1), "EDISS academic letter #1"],
        ["[Lecturer: Compiler Design / Adv. Programming / Complexity]", "Academic", "A- course", "S31, S01", "No", "No", "Not requested", date(2026, 12, 1), "EDISS academic letter #2"],
        ["[SantimPay line manager]", "Professional", "Direct supervisor", "S02, S01", "No", "No", "Not requested", date(2026, 12, 1), "SI work-based reference (template)"],
        ["[Second professional referee]", "Professional", "Kifiya supervisor or senior colleague", "S02", "No", "No", "Not requested", date(2027, 2, 15), "SI template"]]
for i, row in enumerate(REFS):
    for c_, v in enumerate(row, 1):
        x = rt.cell(row=5 + i, column=c_, value=v)
        if isinstance(v, date): x.number_format = "dd-mmm-yyyy"
body(rt, 5, 8, 1, 9, {c_: INP for c_ in (1, 5, 6, 7)})
dv(rt, ["Yes", "No"], "E5:F40"); dv(rt, ["Not requested", "Requested", "Agreed", "Submitted", "Declined"], "G5:G40")

# ============================ Visa Preparation ============================
vp = wb.create_sheet("Visa Preparation")
title(vp, "Visa Preparation (summary)", "Visa rules are separate from admission. Never use false documents or borrowed funds presented as your own.")
header(vp, 4, ["Country", "Permit", "Official link", "Financial requirement", "Notes", "Status"], [16, 24, 44, 40, 46, 12])
VISA = [["Sweden", "Residence permit (studies)", "https://www.migrationsverket.se", "SI award letter; check the Migrationsverket monthly figure", "SI excludes insurance: check university cover", "Not started"],
        ["Estonia / Finland / Portugal (AISS)", "Residence permits per mobility", "Per country", "EM scholarship letter", "Mobility in 3 countries", "Not started"],
        ["Finland (EDISS year 1)", "Residence permit (studies)", "https://migri.fi", "EM scholarship letter", "-", "Not started"],
        ["Rwanda", "Student permit", "https://www.migration.gov.rw", "CMU-Africa letter", "-", "Not started"],
        ["Hungary", "Residence permit (study)", "https://enterhungary.gov.hu", "SH award letter", "-", "Not started"],
        ["Canada", "Study permit", "https://www.canada.ca", "CAD 22,895 + tuition (from 1 Sep 2025)", "Master's at public DLIs exempt from PAL since 1 Jan 2026 (E1 umanitoba.ca)", "Not started"],
        ["Germany", "National (D) visa", "https://addis-abeba.diplo.de/et-en/service/visa-einreise", "Blocked account (Sperrkonto)", "Embassy page verified (E1)", "Not started"],
        ["United Kingdom", "Student visa", "https://www.gov.uk/student-visa/money", "GBP 1,171/month outside London; 1,529 London (E1)", "Visa fee GBP 558 (E1)", "Not started"]]
for i, row in enumerate(VISA):
    for c_, v in enumerate(row, 1):
        x = vp.cell(row=5 + i, column=c_, value=v)
        if c_ == 3 and str(v).startswith("http"): x.hyperlink = v
body(vp, 5, 4 + len(VISA), 1, 6, {6: INP})
dv(vp, ["Not started", "Preparing", "Submitted", "Approved", "Refused", "N/A"], "F5:F30")

# ============================ Interview Preparation ============================
ip = wb.create_sheet("Interview Preparation")
title(ip, "Interview & English Preparation", "Model answers only from verified facts. Log mock results in blue cells.")
header(ip, 4, ["Item", "Target", "Done", "% ready", "Notes"], [44, 10, 10, 10, 60])
QB = [("GPA & academic weaknesses questions", 6, "Prepare first (model answer in plan file)"), ("Professional experience questions", 9, "After HR letters"),
      ("Leadership examples (STAR)", 3, "Real examples only"), ("AISS 10-minute presentation", 1, "Structure in AISS package"),
      ("EDISS 2-min video takes", 3, "Script ready"), ("CMU-Africa video essay takes", 3, "Script ready"), ("IELTS full mock tests", 3, "Target 7.0, no band < 6.5")]
for i, (q, n, note) in enumerate(QB):
    r = 5 + i
    for c_, v in enumerate([q, n, 0, f"=IFERROR(C{r}/B{r},0)", note], 1): ip.cell(row=r, column=c_, value=v)
    ip.cell(row=r, column=4).number_format = "0%"
body(ip, 5, 4 + len(QB), 1, 5, {3: INP})
ip.conditional_formatting.add(f"D5:D{4 + len(QB)}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5"))

# ============================ Research Sources (E1 evidence) ============================
rs = wb.create_sheet("Research Sources")
title(rs, "Official Evidence Register (E1 snapshots in 03_SCHOLARSHIP_RESEARCH/Official_Evidence_2026-11-07)", "Re-check live pages before submitting.")
header(rs, 4, ["Snapshot file", "What it proves", "Official URL", "Captured (UTC)", "SHA-256 (16)"], [24, 64, 60, 18, 18])
rows = []
if os.path.exists(EVIDENCE_INDEX):
    for line in open(EVIDENCE_INDEX, encoding="utf-8"):
        m = re.match(r"\| `([^`]+)` \| (.+?) \| (https?://\S+) \| ([^|]+) \| `?([^`|]+)`? \|", line.strip())
        if m: rows.append([m.group(1), m.group(2), m.group(3), m.group(4).strip(), m.group(5).strip()])
for i, row in enumerate(rows):
    for c_, v in enumerate(row, 1):
        x = rs.cell(row=5 + i, column=c_, value=v)
        if c_ == 3: x.hyperlink = v
body(rs, 5, 4 + max(1, len(rows)), 1, 5)
rs.freeze_panes = "A5"

# ============================ Action Items ============================
ai = wb.create_sheet("Action Items")
title(ai, "Action Items (priority order). Claude never creates accounts, pays, signs, sends or submits.", "Status = blue dropdown. The Command Center shows the top open items automatically.")
header(ai, 4, ["Priority", "Owner", "Action", "Due", "Days left", "Status", "Where", "Open seq (YOU)", "Open seq (CLAUDE)"], [8, 9, 70, 12, 9, 13, 34, 9, 9])
ACTS = [("YOU", "Hand in SantimPay + Kifiya HR requests and the DTU Registrar request", date(2026, 11, 12), "09_APPLICATION_DOCUMENTS/00_Document_Collection_Kit/"),
        ("YOU", "Ask 2 lecturers + 2 managers to be referees", date(2026, 11, 12), "09_APPLICATION_DOCUMENTS/00_Document_Collection_Kit/"),
        ("YOU", "Open the CMU-Africa portal, create the account, paste the essay questions", date(2026, 11, 15), "APPLICATION_EXECUTION_CENTER/07_CMU_AFRICA/"),
        ("YOU", "Approve and send e-mails 1-4 (MoE, DUT, HUST, SMACCs)", date(2026, 11, 15), "APPLICATION_EXECUTION_CENTER/05_READY_TO_SEND_EMAILS.docx"),
        ("YOU", "Choose 4 SI-eligible Swedish programmes (after the 13 Nov list)", date(2026, 11, 20), "APPLICATION_EXECUTION_CENTER/08_SWEDEN_SI/"),
        ("YOU", "Write the 6 STAR stories", date(2026, 11, 22), "APPLICATION_EXECUTION_CENTER/13_INTERVIEW_PREPARATION/"),
        ("YOU", "Build the 2 public demo repositories (no employer code)", date(2026, 11, 30), "APPLICATION_EXECUTION_CENTER/09_AISS/"),
        ("YOU", "Book IELTS Academic (only once AISS is confirmed)", date(2026, 11, 20), "APPLICATION_EXECUTION_CENTER/13_INTERVIEW_PREPARATION/"),
        ("CLAUDE", "Compare the SI 2027/28 templates and list with the package (13 Nov)", date(2026, 11, 14), "APPLICATION_EXECUTION_CENTER/08_SWEDEN_SI/"),
        ("CLAUDE", "Fill CVs and essays from the HR letters within 48 h of receipt", date(2026, 11, 27), "APPLICATION_EXECUTION_CENTER/06_FINAL_CV/"),
        ("CLAUDE", "Adapt CMU-Africa essays to the real questions (2 review rounds)", date(2026, 11, 28), "APPLICATION_EXECUTION_CENTER/07_CMU_AFRICA/"),
        ("CLAUDE", "Mock interviews + video-essay feedback", date(2026, 12, 5), "APPLICATION_EXECUTION_CENTER/13_INTERVIEW_PREPARATION/")]
for i, (own, act, due, where) in enumerate(ACTS):
    r = 5 + i
    for c_, v in enumerate([i + 1, own, act, due, f"=D{r}-TODAY()", "Not started", where], 1):
        x = ai.cell(row=r, column=c_, value=v)
        if c_ == 4: x.number_format = "dd-mmm-yyyy"
    ai.cell(row=r, column=8, value=f'=IF(AND(B{r}="YOU",F{r}<>"Done"),COUNTIFS($B$5:B{r},"YOU",$F$5:F{r},"<>Done"),"")')
    ai.cell(row=r, column=9, value=f'=IF(AND(B{r}="CLAUDE",F{r}<>"Done"),COUNTIFS($B$5:B{r},"CLAUDE",$F$5:F{r},"<>Done"),"")')
AIL = 4 + len(ACTS)
body(ai, 5, AIL, 1, 9, {6: INP})
dv(ai, ["Not started", "In progress", "Done", "Blocked"], f"F5:F{AIL + 20}")
ai.conditional_formatting.add(f"F5:F{AIL}", CellIsRule(operator="equal", formula=['"Done"'], fill=GREEN))
ai.conditional_formatting.add(f"F5:F{AIL}", CellIsRule(operator="equal", formula=['"Blocked"'], fill=RED))
ai.freeze_panes = "A5"

# ============================ Risk Register ============================
rr_ = wb.create_sheet("Risk Register")
title(rr_, "Risk Register", "Likelihood x impact (1-5). Blue = editable.")
header(rr_, 4, ["#", "Risk", "Likelihood", "Impact", "Score", "Mitigation", "Owner"], [5, 52, 10, 10, 8, 62, 9])
RISKS = [("CV dates contradict HR letters / SI forms", 4, 5, "HR letters first; one verified history; Discrepancy Log", "YOU"),
         ("2.62 fails grade-based screens", 4, 3, "Prioritise holistic routes (SI, AISS, CMU-Africa, ELTE exam)", "BOTH"),
         ("Mastercard Foundation award not given (need-based)", 3, 4, "Truthful finances; parallel applications", "YOU"),
         ("IELTS below thresholds (EDISS: no band < 6.0)", 3, 4, "5-week plan; test early to mid December", "YOU"),
         ("Referees late or unwilling", 3, 3, "Ask 4 by 12 Nov; information sheets", "YOU"),
         ("Facts only E2 for blocked sites (CMU-Africa, CSC, Fulbright)", 3, 3, "Check in a browser before submitting", "YOU"),
         ("Compressed deadlines (1 Dec - 24 Feb)", 4, 3, "Action list with buffers", "BOTH"),
         ("Local-only files lost if this cloud container is recycled", 3, 4, "Download files sent in chat; you manage GitHub sync", "YOU")]
for i, (rk, l, im, mit, own) in enumerate(RISKS):
    r = 5 + i
    for c_, v in enumerate([i + 1, rk, l, im, f"=C{r}*D{r}", mit, own], 1): rr_.cell(row=r, column=c_, value=v)
RRL = 4 + len(RISKS)
body(rr_, 5, RRL, 1, 7, {3: INP, 4: INP})
rr_.conditional_formatting.add(f"E5:E{RRL}", CellIsRule(operator="greaterThanOrEqual", formula=["16"], fill=RED))
rr_.conditional_formatting.add(f"E5:E{RRL}", CellIsRule(operator="between", formula=["9", "15"], fill=AMBER))

# ============================ COMMAND CENTER (first sheet) ============================
cc_ = wb.create_sheet("Command Center", 0)
cc_.column_dimensions["A"].width = 4; cc_.column_dimensions["B"].width = 62; cc_.column_dimensions["C"].width = 16
cc_.column_dimensions["D"].width = 14; cc_.column_dimensions["E"].width = 44
cc_["B1"] = "SCHOLARSHIP COMMAND CENTER (2027 intake)"; cc_["B1"].font = Font(name=F, bold=True, size=18, color="1F3864")
cc_["B2"] = f"Applicant: Nebyu Dejenie Tilahun | BSc CS, Debre Tabor University (2022) | CGPA 2.62 (verified) | Last data update {VERIFIED:%d %b %Y}. Updates automatically as you change statuses in the other sheets."
cc_["B2"].font = SUB
DB = "'Master Scholarship Database'"
catcol = f"{DB}!${L['Category']}${R0}:${L['Category']}${RL}"
seqcol = f"{DB}!${L['Cat seq']}${R0}:${L['Cat seq']}${RL}"
namecol = f"{DB}!$G${R0}:$G${RL}"; progcol = f"{DB}!$F${R0}:$F${RL}"
dcol = f"{DB}!${L['Deadline']}${R0}:${L['Deadline']}${RL}"; stcol = f"{DB}!${L['Deadline status']}${R0}:${L['Deadline status']}${RL}"
row = 4

def q(text):
    global row
    x = cc_.cell(row=row, column=2, value=text); x.font = QFONT; x.fill = BLUEF
    for c_ in range(2, 6): cc_.cell(row=row, column=c_).fill = BLUEF
    row += 1

def line(b, c_=None, d=None, e=None, fmt=None, font=BODY):
    global row
    for col, v in ((2, b), (3, c_), (4, d), (5, e)):
        if v is not None:
            x = cc_.cell(row=row, column=col, value=v); x.font = font; x.alignment = WR; x.border = BD
            if fmt and col in (3, 4): x.number_format = fmt
    row += 1

def listcat(cat, k_max, extra_col=None):
    for k in range(1, k_max + 1):
        key = f'"{cat}#{k}"'
        line(f'=IFERROR(INDEX({namecol},MATCH({key},{seqcol},0))&" – "&INDEX({progcol},MATCH({key},{seqcol},0)),"")',
             f'=IFERROR(INDEX({dcol},MATCH({key},{seqcol},0)),"")', None,
             f"=IFERROR(INDEX({DB}!${L['YOUR next action']}${R0}:${L['YOUR next action']}${RL},MATCH({key},{seqcol},0)),\"\")" if extra_col else None, fmt="dd-mmm-yyyy")

q("1. Your five most important actions (top open YOU items)")
line("Action", "Due", "Status", "Where", font=Font(name=F, bold=True, size=10))
for k in range(1, 6):
    line(f"=IFERROR(INDEX('Action Items'!$C$5:$C${AIL},MATCH({k},'Action Items'!$H$5:$H${AIL},0)),\"\")",
         f"=IFERROR(INDEX('Action Items'!$D$5:$D${AIL},MATCH({k},'Action Items'!$H$5:$H${AIL},0)),\"\")",
         f"=IFERROR(INDEX('Action Items'!$F$5:$F${AIL},MATCH({k},'Action Items'!$H$5:$H${AIL},0)),\"\")",
         f"=IFERROR(INDEX('Action Items'!$G$5:$G${AIL},MATCH({k},'Action Items'!$H$5:$H${AIL},0)),\"\")", fmt="dd-mmm-yyyy")
row += 1
q("2. Which scholarship has the nearest deadline? (open, non-eliminated)")
nd = f'_xlfn.MINIFS({dcol},{dcol},">="&TODAY(),{catcol},"<>ELIMINATED",{catcol},"<>NEXT CYCLE")'
line(f'=IFERROR(INDEX({namecol},MATCH({nd},{dcol},0))&" – "&INDEX({progcol},MATCH({nd},{dcol},0)),"none")', f"=IFERROR({nd},\"\")",
     f'=IFERROR({nd}-TODAY()&" days","")', None, fmt="dd-mmm-yyyy", font=Font(name=F, bold=True, size=11, color="9C0006"))
row += 1
q("3. Which applications are READY to submit?")
line(f'=COUNTIF({catcol},"READY")&" ready  (READY = eligible + English done + no blocking discrepancy + progress ≥ 80%)"')
listcat("READY", 3)
row += 1
q("4. Which applications are BLOCKED? (waiting for HR letters / your confirmations)")
line(f'=COUNTIF({catcol},"BLOCKED")&" blocked  |  open BLOCKING discrepancies: "&{BLOCKED}')
listcat("BLOCKED", 8, extra_col=True)
row += 1
q("5. Which documents are missing?")
line(f"=COUNTIF('Required Documents'!$D$5:$D${RDL},\"<>Have*\")&\" of {len(DOCS)} documents not yet in hand – top items:\"")
for k in range(1, 7):
    line(f"=IFERROR(INDEX('Required Documents'!$B$5:$B${RDL},MATCH({k},'Required Documents'!$G$5:$G${RDL},0)),\"\")", None, None,
         f"=IFERROR(INDEX('Required Documents'!$E$5:$E${RDL},MATCH({k},'Required Documents'!$G$5:$G${RDL},0)),\"\")")
row += 1
q("6. What information do I (Claude) need from you?")
for k in range(1, 7):
    line(f"=IFERROR(INDEX('Discrepancy Log'!$B$5:$B${DL},MATCH({k},'Discrepancy Log'!$H$5:$H${DL},0))&\" → \"&INDEX('Discrepancy Log'!$G$5:$G${DL},MATCH({k},'Discrepancy Log'!$H$5:$H${DL},0)),\"\")")
row += 1
q("7. What is Claude doing independently?")
for k in range(1, 6):
    line(f"=IFERROR(INDEX('Action Items'!$C$5:$C${AIL},MATCH({k},'Action Items'!$I$5:$I${AIL},0)),\"\")",
         f"=IFERROR(INDEX('Action Items'!$D$5:$D${AIL},MATCH({k},'Action Items'!$I$5:$I${AIL},0)),\"\")", None, None, fmt="dd-mmm-yyyy")
row += 1
q("8. Which opportunities have been eliminated? (reasons in the database)")
line(f'=COUNTIF({catcol},"ELIMINATED")&" eliminated / low potential"')
for k in range(1, 7):
    key = f'"ELIMINATED#{k}"'
    line(f'=IFERROR(INDEX({namecol},MATCH({key},{seqcol},0))&" – "&INDEX({progcol},MATCH({key},{seqcol},0)),"")', None, None,
         f"=IFERROR(INDEX({DB}!${L['Key note']}${R0}:${L['Key note']}${RL},MATCH({key},{seqcol},0)),\"\")")
line("… full list: filter column 'Category' = ELIMINATED in the Master Scholarship Database", font=SUB)
row += 1
q("9. What is the next deadline or milestone (any type)?")
nm = f"_xlfn.MINIFS('Application Deadlines'!$C$5:$C${MLL},'Application Deadlines'!$C$5:$C${MLL},\">=\"&TODAY(),'Application Deadlines'!$H$5:$H${MLL},\"<>Yes\")"
line(f"=IFERROR(INDEX('Application Deadlines'!$B$5:$B${MLL},MATCH({nm},'Application Deadlines'!$C$5:$C${MLL},0)),\"none\")", f"=IFERROR({nm},\"\")",
     f'=IFERROR({nm}-TODAY()&" days","")', None, fmt="dd-mmm-yyyy", font=Font(name=F, bold=True, size=11))
row += 1
q("10. Next best action")
line(f"=IFERROR(\"→ \"&INDEX('Action Items'!$C$5:$C${AIL},MATCH(1,'Action Items'!$H$5:$H${AIL},0)),\"All your actions are done – check the Deadlines sheet\")", None, None,
     f"=IFERROR(INDEX('Action Items'!$G$5:$G${AIL},MATCH(1,'Action Items'!$H$5:$H${AIL},0)),\"\")", font=Font(name=F, bold=True, size=12, color="006100"))
row += 1
q("Portfolio overview")
for cat in ["IN PREPARATION", "BLOCKED", "READY", "MONITOR", "NEXT CYCLE", "ELIMINATED"]:
    line(cat, f'=COUNTIF({catcol},"{cat}")')
cc_.cell(row=row + 1, column=2, value="Legend: blue text = editable input in other sheets | green = linked | READY needs: GPA-eligible, English done or not needed, no open BLOCKING discrepancy, ≥ 80% progress.").font = SUB

order = ["Command Center", "Master Scholarship Database", "Eligibility Assessment", "Application Deadlines", "Required Documents", "Recommendation Tracking",
         "Application Progress", "Financial Coverage", "Visa Preparation", "Interview Preparation", "Research Sources", "Discrepancy Log", "Action Items", "Risk Register"]
wb._sheets = [wb[n] for n in order]
for s in wb.worksheets: s.sheet_view.zoomScale = 90
wb.active = 0
wb.save(OUT)
print("saved", OUT, "opportunities", N)
