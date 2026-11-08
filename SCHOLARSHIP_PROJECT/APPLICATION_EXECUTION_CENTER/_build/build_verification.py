"""00_VERIFICATION_REGISTER.xlsx: every key fact, separated into VERIFIED (official page read directly),
SEARCH-INDEXED (official page found by search, not opened here) and NOT FOUND / UNVERIFIED."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule

OUT = os.path.join(os.path.dirname(__file__), "..", "00_VERIFICATION_REGISTER.xlsx")
V, S, U = "VERIFIED (E1)", "SEARCH-INDEXED (E2)", "UNVERIFIED / NOT FOUND"
R = [
 # opportunity, fact, value, status, source, checked, what you must confirm
 ("Sweden SI", "Ethiopian work-experience rule", "≥ 3,000 hours before 8 Feb 2027, max 3 employers; internships excluded; otherwise ineligible", V, "si.se scholarship page + FAQ", "8 Nov 2026", "–"),
 ("Sweden SI", "Leadership requirement", "Required; priority to leadership from employment", V, "si.se", "8 Nov 2026", "–"),
 ("Sweden SI", "Scholarship value", "Tuition + SEK 12,000/month + SEK 15,000 travel grant", V, "si.se", "7 Nov 2026", "–"),
 ("Sweden SI", "Dates", "Universities 16 Oct–15 Jan; SI portal 8–24 Feb 2027 (14:59 CET); SI results 22 Apr", V, "si.se", "8 Nov 2026", "–"),
 ("Sweden SI", "Templates and eligible programmes", "Published 13 Nov 2026", V, "si.se", "8 Nov 2026", "Download on 13 Nov"),
 ("Sweden (University Admissions)", "Fee and documents", "SEK 900 by 1 Feb; transcript direct from DTU; Ethiopian degree meets English", V, "universityadmissions.se", "8 Nov 2026", "–"),
 ("Sweden (University Admissions)", "Open programmes + tuition", "100 relevant English-taught programmes for Autumn 2027", V, "universityadmissions.se search API", "8 Nov 2026", "–"),
 ("Sweden programmes", "Entry requirements (LTU, SU, KAU, LiU)", "See SI_Programme_Shortlist.xlsx", S, "ltu.se, su.se, kau.se, liu.se", "8 Nov 2026", "Open each programme page"),
 ("AISS (SITeS-JM)", "Identity", "EACEA catalogue: SITeS-JM 'Sustainable Information Technologies for Societies Joint Masters' → aissprogram.eu; site brands the degree 'AISS Master's program' (SITeS project)", V, "eacea.ec.europa.eu catalogue; aissprogram.eu", "8 Nov 2026", "–"),
 ("AISS (SITeS-JM)", "Assessment", "Documents 50% (min 35) + interview 50% (min 35); threshold 70; transcript assessed", V, "aissprogram.eu/application-procedure", "8 Nov 2026", "–"),
 ("AISS (SITeS-JM)", "Scholarship", "10 awards to top-10 ranked; ≤ 10% per nationality; €1,400/month; tuition waived; no other EMJM scholarship", V, "aissprogram.eu/scholarships", "8 Nov 2026", "–"),
 ("AISS (SITeS-JM)", "Window + English", "2 Dec 2026 – 2 Feb 2027; IELTS 6.5 (≥ 5.5 each) etc.; Ethiopian degree not exempt", V, "aissprogram.eu", "8 Nov 2026", "–"),
 ("CMU-Africa (Kigali, Rwanda)", "Application components", "Essays (3 + optional, ≤ 300 words), 1 video essay (no re-record), CV ≤ 2 pages, 3 recommender contacts, unofficial transcripts", S, "africa.engineering.cmu.edu FAQ / Step 1", "8 Nov 2026", "Open the portal; copy the exact questions"),
 ("CMU-Africa (Kigali, Rwanda)", "Deadlines and fees", "Early 15 Dec 2026; final 15 Jan 2027; no application fee; DET covered for early first-time applicants; English test optional by 15 Dec", S, "africa.engineering.cmu.edu", "8 Nov 2026", "Confirm on the admissions page"),
 ("CMU-Africa (Kigali, Rwanda)", "Tuition for Africans", "US$16,000 (2026–27)", S, "africa.engineering.cmu.edu/admissions/tuition", "7 Nov 2026", "Confirm"),
 ("CMU-Africa + Mastercard Foundation", "Target group", "Talented, service-oriented young people facing the highest social and economic barriers", V, "mastercardfdn.org partner page", "8 Nov 2026", "–"),
 ("SMACCs", "Scholarships", "95 EMJM scholarships over 4 intakes (2024–2027); 35 Heading-6 (31 NDICI, 4 IPA III); ≤ 10% per nationality; €1,400/month + tuition waiver", S, "smaccs.eu/practical/scholarships", "8 Nov 2026", "Confirm Ethiopia's category (e-mail 4)"),
 ("SMACCs", "Selection", "Global merit: degree content, study success, motivation letter, references (≥ 1 academic), work experience, R&D/publications, English", S, "smaccs.eu/admissions/admission-criteria", "8 Nov 2026", "Confirm"),
 ("SMACCs", "Deadline", "1 Feb 2027 (scholarship); interviews 7–18 Apr; results by 9 May", S, "smaccs.eu/admissions/deadlines", "8 Nov 2026", "Confirm on the live page"),
 ("SMACCs", "English", "Standardised test or prior instruction in English", S, "smaccs.eu FAQ", "8 Nov 2026", "Confirm MOI acceptance"),
 ("Italy IYT", "Eligibility (2026/27 call)", "Citizens + permanent residents of 23 countries incl. Ethiopia; born on/after 1 Jan 2000; English B2+", S, "investyourtalentapplication.esteri.it", "8 Nov 2026", "Private age check; wait for the 2027/28 call"),
 ("Italy IYT", "Value", "€10,800 for 9 months + tuition exemption", S, "IYT FAQ; unipi.it", "7 Nov 2026", "Confirm in the 2027/28 call"),
 ("China DUT (Type B)", "Rules", "Under 35; no GPA minimum found; CNY 400 fee; window 1 Nov 2026 – 15 Feb 2027; travel self-funded; CNY 3,000/month", S, "sie.dlut.edu.cn", "8 Nov 2026", "DUT reply (e-mail 2)"),
 ("China HUST", "GPA rule", "2024 guide: 3.0/4 or 75%; 2027 unconfirmed; fee CNY 600 vs 800", S, "iso.hust.edu.cn", "8 Nov 2026", "HUST reply (e-mail 3) before paying"),
 ("China NPU", "GPA rule", "CSC applicants need 80% average", S, "studyat.nwpu.edu.cn 2027 guide", "8 Nov 2026", "–"),
 ("China CSC Type A (Ethiopia)", "2027/28 call + agency number", "Not found", U, "et.china-embassy.gov.cn / moe.gov.et (searched)", "8 Nov 2026", "Ask the Ministry of Education (e-mail 1)"),
 ("EDISS", "Rules", "Academic merit 70/100; IELTS 6.5 no band < 6.0 (in-centre); deadline 12 Jan 2027 15:00 Helsinki", V, "master-ediss.eu", "8 Nov 2026", "–"),
 ("Stipendium Hungaricum", "2027/28 call", "Not yet published; 2026/27 terms: tuition, HUF 43,700/month, dorm, insurance", V, "stipendiumhungaricum.hu", "8 Nov 2026", "Watch the call"),
 ("ICCR India", "2026-27 terms", "130 places for Ethiopia; PG stipend INR 20,000/month; tuition; housing allowance; airfare depends on the scheme", S, "eoiaddisababa.gov.in; a2ascholarships.iccr.gov.in", "8 Nov 2026", "Portal about Feb 2027"),
 ("UK Chevening", "2027/28 window + rule", "4 Aug – 6 Oct 2026 (closed); ≥ 2 years' work experience after the bachelor's", V, "gov.uk news page", "7 Nov 2026", "Next window about Aug–Oct 2027"),
 ("UK Commonwealth", "Eligibility", "Ethiopia is not a Commonwealth country", S, "commonwealthscholarships.org", "7 Nov 2026", "–"),
 ("USA Fulbright", "Ethiopia call", "Run by the US Embassy Addis (blocked here)", U, "et.usembassy.gov (403)", "8 Nov 2026", "Check the embassy page early 2027"),
 ("Germany DAAD EPOS", "Rules", "Upper third of class; ≥ 2 years' work; degree ≤ 6 years; deadlines mostly Aug–Oct 2026", S, "daad.de (blocks automated access) via search", "8 Nov 2026", "Class-rank letter first"),
 ("Canada Manitoba", "Funding + supervisor", "Admitted MSc students funded at a minimum of $20,000/yr; supervisor support required before applying; deadline 1 Dec", V, "umanitoba.ca", "8 Nov 2026", "–"),
 ("Canada Calgary", "Supervisor + fee", "At least one professor must agree to supervise; fee CAD 145", V, "science.ucalgary.ca", "8 Nov 2026", "–"),
 ("IELTS (Ethiopia)", "Price + centre + results", "ETB 32,800, computer-delivered, British Council AAU (6 Kilo); most results within 2 days", V, "ielts.org", "8 Nov 2026", "–"),
]
wb = Workbook(); ws = wb.active; ws.title = "Verification register"
H = ["Opportunity", "Fact", "Value", "Status", "Source", "Checked", "What you must confirm"]
Wd = [28, 26, 70, 22, 40, 12, 34]
HDR = PatternFill("solid", fgColor="1F3864"); th = Side(style="thin", color="BFBFBF"); BD = Border(left=th, right=th, top=th, bottom=th)
for j, (h, w) in enumerate(zip(H, Wd), 1):
    c = ws.cell(row=1, column=j, value=h); c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9); c.fill = HDR; c.border = BD
    ws.column_dimensions[c.column_letter].width = w
for i, r in enumerate(R, 2):
    for j, v in enumerate(r, 1):
        c = ws.cell(row=i, column=j, value=v); c.font = Font(name="Arial", size=9, bold=(j == 4)); c.border = BD; c.alignment = Alignment(wrap_text=True, vertical="top")
n = len(R) + 1
for txt, col in ((V, "C6EFCE"), (S, "FFEB9C"), (U, "FFC7CE")):
    ws.conditional_formatting.add(f"D2:D{n}", FormulaRule(formula=[f'D2="{txt}"'], fill=PatternFill("solid", fgColor=col)))
ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:G{n}"
nt = wb.create_sheet("How to read")
for i, t in enumerate(["VERIFIED (E1): the official page was downloaded and read from this session on the date shown (copies kept as evidence).",
                       "SEARCH-INDEXED (E2): an official page was found and summarised by web search, but this session cannot open the site (network allowlist or bot protection). Treat as likely, and confirm in your browser before paying or submitting.",
                       "UNVERIFIED / NOT FOUND: no official source was found; do not rely on it.",
                       "Sites that cannot be opened from this session: smaccs.eu, africa.engineering.cmu.edu, investyourtalentapplication.esteri.it, Chinese university sites, daad.de, et.usembassy.gov, chevening.org (connection errors)."], 1):
    c = nt.cell(row=i, column=1, value=t); c.font = Font(name="Arial", size=10); c.alignment = Alignment(wrap_text=True)
nt.column_dimensions["A"].width = 140
wb.save(OUT); print("rows", len(R), "| verified", sum(r[3] == V for r in R), "| search-indexed", sum(r[3] == S for r in R), "| unverified", sum(r[3] == U for r in R))
