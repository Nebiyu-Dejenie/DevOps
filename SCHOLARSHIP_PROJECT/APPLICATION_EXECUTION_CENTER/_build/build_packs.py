"""ONE-TIME ASSEMBLER (already run on 8 Nov 2026; its source files are now in 14_ARCHIVE/phase6_superseded). Assembled SUBMISSION_PACK files for SMACCs, AISS and China from the reviewed source documents
(the written responses are taken verbatim from those files, so there is one text to maintain), and writes
the IYT and Germany/USA/Canada/UK assessment packs. Run from APPLICATION_EXECUTION_CENTER/_build."""
import os, re

W = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def section(path, start, stop_level="# "):
    """Return the markdown section beginning with heading text `start` up to the next heading of the same level."""
    s = open(os.path.join(W, path), encoding="utf-8").read()
    i = s.index(start)
    j = s.find("\n" + stop_level, i + len(start))
    body = s[i:j if j > 0 else len(s)]
    return body.split("\n", 1)[1].strip()  # drop the original heading line


def write(path, text):
    with open(os.path.join(W, path), "w", encoding="utf-8") as fh:
        fh.write(text.strip() + "\n")
    print("wrote", path)


HEADER = "**Nothing is submitted, signed, sent or paid for you.**\n\n**Evidence:**\n\n- **E1:** verified on the official site\n- **E2:** search-indexed official page, not opened from this session (confirm in your browser)\n"

# ---------------- SMACCs ----------------
S = "10_SMACCs/SMACCs_Application_Strategy.md"
write("10_SMACCs/SMACCs_SUBMISSION_PACK.md", f"""---
title: "SMACCs (Erasmus Mundus, Smart Cities and Communities): Submission Pack"
subtitle: "Scholarship deadline 1 Feb 2027 (E2). Submit by 25 Jan. ICT specialisation (International Hellenic University)"
---

{HEADER}
**smaccs.eu cannot be opened from this session, so every SMACCs fact is E2.**

# 1. Official links (E2)

| Purpose | URL |
|---|---|
| Programme home | https://www.smaccs.eu/ |
| Application (how and where to apply) | https://www.smaccs.eu/practical/application/ |
| Instructions (documents, formats) | https://www.smaccs.eu/admissions/instructions/ |
| Admission criteria | https://www.smaccs.eu/admissions/admission-criteria/ |
| Deadlines | https://www.smaccs.eu/admissions/deadlines/ |
| Scholarships | https://www.smaccs.eu/practical/scholarships/ |
| Fees | https://www.smaccs.eu/admissions/fees/ |
| FAQ | https://www.smaccs.eu/faq/ |

# 2. Eligibility and scholarship facts (checked)

{section(S, "# 1. Verified facts")}

# 3. Fit and path choice

{section(S, "# 2. Is your profile suitable? Honest fit analysis")}

# 4. Written response: motivation letter (complete draft; rewrite in your own words)

{section(S, "# 3. Motivation letter")}

# 5. Documents and formats

| Document | Format (confirm on the Instructions page) | Status |
|---|---|---|
| Online application with path choice + 2 alternative paths | Web form | You |
| CV | PDF | `06_FINAL_CV/CV_Europe_Academic` (needs HR facts) |
| Transcript + degree certificate | PDF scans of official documents | You (scan) |
| Motivation letter | PDF, 1–2 pages | Section 4 |
| **2 reference letters (at least 1 academic)** | Signed, on letterhead | Drafts in `16_RECOMMENDATIONS/` |
| English proof | Test, **or** proof of prior instruction in English (DTU registrar letter, item 4f) | Ask SMACCs to confirm (e-mail 4) |
| Passport | PDF | You |

# 6. Outstanding items

1. Confirm the facts in your browser (Application, Instructions, Deadlines).
2. HR letters, then fill the CV and the letter.
3. Two signed references (supervisor + SantimPay manager).
4. Registrar letter confirming English medium of instruction.
5. Choose the path names exactly as in the form.

# 7. Submission guide

1. Open the Application page and create the account.
2. Fill in your details; choose the ICT path + 2 alternatives.
3. Upload the documents in the order the form asks.
4. Tick the scholarship option.
5. **Submit by 25 Jan 2027** (deadline 1 Feb).
6. Interviews for the best applicants are held 7–18 April. Scholarship results by 9 May.

# 8. Interview strategy

{section(S, "# 4. Interview strategy")}

# 9. Final audit checklist

- ☐ Facts confirmed on smaccs.eu
- ☐ Path + 2 alternatives chosen
- ☐ Letter ≤ 2 pages, own words, no `[…]` left
- ☐ CV dates match the HR letters
- ☐ 2 references, at least 1 academic, signed
- ☐ English proof accepted
- ☐ Scholarship box ticked
- ☐ Submitted by 25 Jan; confirmation saved
""")

# ---------------- AISS ----------------
A = "09_AISS/AISS_Application_Package.md"
write("09_AISS/AISS_SUBMISSION_PACK.md", f"""---
title: "AISS (Erasmus Mundus): Submission Pack"
subtitle: "Official project name SITeS-JM, Sustainable Information Technologies for Societies Joint Masters. Degree programme branded 'AISS: Artificial Intelligence for Sustainable Societies'. Window 2 Dec 2026 – 2 Feb 2027"
---

{HEADER}

# 1. Identity and official links (E1)

- The EU's Erasmus Mundus catalogue (eacea.ec.europa.eu) lists **SITeS-JM, "Sustainable Information Technologies for Societies Joint Masters"**, with the website **aissprogram.eu**.
- That site presents the degree as the **"AISS Master's program"** within the **SITeS project**, co-funded by the EU through the Erasmus Mundus Joint Masters programme.
- **Partners:** Tallinn University (coordinator, Estonia), Lusófona University (Portugal), Tampere University (Finland).
- **Programme:** https://aissprogram.eu/
- **Application procedure:** https://aissprogram.eu/application-procedure/ (the application itself is filed in Tallinn University's DreamApply; follow the "Apply" link on that page from 2 Dec)
- **Scholarships:** https://aissprogram.eu/scholarships/

# 2. Eligibility and scholarship rules (E1, re-read 8 Nov 2026)

- **Admission:**
  - a recognised bachelor's degree plus an academic, professional or project background that supports the study path
  - education documents and transcript are assessed
  - documents count 50% (minimum 35/50); interview 50% (minimum 35/50); programme threshold 70/100
- **Scholarship:**
  - apply for it during admission
  - awarded only to the **10 highest-ranked** applicants
  - **at most 10% per nationality**
  - not combinable with another EU-funded Erasmus Mundus scholarship
  - you must study the full 24 months and meet the minimum mobility rules
  - **€1,400/month**, covering insurance, mobility travel, housing and food; tuition waived
- **English:** IELTS Academic or IELTS Online 6.5 (≥ 5.5 each); TOEFL iBT; PTE Academic 62 (≥ 59 each); Cambridge.
  - **An Ethiopian degree is not exempt.** The exemption covers only studies in the US, Canada, Australia, New Zealand, UK, Switzerland or the EU/EEA.
- **Your position:** eligible to apply. Competitive only through a strong portfolio, video and interview. **Not GPA-independent.**

# 3. Written and recorded responses (complete drafts)

## 3a. Motivation video script (≤ 1:30)

{section(A, "# 4. Motivation video")}

## 3b. Research interests and thesis directions

{section(A, "# 6. Research interests and thesis directions")}

## 3c. Social impact statement

{section(A, "# 6b. Social impact statement")}

## 3d. Portfolio

The portfolio document is `09_AISS/AISS_Portfolio_Summaries` (5 executed projects, in AISS format). Export it as **one PDF**.

## 3e. Self-funding plan (mandatory)

{section(A, "# 9. Self-funding plan")}

# 4. Documents and formats (E1)

| Document | Format | Status |
|---|---|---|
| Education documents + transcript | PDF scans (issued in English; no translation needed) | You (scan) |
| Passport identification page | PDF | You |
| English evidence | IELTS / PTE / TOEFL result | **Book only once you confirm AISS** |
| Motivation video | **≤ 1:30**, link submitted in DreamApply; you on screen at least half the time; own voice | Script 3a |
| Portfolio | One PDF of summaries (up to 5 projects); links allowed but not as a substitute | `AISS_Portfolio_Summaries` |
| CV | PDF | `06_FINAL_CV/CV_Europe_Academic` |
| Self-funding plan | PDF | 3e |

# 5. Outstanding items

1. Build demo repositories 3 and 4 (5 evenings).
2. Final-year project title + HR facts → portfolio and CV.
3. English test: decide and book.
4. Record the video.
5. Real figures for the self-funding plan.

# 6. Submission guide

1. From **2 Dec**, open the Application procedure page → DreamApply (Tallinn University) and create the account.
2. Choose AISS; **tick the Erasmus Mundus scholarship**.
3. Upload every document in section 4; submit the video link.
4. **Submit by 30 Jan** (deadline 2 Feb). AISS reviewers may give feedback on early applications.
5. If invited: 15-minute interview (10-minute presentation + 5 minutes of questions).

# 7. Interview

{section(A, "# 7. Interview (15 min")}

# 8. Final audit checklist

- ☐ Video ≤ 1:30; on screen at least 50%; no AI voice; not sped up
- ☐ Portfolio is a summary PDF (not a list of links), up to 5 items, each explained
- ☐ No extra files uploaded
- ☐ English result meets 6.5 / each part ≥ 5.5 (IELTS) or the equivalent
- ☐ Self-funding plan included
- ☐ Scholarship box ticked
- ☐ Submitted by 30 Jan; confirmation saved
""")

# ---------------- China (DUT) ----------------
C = "11_CHINA/China_Study_Plan_and_Emails.md"
write("11_CHINA/China_DUT_SUBMISSION_PACK.md", f"""---
title: "China CSC via Dalian University of Technology (Type B): Submission Pack"
subtitle: "Deadline 15 Feb 2027 (E2). University comparison: China_Application_Operation.xlsx"
---

{HEADER}
**DUT and CSC sites cannot be opened from this session, so every fact is E2.**

# 1. Official links (E2)

| Purpose | URL |
|---|---|
| DUT 2027 English-taught master's programmes | https://sie.dlut.edu.cn/info/1180/2485.htm |
| DUT CSC Chinese University Program (Type B) | https://sie.dlut.edu.cn/English/Scholarship1/CSC_Scholarship/Chinese_University_Program.htm |
| DUT Bilateral Program (Type A) | https://sie.dlut.edu.cn/English/Scholarship1/CSC_Scholarship/Bilateral_Program.htm |
| DUT application portal | http://iso.dlut.edu.cn |
| CSC online application system | https://www.campuschina.org/ (opens in your browser) |
| DUT admissions e-mail | dutsice@dlut.edu.cn |

# 2. Eligibility and coverage (E2)

| Item | Fact |
|---|---|
| Requirements | Non-Chinese citizen, bachelor's degree, under 35, good health. No minimum GPA found on any DUT page |
| English | TOEFL 80 / IELTS 5.5, **or** an English-medium certificate from your university (page versions differ: e-mail 2 asks) |
| Coverage | Tuition, free campus housing, CNY 3,000/month, insurance. **International travel is self-funded** |
| Fee | **CNY 400** for CSC applicants |
| Window | 1 Nov 2026 – **15 Feb 2027**; admission notices about Jul–Aug 2027 |
| Agency number | Type B: DUT's own code, **to be confirmed by DUT** (e-mail 2). **Never use another country's embassy number** |
| Programme | DUT's CS school lists Computer Science and Technology and Computer Technology; **which are taught in English must be confirmed** (e-mail 2) |

**Not recommended:**

- NPU: CSC applicants need an 80% average.
- HUST: on hold until it confirms whether 2.62/4.00 is acceptable.

# 3. Written responses (complete drafts)

## 3a. Study plan / research proposal

{section(C, "# 1. Study plan / research proposal")}

## 3b. E-mails (admissions and optional professor)

{section(C, "# 2. Pre-admission request e-mails")}

# 4. Documents and formats

{section(C, "# 4. Document checklist")}

# 5. Outstanding items

1. DUT's reply (English-taught CS master's, English proof, agency number).
2. 2 professor or associate-professor letters (DTU).
3. Notarised degree and transcript; physical examination form; police clearance.
4. Study plan names the confirmed programme and a DUT lab or professor.

# 6. Submission guide

{section(C, "# 5. DUT step by step (Type B)")}

# 7. Final audit checklist

- ☐ Programme confirmed English-taught
- ☐ Agency number confirmed by DUT (Type B)
- ☐ Same agency number on the CSC form and the DUT form
- ☐ Study plan tailored to DUT
- ☐ 2 letters from professors / associate professors
- ☐ Medical form and police clearance valid
- ☐ CNY 400 paid only after eligibility confirmed
- ☐ Submitted by 15 Feb; DUT application number and CSC number saved
""")

# ---------------- Italy IYT ----------------
write("12_OTHER_SCHOLARSHIPS/Italy_IYT_SUBMISSION_PACK.md", """---
title: "Italy: Invest Your Talent (IYT): Submission Pack (conditional)"
subtitle: "The 2027/28 call is not yet published; the 2026/27 call ran 20 Apr – 11 May 2026 (E2)"
---

**Evidence:** all facts are **E2**, from the official IYT call and university pages found by search (investyourtalentapplication.esteri.it cannot be opened from this session).

**Nothing is submitted for you.**

# 1. Official links (E2)

| Purpose | URL |
|---|---|
| Current call | https://investyourtalentapplication.esteri.it/SitoIYT/EN/current-call |
| How it works | https://investyourtalentapplication.esteri.it/SitoIYT/EN/how-does-it-work |
| FAQ | https://investyourtalentapplication.esteri.it/SitoIYT/EN/frequently-asked-questions |
| Regulations | https://investyourtalentapplication.esteri.it/SitoIYT/EN/pdf/Regulations_of_the_IYT_Scholaship.pdf |
| Example participating university (Pisa) | https://www.unipi.it/en/education/registration/enrolment-and-registration/enrolment-for-international-students/invest-your-talent-in-italy-scholarships/ |

# 2. Eligibility (2026/27 call; the 2027/28 rules may change)

| Rule | Your position |
|---|---|
| Citizens **and permanent residents** of one of 23 countries, including **Ethiopia** | Meets |
| **Age: born on or after 1 January 2000** (except renewals) | **Check privately against your passport.** If you were born before 2000, IYT is closed to you |
| Qualification valid for enrolment in the chosen Italian master's (engineering / high-tech, management, design) | Plausible: confirm with the chosen university |
| English certificate at **B2 or above** | IELTS (or a test the call accepts) |
| Not a current or former student of an Italian university | Meets |
| Only one programme may be chosen | – |
| **Value:** €10,800 for 9 months + tuition exemption | Nearly full (travel/insurance not stated) |

# 3. Written and recorded responses (drafts to use if the age rule allows)

**Motivational video (≤ 59 seconds).** About 130 words:

> I'm [name], a Computer Science graduate from Ethiopia. For [3½] years I have kept digital-payment systems running for merchants and banks in Addis Ababa. Italy's engineering universities combine strong computer engineering with close links to industry. In the MSc in [Computer Engineering] at [university] I want to learn to design secure, reliable networked systems, then bring that knowledge back to Ethiopia's fast-growing digital economy, and build lasting links between Italian and Ethiopian technology companies. Thank you.

**One-page CV:** a shortened `06_FINAL_CV/CV_Europe_Academic`.

# 4. Documents (2026/27 call)

- Transcript with final grade/GPA
- English certificate (B2+)
- Video (≤ 59 s)
- One-page CV
- Passport

All in English, French or Spanish, or officially translated.

# 5. Outstanding items

1. **Private age check** (decisive).
2. Watch for the 2027/28 call (expected about April 2027).
3. Choose one participating programme (e.g. Pisa MSc Computer Engineering) and check its admission requirements.

# 6. Final audit checklist

- ☐ Age rule met (private passport check)
- ☐ Programme on the 2027/28 participating list
- ☐ English B2+ certificate
- ☐ Video ≤ 59 s
- ☐ One-page CV
- ☐ Submitted before the call deadline
""")

# ---------------- Germany / USA / Canada / UK ----------------
write("12_OTHER_SCHOLARSHIPS/Germany_USA_Canada_UK_Assessment.md", """---
title: "Fully Funded Options in Germany, USA, Canada and the UK: Honest Assessment for the 2027 Intake"
subtitle: "Checked 8 Nov 2026. E1 = official page read; E2 = official page found by search; E3 = secondary"
---

**Bottom line:**

- For a **September 2027 start**, none of these four countries has an open, fully funded master's scholarship whose published rules you clearly meet.
- The realistic routes are **Chevening (UK) and Fulbright (USA) in the next cycle (2028 start)**, and **partial funding in Canada** if a professor agrees to supervise you.
- Nothing below is invented. Where a rule is not met, it says so.

| Country | Opportunity | Funding | Published rules vs you | 2027 intake? | Next step | Evidence |
|---|---|---|---|---|---|---|
| **UK** | **Chevening** | Full (tuition, stipend, flights, visa) for a 1-year master's | Requires **≥ 2 years' work experience after the bachelor's**: your 3½ years would qualify once the HR letters prove it. Also requires an offer from a UK university, so check that each course accepts a 2.62/4.00 | **Closed** (2027/28 window was 4 Aug – 6 Oct 2026) | Prepare for the **2028/29 window, expected Aug–Oct 2027**: 3 UK course choices, essays, references | E1 (gov.uk) |
| UK | Commonwealth Scholarships | Full | **Ethiopia is not a Commonwealth country**: not eligible | – | None | E1/E2 |
| UK | GREAT / university awards | Partial | Partial tuition only | – | Not recommended (not fully funded) | E3 |
| **USA** | **Fulbright Foreign Student Program (Ethiopia)** | Full | Run by the US Embassy in Addis (its page blocks automated access). English test and a strong academic/professional profile required | Not for 2027 | Check the US Embassy Addis Fulbright announcement (usually early in the year, for study starting the following year): 2028 | E2/E3 |
| USA | University assistantships (TA/RA) | Variable | Most master's assistantships are competitive and expect a GPA ≥ 3.0 | Low probability | Not recommended | E3 |
| **Canada** | **Manitoba MSc (thesis)** | Partial: **minimum $20,000/yr** for admitted MSc students | Needs a supervisor **before applying**; your last-two-years average (2.98) is borderline | Deadline 1 Dec 2026 | Only if a professor agrees now | E1 |
| Canada | Calgary MSc (thesis) | Partial (awards/assistantships) | A professor must agree to supervise; GPA is assessed | Early review 15 Jan 2027 | Optional | E1 |
| Canada | Mastercard Foundation at UBC / McGill | Full | Graduate admission minimums at or above 3.0 | Not realistic | None | E2 |
| **Germany** | **DAAD EPOS** (development-related master's) | Full | Requires **an above-average degree (upper third of your class)**, ≥ 2 years' work experience, degree ≤ 6 years old. Most course deadlines for 2027 were **Aug–Oct 2026**; no ICT course was confirmed; Ethiopia's eligibility for 2027/28 is unverified (DAAD pages block automated access) | Mostly closed | **Ask DTU for a class-rank letter.** If you are in the upper third, EPOS becomes a real 2028 option | E2/E3 |
| Germany | Political-foundation scholarships (e.g. Heinrich Böll, Friedrich Ebert, Konrad Adenauer) | Full | Usually require admission to a German university first and above-average grades | Not realistic for 2027 | Revisit after admission to a tuition-free German master's | E3 |
| Germany | Tuition-free public universities | Self-funded (blocked account ≈ living costs) | – | – | Not fully funded | E3 |

**What would change this picture:**

1. A **class-rank letter** showing the upper third (opens DAAD EPOS for 2028).
2. **HR letters** proving 2+ years of post-degree work (Chevening 2028/29).
3. A **professor's agreement** in Canada (Manitoba/Calgary, partial funding).
""")
print("done")
