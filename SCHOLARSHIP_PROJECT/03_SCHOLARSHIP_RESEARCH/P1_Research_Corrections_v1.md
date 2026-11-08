---
title: "Research Corrections (Phase 5, Priority 1)"
subtitle: "AISS, EDISS, CMU-Africa, Sweden SI and age handling. 7 November 2026"
---

**Evidence codes:**

- **E0:** your own document (here, the official DTU transcript)
- **E1:** read directly on the official website
- **E2:** official website found by search but not opened directly (CMU-Africa's site fails its security check from this machine). Check E2 items in your browser.

**Labels:**

- **A:** strong and realistic
- **B:** possible, but needs evidence
- **C:** low probability
- **D:** not eligible

These labels are not admission probabilities.

# Summary of what changed

| Programme | What the earlier files said | What is correct | New label |
|---|---|---|---|
| **AISS** (Erasmus Mundus) | "GPA is not a scored criterion." Priority A | Your transcript is one of the documents assessed, and documents are 50% of the score. Only the **10 highest-ranked** applicants get scholarships, and no more than 10% may share a nationality, so **at most one Ethiopian** | **B** |
| **EDISS** (Erasmus Mundus) | "Eligible; competitive" | **70 of 100 points** come from grades, with special attention to the areas where your grades are lowest. EDISS states that IT work experience does not compensate | **C** |
| **CMU-Africa + Mastercard Foundation** | "Eligible for admission (no GPA cut-off)" | Three separate decisions: admission, need-based aid and Scholar selection. No cut-off exists, but most admitted students have 3.0 or above (E2), so **2.62 is a weakness to offset** | **B** |
| **Sweden SI** | 3,000 hours as a simple sum | Each month now counts once even if two jobs overlap. Up to 3 work forms and 2 leadership forms | **B** until the HR letters arrive |
| **Türkiye Burslari** | Age-based reasoning | The **75% academic minimum** decides it. Age is checked only privately and not used | **D** |

# 1. AISS: corrected assessment

Source: aissprogram.eu application procedure and scholarship rules pages (E1, saved 7 Nov 2026).

## Official facts

- **Checklist item 01** is "Education documents and transcript". The page states: *"Your education documents and your complete application are assessed during the admissions process."*
- **Score structure:**
  - Documents: **50%** (minimum 35, maximum 50)
  - Interview: **50%** (minimum 35, maximum 50)
  - Programme threshold: **70/100**
  - Not everyone who passes the documents round is interviewed.
- **Documents criteria:**
  - 1a: technical soundness (video within the time limit, no extra files, the portfolio summarised and explained without needing links)
  - 2a: video shows your intention, interests, research topics and connection to the programme
  - 2b: portfolio of **projects actually executed**, with a summary and the skills used
  - 2c: originality
  - 2d: CV covering work, education, projects and social activities
- **Interview criteria (15 minutes: 10-minute presentation + 5 minutes of questions):**
  - 3a: punctual and follows the structure
  - 3b: convincing and authentic, with no reading or prepared written answers
  - 3c: able to express yourself
- **Scholarships:**
  - *"The published 2027 offer provides 10 scholarships"* (tuition waived + €1,400/month for 24 months).
  - *"Allocate awards to the ten highest-ranked eligible candidates."*
  - *"No more than 10% of the total number of scholarships are awarded to students from the same nationality."*
- **Self-funded tuition:** €2,000 per semester (€8,000 in total). The self-funding plan is **mandatory**; without it, the application is rejected.
- **English (one of these):**
  - IELTS Academic or IELTS Online: 6.5 overall, ≥ 5.5 in each part
  - TOEFL iBT: 92 (no section below 20) before 21 Jan 2026, or 4 (at least 4 in each section) from that date
  - **PTE Academic: 62 overall, ≥ 59 in each part**
  - Cambridge: 169

## What was wrong

The criteria list does not mention GPA, but the transcript is part of the documents that are scored. AISS does not publish how much grades weigh, so assume your 2.62 **can lower** your documents score.

## Competitiveness: B (possible, needs strong evidence)

**For you:**

- Most of the named document criteria (video, portfolio, originality, CV) reward real, explained projects and professional work, which is where your profile is strongest.
- The interview is half the score.

**Against you:**

- Only 10 scholarships in total, and in practice no more than one for an Ethiopian.
- Your transcript is visible to the assessors.
- Scoring 35/50 in each part is needed even for admission.

## Plan built only from real work

| Element | Built from | Missing evidence |
|---|---|---|
| Portfolio (up to 5 projects) | See `APPLICATION_EXECUTION_CENTER/06_FINAL_CV/Technical_Project_Portfolio` (P5). Use only items with evidence: (1) final-year project (A−, two parts); (2) SantimPay payment-infrastructure operations; (3) POS fleet / merchant onboarding, if your role is confirmed; (4) 1–2 public demo repositories you build (CI/CD and monitoring) | Project title (DR-17); HR letter (DR-06); your exact role in POS and merchant work |
| Video (≤ 1 min 30 s) | Script in `APPLICATION_EXECUTION_CENTER/09_AISS/AISS_Application_Package` | None; record after the portfolio is final |
| CV | CV master v0.9, academic variant | HR letters |
| Research interests | Reliable, low-cost infrastructure for AI services in low-resource settings; anomaly detection on payment telemetry for financial inclusion. **Both must connect to work you actually did** | Your confirmation of which topics you can talk about with depth |
| Interview | 10-minute presentation structure + question bank with answer outlines (package section 7) | Mock interviews (you) |

# 2. EDISS: course-by-course comparison

Sources: master-ediss.eu admission requirements, evaluation criteria and FAQ (E1, 7 Nov 2026). Grades: your official DTU transcript (E0, read visually; DTU scale A/A+ = 4, A− = 3.75, B+ = 3.5, B = 3, B− = 2.75, C+ = 2.5, C = 2, C− = 1.75, D = 1).

## What EDISS asks for

- A bachelor's degree of at least 180 ECTS in computer science, software engineering or an equivalent subject.
- Applicants are *"expected to have passed courses in several of the following areas: computer science (e.g. programming, algorithms and data structures), basic computer architecture, computer networks, software engineering and mathematics."*
- From the FAQ: *"Work experience in IT will not compensate for a lack of theoretical knowledge demonstrated through university studies."*

**Scoring (100 points):**

| Criterion | Points |
|---|---|
| Academic merit: *"grades… Special attention… programming, algorithms and data structures, computer architecture, computer networks, software engineering, mathematics and thesis work"* | 0–70 |
| Documented relevant experience | 0–15 |
| Motivation video (three set questions) | 0–10 |
| Two academic recommendation letters | 0–5 |

Ties are broken by academic merit.

## Your transcript against each area

| Area | EDISS reference | Your courses (credits, grade) | Credit-weighted average | Assessment |
|---|---|---|---|---|
| Programming | Named | Fundamentals of Programming I (5, B); Fundamentals of Programming II (5, C+); OOP (6, C+); **Advanced Programming (6, A−)**; Internet Programming (7, B−); Windows Programming (5, B) | **2.92** | Covered; your second-strongest area |
| Algorithms and data structures | Named | Data Structures & Algorithms (6, **C−**); Analysis of Algorithms (5, C) | **1.86** | Covered but **weak**. This is the area EDISS highlights first |
| Theory of computation | Part of "computer science" | **Complexity Theory (5, A−)**; **Compiler Design (6, A−)**; Formal Languages & Automata (5, B−); Logic in CS (4, C+) | **3.25** | Strong; your best evidence of theoretical ability |
| Mathematics | Named | Linear Algebra (5, C−); Applied Mathematics for CS (6, C); Discrete Mathematics (5, Fx then **D** at re-sit); Probability & Statistics (5, F, **retaken C+**); Numerical Analysis (5, B) | **2.05** | Covered but **weak**: two initial fails. **No calculus course** appears on the transcript |
| Software engineering | Named | Fundamentals of SE (5, C+); Object-Oriented SE (5, C+); HCI (5, B+) | **2.83** | Covered. No separate software testing, requirements or quality course |
| Databases | Not named | Fundamentals of Database (6, C+); Advanced Database (6, C+) | **2.50** | Covered. Relevant to EDISS's "data-intensive" focus |
| Computer systems and architecture | Named ("basic computer architecture") | Digital Electronics (5, C+); Computer Organization & Architecture (5, C+); Microprocessor & Assembly (5, C); Operating Systems (6, C+) | **2.38** | Covered, below average |
| Computer networks | Named | Data Communication & Networks (6, C); Network & System Administration (5, B−); Wireless & Mobile Computing (5, B+); Computer Security (5, B) | **2.77** | Covered |
| Distributed systems | Not named | Introduction to Distributed Systems (6, B−) | **2.75** | Covered |
| Thesis work | Named | **Final Project I (6, A−); Final Project II (7, A−)** | **3.75** | **Strongest area.** Title still needed (DR-17) |
| AI / data | Programme focus | Introduction to AI (6, B) | **3.00** | One course. No machine-learning or data-engineering course |

**Combined:**

- All EDISS-named areas, including the thesis: **2.72**
- The same areas without the thesis: **2.63**

*Year-1 "Logic" (C−) is left out of mathematics because its title suggests a general-education course. Including it lowers mathematics to 2.00.*

## Prerequisites

**None of the named areas is missing**, so you meet the formal requirement ("several of the following areas"). The weaknesses an assessor will see:

1. Low grades in core algorithms and data structures (C−, C).
2. Mathematics with two initial fails and no calculus.
3. No dedicated machine-learning or data-engineering course, although these are central to a "data-intensive intelligent software systems" degree.

## English

- **IELTS Academic 6.5 with no band below 6.0, taken in a test centre** (IELTS Online is not accepted).
- **Or TOEFL iBT:** 93 if taken before 21 Jan 2026; from that date, 4.5 with no band below 4. TOEFL Home Edition is not accepted.
- **Your Ethiopian degree does not exempt you.** The exemption covers only degrees completed on-site in Australia, Canada, New Zealand, the UK, the USA, Switzerland or the EU/EEA.

## Realism: C (low probability)

EDISS does not publish how grades are converted to its 70 points, so the figures below are an **illustration only**, using a simple proportional conversion, not EDISS's own formula:

| | Points |
|---|---|
| Your EDISS-area average 2.72/4.00 × 70 | about **48 of 70** |
| Full marks on everything else | 15 + 10 + 5 = **30** |
| Illustrative maximum total | about **78** |
| Applicant averaging 3.6 in these areas, before any other points | about **63** |

Even with perfect non-academic points, you would trail such an applicant who also scores well on the rest. EDISS breaks ties by academic merit and states that work experience does not compensate.

**Recommendation:**

- Apply in the **first round only** (16 Nov 2026 – 12 Jan 2027, 15:00 Helsinki), and only because the extra cost is small: you need IELTS for AISS anyway, and DTU academic letters for other applications.
- Spend no more than about one day on it. Do not let it take time from SI, AISS or CMU-Africa.
- Rounds 2 (tuition waiver only) and 3 (self-funded) are **not fully funded**.

# 3. CMU-Africa: three separate decisions

CMU-Africa pages are E2 (found by search; the site fails its security check from this machine). Mastercard Foundation pages are E1. **Check every E2 item in your browser before relying on it.**

| | 1. Admission to the MSIT | 2. CMU-Africa need-based financial aid | 3. Mastercard Foundation Scholars Program |
|---|---|---|---|
| **Who decides** | CMU-Africa admissions | CMU-Africa financial aid | CMU-Africa, under its Mastercard Foundation partnership |
| **Who can be considered** | Bachelor's graduates (or final-year students) | Admitted students; aid is described as being for admitted candidates of **African nationality** (E2, French FAQ) | Admitted applicants only. Target group: *"highly talented, service-oriented young people who face the highest social and economic barriers to opportunity"* (E1) |
| **Criteria** | Strong programming skills; high proficiency in English; transcript; written essays + video essay; CV; **three recommenders**, whom CMU contacts directly (MSIT applicants do not upload letters) (all E2) | Financial need, assessed from a financial-aid form that appears on your checklist **after** you submit the application (E2) | Academic talent, commitment to giving back, and potential to lead (E2, CMU-Africa MCF page) |
| **GPA** | No cut-off, but *"most students have graduated with a GPA of 3.0 or above"* (E2). Report 2.62/4.00 **exactly as printed**; do not recalculate or round (E2) | Not a stated factor | "Academic talent" is assessed. 2.62 **does not on its own show it**; the final-year GPA of 3.20 and your A− courses must carry this |
| **Cost and coverage** | No application fee. DET fee covered for first-time applicants who apply by **15 Dec** (E2) | 2026–27 tuition for African students: **US$16,000** (E2; another page gives a different international figure) | Full cost if selected (tuition, living, travel, insurance); Class of 2026 had **57 Scholars** (E2) |
| **English** | Tests are listed as **optional**, but after 15 Dec only applications with a valid DET, IELTS or TOEFL are considered (E2). Recommended scores differ between pages (TOEFL 81, 87 or 93; IELTS 6.5; DET 120) | – | – |
| **Evidence you need** | Transcript and grading key (have); CV with HR-verified facts; 2–3 public code repositories; essays | Honest household and personal income figures, dependants, savings, other funding | Real leadership examples, community service and a believable plan to give back; honest description of barriers you faced |

## Strategy for each scored theme

**Academic:**

- Do not hide the 2.62.
- Show the upward trend: final year **3.20**, with **A−** in Advanced Programming, Compiler Design, Complexity Theory and both parts of the final project.
- Prove programming ability with public repositories that admissions staff can open.

**Financial need:**

- Describe it accurately. You are employed in fintech, and CMU will see your salary.
- The honest case is the gap between an Ethiopian salary and US$16,000 tuition plus living costs in Kigali, plus any real family obligations.
- Do not exaggerate. The Mastercard Foundation verifies.

**Leadership:**

- Use only examples you can confirm: who you led or trained, what you decided, and the result.
- Dereja co-training (2022) is a candidate example once the dates and numbers are confirmed.

**Community impact:**

- Link your payments work to the people who use it (merchants, POS users), but claim only what you did.
- Give a concrete plan for after the degree: a role, and an Ethiopian problem in digital payments infrastructure you will work on.

# 4. Sweden SI: work hours without double counting

**Official rules (E1, si.se, 7 Nov 2026):**

- At least **3,000 hours before 8 February 2027**, from **at most three employers**.
- Valid work: full-time, part-time, freelancing or entrepreneurship. Internships do not count.
- Formula: hours per week × weeks per year × years.
- Up to **three work-experience forms** and **two leadership forms**, merged into **one PDF of no more than 5 pages**: work forms first, then leadership forms, newest first, one page each.
- Each form is typed, printed, signed and **stamped** by the employer.
- Use only the SI 2027/2028 template, which is published on 13 Nov 2026.

**Overlapping jobs:** SI's FAQ does not say whether hours from two jobs held at the same time may be added. To be safe, the new sheet **"SI hours - no double count"** in `Employment_Evidence_Collection_Form_v2.xlsx` counts each month only once, using the job with more hours that month.

- It also shows the higher "added together" figure, labelled *ask SI before relying on it*.
- It counts only rows with E0 or E1 evidence.
- It checks the three-employer limit.
- **Test with sample data:** a full-time job (40 h × 48 weeks) overlapping a part-time job for six months gave 8,360 safe hours against 8,840 added.

**Rule of thumb:** at 40 hours a week for 48 weeks a year (160 hours a month), you reach 3,000 hours after about **19 months** of full-time work before 8 Feb 2027. Your result depends on the HR letters, which must state for each job title:

- the dates
- the weekly hours
- whether the job was full-time

**Still to check on 13 Nov:** whether SantimPay's three job titles go on **one** form (one employer) or need separate forms. The template instructions will say. Planning assumption: one form per employer.

# 5. Personal information and age

**What I fixed:**

- Three project files contained a sentence that revealed your date of birth:
  - the master profile v3
  - the opportunity register v3 (.md/.docx/.pdf)
  - the Command Center build script and workbook
- The sentence is removed and the files are rebuilt.
- It still exists in **local commit a5c7983, which has not been pushed**. Removing it from history needs your approval.

**Rule from now on:**

- Your passport is the private source for age checks. Reports do not record your date of birth.
- No eligibility decision uses age unless you have confirmed the check privately.

**Türkiye Burslari (E1, criteria page and all six FAQ categories read on 7 Nov 2026):**

- **Age:** the rule for master's applicants is *"being under the age of 30"*. The site does **not** say on which date age is measured. **Not used.**
- **Academic minimum:** 75% for graduate applicants. Your 2.62/4.00 is 65.5% on a straight conversion. The portal applies its own conversion when you enter your grades.
- **Language year:** a one-year Turkish preparatory course is normally required.
- **Verdict: D** (the academic minimum is not met), regardless of age.

**Other age rules to check privately:**

- The CSC master's age limit (covered in the CSC deep-dive).
- The Pan African University (call closed).
