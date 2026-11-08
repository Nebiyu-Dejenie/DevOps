// Builds the Document Collection Kit (.docx). Run: node build_kit.js  (from this _build folder)
// All factual statements reflect verified evidence (E0/E1) as of 7 Nov 2026. [Brackets] = applicant must fill.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, LevelFormat, PageBreak, Footer, PageNumber,
} = require("docx");

const OUT = path.resolve(__dirname, "..");
const W = 9638; // A4 content width in DXA with 2 cm margins
const FONT = "Arial";

const numbering = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 270 } } } },
      { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 1080, hanging: 270 } } } }] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] },
  ],
};
let numInstance = 0;

const styles = {
  default: { document: { run: { font: FONT, size: 21 } } },
  paragraphStyles: [
    { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", run: { font: FONT, size: 34, bold: true, color: "1F3864" },
      paragraph: { spacing: { before: 0, after: 120 }, keepNext: true } },
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: FONT, size: 27, bold: true, color: "1F3864" }, paragraph: { spacing: { before: 300, after: 120 }, keepNext: true, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: FONT, size: 23, bold: true, color: "2E5597" }, paragraph: { spacing: { before: 220, after: 80 }, keepNext: true, outlineLevel: 1 } },
  ],
};

// ---------- helpers ----------
const runs = (t, o = {}) => {
  // supports **bold** segments
  const parts = String(t).split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
  return parts.map((p) => p.startsWith("**") ? new TextRun({ text: p.slice(2, -2), bold: true, font: FONT, ...o })
                                              : new TextRun({ text: p, font: FONT, ...o }));
};
const P = (t, o = {}) => new Paragraph({ children: runs(t, o.run || {}), spacing: { after: o.after ?? 100, before: o.before ?? 0 }, alignment: o.align });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT })] });
const TITLE = (t) => new Paragraph({ style: "Title", children: [new TextRun({ text: t, font: FONT })] });
const B = (t, level = 0) => new Paragraph({ numbering: { reference: "bullets", level }, children: runs(t), spacing: { after: 60 } });
const NUMS = (items) => { numInstance++; const inst = numInstance;
  return items.map((t) => new Paragraph({ numbering: { reference: "numbers", level: 0, instance: inst }, children: runs(t), spacing: { after: 60 } })); };
const BR = () => new Paragraph({ children: [new PageBreak()] });
const SP = () => new Paragraph({ children: [new TextRun("")], spacing: { after: 60 } });
const border = { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
const borders = { top: border, bottom: border, left: border, right: border };

// shaded "copy-ready" box: array of paragraph strings
const BOX = (label, lines, fill = "F2F6FC") => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [W],
  rows: [
    new TableRow({ children: [new TableCell({ width: { size: W, type: WidthType.DXA }, borders,
      shading: { fill: "1F3864", type: ShadingType.CLEAR, color: "auto" }, margins: { top: 60, bottom: 60, left: 140, right: 140 },
      children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, color: "FFFFFF", font: FONT, size: 20 })] })] })] }),
    new TableRow({ children: [new TableCell({ width: { size: W, type: WidthType.DXA }, borders,
      shading: { fill, type: ShadingType.CLEAR, color: "auto" }, margins: { top: 120, bottom: 120, left: 160, right: 160 },
      children: lines.map((l) => new Paragraph({ children: runs(l), spacing: { after: 80 } })) })] }),
  ],
});
// two-column form table: rows = [label, value]
const FORM = (rows, c1 = 3600) => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [c1, W - c1],
  rows: rows.map(([a, b], i) => new TableRow({ children: [
    new TableCell({ width: { size: c1, type: WidthType.DXA }, borders, shading: { fill: "EAF0F8", type: ShadingType.CLEAR, color: "auto" },
      margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: runs(a, { bold: true, size: 20 }) })] }),
    new TableCell({ width: { size: W - c1, type: WidthType.DXA }, borders, margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: runs(b || " ", { size: 20 }) })] }),
  ] })),
});
// multi-column grid table with header
const GRID = (headers, rows, widths) => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: widths,
  rows: [
    new TableRow({ tableHeader: true, children: headers.map((h, i) => new TableCell({ width: { size: widths[i], type: WidthType.DXA }, borders,
      shading: { fill: "1F3864", type: ShadingType.CLEAR, color: "auto" }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: "FFFFFF", font: FONT, size: 18 })] })] })) }),
    ...rows.map((r) => new TableRow({ children: r.map((c, i) => new TableCell({ width: { size: widths[i], type: WidthType.DXA }, borders,
      margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [new Paragraph({ children: runs(c || " ", { size: 18 }) })] })) })),
  ],
});

function save(name, title, children) {
  const doc = new Document({
    creator: "Scholarship Project", title, styles, numbering,
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
        new TextRun({ text: `${title}  |  page `, font: FONT, size: 16, color: "808080" }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "808080" })] })] }) },
      children }],
  });
  return Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(path.join(OUT, name), buf); console.log("written", name); });
}

const HOWTO_COMMON = [
  "Replace every [bracketed] field with your real information. Delete anything that does not apply.",
  "Nothing in these drafts states a date, title or result on the organisation's behalf: **they must state the facts from their own records.**",
  "If an official letter differs from your CV, **the official letter is correct** and your CV will be changed to match it.",
  "Keep the original paper letter. Scan it in colour to PDF (300 dpi) and save it in 01_ACADEMIC_DOCUMENTS or 02_PROFESSIONAL_CV.",
];

// =====================================================================================
// 1. SantimPay HR request
// =====================================================================================
const hrFields = (company) => [
  `Full name of the employee (as in company records)`,
  `Exact date of joining ${company}: day/month/year in the **Gregorian calendar** (the Ethiopian-calendar date may be added in brackets)`,
  `**Every job title held**, each with its own start and end date (promotion or transfer history)`,
  `Current status: still employed, or the exact last working day`,
  `Employment type: full-time or part-time (and permanent, contract or other), with the **normal working hours per week**`,
  `A short description of the main duties in each position`,
  `Issuing officer's name, title, signature, official contact (email/phone), date of issue, and the **official company stamp** on company letterhead`,
];
const santim = [
  TITLE("SantimPay HR: Employment Certificate Request"),
  P("Purpose: obtain the official employer evidence needed for SI (Sweden), CMU-Africa, Erasmus Mundus and other applications. This letter is the most important document to resolve discrepancies DR-06 and DR-11.", { run: { italics: true, color: "404040" } }),
  H1("How to use this file"),
  ...HOWTO_COMMON.map((t) => B(t)),
  B("Ask for the certificate **in English**. Scholarship bodies require English documents, and a translation adds cost and delay."),
  B("The Swedish Institute will publish its own **Proof of Work and Leadership Experience** form on **13 November 2026**. Bring that one-page form to HR as well. It must be typed, printed, signed and stamped (E1, si.se). HR may sign it on behalf of the company."),
  H1("A. Formal letter (print, sign and hand in to HR)"),
  BOX("FORMAL LETTER: copy into a document, print and sign", [
    "[Your full name as in your passport]",
    "[Your current position], [Department]",
    "[Employee ID, if any]",
    "Addis Ababa, [date, e.g. 10 November 2026]",
    " ",
    "To: The Human Resources Department",
    "SantimPay Financial Solution S.C., Addis Ababa",
    " ",
    "**Subject: Request for an employment certificate for graduate scholarship applications**",
    " ",
    "Dear [Ms./Mr. HR Manager's name / Sir or Madam],",
    "I am applying for fully funded master's degree scholarships for the 2027 academic year. The scholarship bodies require an official employment certificate issued by my employer. I would be grateful if Human Resources could issue a certificate, on SantimPay letterhead and in English, stating:",
    ...hrFields("SantimPay").map((t, i) => `${i + 1}. ${t}`),
    "Salary information is not needed at this stage. I will ask for it separately only if a scholarship specifically requires it.",
    "One of the scholarships (Swedish Institute) also requires its own one-page form, \"Proof of Work and Leadership Experience\". It must be completed on the Swedish Institute's template, signed and stamped by the employer. The template will be published on 13 November 2026, and I will bring it to you as soon as it is available. It may be signed by the HR Manager, or by another person who can confirm my work on behalf of the company.",
    "Because my first application deadline is on 15 December 2026, I would be very grateful to receive the certificate by [20 November 2026]. I am happy to provide any information that would help.",
    "Thank you for your support.",
    "Sincerely,",
    "[Signature]",
    "[Your full name]  |  [Phone]  |  [Email]",
  ]),
  H1("B. Email version"),
  BOX("EMAIL: copy and paste", [
    "**To:** [HR email address]",
    "**Subject:** Request for employment certificate (graduate scholarship applications) – [Your full name]",
    " ",
    "Dear [Name],",
    "I am applying for fully funded master's scholarships for 2027 and need an official employment certificate from SantimPay, in English, on letterhead, signed and stamped. Could it please state:",
    "(1) my date of joining (Gregorian calendar); (2) every job title I have held, each with start and end dates; (3) my current status; (4) full-time or part-time and normal weekly hours; (5) my main duties in each role; and (6) the issuing officer's name, title, contact, date and the company stamp.",
    "A separate one-page Swedish Institute form (\"Proof of Work and Leadership Experience\") will follow after 13 November; I will bring it in person.",
    "My first deadline is 15 December 2026, so receiving the certificate by [20 November 2026] would help me greatly. I have also handed in a signed request letter.",
    "Thank you very much,",
    "[Your full name]  |  [Position]  |  [Phone]",
  ]),
  H1("C. Short message (Telegram / WhatsApp)"),
  BOX("SHORT MESSAGE", [
    "Hello [Name], I'm applying for master's scholarships and need an official SantimPay employment certificate (in English): joining date, all my job titles with dates, full-time/hours per week, main duties, signed and stamped. I've emailed the details and handed in a letter. Could HR issue it by [20 Nov]? Thank you! – [Your name]",
  ]),
  BR(),
  H1("Annex 1: Content checklist for HR (blank, for HR to complete)"),
  P("You may attach this page to the request so HR knows exactly what the scholarship bodies need. **Leave every value blank. HR fills it in from company records.**"),
  FORM([
    ["Employee full name", ""], ["Date of joining (dd/mm/yyyy, Gregorian)", ""],
    ["Position 1: title / from / to", ""], ["Position 2: title / from / to", ""], ["Position 3: title / from / to", ""],
    ["Employment type (full-time / part-time; permanent / contract)", ""], ["Normal working hours per week", ""],
    ["Current status (employed / last working day)", ""], ["Main duties (brief)", ""],
    ["Issued by (name, title)", ""], ["Official contact (email / phone)", ""], ["Date of issue", ""], ["Signature and company stamp", ""],
  ]),
  H1("Annex 2: Check when you receive the letter"),
  ...NUMS([
    "On company letterhead, dated, signed, with the official stamp.",
    "Your name is spelled **exactly** as in your passport. If not, ask for a correction before using it.",
    "The joining date is in the **Gregorian** calendar (dd/mm/yyyy).",
    "**Each** job title has its own dates. Record them in Employment_Evidence_Collection_Form.xlsx.",
    "Full-time/part-time and weekly hours are stated (needed for SI's 3,000-hour rule: hours/week × weeks/year × years, from at most 3 employers).",
    "It is in English. Scan it in colour to PDF and keep the original safe.",
    "Update the Discrepancy Log (DR-06) and send me the scan. I will then finalise the CV.",
  ]),
];

// =====================================================================================
// 2. Kifiya HR request
// =====================================================================================
const kifiya = [
  TITLE("Kifiya HR: Employment Certificate Request"),
  P("Purpose: confirm your exact Kifiya Financial Technology PLC employment record. Your CV versions show three different date ranges, and the period overlaps your SantimPay employment (DR-07). Only Kifiya's own record can settle this.", { run: { italics: true, color: "404040" } }),
  H1("How to use this file"),
  ...HOWTO_COMMON.map((t) => B(t)),
  B("Kifiya does **not** need to explain anything about SantimPay. The **employment type and weekly hours** on each employer's letter will show, factually, how the two roles fitted together."),
  B("If you no longer work at Kifiya, address the request to HR as a former employee. The same content applies."),
  H1("A. Formal letter (print, sign and send or hand in)"),
  BOX("FORMAL LETTER: copy into a document, print and sign", [
    "[Your full name as in your passport]",
    "[Address in Addis Ababa]  |  [Phone]  |  [Email]",
    "Addis Ababa, [date]",
    " ",
    "To: The Human Resources Department",
    "Kifiya Financial Technology PLC, Addis Ababa",
    " ",
    "**Subject: Request for an employment (experience) certificate for scholarship applications**",
    " ",
    "Dear [Ms./Mr. HR Manager's name / Sir or Madam],",
    "I [am / was] employed at Kifiya Financial Technology PLC and am now applying for fully funded master's degree scholarships for 2027. The scholarship bodies require an official employment certificate. I would be grateful if Human Resources could issue a certificate, on Kifiya letterhead and in English, stating:",
    "1. My full name as in company records;",
    "2. My exact start date and, if applicable, my exact end date (day/month/year, Gregorian calendar);",
    "3. Each job title I held, with the dates of each;",
    "4. The type of engagement (for example: permanent full-time employee, part-time, fixed-term contract, consultant, or assignment at client sites) and my normal working hours per week;",
    "5. Whether the role was performed on-site, remotely or in a hybrid arrangement;",
    "6. A short description of my main duties;",
    "7. The issuing officer's name, title, signature, official contact details, the date, and the official company stamp.",
    "Salary information is not required.",
    "[If still employed:] The Swedish Institute will also require its own one-page \"Proof of Work and Leadership Experience\" form, signed and stamped by the employer. I will bring it after it is published on 13 November 2026.",
    "I would be very grateful to receive the certificate by [20 November 2026], as my first deadline is in mid-December 2026.",
    "Thank you for your assistance.",
    "Sincerely,",
    "[Signature]",
    "[Your full name]",
  ]),
  H1("B. Email version"),
  BOX("EMAIL: copy and paste", [
    "**To:** [Kifiya HR email]",
    "**Subject:** Employment certificate request – [Your full name] – scholarship applications",
    " ",
    "Dear [Name],",
    "I am applying for master's scholarships for 2027 and need an official employment certificate from Kifiya, in English, on letterhead, signed and stamped. Could it please state: my start date and (if applicable) end date in the Gregorian calendar; each job title with its dates; the type of engagement and normal weekly hours; whether the work was on-site, remote or hybrid; my main duties; and the issuing officer's details, date and stamp.",
    "I would be grateful to receive it by [20 November 2026]. Thank you very much for your help.",
    "Kind regards,",
    "[Your full name]  |  [Phone]",
  ]),
  H1("C. Short message (Telegram / WhatsApp)"),
  BOX("SHORT MESSAGE", [
    "Hello [Name], I'm applying for master's scholarships and need an official Kifiya employment certificate in English: start/end dates, job title(s), employment type and weekly hours, main duties, signed and stamped. I've sent an email with details. Could HR help by [20 Nov]? Thank you! – [Your name]",
  ]),
  BR(),
  H1("Annex: Check when you receive the letter"),
  ...NUMS([
    "Letterhead, date, signature, official stamp, contact details.",
    "Name spelled exactly as in your passport.",
    "Start and end dates in the Gregorian calendar.",
    "Employment type and **weekly hours** stated. These explain the overlap with SantimPay factually.",
    "Compare with the SantimPay certificate. If the combined hours look implausible, talk to me before using either letter in an application.",
    "Record the data in Employment_Evidence_Collection_Form.xlsx and update DR-07.",
  ]),
];

// =====================================================================================
// 3. University Registrar request
// =====================================================================================
const registrar = [
  TITLE("Debre Tabor University: Registrar Request Pack"),
  P("Purpose: obtain the official academic documents that several programmes require **directly from the university**. These rules were verified on the official programme websites on 7 Nov 2026.", { run: { italics: true, color: "404040" } }),
  H1("Why this matters (verified rules)"),
  GRID(["Programme", "What the university must do", "Evidence"], [
    ["Sweden: University Admissions (SI scholarship)", "Send the **official transcript directly**: from an official registrar e-mail to **verifications@uhr.se**, **or** in a sealed envelope to: PostNord Strålfors AB, Att: University Admissions in Sweden, SE-190 81 Rosersberg, Sweden. Envelope must contain only the transcript; no cover sheet. Transcripts must be dated. Supporting-document deadline **1 Feb 2027**.", "E1 universityadmissions.se"],
    ["EDISS (Erasmus Mundus)", "For African degrees, an **official transcript mailed by the university**, stamped and signed, in an envelope sealed with the **university's official stamp across the back flap** (generic \"sealed\" stamps rejected) to: Åbo Akademi University / Kiinteistöpalvelut, EDISS programme office, Dept. of Information Technology – Bldg. Arken, Porthaninkatu 3, FIN-20500 Turku, Finland. Arrival by 30 Jul 2027.", "E1 master-ediss.eu"],
    ["EDISS application form", "You must state your **CGPA, the maximum CGPA and the grading scale** (highest/lowest, pass/fail).", "E1 master-ediss.eu"],
    ["Korea GKS", "A rank/percentile certificate helps only if you are in the **top 20%**.", "E1 studyinkorea.go.kr"],
    ["Univ. of Manitoba / Calgary", "GPA of the most recent two years is assessed. A certified last-two-years GPA helps.", "E1 umanitoba.ca"],
    ["Hungary (Szeged)", "May request **detailed course syllabuses**.", "E1 stipendiumhungaricum.hu"],
  ], [2400, 5600, 1638]),
  H1("A. Formal request letter (print, sign and submit to the Registrar)"),
  BOX("FORMAL LETTER: to the Registrar and Alumni Directorate", [
    "NEBYU DEJENIE TILAHUN [use exactly the name on your transcript]",
    "BSc in Computer Science (Extension programme), graduated 14 July 2022 (7 Hamle 2014 E.C.)",
    "Student ID: [your ID number, as on the transcript]",
    "[Phone]  |  [Email]  |  Addis Ababa, [date]",
    " ",
    "To: The Director, Registrar and Alumni Directorate",
    "Debre Tabor University, Debre Tabor",
    " ",
    "**Subject: Request for official academic documents for international master's applications**",
    " ",
    "Dear Sir/Madam,",
    "I graduated from Debre Tabor University with a Bachelor of Science in Computer Science on 14 July 2022. I am applying for master's degree scholarships abroad for 2027, and the universities require official documents issued by, and in some cases sent directly from, the Registrar. I respectfully request the following, and I will pay any applicable service fees:",
    "1. My permanent (printed) degree diploma, if it is ready. I hold the temporary certificate of graduation.",
    "2. [Three] official copies of my transcript (Student Academic Record), signed, stamped and dated, each in an envelope sealed with the University's official stamp across the back flap.",
    "3. Direct submission of my official transcript to University Admissions in Sweden, preferably by e-mail from the Registrar's official e-mail address to verifications@uhr.se, or by post in a sealed envelope to: PostNord Strålfors AB, Att: University Admissions in Sweden, SE-190 81 Rosersberg, Sweden. (The envelope should contain only the transcript.)",
    "4. An official letter (To Whom It May Concern) confirming: (a) the degree and the date of graduation; (b) my cumulative GPA of 2.62 on a maximum scale of 4.00; (c) the University's grading scale (letter grades, grade points, descriptors, and which grades are passing or failing); (d) that the academic years shown on the transcript (2017/2018 to 2021/2022) follow the Gregorian calendar; (e) that the Extension programme leads to the same Bachelor of Science degree, with a normal duration of five academic years; and (f) that the medium of instruction was English.",
    "5. If the University issues one: a certificate of my class rank or percentile within my graduating cohort.",
    "6. If possible: a certified statement of my GPA for the last two academic years (2020/2021 and 2021/2022).",
    "7. Guidance on whether my documents require authentication by the Ministry of Education or other authority for use abroad, and the procedure.",
    "Later in the process, I may also request that an official transcript be mailed directly to one admitting university in Finland. I will provide the address and form at that time.",
    "Because the first deadlines fall in December 2026 and January 2027, I would be very grateful if items 2–4 could be completed by [5 December 2026].",
    "Thank you for your assistance.",
    "Respectfully,",
    "[Signature]",
    "[Full name]",
  ]),
  H1("B. Letter to the Head, Department of Computer Science"),
  BOX("LETTER: final-year project and course descriptions", [
    "To: The Head, Department of Computer Science, Faculty of Technology, Debre Tabor University",
    "**Subject: Request for final-year project confirmation and course descriptions**",
    "Dear [Dr./Mr./Ms. Name],",
    "I graduated in Computer Science in July 2022 and am applying for master's programmes abroad. I respectfully request: (1) a short letter confirming the title of my final-year project (Final Project I and II), the supervisor's name, and the grades recorded; and (2) official course descriptions (syllabi) for the courses listed below, which some universities request to assess prerequisites:",
    "Data Structures and Algorithms; Analysis of Algorithms; Operating System; Data Communication and Computer Networking; Network and System Administration; Computer Security; Introduction to Distributed Systems; Compiler Design; Complexity Theory; Advanced Programming; Fundamentals of Database; Advanced Database System; Fundamentals of Software Engineering; Object Oriented Software Engineering; Introduction to Artificial Intelligence; Discrete Mathematics and Combinatorics; Linear Algebra; Numerical Analysis; Probability and Statistics.",
    "Thank you very much for your support.",
    "Respectfully,",
    "[Signature]  [Full name]  [Phone]  [Email]",
  ]),
  H1("C. Email to the Registrar (if they accept e-mail requests)"),
  BOX("EMAIL", [
    "**To:** [Registrar's official e-mail]",
    "**Subject:** Request for official transcripts and direct submission to University Admissions in Sweden – NEBYU DEJENIE TILAHUN, BSc Computer Science (2022)",
    "Dear Sir/Madam,",
    "I graduated in BSc Computer Science on 14 July 2022 (Student ID [ID]). For master's applications abroad, I kindly request: official sealed transcripts; direct submission of my official transcript to University Admissions in Sweden (official e-mail to verifications@uhr.se, or sealed post to PostNord Strålfors AB, Att: University Admissions in Sweden, SE-190 81 Rosersberg, Sweden); and an official letter confirming my CGPA (2.62 out of 4.00), the grading scale, the Gregorian academic calendar used on the transcript, the Extension programme and its duration, and English as the medium of instruction. I have attached my formal request letter and will pay the applicable fees.",
    "Thank you very much.",
    "[Full name]  |  [Phone]",
  ]),
  BR(),
  H1("Annex: Practical checklist"),
  ...NUMS([
    "Go in person if possible: Ethiopian registrar offices often require a written application, an ID and a fee receipt. Ask about processing time and fees.",
    "Bring copies of your temporary certificate and transcript, and your ID.",
    "Ask the Registrar to **date** every transcript (University Admissions Sweden requires dated transcripts).",
    "Confirm in writing (e-mail or receipt) when the transcript is sent to University Admissions Sweden. Keep proof for the 1 Feb 2027 deadline.",
    "Check that the name on all documents is identical (UA Sweden: names and birthdates must match on all documents, or you need official documentation explaining the difference).",
    "Scan everything in colour to PDF and store it in 01_ACADEMIC_DOCUMENTS.",
  ]),
];

// =====================================================================================
// 4. Lecturer recommendation request
// =====================================================================================
const lecturer = [
  TITLE("Lecturer Recommendation Request Pack"),
  P("Purpose: get **new, accurate academic recommendation letters** written by your lecturers themselves. The 2024 letter cannot be reused because it overstates grades (DR-14).", { run: { italics: true, color: "404040" } }),
  H1("Who to ask (based on your transcript, E0)"),
  GRID(["Priority", "Referee", "Why (verified grade)"], [
    ["1", "[Supervisor of Final Project I and II]", "Final Project I A−, Final Project II A−. Strongest academic evidence"],
    ["2", "[Lecturer of Compiler Design / Complexity Theory / Advanced Programming]", "A− in each"],
    ["3", "[Lecturer of Selected Topics in Computer Science]", "A"],
    ["Backup", "[Lecturer of Computer Security / Distributed Systems]", "B / B− (relevant to your cloud and security goals)"],
  ], [1100, 4600, 3938]),
  H1("Where the letters go (verified requirements)"),
  GRID(["Programme", "Requirement", "Deadline guidance"], [
    ["EDISS (Erasmus Mundus)", "**Two academic letters**, signed; university letterhead and stamp preferred. You may upload them, or the referee may e-mail them to ediss@abo.fi within the deadline (follow EDISS's subject-line rule).", "Before 12 Jan 2027 (15:00 Helsinki)"],
    ["Sweden SI", "Two references on the **SI 2027/28 template**, typed, printed, signed and **stamped**; one must be work-based; the second may be academic. The referee must answer SI e-mails until 22 Apr 2027.", "By 15 Feb 2027 (portal closes 24 Feb)"],
    ["CMU-Africa", "Three referees submit through the online portal (you enter their contact details).", "Enter by 1 Dec 2026 (early deadline 15 Dec)"],
    ["Hungary (Szeged)", "Optionally two letters from recognised academics.", "With the SH application (expected Jan 2027)"],
  ], [2300, 5000, 2338]),
  H1("A. Request e-mail (personalise for each lecturer)"),
  BOX("EMAIL: copy and paste", [
    "**Subject:** Request for an academic recommendation – NEBYU DEJENIE TILAHUN (BSc Computer Science, 2022)",
    "Dear [Dr./Mr./Ms. Name],",
    "I hope you are well. I was your student in [course name, academic year] at Debre Tabor University (BSc Computer Science, graduated July 2022). Since graduating I have worked in Ethiopia's fintech sector on [your verified role, e.g. IT infrastructure and operations].",
    "I am applying for fully funded master's programmes for 2027, including [EDISS (Erasmus Mundus, Finland/Spain/Italy/Sweden), the Swedish Institute Scholarship, and Carnegie Mellon University Africa]. Would you be willing to write an academic recommendation letter for me?",
    "The letter would need your signature and, if possible, university letterhead and stamp. The first deadline is [1 December 2026 for the online referee form] and the last is [15 February 2027]. I attach a one-page information sheet with my grades in your course, the programmes and deadlines, and how the letter is submitted.",
    "I would be grateful if the letter reflected only what you personally observed of my work.",
    "Thank you very much for considering this.",
    "With respect,",
    "[Full name]  |  [Phone]  |  [Email]",
  ]),
  H1("B. Short message (Telegram / WhatsApp)"),
  BOX("SHORT MESSAGE", [
    "Good day [Dr./Mr./Ms. Name], this is [your name], your former student in [course] (BSc CS, 2022). I'm applying for fully funded master's programmes and would be honoured if you could write an academic recommendation. I've sent details by e-mail. Thank you!",
  ]),
  BR(),
  H1("C. Recommender information sheet (attach to the e-mail)"),
  P("Facts only. Complete the bracketed items from your verified records before sending.", { run: { italics: true } }),
  FORM([
    ["Student", "NEBYU DEJENIE TILAHUN [use passport spelling if different]"],
    ["Degree", "BSc Computer Science, Debre Tabor University (Extension programme), graduated 14 July 2022"],
    ["Your course and my grade", "[Course name] – [grade from transcript]"],
    ["Overall record", "CGPA 2.62/4.00; final-year (2021/22) semester averages 3.05 and 3.35; final project A− (both parts)"],
    ["Final-year project", "[Title] – [one-line description] [to confirm with the department]"],
    ["Current work", "[Verified title(s) and employer(s) from HR letters]"],
    ["Programmes", "[List the programmes you will actually apply to]"],
    ["Submission method", "[Portal link / e-mail / signed scan, per programme]"],
    ["Deadlines", "[Dates]"],
    ["What you could speak to (only if true)", "[e.g. performance and initiative in the final project; problem-solving in compiler design]"],
  ]),
  SP(),
  P("**Integrity rule:** never write the letter for the referee and never sign on their behalf. If a referee asks you to draft, offer this information sheet instead."),
];

// =====================================================================================
// 5. Manager recommendation request
// =====================================================================================
const manager = [
  TITLE("Manager Recommendation Request Pack"),
  P("Purpose: get **work-based recommendations** from supervisors who know your real responsibilities. SI requires at least one work-based reference. Mastercard Foundation and SI both weigh leadership.", { run: { italics: true, color: "404040" } }),
  H1("Who to ask"),
  ...NUMS([
    "[Your current line manager at SantimPay]: knows your current responsibilities.",
    "[A second professional referee]: your Kifiya supervisor, or a senior SantimPay colleague who observed a project you led or coordinated.",
  ]),
  H1("What the scholarship bodies look for (verified)"),
  GRID(["Programme", "Requirement", "Evidence"], [
    ["Sweden SI", "Two references on the SI 2027/28 template (published 13 Nov 2026), typed, printed, signed and stamped. **One must be work-based.** Leadership means leading colleagues, allocating tasks, influencing strategy or decisions. The referee must reply to SI e-mails until 22 Apr 2027.", "E1 si.se"],
    ["Sweden SI (separate form)", "Proof of Work and Leadership Experience form: signed and stamped by someone who can confirm your work on behalf of the employer (HR or head of organisation acceptable).", "E1 si.se FAQ"],
    ["CMU-Africa / Mastercard Foundation", "3 referees via portal; the scholarship values leadership and giving back.", "E2 / E1"],
  ], [2300, 5600, 1738]),
  H1("A. Request e-mail"),
  BOX("EMAIL: copy and paste", [
    "**Subject:** Request for a professional reference – master's scholarship applications",
    "Dear [Name],",
    "I am applying for fully funded master's programmes for 2027, including the Swedish Institute Scholarship for Global Professionals and Carnegie Mellon University Africa. These scholarships require references from supervisors who know my work. Would you be willing to act as one of my professional referees?",
    "It would involve: (1) completing and signing the Swedish Institute's reference template (with the company stamp) after it is published on 13 November; (2) answering a short online form from Carnegie Mellon University Africa; and (3) possibly replying to a verification e-mail from the Swedish Institute by April 2027.",
    "I attach a one-page summary of the work you supervised, so you can check the facts. Please include only what you have observed yourself.",
    "Thank you very much for considering this.",
    "Best regards,",
    "[Full name]",
  ]),
  H1("B. Short message (Telegram / WhatsApp)"),
  BOX("SHORT MESSAGE", [
    "Hi [Name], I'm applying for fully funded master's scholarships (Sweden, CMU-Africa). Would you be willing to be a professional referee? It needs a signed, stamped reference form and possibly a verification e-mail later. Details are in my e-mail. Thank you!",
  ]),
  BR(),
  H1("C. Professional referee information sheet (attach)"),
  P("Fill this in **only** with facts your manager can confirm. Leave anything uncertain blank.", { run: { italics: true } }),
  FORM([
    ["Employee", "[Full name as in passport]"],
    ["Position(s) and dates", "[From the HR certificate: title – from – to]"],
    ["Main systems you worked on", "[e.g. POS terminal management, payment gateway, internal infrastructure: only what applies]"],
    ["Responsibilities the manager observed", "[e.g. CI/CD pipelines, Ansible automation, monitoring, incident response]"],
    ["A concrete example", "[One real event: situation, what you did, result. Numbers only if known]"],
    ["Leadership (only if real)", "[e.g. coordinated a rollout team of N people; trained colleagues; led incident response]"],
    ["Programmes and deadlines", "[List]"],
    ["How to submit", "[SI template (signed + stamped) / CMU portal link]"],
  ]),
  SP(),
  P("**Integrity rule:** do not ask a manager to confirm metrics, titles or leadership they did not observe. A reference that contradicts the HR certificate will damage the whole application."),
];

// =====================================================================================
// 6. Personal information confirmation form
// =====================================================================================
const personal = [
  TITLE("Personal Information Confirmation Form"),
  P("Fill this in once. Your answers become **E4 (personal confirmation)** evidence. Official documents (E0/E1) still override this form. **Do not write your passport number, ID numbers or bank details here.**", { run: { italics: true, color: "404040" } }),
  H1("1. Identity"),
  FORM([["Full name exactly as in passport", ""], ["Passport expiry date (dd/mm/yyyy)", ""], ["Spelling you use on LinkedIn / e-mail (if different)", ""]]),
  H1("2. Education"),
  FORM([
    ["Month/year your first semester started", ""],
    ["Have you collected the printed (permanent) diploma? Where is it?", ""],
    ["Cumulative GPA used in every application: 2.62/4.00 (official transcript; confirmed 8 Nov 2026). Does any official document print a different CGPA? (Yes/No – which?)", ""],
    ["Final-year project title", ""], ["Project: technology, team size, your part", ""], ["Project supervisor (name, title)", ""],
  ]),
  H1("3. Employment (facts only; HR letters will be compared)"),
  GRID(["Employer", "Title(s)", "Start (dd/mm/yyyy)", "End / current", "FT/PT + hours/week", "Evidence you can get"], [
    ["SantimPay Financial Solution S.C.", "", "", "", "", ""],
    ["Kifiya Financial Technology PLC", "", "", "", "", ""],
    ["Infonans Solutions PLC", "", "", "", "", ""],
    ["Gebeya Inc. (freelance?)", "", "", "", "", ""],
    ["Dereja (co-trainer / programme)", "", "", "", "", ""],
    ["GAO Tek Inc. (unpaid internship)", "", "", "", "", ""],
  ], [2300, 1500, 1350, 1350, 1500, 1638]),
  SP(),
  FORM([
    ["Plan International Ethiopia: was this listed by mistake? (Yes / No + evidence)", ""],
    ["Yandex (Odoo developer): evidence available? (contract / payment / platform record)", ""],
    ["How did the SantimPay and Kifiya roles run at the same time? (hours, arrangement)", ""],
  ], 4600),
  H1("4. Leadership examples (only real ones)"),
  GRID(["Situation", "Your role", "People involved", "Result", "Who can confirm"], [["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]],
    [2400, 1800, 1500, 2200, 1738]),
  H1("5. Technical tools: keep only what you can demonstrate in an interview"),
  GRID(["Tool", "Used? (Y/N)", "Period / employer", "Could you demo it? (Y/N)"], [
    ["Jenkins", "", "", ""], ["GitHub Actions", "", "", ""], ["Azure DevOps", "", "", ""], ["Ansible", "", "", ""], ["Terraform", "", "", ""],
    ["Docker", "", "", ""], ["Kubernetes (operations)", "", "", ""], ["Prometheus", "", "", ""], ["Grafana", "", "", ""], ["Splunk", "", "", ""],
    ["ELK / Fluentd", "", "", ""], ["VMware ESXi", "", "", ""], ["Proxmox VE", "", "", ""], ["n8n / AI-agent automation", "", "", ""],
    ["IAM/MFA, WAF/WAAP, PAM/DAM", "", "", ""],
  ], [3200, 1500, 2900, 2038]),
  H1("6. Certifications"),
  FORM([["CKA certificate ID / verification link", ""], ["Azure Fundamentals transcript link", ""], ["Other (name, issuer, ID)", ""]]),
  H1("7. English tests"),
  FORM([["Any English test taken before? (which, score, date)", ""], ["Can you sit IELTS Academic in early to mid December 2026? (ETB 32,800 at BC Addis Ababa University)", ""]], 4600),
  H1("8. Finances (needed for need-based Mastercard Foundation consideration; answer honestly)"),
  FORM([["Could you or your family pay any part of tuition or living costs? (approximate)", ""], ["Dependants you support (number)", ""], ["Budget available now for tests and fees (ETB)", ""]], 4600),
  H1("9. Preferences"),
  FORM([
    ["Countries you accept (Sweden, Finland/Estonia/Portugal, Rwanda, Hungary, China, Korea, Canada …)", ""],
    ["Do you accept a commitment to return home after study? (Mastercard Foundation, Chevening, Fulbright)", ""],
    ["Pursue Manitoba (deadline 1 Dec; needs a supervisor) – Yes/No", ""],
  ], 4600),
  H1("10. Referees"),
  GRID(["Name", "Title / organisation", "E-mail", "Relationship / course", "Agreed?"], [["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]],
    [1900, 2300, 2100, 2100, 1238]),
  H1("Declaration"),
  P("I confirm that the information above is accurate and complete to the best of my knowledge. I understand that official documents take precedence over this form."),
  P("Name: ______________________    Signature: ______________________    Date: ____________"),
];

(async () => {
  // The two HR requests were superseded in Phase 6 (now in APPLICATION_EXECUTION_CENTER/15_EMPLOYMENT_EVIDENCE;
  // old copies in 14_ARCHIVE/phase6_superseded). Set BUILD_SUPERSEDED=1 only to regenerate those archived copies.
  if (process.env.BUILD_SUPERSEDED === "1") {
    await save("SantimPay_HR_Request.docx", "SantimPay HR Request", santim);
    await save("Kifiya_HR_Request.docx", "Kifiya HR Request", kifiya);
  }
  await save("University_Registrar_Request.docx", "University Registrar Request", registrar);
  await save("Lecturer_Recommendation_Request.docx", "Lecturer Recommendation Request", lecturer);
  await save("Manager_Recommendation_Request.docx", "Manager Recommendation Request", manager);
  await save("Personal_Information_Confirmation_Form.docx", "Personal Information Confirmation Form", personal);
})();
