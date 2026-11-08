---
title: "Technical Project Portfolio v1 (working draft)"
subtitle: "Nebyu Dejenie Tilahun. Real projects only. 7 November 2026"
---

**Status:** every project below is real. **None has a measured result yet**, so "Results" fields stay empty until you supply numbers you can defend in an interview. Do not fill a gap with an estimate.

- `[⏳ …]` marks something only you or your employer can provide.
- **Employer confidentiality:** use no client names, internal hostnames, IP addresses, code or screenshots containing customer data without your employer's written permission. Where a client is named in your documents, this portfolio writes "a national supermarket chain".

**Evidence codes:**

- **E0:** your own document
- **E4:** your confirmation only
- **Pending:** awaiting the HR letter or your confirmation

# Which projects go where

| # | Project | AISS (max 5) | EDISS | CMU-Africa | SI | CSC | Evidence today |
|---|---|---|---|---|---|---|---|
| P1 | Final-year project (A−, A−) | ✔ | ✔ (thesis work) | ✔ | – | ✔ | E0 grades; **title missing (DR-17)** |
| P2 | Payment-infrastructure operations at SantimPay | ✔ | ✔ | ✔ | ✔ | ✔ | Company services E0; **your role pending (DR-06)** |
| P3 | POS terminal fleet and merchant onboarding operations | ✔ if your role is confirmed | – | ✔ | ✔ (leadership) | – | Tracking sheets 2025–26 (authorship **pending**) |
| P4 | Retail QR, loyalty and gift-card integration | optional | ✔ (software) | ✔ | – | – | Project documents 2024–25; bank letters are company-level; **your role pending** |
| P5 | Public demo: CI/CD + configuration automation (to build) | ✔ | ✔ | ✔ | – | ✔ | **Not built yet** |
| P6 | Public demo: monitoring and alerting stack (to build) | ✔ | ✔ | ✔ | – | ✔ | **Not built yet** |
| P7 | Security operations support at Kifiya | – | – | ✔ | ✔ | – | **Role and tools pending (DR-07, DR-13)** |
| P8 | Dereja co-training (2022) | ✔ (social activity) | – | ✔ (giving back) | ✔ (civil society) | – | **Dates and numbers pending** |

**Not usable yet:** code archives found in your Drive (school payments, loyalty, events, web backend). Their **authorship is unknown** and they may be employer property. Do not publish them or describe them as yours unless you wrote them **and** your employer agrees.

# P1. Final-year project, BSc Computer Science (Debre Tabor University, 2021/22)

| Field | Content |
|---|---|
| Title | [⏳ exact title from the project report, DR-17] |
| Problem | [⏳ one sentence: who had what problem] |
| Architecture | [⏳ main components: e.g. client, server, database; draw a simple box diagram] |
| Your responsibilities | [⏳ your part if it was a team project; team size] |
| Technologies | [⏳ languages, frameworks, database] |
| Implementation | [⏳ 3–4 sentences: how it worked] |
| Results | **Grade: A− (Final Project I, 6 credits) and A− (Final Project II, 7 credits)** (E0 transcript). [⏳ any test result, users or demo] |
| Evidence | Transcript (E0); [⏳ report PDF; supervisor's name for the reference] |
| Relevance | Your strongest academic evidence of building software end to end. EDISS scores "thesis work" inside its 70 academic points; AISS scores executed projects |
| Missing | Title, report, your role, screenshots |

# P2. Payment-infrastructure operations (SantimPay Financial Solution S.C.)

| Field | Content |
|---|---|
| Title | Keeping digital-payment services running for Ethiopian banks and merchants |
| Problem | Payment services (POS terminal management and monitoring, QR merchant payments, school payments, a payment gateway and loyalty services, offered with partner banks) must stay available and secure; when they fail, merchants cannot take payments |
| Architecture | Describe at a high level only: Linux servers [⏳ physical / VMware ESXi / Proxmox], container platform [⏳ Docker / Kubernetes], deployment pipeline [⏳ Jenkins / GitHub Actions / Azure DevOps: which, when], configuration automation (Ansible), monitoring [⏳ Prometheus / Grafana / Splunk] |
| Your responsibilities | [⏳ exactly as the HR letter states, per title and period] |
| Technologies | Confirmed in your documents: Linux, Ansible, CI/CD, Docker, Kubernetes operations, monitoring. **Keep only the ones you can demonstrate** (DR-13) |
| Implementation | [⏳ 3–4 sentences: e.g. "I wrote Ansible playbooks for … ; I maintained the … pipeline for … services"] |
| Results | [⏳ only real numbers: servers or services managed, deployments per week, incidents handled, time saved. Label approximations "about"] |
| Evidence | Partner-bank letters describe SantimPay's services (company level, E0); HR letter (pending); manager reference (pending) |
| Relevance | The core of your profile: production operations for financial services, linked to SDG 8/9 (SI), "service-oriented" impact (Mastercard Foundation), and data-intensive operations (EDISS) |
| Missing | HR letter with titles, dates and duties; 1 incident story; 1 automation story with a number |

**Incident story template (STAR, 120 words):** Situation (what failed, when, who was affected) → Task (your responsibility) → Action (how you found the cause: logs, metrics, steps) → Result (time to recover, what changed afterwards so it did not recur).

# P3. POS terminal fleet and merchant onboarding operations

| Field | Content |
|---|---|
| Title | Rolling out and supporting card/QR payment terminals for merchants |
| Problem | Merchants need working terminals, configuration and support to accept digital payments |
| Architecture | Terminal management system → terminals in the field → payment platform → partner banks (high level only) |
| Your responsibilities | [⏳ led / coordinated / supported? team size you directed?] |
| Technologies | Terminal management and monitoring tools [⏳ names you can state publicly] |
| Implementation | [⏳ how onboarding worked; your steps] |
| Results | [⏳ terminals deployed or managed; merchants onboarded; regions. Only from records you are allowed to cite] |
| Evidence | 2025–26 tracking spreadsheets in your Drive (**confirm you created or maintained them**); manager confirmation |
| Relevance | SI leadership (if you allocated tasks or led a rollout); Mastercard Foundation "service-oriented" impact; AISS "participatory" angle (merchants as users) |
| Missing | Your exact role; numbers you may disclose; who confirms |

# P4. Retail QR, loyalty and gift-card integration

| Field | Content |
|---|---|
| Title | Digital payments, loyalty and gift cards for a national supermarket chain |
| Problem | A retailer wanted to accept QR payments and run loyalty and gift-card programmes with partner banks |
| Architecture | [⏳ high level: point-of-sale integration ↔ payment platform ↔ bank] |
| Your responsibilities | [⏳ integration, testing, deployment, support?] |
| Technologies | [⏳] |
| Results | [⏳ go-live date; stores; only what you can cite] |
| Evidence | Project documents 2024–25 (your Drive). Bank letters cover the company, not you |
| Relevance | Software integration and testing (EDISS software-engineering angle; CMU-Africa) |
| Missing | Your role; permission to describe the client |

# P5. Public demo: CI/CD pipeline + configuration automation (you build it)

**Purpose:** give reviewers a portfolio item they can open, built from scratch with **no employer code or secrets**. This is new work you do now; describe it as a demonstration, not as production work.

| Field | Specification |
|---|---|
| Title | `ci-cd-and-config-demo`: automated build, test and deployment of a small web service |
| Problem shown | How to ship changes safely: every commit is tested, packaged and deployed the same way |
| Architecture | A tiny web API (Python/Flask or Go) → unit tests → container image → pipeline (GitHub Actions) → deploy to a local VM or container host configured by an Ansible playbook |
| What you implement | 1) the service + 3–5 tests; 2) Dockerfile; 3) pipeline with lint, test, build and deploy stages; 4) Ansible role that installs Docker and runs the service; 5) README with a diagram and "how to run" |
| Done when | A reviewer can read the README in 3 minutes and see a green pipeline run |
| Evidence | Public repository link [⏳]; screenshots of a successful run |
| Relevance | AISS portfolio criterion 2b; CMU-Africa "strong programming skills"; EDISS software engineering; CSC study plan |
| Time | About 2–3 evenings |

# P6. Public demo: monitoring and alerting stack (you build it)

| Field | Specification |
|---|---|
| Title | `observability-demo`: metrics, dashboards and alerts for a small service |
| Problem shown | How to detect failures before users report them |
| Architecture | The P5 service exposes metrics → Prometheus scrapes them → Grafana dashboard → alert rule (e.g. error rate or latency) → notification (e-mail or webhook) |
| What you implement | docker-compose with Prometheus and Grafana; 1 dashboard; 2 alert rules with written reasons for each threshold; a simulated failure and a screenshot of the alert firing |
| Done when | The README explains why each alert exists and what the on-call person should do |
| Evidence | Public repository link [⏳]; dashboard screenshot |
| Relevance | Data-driven operations (EDISS "data-intensive"); AISS reliability angle; CMU-Africa |
| Time | About 2 evenings |

# P7. Security operations support (Kifiya Financial Technology PLC)

| Field | Content |
|---|---|
| Title | Security controls for financial-service clients |
| Your responsibilities | [⏳ configured yourself vs supported: IAM/MFA, WAF/WAAP, PAM/DAM, firewall rules] |
| Results | [⏳ projects taken to go-live; your role] |
| Evidence | HR letter (pending, DR-07) |
| Relevance | Cybersecurity master's options; SI work hours |
| Missing | Dates, type of contract, tools you configured yourself |

# P8. Dereja co-training (2022)

| Field | Content |
|---|---|
| Title | Supporting trainees in a youth employability programme |
| Your responsibilities | Supported trainees on the learning management system; prepared weekly progress reports (from your CV, **E4**) |
| Results | [⏳ number of trainees; period] |
| Evidence | [⏳ certificate or letter from Dereja] |
| Relevance | Giving back (Mastercard Foundation), social activities (AISS CV criterion 2d), civil-society leadership (SI) |
| Missing | Certificate, dates, numbers; whether it was employment or volunteering |

# Description template for AISS (80–120 words per project)

> **[Title]** ([year], [your role]). *Context:* … *What I did:* … *Result:* … *Skills used:* … *Connection to AISS:* …

# What you need to do (in order)

1. Find the final-year project title and report (P1).
2. HR letters (P2, P7), then fill in the "Your responsibilities" rows exactly as the letters state.
3. Write one incident story and one automation story (P2) using the STAR template.
4. Build P5 and P6 by **30 Nov** (about 5 evenings in total).
5. Confirm P3, P4 and P8 roles, and ask the employer's permission before naming any client.
