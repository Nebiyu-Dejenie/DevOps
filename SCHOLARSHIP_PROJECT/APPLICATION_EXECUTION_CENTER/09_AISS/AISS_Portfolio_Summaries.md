---
title: "Portfolio of Executed Projects"
subtitle: "[Passport name]. Application to Artificial Intelligence for Sustainable Societies (SITeS-JM), Erasmus Mundus, September 2027"
---

*AISS assesses whether "the portfolio has a summary of the projects and is not just a collection of articles or links, and the items are explained clearly and understandably without a need to go through extra links and files" (criterion 1a). Each item must be an "actual project that has been executed", with "a clear summary and connection to the program, explaining the skill set needed to select and execute these" (2b). Originality is assessed (2c).*

**Before you submit:**

- Replace every `[⏳ …]`.
- Delete any item you cannot prove.
- Keep each summary to 80–150 words.
- Export as one PDF.

---

## 1. Final-year project: [⏳ project title]

**Context.** Two-part final-year project in the BSc Computer Science at Debre Tabor University (2021/22). [⏳ One sentence: the problem and who had it.]

**My role.** [⏳ Your part; team size.]

**What I did.** [⏳ 2–3 sentences: design, implementation, testing; languages and tools.]

**Result.** Graded **A− in both parts** (Final Project I and II, 13 credits). [⏳ Demo, users or test results, if any.]

**Skills used.** [⏳ e.g. requirements analysis, database design, programming in …, teamwork]

**Connection to AISS.** [⏳ One honest sentence, e.g. how the system served its users and what you would now design differently with them.]

*Evidence:* transcript (official); [⏳ report or screenshots].

---

## 2. Keeping digital-payment services running (SantimPay Financial Solution S.C.)

**Context.** SantimPay provides POS, QR, school-payment, payment-gateway and loyalty services together with Ethiopian banks. When these services fail, merchants cannot sell and families cannot pay.

**My role.** [⏳ Title(s) and period exactly as in the HR letter.]

**What I did.** [⏳ 2–3 verified duties, e.g. operate the Linux servers that run payment services; automate their configuration with Ansible; maintain the deployment pipeline; monitor service health and respond to incidents.]

**Result.** [⏳ One real outcome, e.g. an incident you resolved and what changed afterwards; time saved by an automation. No estimates.]

**Skills used.** Linux operations, configuration automation, deployment pipelines, monitoring, incident response [⏳ keep only those you can demonstrate].

**Connection to AISS.** Digital payments are becoming basic public infrastructure in Ethiopia. Making them reliable and trustworthy is a precondition for any AI that will later run inside them, such as fraud detection or credit decisions. That is the socio-technical problem I want to study.

*Evidence:* HR letter; manager reference.

---

## 3. Public demonstration: automated build, test and deployment (CI/CD + configuration automation)

**Context.** A small, self-built project showing how changes to a service can be tested and deployed safely and the same way every time. **It contains no employer code or data.**

**What I did.** [⏳ After you build it:]

- a small web service with automated tests
- a container image
- a pipeline (lint → test → build → deploy)
- an Ansible role that prepares the server and runs the service

**Result.** [⏳ Repository link; a green pipeline run; what you learned.]

**Skills used.** Software testing, automation, reproducible infrastructure.

**Connection to AISS.** Responsible AI services need the same discipline: every change is tested and traceable before it reaches users.

---

## 4. Public demonstration: monitoring and alerting for a live service

**Context.** A self-built observability setup for the service in item 3.

**What I did.** [⏳ After you build it:]

- service metrics collected by Prometheus
- a Grafana dashboard
- two alert rules, with the reason for each threshold written down
- a simulated failure to show the alert firing

**Result.** [⏳ Repository link; screenshot of the dashboard and the alert.]

**Skills used.** Observability, data-driven operations, alert design.

**Connection to AISS.** Monitoring data is the raw material for AI that predicts failures. Designing alerts that people can act on is a human-centred design problem, not only a technical one.

---

## 5. Choose one: merchant terminal operations **or** Dereja co-training

**5a. POS terminals and merchant onboarding** (only if your role is confirmed). [⏳ Your role, the people you worked with, the result.] *Connection:* designing digital services with and for the merchants who use them.

**5b. Dereja co-training (2022)** (only if confirmed with dates or a certificate). Supported trainees on the learning platform and prepared weekly progress reports. *Connection:* inclusion and capacity-building, central to sustainable societies.

---

## How to build items 3 and 4 yourself (about 5 evenings)

- Build them yourself, on your own GitHub account, from scratch.
- **If you use tutorials or AI assistance, that is fine, but you must understand and be able to explain every file.** The AISS interview checks authenticity ("does not read from the screen or provide predefined answers").
- Write the README in your own words, with a diagram.

**Acceptance checklist:**

- [ ] The README explains the purpose in 3 sentences.
- [ ] A diagram is included.
- [ ] "How to run" steps work on a clean machine.
- [ ] The pipeline is green.
- [ ] The alert fires during the simulated failure.
- [ ] No secrets in the repository.
