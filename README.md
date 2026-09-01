# Tender Project Snapshot

An AI-assisted Skill for reviewing tender documents before BID drafting and generating a concise, traceable, and stable Excel project snapshot.

## What this project does

Tender Project Snapshot helps users understand a purchased or otherwise obtained tender package before they begin drafting a BID.

It reads the complete tender document and relevant attachments, extracts the information that affects whether and how the project should proceed, and produces a professionally formatted `.xlsx` workbook. The goal is to reduce repetitive document review while keeping important conclusions traceable to the source text.

This project is designed to assist human work, not replace it. AI performs extraction, classification, comparison, and organization; the user remains responsible for qualification decisions, clarification requests, commercial judgment, and the final BID.

## Current status

The project is currently in **Stage 1: Pre-BID Review**.

Stage 1 is implemented and focuses on reviewing the tender package before BID drafting begins. The following stages are still under development and are not yet complete:

- **Stage 2: BID Drafting** - extract the required BID format and basic fields, classify historical content as reusable, reusable after update, or project-specific, and organize scoring and response requirements.
- **Stage 3: Post-BID Check** - provide a final yes/no checklist for dates, signatures, seals, attachments, printing, binding, packaging, sealing, and scanned PDF delivery.

## Stage 1 capabilities

The current Skill generates one Excel Sheet named `项目介绍` and includes five sections:

1. **Enterprise-level eligibility review**
   - Extracts only requirements that determine whether the bidding entity itself is eligible.
   - Excludes registration, tender-document purchase, personnel deployment, equipment configuration, and service-plan requirements.
   - Splits compound requirements into one factual yes/no question per row.
   - Provides Excel dropdowns for user confirmation and formula-based status results.

2. **30-second project overview**
   - Summarizes the project name, project number, purchaser, procurement agent, procurement method, award count, deadline, opening location, and a one-sentence project definition.
   - Separates the actual project scenario from pricing or calculation benchmarks.

3. **Project scope and key requirements**
   - Explains what is being purchased, who receives the service or deliverable, how the project is expected to operate, and the relevant time, location, scale, standards, deliverables, and acceptance requirements.
   - Uses `未说明` (`Not specified`) when the source document does not provide an answer instead of guessing.

4. **Contract comparison**
   - Compares the contract template with the project requirements whenever a contract is included.
   - Keeps the Excel output compact when the two are consistent.
   - Expands only conflicts, omissions, uncertain differences, and items requiring human confirmation.
   - Extracts statements about whether the contract is only a reference template and which document takes priority if provisions conflict.

5. **Quotation restriction reminders**
   - Extracts the maximum price or other quotation limits, tax requirements, consequences of non-compliance, and the location of the official quotation form.
   - Does not rebuild the purchaser's quotation form or invent a new pricing structure.

## What Stage 1 does not do

- It does not handle project registration or tender-document purchase procedures.
- It does not draft the BID, write service plans, or produce scoring responses.
- It does not replace the official quotation form.
- It does not make eligibility decisions on behalf of the user.
- It does not guarantee that AI has identified every semantic conflict or document defect.
- It is not legal, procurement, financial, or compliance advice.

## Output design

The formal deliverable is always a real `.xlsx` workbook, not TXT, Markdown, CSV, JSON, HTML, or a chat table.

The Skill separates document understanding from spreadsheet rendering:

```text
Tender document and attachments
            |
            v
AI extraction, classification, and contract comparison
            |
            v
Structured internal data
            |
            v
Deterministic Python renderer
            |
            v
Validated single-Sheet Excel workbook
```

The bundled renderer calculates the complete layout before writing the workbook. It validates the Sheet structure, formulas, dropdown ranges, conditional formatting, initial blank state, and contract-comparison completeness before replacing the final file.

This design reduces common spreadsheet-generation failures such as shifted columns, misplaced headers, leftover test selections, unreadable status colors, and inconsistent formatting.

## Installation and use

Use the complete Skill package. Do not provide only `SKILL.md`, because the renderer, field rules, input schema, and Excel style reference are required runtime assets.

Requirements:

- Python 3
- `openpyxl`
- An AI agent environment capable of reading DOCX or PDF files and executing the bundled Python renderer
- Excel, WPS, LibreOffice, or another compatible application for opening the output

Run the package preflight check first:

```bash
cd tender-project-snapshot
python3 scripts/build_project_snapshot.py --preflight
```

Example request:

```text
Use $tender-project-snapshot to read this complete tender package and generate the Stage 1 pre-BID review workbook. Compare the contract template with the project requirements and deliver a real .xlsx file.
```

For best results, provide the main tender document together with separately issued attachments, BID formats, quotation forms, clarification documents, and amendments.

## Project structure

```text
bid-skills/
├── README.md
└── tender-project-snapshot/
    ├── SKILL.md
    ├── agents/
    │   └── openai.yaml
    ├── assets/
    │   └── 项目介绍_通用空白模板.xlsx
    ├── references/
    │   ├── pre-bid-schema.md
    │   └── renderer-input-schema.md
    └── scripts/
        └── build_project_snapshot.py
```

## Design principles

- Read the complete document before drawing conclusions.
- Record explicit facts and write `未说明` (`Not specified`) when information is absent.
- Keep the framework industry-neutral and populate specific requirements only from the current tender package.
- Keep qualification review limited to enterprise-level eligibility.
- Preserve searchable source locations for important conclusions and anomalies.
- Use deterministic code for spreadsheet layout and validation.
- Keep the output concise enough for rapid human review.
- Fail clearly when required assets or runtime capabilities are unavailable.

## Limitations and human review

The renderer can verify mechanical rules such as file integrity, Sheet structure, formulas, dropdowns, formatting, and blank initial states. It cannot prove that the AI has understood every clause or discovered every inconsistency.

OCR quality, complex tables, embedded objects, and unstable DOCX pagination may affect source extraction. Before using the result in a real procurement process, manually review all eligibility items, source references, contract anomalies, quotation limits, and clarification questions.

## Roadmap

- [x] Stage 1: Pre-BID Review
- [ ] Stage 2: BID Drafting
- [ ] Stage 3: Post-BID Check

The project is still evolving. Real tender-package tests, issue reports, and generalizable workflow suggestions are welcome.
