# _tools: document rendering helpers

Used by `APPLICATION_EXECUTION_CENTER/_build/render_all.sh` and `render_packs.sh`.

- `md2pdf.sh file.md [landscape]`: Markdown → HTML (pandoc, `doc.css`) → PDF (Chromium via Playwright, `md2pdf.js`).
- Word copies are made with pandoc and `00_MASTER_PROFILE/generator/reference_compact.docx`.
- Spreadsheets are recalculated with LibreOffice after every build, so cached values exist (0 formula errors is the acceptance check).

**Requirements:** pandoc, Node.js with Playwright and Chromium, poppler-utils (`pdfinfo`), LibreOffice.
