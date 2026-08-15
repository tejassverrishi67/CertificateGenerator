# Certificate Generator

A browser-based tool for bulk-generating certificates from a single Word (DOCX) template. Upload a template containing placeholder tags (e.g. `{NAME}`, `{Designation}`, `{Paper Title}`, `{DOI}`, `{vol}`, `{issue}`, `{year}`, `{month}`), paste in a block of recipient data, and the app fills in a copy of the DOCX per recipient, combining every certificate into a single multi-page Word document (via docxtemplater/PizZip). A companion local server (`server.js`) can convert that combined DOCX to a pixel-faithful PDF using Microsoft Word itself (COM automation via `convert-docx-to-pdf.ps1`).

## Running it

This is a static site with no build step or backend. To run it locally, either:

- Open `index.html` directly in a browser, or
- Serve the folder with any static file server, e.g.:
  ```
  npx serve .
  ```
  or
  ```
  python -m http.server
  ```
  then visit the printed local URL.

## How it works

1. **Upload a DOCX template** containing placeholder tags: `{NAME}`, `{Designation}`, `{Paper Title}` / `{PaperTitle}`, `{DOI}`, `{vol}`, `{issue}`, `{year}`, and `{month}` (docxtemplater performs the substitution, matching several case/spacing variants).

2. **Paste mapping data** into the text box, one category per line. Each line is parsed by `parseDataInput`/`parseMappingLine` in `app.js`:
   - **Names**: a comma-separated list of names, each with a numeric suffix indicating which certificate(s) it maps to, e.g. `John Smith1, Jane Doe2`. A range suffix like `3-5` applies the same value to certificates 3 through 5, and a comma-separated list like `1,2,3` (or a mix, e.g. `1,3-5`) applies to exactly those certificates. The name line determines how many certificates are generated (the highest suffix number). Superscript numerals (e.g. `Name¹²`) are also recognized as suffix digits.
   - **Designations**: same comma+suffix format, e.g. `Professor, Example University1-2`. Lines are heuristically classified as a designation if they contain keywords such as "student", "professor", "department", "college", "university", etc.
   - **Paper Title**: a line without keyword matches is treated as a paper title. If it has no numeric suffixes it's applied as a constant to every certificate; if it has suffixes, it's mapped per-certificate the same way as names.
   - Use the **"Load Example Data"** button to see a sample of the expected input format.

3. **DOI**: enter the DOI in its own dedicated field (separate from the mapping text box). It's applied as a constant value to every generated certificate, and its final numeric segment is also parsed into `{vol}` (2 digits), `{issue}` (1 digit), and a paper/file number (last 2 digits), with `{year}` read from the numeric segment before it — e.g. `10.17148/IJARCCE.2026.15817` → year `2026`, vol `15`, issue `8`, file number `17`. Each output file is named `"{paper number} {author list} {certificate index}"`; if no DOI is entered, the paper number is omitted from the filename.

4. **Month**: an optional free-text field (not derivable from the DOI) mapped to `{month}`, applied as a constant to every certificate.

5. **Export**: generates every certificate and combines them into a single multi-page Word (`.docx`) file for download (`buildCombinedDocxBlob` in `app.js`). Optionally, running the local server (`node server.js`) exposes `POST /api/convert`, which converts a batch of DOCX files to pixel-faithful PDFs using Microsoft Word's own COM automation (`convert-docx-to-pdf.ps1`) — this requires Windows with Microsoft Word installed.

## Dependencies

- [PizZip](https://github.com/open-xml-templating/pizzip) + [docxtemplater](https://docxtemplater.com/) (loaded from CDN) — filling placeholders inside the DOCX template
- [Font Awesome](https://fontawesome.com/) and [Google Fonts](https://fonts.google.com/) (loaded from CDN) — UI typography
- Node.js standard library only for `server.js` — no npm dependencies
- Microsoft Word (Windows, COM automation) for the optional PDF conversion step
