# Certificate Generator

A browser-based tool for bulk-generating certificates from pre-loaded Word (DOCX) templates. Pick one of the four journals (**IJARCCE**, **IARJSET**, **IJIREEICE**, **IMRJR** — internally keyed as CCE/SET/ICE/RJR), paste in a block of recipient data, and download the result as a combined file, individual files, or a single group certificate (via docxtemplater/PizZip). A companion local server (`server.js`) can convert a generated DOCX to a pixel-faithful PDF using Microsoft Word itself (COM automation via `convert-docx-to-pdf.ps1`).

There is no upload step — the templates ship with the app in `templates/`.

## Running it

This is a static site with no build step or backend, but it **must be served over HTTP** — templates are loaded with `fetch()`, which browsers block on `file://`, so opening `index.html` directly will not work.

```
node server.js
```
then visit http://localhost:8080. Any static file server works equally well (`npx serve .`, `python -m http.server`).

## Templates

The four journal templates live in `templates/` (`CCE.docx`, `SET.docx`, `ICE.docx`, `RJR.docx`) and are fetched on demand — only the template you click is downloaded, and it's cached for the rest of the session. Each journal's logo lives in `logos/` (`CCE.jpg`, `SET.jpg`, `ICE.jpg`, `RJR.jpg`).

The `TEMPLATES` array at the top of `app.js` is the single source of truth for the picker — `index.html` just holds an empty `<div id="template-grid">`, and `renderTemplateGrid()` builds the four cards (logo + full journal name) from the array on load.

- **To update a template or logo**: replace the file in `templates/` or `logos/`, keeping the same name.
- **To add a journal**: drop the `.docx` in `templates/` and the logo in `logos/`, then add one `{ key, fullName, file, logo }` entry to `TEMPLATES` in `app.js` — no HTML changes needed.

Placeholder tags are matched case- and spacing-insensitively (`{Paper Title}`, `{paper_title}`, and `{PAPERTITLE}` are equivalent), so a template with an inconsistently-spelled tag still fills correctly — the RJR template, for instance, spells its name tag `{NAMe}`. An unrecognized tag renders as empty text rather than breaking the export.

## How it works

1. **Choose a journal template** — IJARCCE, IARJSET, IJIREEICE, or IMRJR. All four contain the same placeholder tags: `{NAME}`, `{Designation}`, `{Paper Title}`, `{DOI}`, `{vol}`, `{issue}`, `{year}`, and `{month}` (docxtemplater performs the substitution). You can switch templates at any point without reloading; the next export uses the newly selected one.

2. **Paste mapping data** into the text box, one category per line. Each line is parsed by `parseDataInput`/`parseMappingLine` in `app.js`:
   - **Names**: a comma-separated list of names, each with a numeric suffix indicating which certificate(s) it maps to, e.g. `John Smith1, Jane Doe2`. A range suffix like `3-5` applies the same value to certificates 3 through 5, and a comma-separated list like `1,2,3` (or a mix, e.g. `1,3-5`) applies to exactly those certificates. The name line determines how many certificates are generated (the highest suffix number). Superscript numerals (e.g. `Name¹²`) are also recognized as suffix digits.
   - **Designations**: same comma+suffix format, e.g. `Professor, Example University1-2`. Lines are heuristically classified as a designation if they contain keywords such as "student", "professor", "department", "college", "university", etc.
   - **Paper Title**: a line without keyword matches is treated as a paper title. If it has no numeric suffixes it's applied as a constant to every certificate; if it has suffixes, it's mapped per-certificate the same way as names.
   - Use the **"Load Example Data"** button to see a sample of the expected input format.

3. **Publication details** — volume, issue, year, month and the file number, supplied in one of two modes:

   - **With DOI** (default): paste the DOI and everything else is derived from it. The final numeric segment packs **volume (2 digits) + issue + file number (2 digits)**, and the segment before it is the year. The issue is whatever sits between volume and file number: one digit for issues 1–9, two digits for issues 10–12. So `10.17148/IJARCCE.2026.15817` → `{year}` `2026`, `{vol}` `15`, `{issue}` `8`, `{month}` `August`, file number `17`.
   - **Without DOI**: enter **volume**, **issue**, **file number** and **year** by hand. `{DOI}` renders empty on the certificate.

   In both modes the issue number *is* the month of publication, so `{month}` is always derived rather than typed — issue 7 is July, issue 9 is September. `getPublicationInfo()` in `app.js` resolves all of this for whichever mode is active, so the rest of the app never has to care which was used.

4. **Download** — one of three modes (all disabled until a template is picked and data is parsed):
   - **All in One**: every certificate combined into a single multi-page Word file (`buildCombinedDocxBlob` + `handleDownloadAllInOne`), named `"{file number} {author list}"` — e.g. `17 Author1, Author2.docx` (`buildCombinedFilename`).
   - **Multiple Downloads**: one separate file per recipient (`renderDocxBlobForRecord` + `handleDownloadMultiple`), triggered as sequential browser downloads staggered ~250ms apart. Each filename shares the same base as "All in One" but ends in that recipient's 1-based certificate index — e.g. `17 Author1, Author2 1.docx`, `17 Author1, Author2 2.docx` (`buildIndividualFilename`). Chrome may show a one-time "allow multiple downloads" prompt after the second file.
   - **Group Download**: every recipient combined onto a single certificate instead of one each (`buildGroupRecord` + `handleDownloadGroup`) — names are comma-joined, and Designation/Paper Title are de-duplicated (each distinct value appears once, in first-occurrence order) then comma-joined, since the same designation is often shared by several recipients. Same filename pattern as "All in One".

   The file number is omitted from any filename if it isn't available in the active mode (e.g. no DOI entered and no manual file number given).

Optionally, running the local server (`node server.js`) exposes `POST /api/convert`, which converts a batch of DOCX files to pixel-faithful PDFs using Microsoft Word's own COM automation (`convert-docx-to-pdf.ps1`) — this requires Windows with Microsoft Word installed.

## Dependencies

- [PizZip](https://github.com/open-xml-templating/pizzip) + [docxtemplater](https://docxtemplater.com/) (loaded from CDN) — filling placeholders inside the DOCX template
- [Font Awesome](https://fontawesome.com/) and [Google Fonts](https://fonts.google.com/) (loaded from CDN) — UI typography
- Node.js standard library only for `server.js` — no npm dependencies
- Microsoft Word (Windows, COM automation) for the optional PDF conversion step
