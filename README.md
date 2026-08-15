# Certificate Generator

A browser-based tool for bulk-generating certificates from pre-loaded Word (DOCX) templates. Pick one of the four journal templates (**CCE**, **SET**, **ICE**, **RJR**), paste in a block of recipient data, and the app fills in a copy of the DOCX per recipient, combining every certificate into a single multi-page Word document (via docxtemplater/PizZip). A companion local server (`server.js`) can convert that combined DOCX to a pixel-faithful PDF using Microsoft Word itself (COM automation via `convert-docx-to-pdf.ps1`).

There is no upload step — the templates ship with the app in `templates/`.

## Running it

This is a static site with no build step or backend, but it **must be served over HTTP** — templates are loaded with `fetch()`, which browsers block on `file://`, so opening `index.html` directly will not work.

```
node server.js
```
then visit http://localhost:8080. Any static file server works equally well (`npx serve .`, `python -m http.server`).

## Templates

The four journal templates live in `templates/` (`CCE.docx`, `SET.docx`, `ICE.docx`, `RJR.docx`) and are fetched on demand — only the template you click is downloaded, and it's cached for the rest of the session.

- **To update a template**: replace the file in `templates/`, keeping the same name.
- **To add a journal**: drop the `.docx` in `templates/` and add an entry to the `TEMPLATES` array at the top of `app.js`, plus a matching `<button class="template-card" data-template="...">` in `index.html`.

Placeholder tags are matched case- and spacing-insensitively (`{Paper Title}`, `{paper_title}`, and `{PAPERTITLE}` are equivalent), so a template with an inconsistently-spelled tag still fills correctly — the RJR template, for instance, spells its name tag `{NAMe}`. An unrecognized tag renders as empty text rather than breaking the export.

## How it works

1. **Choose a journal template** — CCE, SET, ICE, or RJR. All four contain the same placeholder tags: `{NAME}`, `{Designation}`, `{Paper Title}`, `{DOI}`, `{vol}`, `{issue}`, `{year}`, and `{month}` (docxtemplater performs the substitution). You can switch templates at any point without reloading; the next export uses the newly selected one.

2. **Paste mapping data** into the text box, one category per line. Each line is parsed by `parseDataInput`/`parseMappingLine` in `app.js`:
   - **Names**: a comma-separated list of names, each with a numeric suffix indicating which certificate(s) it maps to, e.g. `John Smith1, Jane Doe2`. A range suffix like `3-5` applies the same value to certificates 3 through 5, and a comma-separated list like `1,2,3` (or a mix, e.g. `1,3-5`) applies to exactly those certificates. The name line determines how many certificates are generated (the highest suffix number). Superscript numerals (e.g. `Name¹²`) are also recognized as suffix digits.
   - **Designations**: same comma+suffix format, e.g. `Professor, Example University1-2`. Lines are heuristically classified as a designation if they contain keywords such as "student", "professor", "department", "college", "university", etc.
   - **Paper Title**: a line without keyword matches is treated as a paper title. If it has no numeric suffixes it's applied as a constant to every certificate; if it has suffixes, it's mapped per-certificate the same way as names.
   - Use the **"Load Example Data"** button to see a sample of the expected input format.

3. **DOI**: enter the DOI in its own dedicated field (separate from the mapping text box). It's applied as a constant value to every generated certificate, and everything else is derived from it — there are no other fields to fill in.

   The final numeric segment packs **volume (2 digits) + issue + file number (2 digits)**, and the segment before it is the year. The issue is whatever sits between volume and file number: one digit for issues 1–9, two digits for issues 10–12. Since the issue number *is* the month of publication, `{month}` is derived from it rather than typed — issue 7 is July, issue 9 is September.

   So `10.17148/IJARCCE.2026.15817` → `{year}` `2026`, `{vol}` `15`, `{issue}` `8`, `{month}` `August`, file number `17`. Each output file is named `"{paper number} {author list} {certificate index}"`; if no DOI is entered, the paper number is omitted from the filename.

4. **Export**: generates every certificate and combines them into a single multi-page Word (`.docx`) file for download (`buildCombinedDocxBlob` in `app.js`). Optionally, running the local server (`node server.js`) exposes `POST /api/convert`, which converts a batch of DOCX files to pixel-faithful PDFs using Microsoft Word's own COM automation (`convert-docx-to-pdf.ps1`) — this requires Windows with Microsoft Word installed.

## Dependencies

- [PizZip](https://github.com/open-xml-templating/pizzip) + [docxtemplater](https://docxtemplater.com/) (loaded from CDN) — filling placeholders inside the DOCX template
- [Font Awesome](https://fontawesome.com/) and [Google Fonts](https://fonts.google.com/) (loaded from CDN) — UI typography
- Node.js standard library only for `server.js` — no npm dependencies
- Microsoft Word (Windows, COM automation) for the optional PDF conversion step
