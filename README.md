# Certificate Generator

A browser-based tool for bulk-generating certificates from a single Word (DOCX) template. Upload a template containing placeholder tags (e.g. `{NAME}`, `{Designation}`, `{Paper Title}`, `{DOI}`), paste in a block of recipient data, and the app fills in a copy of the DOCX per recipient. Each certificate is exported both as a native DOCX (via docxtemplater) and as a PDF rendered directly from that same generated DOCX (via docx-preview + html2canvas + jsPDF) — so the PDF is a faithful rendering of the exact DOCX, not a separately hand-drawn version, and the two always visually match.

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

1. **Upload a DOCX template** containing placeholder tags: `{NAME}`, `{Designation}`, `{Paper Title}` / `{PaperTitle}`, and `{DOI}` (docxtemplater performs the substitution, matching several case/spacing variants).

2. **Paste mapping data** into the text box, one category per line. Each line is parsed by `parseDataInput`/`parseMappingLine` in `app.js`:
   - **Names**: a comma-separated list of names, each with a numeric suffix indicating which certificate(s) it maps to, e.g. `John Smith1, Jane Doe2`. A range suffix like `3-5` applies the same value to certificates 3 through 5, and a comma-separated list like `1,2,3` (or a mix, e.g. `1,3-5`) applies to exactly those certificates. The name line determines how many certificates are generated (the highest suffix number). Superscript numerals (e.g. `Name¹²`) are also recognized as suffix digits.
   - **Designations**: same comma+suffix format, e.g. `Professor, Example University1-2`. Lines are heuristically classified as a designation if they contain keywords such as "student", "professor", "department", "college", "university", etc.
   - **Paper Title**: a line without keyword matches is treated as a paper title. If it has no numeric suffixes it's applied as a constant to every certificate; if it has suffixes, it's mapped per-certificate the same way as names.
   - Use the **"Load Example Data"** button to see a sample of the expected input format.

3. **DOI**: enter the DOI in its own dedicated field (separate from the mapping text box). It's applied as a constant value to every generated certificate. Each output file is named `"{paper number} {author list} {certificate index}"`, where the paper number is the last 2 digits of the DOI's final numeric segment (e.g. `10.17148/IJARCCE.2026.15817` → `17`); if no DOI is entered, the paper number is omitted from the filename.

4. **Preview and export**: a live preview renders the actual generated DOCX for the current record (via docx-preview) directly in the browser. Export as a ZIP containing both a `.docx` and a `.pdf` for every record, or download a single combined multi-page PDF (one page per record).

## Dependencies

All dependencies are loaded from CDNs — no local install/build required:

- [PizZip](https://github.com/open-xml-templating/pizzip) + [docxtemplater](https://docxtemplater.com/) — filling placeholders inside the DOCX template
- [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) — rendering the generated DOCX to HTML/DOM for both the live preview and PDF conversion
- [html2canvas](https://html2canvas.hertzen.com/) — rasterizing the rendered DOCX to an image for PDF conversion
- [jsPDF](https://github.com/parallax/jsPDF) — assembling the rasterized pages into PDF files
- [JSZip](https://stuk.github.io/jszip/) — bundling generated certificates into ZIP archives
- [Font Awesome](https://fontawesome.com/) and [Google Fonts](https://fonts.google.com/) — UI typography
