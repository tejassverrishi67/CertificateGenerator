// Zero-Styling Certificate Generator Logic (pre-loaded DOCX templates -> single combined DOCX)

// The four journal templates shipped with the app. They live in templates/ as static assets and
// are fetched on demand when the user picks one, so only the selected template is ever downloaded
// (SET alone is ~1.9 MB). This array is the single source of truth for the template picker -- the
// grid in index.html is generated from it (see renderTemplateGrid) rather than hardcoded, so
// adding a journal only means adding an entry here plus its .docx/logo files.
const TEMPLATES = [
    { key: 'CCE', fullName: 'IJARCCE',   file: 'templates/CCE.docx', logo: 'logos/CCE.jpg' },
    { key: 'SET', fullName: 'IARJSET',   file: 'templates/SET.docx', logo: 'logos/SET.jpg' },
    { key: 'ICE', fullName: 'IJIREEICE', file: 'templates/ICE.docx', logo: 'logos/ICE.jpg' },
    { key: 'RJR', fullName: 'IMRJR',     file: 'templates/RJR.docx', logo: 'logos/RJR.jpg' }
];

const state = {
    docxLoaded: false,       // Whether a DOCX template has been loaded
    docxBuffer: null,        // Raw ArrayBuffer of the selected DOCX template
    templateKey: '',         // Which journal template is currently selected
    buffers: {},             // Cache of already-fetched template ArrayBuffers, keyed by journal
    doiMode: 'with',         // 'with' = derive everything from the DOI, 'without' = manual entry
    records: [],             // Parsed recipient row data
    authorListLine: ''       // Original (superscript-preserved) author-list text, used in output filenames
};

const el = {
    templateGrid: document.getElementById('template-grid'),
    templateStatus: document.getElementById('template-status'),
    dataInput: document.getElementById('data-input'),
    doiInput: document.getElementById('doi-input'),
    modeToggle: document.getElementById('mode-toggle'),
    doiFields: document.getElementById('doi-fields'),
    manualFields: document.getElementById('manual-fields'),
    volInput: document.getElementById('vol-input'),
    issueInput: document.getElementById('issue-input'),
    fileNumInput: document.getElementById('filenum-input'),
    yearInput: document.getElementById('year-input'),
    parserStatus: document.getElementById('parser-status'),
    btnLoadDemo: document.getElementById('btn-load-demo'),
    btnDownloadAll: document.getElementById('btn-download-all'),
    btnDownloadMulti: document.getElementById('btn-download-multi'),
    btnDownloadGroup: document.getElementById('btn-download-group'),
    logsContainer: document.getElementById('logs-container')
};

// Logger utility
function log(msg, type = 'system') {
    const entry = document.createElement('div');
    entry.className = `log-entry ${type}`;

    let icon = 'fa-info-circle';
    if (type === 'success') icon = 'fa-circle-check';
    if (type === 'error') icon = 'fa-circle-exclamation';
    if (type === 'info') icon = 'fa-magnifying-glass';

    const iconEl = document.createElement('i');
    iconEl.className = `fa-solid ${icon}`;
    const textEl = document.createElement('span');
    textEl.textContent = msg;
    entry.appendChild(iconEl);
    entry.appendChild(document.createTextNode(' '));
    entry.appendChild(textEl);
    el.logsContainer.appendChild(entry);
    el.logsContainer.scrollTop = el.logsContainer.scrollHeight;
}

// Initialize listeners
function init() {
    renderTemplateGrid();
    setupEventListeners();
}

// Build the template picker from TEMPLATES -- the array is the only place a journal is defined.
function renderTemplateGrid() {
    el.templateGrid.innerHTML = '';
    TEMPLATES.forEach(template => {
        const card = document.createElement('button');
        card.type = 'button';
        card.className = 'template-card';
        card.dataset.template = template.key;
        card.innerHTML = `<img src="${template.logo}" alt="${template.fullName}"><span>${template.fullName}</span>`;
        el.templateGrid.appendChild(card);
    });
}

function setupEventListeners() {
    el.templateGrid.addEventListener('click', (e) => {
        const card = e.target.closest('.template-card');
        if (card) selectTemplate(card.dataset.template);
    });

    el.btnLoadDemo.addEventListener('click', loadDemoData);
    el.dataInput.addEventListener('input', handleDataInput);
    el.doiInput.addEventListener('input', handleDataInput);
    el.btnDownloadAll.addEventListener('click', handleDownloadAllInOne);
    el.btnDownloadMulti.addEventListener('click', handleDownloadMultiple);
    el.btnDownloadGroup.addEventListener('click', handleDownloadGroup);

    el.modeToggle.addEventListener('click', (e) => {
        const btn = e.target.closest('.mode-btn');
        if (btn) setDoiMode(btn.dataset.mode);
    });
    [el.volInput, el.issueInput, el.fileNumInput, el.yearInput].forEach(input => {
        input.addEventListener('input', handleDataInput);
    });
}

// Switch between deriving publication details from a DOI and entering them by hand.
function setDoiMode(mode) {
    state.doiMode = mode === 'without' ? 'without' : 'with';

    el.modeToggle.querySelectorAll('.mode-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.mode === state.doiMode);
    });
    el.doiFields.hidden = state.doiMode !== 'with';
    el.manualFields.hidden = state.doiMode !== 'without';

    handleDataInput();
}

// Load Example Data
function loadDemoData() {
    const demo = [
        'A Study on Sample Data Processing Techniques',
        'Mr.JOHN SMITH1, Ms.JANE DOE2',
        'Assistant Professor, Department of Computer Science, Example Institute of Technology, Example City, Example State, Example Country1',
        'Student, Department of Computer Science, Example Institute of Technology, Example City, Example State, Example Country2'
    ].join('\n');
    el.dataInput.value = demo;
    el.doiInput.value = '10.9999/example.2026.15817';
    handleDataInput();
}

// 1. Select one of the pre-loaded journal templates
function markActiveCard(key) {
    el.templateGrid.querySelectorAll('.template-card').forEach(card => {
        card.classList.toggle('active', card.dataset.template === key);
    });
}

async function selectTemplate(key) {
    const template = TEMPLATES.find(t => t.key === key);
    if (!template) return;

    markActiveCard(key);

    // Re-selecting an already-fetched template is instant: swap the cached buffer back in.
    if (state.buffers[key]) {
        state.docxBuffer = state.buffers[key];
        state.docxLoaded = true;
        state.templateKey = key;
        el.templateStatus.textContent = `${key} template selected`;
        log(`Switched to the ${key} template.`, 'success');
        handleDataInput();
        return;
    }

    el.templateStatus.textContent = `Loading ${key} template...`;
    log(`Loading the ${key} certificate template...`, 'system');

    try {
        const res = await fetch(template.file);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const buffer = await res.arrayBuffer();

        state.buffers[key] = buffer;
        state.docxBuffer = buffer;
        state.docxLoaded = true;
        state.templateKey = key;

        el.templateStatus.textContent = `${key} template selected`;
        log(`${key} template loaded successfully.`, 'success');
        handleDataInput();
    } catch (err) {
        // Leave docxLoaded false so the export button stays disabled.
        state.docxLoaded = false;
        markActiveCard('');
        el.templateStatus.textContent = 'No template selected';
        log(`Could not load the ${key} template: ${err.message}`, 'error');
        toggleButtons(false);
    }
}

// Helper to disable/enable export buttons
function toggleButtons(enabled) {
    el.btnDownloadAll.disabled = !enabled;
    el.btnDownloadMulti.disabled = !enabled;
    el.btnDownloadGroup.disabled = !enabled;
}

// 2. Data Inputs Parser
function handleDataInput() {
    const text = el.dataInput.value.trim();
    if (!text) {
        state.records = [];
        el.parserStatus.textContent = 'No data input';
        toggleButtons(false);
        return;
    }

    try {
        state.records = parseDataInput(text);
        el.parserStatus.textContent = `${state.records.length} records parsed successfully`;

        if (state.docxLoaded) {
            toggleButtons(true);
        }
    } catch (err) {
        el.parserStatus.textContent = `Parse error: ${err.message}`;
        toggleButtons(false);
    }
}

function parseMappingLine(line) {
    // Split on commas, then reassemble: a piece with no trailing digit suffix continues
    // the current value (values may legitimately contain commas, e.g. "Dept, City, Country1-2").
    // A piece that is PURELY numeric/range (e.g. "2", "3-4") is an additional suffix for the
    // most recent item, supporting discontinuous suffix lists like "India1,2,3".
    const pieces = line.split(',').map(p => p.trim());
    const items = [];
    let bufferVal = null;

    pieces.forEach(piece => {
        if (!piece) return;

        if (/^\d+(?:-\d+)?$/.test(piece) && items.length > 0) {
            items[items.length - 1].suffix += ',' + piece;
            return;
        }

        const match = piece.match(/\d+(?:-\d+)?$/);
        if (match) {
            const val = piece.substring(0, match.index).trim();
            const suffix = match[0];
            const fullVal = bufferVal ? `${bufferVal}, ${val}` : val;
            items.push({ val: fullVal, suffix });
            bufferVal = null;
        } else {
            bufferVal = bufferVal ? `${bufferVal}, ${piece}` : piece;
        }
    });

    return items;
}

// Convert superscript numeral characters (e.g. "Name¹²") to plain ASCII digits ("Name12")
const SUPERSCRIPT_MAP = { '⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4', '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9' };
function normalizeSuperscripts(text) {
    return text.replace(/[⁰¹²³⁴⁵⁶⁷⁸⁹]/g, ch => SUPERSCRIPT_MAP[ch]);
}

function parseDataInput(text) {
    // Keep original (pre-normalization) lines alongside normalized ones so the original
    // superscript-formatted author list can be reused verbatim (e.g. for output filenames),
    // while suffix/index parsing runs against the ASCII-digit normalized copy.
    const origLines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    const normLines = origLines.map(normalizeSuperscripts);

    // Classify lines dynamically
    const classification = {
        NAME: [],          // normalized text, used for suffix/index parsing
        NAME_ORIGINAL: [],  // original text (superscripts preserved), used for display/filenames
        Designation: [],
        PaperTitle: []
    };

    normLines.forEach((line, i) => {
        const trimmed = line.trim();
        if (!trimmed) return;

        // 1. Name Classify (e.g. contains Mr/Ms prefix, or multiple comma items with numeric suffixes)
        const isNameList = trimmed.includes('Mr.') || trimmed.includes('Ms.') || trimmed.includes('Mrs.') || trimmed.includes('Dr.');
        const parsed = parseMappingLine(trimmed);
        const allShort = parsed.length > 0 && parsed.every(item => item.val.length < 25);
        if (isNameList || (parsed.length > 1 && allShort)) {
            classification.NAME.push(trimmed);
            classification.NAME_ORIGINAL.push(origLines[i]);
            return;
        }

        // 2. Designation Classify (contains academic/corporate keywords)
        const keywords = ['student', 'professor', 'lecturer', 'department', 'college', 'university', 'researcher', 'studies', 'faculty', 'india', 'scholar'];
        const lower = trimmed.toLowerCase();
        const hasKeyword = keywords.some(kw => lower.includes(kw));
        if (hasKeyword) {
            classification.Designation.push(trimmed);
            return;
        }

        // 3. Default to Paper Title
        classification.PaperTitle.push(trimmed);
    });

    // We must have at least one NAME line to determine the count
    if (classification.NAME.length === 0) {
        throw new Error("Could not detect any Name line containing certificate mapping suffixes (e.g., Name1, Name2).");
    }

    // Store the original (superscript-preserved) author-list line(s) for filename generation
    state.authorListLine = classification.NAME_ORIGINAL.join(', ');

    // Parse names to determine maxIdx
    const nameItems = [];
    classification.NAME.forEach(line => {
        nameItems.push(...parseMappingLine(line));
    });

    let maxIdx = 0;
    nameItems.forEach(item => {
        const indexes = parseSuffix(item.suffix);
        indexes.forEach(idx => {
            if (idx > maxIdx) maxIdx = idx;
        });
    });

    if (maxIdx === 0) {
        throw new Error("Could not resolve certificate page mapping indexes from Name line.");
    }

    // Initialize records
    const records = Array.from({ length: maxIdx }, () => ({}));

    // Map Names
    nameItems.forEach(item => {
        const indexes = parseSuffix(item.suffix);
        indexes.forEach(idx => {
            if (idx - 1 < maxIdx) {
                records[idx - 1].NAME = item.val;
            }
        });
    });

    // Map Designations (supports multiple lines targeting different pages)
    classification.Designation.forEach(line => {
        const items = parseMappingLine(line);
        items.forEach(item => {
            const indexes = parseSuffix(item.suffix);
            indexes.forEach(idx => {
                if (idx - 1 < maxIdx) {
                    records[idx - 1].Designation = item.val;
                }
            });
        });
    });

    // Map Paper Titles (constant if no suffix, mapped if suffix is present)
    classification.PaperTitle.forEach(line => {
        const items = parseMappingLine(line);
        if (items.length > 0) {
            items.forEach(item => {
                const indexes = parseSuffix(item.suffix);
                indexes.forEach(idx => {
                    if (idx - 1 < maxIdx) {
                        records[idx - 1].PaperTitle = item.val;
                    }
                });
            });
        } else {
            // Apply as constant to all certificates
            records.forEach(record => {
                record.PaperTitle = line;
            });
        }
    });

    // Apply the publication details (DOI-derived or hand-entered) as constants to every certificate.
    const pub = getPublicationInfo();
    records.forEach(record => {
        record.DOI = pub.doi;
        record.Volume = pub.vol;
        record.Issue = pub.issue;
        record.Year = pub.year;
        record.Month = pub.month;
    });

    return records;
}

function parseSuffix(suffix) {
    const indexes = [];
    // Supports comma-separated lists of numbers and/or ranges, e.g. "1,3-5,7"
    suffix.split(',').forEach(part => {
        part = part.trim();
        if (!part) return;
        if (part.includes('-')) {
            const parts = part.split('-');
            const start = parseInt(parts[0], 10);
            const end = parseInt(parts[1], 10);
            if (!isNaN(start) && !isNaN(end)) {
                for (let i = start; i <= end; i++) {
                    indexes.push(i);
                }
            }
        } else {
            const idx = parseInt(part, 10);
            if (!isNaN(idx)) {
                indexes.push(idx);
            }
        }
    });
    return indexes;
}

// Extract the journal "file number" from a DOI, e.g. "10.17148/IJARCCE.2026.15817" -> "17"
// (the last 2 digits of the final numeric segment: VOLUME(15) + ISSUE(8) + FILE(17)).
function extractPaperNumber(doi) {
    if (!doi) return '';
    const runs = doi.match(/\d+/g);
    if (!runs || runs.length === 0) return '';
    return runs[runs.length - 1].slice(-2);
}

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June',
                     'July', 'August', 'September', 'October', 'November', 'December'];

// Extract Volume/Issue/Year/Month encoded in a DOI, e.g. "10.17148/IJARCCE.2026.15817" ->
// { vol: "15", issue: "8", year: "2026", month: "August" }. The final numeric segment packs
// VOLUME(2 digits) + ISSUE + FILE NUMBER(2 digits), so the issue is whatever sits between them
// -- one digit for issues 1-9, two for issues 10-12. The segment before it is the year.
//
// The issue number IS the month of publication (issue 7 = July, issue 9 = September), so the
// month is derived rather than entered by hand.
function parseDoiParts(doi) {
    const empty = { vol: '', issue: '', year: '', month: '' };
    if (!doi) return empty;
    const runs = doi.match(/\d+/g);
    if (!runs || runs.length < 2) return empty;
    const last = runs[runs.length - 1];
    const year = runs[runs.length - 2];
    if (last.length < 5) return { vol: '', issue: '', year, month: '' };

    const issueNum = parseInt(last.slice(2, -2), 10);
    return {
        vol: last.slice(0, 2),
        issue: isNaN(issueNum) ? '' : String(issueNum),
        year,
        month: MONTH_NAMES[issueNum - 1] || ''
    };
}

// The single source of truth for volume/issue/year/month/file number, resolved from whichever
// input mode is active. In "with DOI" mode everything is decoded from the DOI; in "without DOI"
// mode the user supplies volume, issue, file number and year by hand. The issue number sets the
// month in both modes, so that rule never has to be restated by the user.
function getPublicationInfo() {
    if (state.doiMode === 'with') {
        const doi = el.doiInput.value.trim();
        const parts = parseDoiParts(doi);
        return { doi, ...parts, fileNumber: extractPaperNumber(doi) };
    }

    const issueNum = parseInt(el.issueInput.value.trim(), 10);
    return {
        doi: '',
        vol: el.volInput.value.trim(),
        issue: isNaN(issueNum) ? '' : String(issueNum),
        year: el.yearInput.value.trim(),
        month: MONTH_NAMES[issueNum - 1] || '',
        fileNumber: el.fileNumInput.value.trim()
    };
}

// Only strip characters that are actually invalid in filenames; keep spaces, commas, superscripts.
function sanitizeFilename(name) {
    return name.replace(/[\\/:*?"<>|]/g, '_');
}

// Build the output filename (without extension) shared by the "All in One" and "Group Download"
// modes, both of which produce a single file: "{fileNumber} {full author-list line}",
// e.g. "17 Author1, Author2". The file number comes from the DOI or the manual field depending on
// the active mode, and is omitted when not available.
function buildCombinedFilename() {
    const { fileNumber } = getPublicationInfo();
    const parts = [];
    if (fileNumber) parts.push(fileNumber);
    if (state.authorListLine) parts.push(state.authorListLine);
    return sanitizeFilename(parts.join(' ').trim() || 'certificates');
}

// Build the output filename (without extension) for one certificate in "Multiple Downloads" mode:
// "{fileNumber} {full author-list line} {certIndex}" -- every file shares the same base name and
// differs only by the trailing 1-based certificate index.
function buildIndividualFilename(index) {
    const { fileNumber } = getPublicationInfo();
    const parts = [];
    if (fileNumber) parts.push(fileNumber);
    if (state.authorListLine) parts.push(state.authorListLine);
    parts.push(String(index + 1));
    return sanitizeFilename(parts.join(' ').trim());
}

// Collapse a list of values to their distinct entries (first-occurrence order, empties dropped),
// comma-joined. Used by Group Download to combine every record's Designation/PaperTitle onto the
// single combined certificate without repeating a value that's shared by multiple authors.
function dedupeJoin(values) {
    const seen = [];
    values.forEach(v => {
        const trimmed = (v || '').trim();
        if (trimmed && !seen.includes(trimmed)) seen.push(trimmed);
    });
    return seen.join(', ');
}

// Build the single synthetic record used by Group Download: every recipient's name and
// designation combined onto one certificate, rather than one certificate per recipient.
function buildGroupRecord() {
    const first = state.records[0];
    return {
        NAME: state.records.map(r => r.NAME).join(', '),
        Designation: dedupeJoin(state.records.map(r => r.Designation)),
        PaperTitle: dedupeJoin(state.records.map(r => r.PaperTitle)),
        DOI: first.DOI,
        Volume: first.Volume,
        Issue: first.Issue,
        Year: first.Year,
        Month: first.Month
    };
}

// Trigger a browser download of a blob, then release the object URL once the click is queued.
function triggerDownload(blob, filename) {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
}

// Reduce a placeholder tag to a canonical form so lookups tolerate case and spacing differences:
// "Paper Title", "paper_title" and "PAPERTITLE" all collapse to "papertitle".
function normalizeTag(tag) {
    return String(tag).toLowerCase().replace(/[\s_]+/g, '');
}

// 3. Generate a filled DOCX package (PizZip instance) for one record
function renderDocxZipForRecord(record) {
    const docZip = new window.PizZip(state.docxBuffer);

    // Canonical values, keyed by normalized tag name. This backs the nullGetter below, which
    // catches placeholders whose spelling doesn't exactly match a key in setData -- notably the
    // RJR template, which spells its name tag "{NAMe}". Without this, docxtemplater's default
    // nullGetter would stamp the literal text "undefined" onto the certificate. Unknown tags
    // resolve to an empty string rather than failing the render.
    const canonical = {
        name: record.NAME ? record.NAME.toUpperCase() : '',
        designation: record.Designation || '',
        papertitle: record.PaperTitle || '',
        doi: record.DOI || '',
        vol: record.Volume || '',
        volume: record.Volume || '',
        issue: record.Issue || '',
        year: record.Year || '',
        month: record.Month || ''
    };

    const doc = new window.docxtemplater(docZip, {
        paragraphLoop: true,
        linebreaks: true,
        nullGetter(part) {
            const value = canonical[normalizeTag(part.value)];
            return value === undefined ? '' : value;
        },
    });

    // Map data (support both space, no space, underscore, and case variants for absolute safety)
    doc.setData({
        NAME: record.NAME ? record.NAME.toUpperCase() : '',
        name: record.NAME ? record.NAME.toUpperCase() : '',
        Designation: record.Designation || '',
        designation: record.Designation || '',
        "Paper Title": record.PaperTitle || '',
        "paper title": record.PaperTitle || '',
        PaperTitle: record.PaperTitle || '',
        papertitle: record.PaperTitle || '',
        DOI: record.DOI || '',
        doi: record.DOI || '',
        vol: record.Volume || '',
        Volume: record.Volume || '',
        issue: record.Issue || '',
        Issue: record.Issue || '',
        year: record.Year || '',
        Year: record.Year || '',
        month: record.Month || '',
        Month: record.Month || ''
    });

    doc.render();

    return doc.getZip();
}

function renderDocxBlobForRecord(record) {
    return renderDocxZipForRecord(record).generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
}

// Extract just the paragraph content of a rendered record's body (everything between
// <w:body> and its trailing <w:sectPr>, exclusive).
function extractRecordBodyContent(xml) {
    const bodyOpenIdx = xml.indexOf('<w:body>') + '<w:body>'.length;
    const sectPrIdx = xml.lastIndexOf('<w:sectPr');
    return xml.slice(bodyOpenIdx, sectPrIdx);
}

// Combine every record's rendered DOCX into a single multi-page DOCX file, one certificate
// per page. Records are separated by a real section break (an empty paragraph carrying a
// copy of the template's own <w:sectPr>) rather than a manual page break: since the template
// uses <w:titlePg/> (a distinct "first page" header/layout), a manual page break would push
// every certificate after the first onto the section's "default" (non-first) page, which
// renders differently (visible as stray whitespace/misalignment). Giving each certificate its
// own one-page section means every certificate consistently gets the section's "first page"
// treatment, matching how page 1 renders.
function buildCombinedDocxBlob(records) {
    const baseZip = renderDocxZipForRecord(records[0]);
    const xml0 = baseZip.file('word/document.xml').asText();

    const bodyOpenIdx = xml0.indexOf('<w:body>') + '<w:body>'.length;
    const sectPrStart = xml0.lastIndexOf('<w:sectPr');
    const sectPrEnd = xml0.indexOf('</w:sectPr>', sectPrStart) + '</w:sectPr>'.length;
    const sectPrXml = xml0.slice(sectPrStart, sectPrEnd);
    const prefix = xml0.slice(0, bodyOpenIdx);
    const suffix = xml0.slice(sectPrEnd); // "</w:body></w:document>"

    const bodies = [extractRecordBodyContent(xml0)];
    for (let i = 1; i < records.length; i++) {
        const recordZip = renderDocxZipForRecord(records[i]);
        const recordXml = recordZip.file('word/document.xml').asText();
        bodies.push(extractRecordBodyContent(recordXml));
    }

    const sectionBreakParagraph = `<w:p><w:pPr>${sectPrXml}</w:pPr></w:p>`;
    const combinedXml = prefix + bodies.join(sectionBreakParagraph) + sectPrXml + suffix;

    baseZip.file('word/document.xml', combinedXml);
    return baseZip.generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
}

// Shared "disable every download button, show a spinner on the clicked one, restore afterward"
// wrapper so a user can't fire two exports at once, and errors from any mode are logged the
// same way.
async function runDownload(btnEl, task) {
    if (state.records.length === 0 || !state.docxLoaded) return;

    const initialHTML = btnEl.innerHTML;
    toggleButtons(false);
    btnEl.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Generating...';

    try {
        await task();
    } catch (err) {
        log(`Failed generation: ${err.message}`, 'error');
        console.error(err);
    } finally {
        btnEl.innerHTML = initialHTML;
        toggleButtons(state.docxLoaded && state.records.length > 0);
    }
}

// 4a. All in One: every certificate combined into a single multi-page Word file
async function handleDownloadAllInOne() {
    await runDownload(el.btnDownloadAll, async () => {
        log(`Generating ${state.records.length} certificate(s) into a single Word document...`, 'system');
        const blob = buildCombinedDocxBlob(state.records);
        triggerDownload(blob, `${buildCombinedFilename()}.docx`);
        log(`Successfully generated and downloaded the combined Word document!`, 'success');
    });
}

// 4b. Multiple Downloads: one separate Word file per recipient, triggered in sequence. Each
// download is staggered slightly so Chrome doesn't drop downloads fired in the same tick; the
// browser may show a one-time "this site wants to download multiple files" prompt after the
// second file, which is expected.
async function handleDownloadMultiple() {
    await runDownload(el.btnDownloadMulti, async () => {
        const total = state.records.length;
        log(`Generating ${total} individual certificate(s)...`, 'system');
        for (let i = 0; i < total; i++) {
            const blob = renderDocxBlobForRecord(state.records[i]);
            triggerDownload(blob, `${buildIndividualFilename(i)}.docx`);
            log(`Downloaded certificate ${i + 1} of ${total}.`, 'success');
            if (i < total - 1) await new Promise(resolve => setTimeout(resolve, 250));
        }
        log(`All ${total} certificates downloaded individually.`, 'success');
    });
}

// 4c. Group Download: every recipient combined onto one certificate (one page, one file).
async function handleDownloadGroup() {
    await runDownload(el.btnDownloadGroup, async () => {
        log(`Generating one combined certificate for all ${state.records.length} author(s)...`, 'system');
        const blob = buildCombinedDocxBlob([buildGroupRecord()]);
        triggerDownload(blob, `${buildCombinedFilename()}.docx`);
        log(`Successfully generated and downloaded the group certificate!`, 'success');
    });
}

// DOM trigger
document.addEventListener('DOMContentLoaded', init);
if (document.readyState === 'interactive' || document.readyState === 'complete') {
    init();
}
