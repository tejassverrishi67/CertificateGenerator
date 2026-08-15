// Zero-Styling Certificate Generator Logic (DOCX template -> DOCX + PDF rendered from that DOCX)

const state = {
    docxLoaded: false,       // Whether a DOCX template has been loaded
    docxBuffer: null,        // Raw ArrayBuffer of the DOCX template
    records: [],             // Parsed recipient row data
    authorListLine: ''       // Original (superscript-preserved) author-list text, used in output filenames
};

const el = {
    uploadZoneDocx: document.getElementById('upload-zone-docx'),
    docxUpload: document.getElementById('docx-upload'),
    docxFileName: document.getElementById('docx-file-name'),
    dataInput: document.getElementById('data-input'),
    doiInput: document.getElementById('doi-input'),
    parserStatus: document.getElementById('parser-status'),
    btnLoadDemo: document.getElementById('btn-load-demo'),
    btnGenerateZip: document.getElementById('btn-generate-zip'),
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
    setupEventListeners();
}

function setupEventListeners() {
    setupUploadZone(el.uploadZoneDocx, el.docxUpload, handleDocxFileSelect);

    el.btnLoadDemo.addEventListener('click', loadDemoData);
    el.dataInput.addEventListener('input', handleDataInput);
    el.doiInput.addEventListener('input', handleDataInput);
    el.btnGenerateZip.addEventListener('click', handleZipExport);
}

function setupUploadZone(zoneEl, inputEl, onFile) {
    zoneEl.addEventListener('click', () => inputEl.click());
    inputEl.addEventListener('change', () => {
        if (inputEl.files.length > 0) onFile(inputEl.files[0]);
    });

    zoneEl.addEventListener('dragover', (e) => {
        e.preventDefault();
        zoneEl.classList.add('dragover');
    });
    zoneEl.addEventListener('dragleave', () => {
        zoneEl.classList.remove('dragover');
    });
    zoneEl.addEventListener('drop', (e) => {
        e.preventDefault();
        zoneEl.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            inputEl.files = e.dataTransfer.files;
            onFile(e.dataTransfer.files[0]);
        }
    });
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
    el.doiInput.value = '10.9999/example.2026.00001';
    handleDataInput();
}

// 1. Load Template File (DOCX)
function handleDocxFileSelect(file) {
    if (!file) return;
    el.docxFileName.textContent = file.name;
    log(`Loading DOCX file: ${file.name}...`, 'system');

    const reader = new FileReader();
    reader.onload = function(e) {
        state.docxBuffer = e.target.result;
        state.docxLoaded = true;
        log(`Word Document (DOCX) loaded successfully.`, 'success');
        log(`Format mappings to replace {NAME}, {Designation}, {PaperTitle}, and {DOI} tags inside your Word document.`, 'info');
        handleDataInput();
    };
    reader.readAsArrayBuffer(file);
}

// Helper to disable/enable export buttons
function toggleButtons(enabled) {
    el.btnGenerateZip.disabled = !enabled;
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

    // Map DOI (from the dedicated DOI input field, constant applied to all certificates)
    const doiValue = el.doiInput.value.trim();
    if (doiValue) {
        records.forEach(record => {
            record.DOI = doiValue;
        });
    }

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

// Build the per-certificate output filename (without extension):
// "{paperNumber} {full author-list line} {certIndex}" when a DOI is provided,
// otherwise "{full author-list line} {certIndex}".
function buildOutputFilename(index) {
    const paperNumber = extractPaperNumber(el.doiInput.value.trim());
    const parts = [];
    if (paperNumber) parts.push(paperNumber);
    if (state.authorListLine) parts.push(state.authorListLine);
    parts.push(String(index + 1));

    let name = parts.join(' ').trim();
    if (!name) name = `certificate_${index + 1}`;
    // Only strip characters that are actually invalid in filenames; keep spaces, commas, superscripts.
    return name.replace(/[\\/:*?"<>|]/g, '_');
}

// 3. Generate a filled DOCX package (PizZip instance) for one record
function renderDocxZipForRecord(record) {
    const docZip = new window.PizZip(state.docxBuffer);
    const doc = new window.docxtemplater(docZip, {
        paragraphLoop: true,
        linebreaks: true,
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
        doi: record.DOI || ''
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

// 4. Download all certificates combined into a single Word (DOCX) file
async function handleZipExport() {
    if (state.records.length === 0 || !state.docxLoaded) return;

    const initialText = el.btnGenerateZip.innerHTML;
    el.btnGenerateZip.disabled = true;
    el.btnGenerateZip.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Generating...';

    try {
        log(`Generating ${state.records.length} certificate(s) into a single Word document...`, 'system');
        const combinedBlob = buildCombinedDocxBlob(state.records);
        const fileName = state.records.length === 1 ? buildOutputFilename(0) : (state.authorListLine || 'certificates');

        const link = document.createElement('a');
        link.href = URL.createObjectURL(combinedBlob);
        link.download = `${fileName.replace(/[\\/:*?"<>|]/g, '_')}.docx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(link.href);

        log(`Successfully generated and downloaded the combined Word document!`, 'success');
    } catch (err) {
        log(`Failed generation: ${err.message}`, 'error');
        console.error(err);
    } finally {
        el.btnGenerateZip.disabled = false;
        el.btnGenerateZip.innerHTML = initialText;
    }
}

// DOM trigger
document.addEventListener('DOMContentLoaded', init);
if (document.readyState === 'interactive' || document.readyState === 'complete') {
    init();
}
