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

// Upper bound on how many certificates a single paste can request. Without this, a typo'd suffix
// range (e.g. "1-500" meant to be "1-5") silently builds hundreds of records and can freeze the tab.
const MAX_CERTIFICATES = 200;

const state = {
    docxLoaded: false,       // Whether a DOCX template has been loaded
    docxBuffer: null,        // Raw ArrayBuffer of the selected DOCX template
    templateKey: '',         // Which journal template is currently selected
    buffers: {},             // Cache of already-fetched template ArrayBuffers, keyed by journal
    doiMode: 'with',         // 'with' = derive everything from the DOI, 'without' = manual entry
    records: [],             // Parsed recipient row data
    authorListLine: '',      // Original (superscript-preserved) author-list text, used in output filenames
    loadToken: 0,             // Bumped on every template pick; guards against a stale fetch() resolving late
    librariesMissing: false, // True if PizZip/docxtemplater failed to load from the CDN
    _lastParseError: null,
    _lastWarningsKey: ''
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
    previewCard: document.getElementById('preview-card'),
    parseWarnings: document.getElementById('parse-warnings'),
    previewTableBody: document.getElementById('preview-table-body'),
    previewPubInfo: document.getElementById('preview-pub-info'),
    livePreview: document.getElementById('live-preview'),
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

    // The DOCX libraries load from a CDN (index.html); if that request is blocked or fails,
    // every download would previously fail only once the user clicked a button, with an opaque
    // "PizZip is not a constructor" error. Detect it up front instead.
    if (typeof window.PizZip !== 'function' || typeof window.docxtemplater !== 'function') {
        state.librariesMissing = true;
        toggleButtons(false);
        log('DOCX libraries (PizZip/docxtemplater) failed to load from the CDN -- downloads are disabled. Check your network connection and reload the page.', 'error');
    }

    // The Live Preview needs docx-preview (+ JSZip). It's non-critical -- downloads still work
    // and the panel falls back to a text summary -- so this is a note, not an error.
    if (!docxPreviewAvailable()) {
        log('Live preview renderer (docx-preview) failed to load from the CDN -- the preview panel will show a plain-text summary instead.', 'system');
    }
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

    // Leaving "With DOI" clears the DOI box so a stale value can't reappear (or be picked up
    // again) if the user switches back and forth. In "without" mode the DOI renders empty.
    if (state.doiMode === 'without') el.doiInput.value = '';

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
    // Guard against a slower, earlier fetch() resolving AFTER a later click: only the most
    // recent selectTemplate() call is allowed to commit its result to state.
    const myToken = ++state.loadToken;

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

        if (myToken !== state.loadToken) return; // superseded by a newer selection

        state.buffers[key] = buffer;
        state.docxBuffer = buffer;
        state.docxLoaded = true;
        state.templateKey = key;

        el.templateStatus.textContent = `${key} template selected`;
        log(`${key} template loaded successfully.`, 'success');
        handleDataInput();
    } catch (err) {
        if (myToken !== state.loadToken) return; // superseded by a newer selection

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
    const canEnable = enabled && !state.librariesMissing;
    el.btnDownloadAll.disabled = !canEnable;
    el.btnDownloadMulti.disabled = !canEnable;
    el.btnDownloadGroup.disabled = !canEnable;
}

// 2. Data Inputs Parser
function handleDataInput() {
    const text = el.dataInput.value.trim();
    if (!text) {
        state.records = [];
        state.authorListLine = '';
        el.parserStatus.textContent = 'No data input';
        toggleButtons(false);
        renderPreview([], []);
        renderLivePreview({ records: [] });
        return;
    }

    try {
        const result = parseDataInput(text);
        state.records = result.records;
        state.authorListLine = result.authorListLine;
        el.parserStatus.textContent = `${state.records.length} record(s) parsed`
            + (state.docxLoaded ? '' : ' — pick a template to enable downloads');
        toggleButtons(state.docxLoaded && state.records.length > 0);
        state._lastParseError = null;
        renderPreview(result.records, result.warnings);
        renderLivePreview({ records: result.records });

        // Log each distinct set of warnings once, rather than re-logging on every keystroke.
        const warningsKey = result.warnings.join('\n');
        if (warningsKey && warningsKey !== state._lastWarningsKey) {
            result.warnings.forEach(w => log(w, 'error'));
        }
        state._lastWarningsKey = warningsKey;
    } catch (err) {
        state.records = [];
        el.parserStatus.textContent = `Parse error: ${err.message}`;
        toggleButtons(false);
        renderPreview([], []);
        renderLivePreview({ records: [], error: err.message });
        // Surface it in the visible log too (once per distinct message) so a failed parse
        // isn't just a greyed-out button with no explanation.
        if (state._lastParseError !== err.message) {
            log(`Could not read the recipient data: ${err.message}`, 'error');
            state._lastParseError = err.message;
        }
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

    // A trailing piece with no suffix of its own (e.g. the ", Chennai" in "Professor, Anna
    // University1, Chennai") has nowhere else to go -- fold it onto the last mapped item
    // instead of silently dropping it. A line with no suffixed piece at all (a plain
    // constant-title line) is intentionally left as an empty items array; its raw text is used
    // verbatim by the caller, so there's nothing to flush in that case.
    if (bufferVal && items.length > 0) {
        items[items.length - 1].val += `, ${bufferVal}`;
    }

    return items;
}

// Convert superscript numeral characters (e.g. "Name¹²") to plain ASCII digits ("Name12").
// The superscript minus/plus are included so a superscript range like "India¹⁻²" normalizes to
// the ASCII "India1-2" that parseMappingLine understands, instead of corrupting the value.
const SUPERSCRIPT_MAP = {
    '⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4', '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9',
    '⁻': '-', '⁺': '+'
};
const SUPERSCRIPT_CHARS = /[⁰¹²³⁴⁵⁶⁷⁸⁹⁻⁺]/;
function normalizeSuperscripts(text) {
    return text.replace(new RegExp(SUPERSCRIPT_CHARS, 'g'), ch => SUPERSCRIPT_MAP[ch]);
}

// Affiliation/designation keywords, and the generic words that mark a short suffixed value as a
// paper title rather than a name (e.g. the app's own "Sample Paper Title1-2" placeholder). Both
// are matched as WHOLE WORDS, not substrings -- a substring match would flag ordinary prose like
// "Indian Universities" (contains "india"/"university" as substrings) as an affiliation line.
const NAME_KEYWORDS = ['student', 'professor', 'lecturer', 'department', 'college', 'university', 'researcher', 'studies', 'faculty', 'india', 'scholar'];
const NAME_KEYWORD_RE = new RegExp(`\\b(?:${NAME_KEYWORDS.join('|')})\\b`, 'i');
const NAME_KEYWORD_RE_G = new RegExp(`\\b(?:${NAME_KEYWORDS.join('|')})\\b`, 'gi');
const TITLE_HINT_WORDS = ['title', 'paper', 'study', 'analysis', 'survey', 'review', 'research', 'abstract', 'thesis'];
const TITLE_HINT_RE = new RegExp(`\\b(?:${TITLE_HINT_WORDS.join('|')})\\b`, 'i');

function parseDataInput(text) {
    // Keep original (pre-normalization) lines alongside normalized ones so the original
    // superscript-formatted author list can be reused verbatim (e.g. for output filenames),
    // while suffix/index parsing runs against the ASCII-digit normalized copy.
    const rawLines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);

    // Journal author blocks routinely wrap one affiliation across two lines. A line that ends
    // with a comma is a continuation, so stitch it onto the next line before classifying --
    // otherwise the tail ("Tamil Nadu, India1") is parsed as its own line while the head
    // ("Professor and Head, Department of Commerce,") loses its mapping suffix and is dropped.
    const origLines = [];
    rawLines.forEach(line => {
        const prev = origLines[origLines.length - 1];
        if (prev && /,$/.test(prev)) {
            origLines[origLines.length - 1] = prev + ' ' + line;
        } else {
            origLines.push(line);
        }
    });
    const normLines = origLines.map(normalizeSuperscripts);

    // Classify lines dynamically
    const classification = {
        NAME: [],          // normalized text, used for suffix/index parsing
        NAME_ORIGINAL: [],  // original text (superscripts preserved), used for display/filenames
        Designation: [],
        PaperTitle: []
    };

    const addName = (norm, orig) => {
        classification.NAME.push(norm);
        classification.NAME_ORIGINAL.push(orig);
    };

    // A single short (<=3 word) suffixed value with no commas is inherently ambiguous: it could
    // be a lone recipient's name ("Ravi Kumar1") or a short paper title mapped to several
    // certificates via a range ("Deep Learning Advances1-3"). Rather than guess per-line, every
    // such value is queued here and resolved in one pass after all lines are classified (below),
    // based on whether a STRONGER, unambiguous name signal (an honorific, a comma-separated list
    // of short values, or a superscript-tagged author) turned up anywhere else in the input.
    const ambiguousShortValues = [];

    normLines.forEach((line, i) => {
        const trimmed = line.trim();
        if (!trimmed) return;

        const parsed = parseMappingLine(trimmed);
        const hasSuffix = parsed.length > 0 && parsed.some(item => /\d/.test(item.suffix));
        const allShort = parsed.length > 0 && parsed.every(item => item.val.length < 40);
        const noInnerCommas = parsed.every(item => !item.val.includes(','));
        const hadSuperscript = SUPERSCRIPT_CHARS.test(origLines[i] || '');
        const singleValWords = parsed.length === 1 ? parsed[0].val.split(/\s+/).filter(Boolean).length : 0;

        // Keyword/title-hint matching ignores trailing suffix digits (so "India1" still reads
        // as the keyword "india") but keeps whole-word boundaries elsewhere.
        const keywordProbe = trimmed.replace(/\d+/g, ' ');
        const keywordMatches = keywordProbe.match(NAME_KEYWORD_RE_G) || [];
        const hasKeyword = keywordMatches.length > 0;
        const firstWord = (trimmed.split(/\s+/)[0] || '').replace(/[.,;:]+$/, '').replace(/\d+/g, '');
        const startsWithKeyword = NAME_KEYWORD_RE.test(firstWord);
        const commaCount = (trimmed.match(/,/g) || []).length;
        const totalWords = trimmed.split(/\s+/).filter(Boolean).length;

        // 1. An honorific ("Mr.", "Dr ", "Prof.") at the START of a comma-separated author
        //    segment that also carries a mapping suffix ("Dr. Arunpriya S1") is an unambiguous
        //    name marker. Requiring it at the segment start -- not anywhere in the line -- keeps
        //    an institution whose name contains an honorific ("Dr. N.G.P. Arts and Science
        //    College") from turning its whole affiliation line into a spurious name. The dot or
        //    space after the honorific keeps ordinary names ("Mrinal1", "Drithi1") from matching.
        const honorificName = parsed.some(item =>
            /^\(?(Mr|Mrs|Ms|Dr|Prof)(\.|\s)/i.test(item.val.trim()) && /\d/.test(item.suffix));
        if (honorificName) {
            addName(trimmed, origLines[i]);
            return;
        }

        // 2. A suffixed, comma-free, comma-separated list of 2+ short values, or ANY
        //    superscript-tagged suffixed line, reads unambiguously as a name list -- a real
        //    paper title is essentially never formatted this way. This runs BEFORE any
        //    keyword/title heuristic so an ordinary line like "Ravi Kumar1, Anita Rao2" is
        //    never mistaken for prose.
        if (hasSuffix && noInnerCommas && ((parsed.length >= 2 && allShort) || hadSuperscript)) {
            addName(trimmed, origLines[i]);
            return;
        }

        // 3. A short suffixed value that names itself as a title ("Sample Paper Title1-2", the
        //    app's own example) -- caught before the ambiguous single-value rule below, and
        //    before the prose-title rule since it's under that rule's 4-word floor.
        const singleValueTitleHint = parsed.length === 1 && TITLE_HINT_RE.test(parsed[0].val);
        if (hasSuffix && noInnerCommas && parsed.length === 1 && singleValueTitleHint) {
            classification.PaperTitle.push(trimmed);
            return;
        }

        // 3b. A single short (<=3 word) suffixed value with no commas -- e.g. "Ravi Kumar1" or
        //     "Deep Learning Advances1-3" -- is ambiguous between a lone recipient's name and a
        //     short paper title. Deferred to ambiguousShortValues for resolution after the loop
        //     (see below) instead of guessed here.
        if (hasSuffix && noInnerCommas && parsed.length === 1 &&
            singleValWords >= 1 && singleValWords <= 3 && parsed[0].val.length < 28 && !singleValueTitleHint) {
            ambiguousShortValues.push({ trimmed, orig: origLines[i] });
            return;
        }

        // 4. Prose paper title: few commas, several words, and not an affiliation line. This
        //    runs before the keyword check below so an ordinary title that happens to contain
        //    one affiliation word ("...Student Performance in Indian Universities") isn't
        //    swallowed as a Designation -- but a line that STARTS with a keyword ("Department
        //    of...") or piles up 2+ of them still reads as an affiliation.
        if (commaCount <= 1 && totalWords >= 4 && !startsWithKeyword && keywordMatches.length < 2) {
            classification.PaperTitle.push(trimmed);
            return;
        }

        // 5. Academic/corporate keywords mark an affiliation/designation line.
        if (hasKeyword) {
            classification.Designation.push(trimmed);
            return;
        }

        // 6. Any other suffixed line (short, or with internal commas) is an affiliation.
        if (hasSuffix) {
            classification.Designation.push(trimmed);
            return;
        }

        // 7. No suffix at all -> a constant paper title.
        classification.PaperTitle.push(trimmed);
    });

    // Resolve the deferred single-short-value lines (rule 3b): if a stronger, unambiguous name
    // signal was found elsewhere, these are almost certainly short paper titles -- each mapped
    // by its own suffix (or applied to every certificate if its suffix doesn't resolve). If
    // nothing else looked like a name, they're promoted to NAME instead, so a lone recipient on
    // a line by itself ("Ravi Kumar1") still works.
    if (classification.NAME.length > 0) {
        ambiguousShortValues.forEach(({ trimmed }) => classification.PaperTitle.push(trimmed));
    } else {
        // Two ambiguous single-value lines can still target the SAME certificate index (e.g. a
        // title mistakenly suffixed "1" sitting next to a name also suffixed "1") -- promoting
        // both to NAME would let the second silently overwrite the first. Only the first line to
        // claim a given index is promoted; anything colliding with an already-claimed index
        // falls back to PaperTitle instead of clobbering it.
        const claimedIndexes = new Set();
        ambiguousShortValues.forEach(({ trimmed, orig }) => {
            const indexes = parseMappingLine(trimmed).reduce((acc, it) => acc.concat(parseSuffix(it.suffix)), []);
            const collides = indexes.some(idx => claimedIndexes.has(idx));
            if (collides) {
                classification.PaperTitle.push(trimmed);
            } else {
                indexes.forEach(idx => claimedIndexes.add(idx));
                addName(trimmed, orig);
            }
        });
    }

    // Fallback: nothing looked like a name, but a line carries mapping suffixes -- promote the
    // first such line (affiliations are checked before titles) rather than failing outright.
    if (classification.NAME.length === 0) {
        for (const pool of [classification.Designation, classification.PaperTitle]) {
            const idx = pool.findIndex(l => parseMappingLine(l).some(it => /\d/.test(it.suffix)));
            if (idx !== -1) {
                const [line] = pool.splice(idx, 1);
                addName(line, line);
                break;
            }
        }
    }

    // We must have at least one NAME line to determine the count
    if (classification.NAME.length === 0) {
        throw new Error("Could not detect any Name line. Add a numeric suffix to at least one name, e.g. \"Ravi Kumar1\".");
    }

    // The original (superscript-preserved) author-list line(s), used for filenames.
    const authorListLine = classification.NAME_ORIGINAL.join(', ');

    // Parse names to determine maxIdx
    const nameItems = [];
    classification.NAME.forEach(line => {
        nameItems.push(...parseMappingLine(line));
    });

    let maxIdx = 0;
    nameItems.forEach(item => {
        parseSuffix(item.suffix).forEach(idx => {
            if (idx > maxIdx) maxIdx = idx;
        });
    });

    if (maxIdx === 0) {
        throw new Error("Could not resolve certificate page mapping indexes from Name line.");
    }
    if (maxIdx > MAX_CERTIFICATES) {
        throw new Error(`Name suffixes imply ${maxIdx} certificates, which is over the ${MAX_CERTIFICATES}-certificate limit. Check for a typo in a suffix range (e.g. "1-500" instead of "1-5").`);
    }

    // Initialize records
    const records = Array.from({ length: maxIdx }, () => ({}));

    const inRange = idx => idx >= 1 && idx <= maxIdx;

    // Map Names
    nameItems.forEach(item => {
        parseSuffix(item.suffix).forEach(idx => {
            if (inRange(idx)) records[idx - 1].NAME = item.val;
        });
    });

    // Map a "field, suffix" line onto records.
    function applyMappedLine(line, field) {
        const items = parseMappingLine(line);
        let mapped = 0;
        items.forEach(item => {
            parseSuffix(item.suffix).forEach(idx => {
                if (inRange(idx)) { records[idx - 1][field] = item.val; mapped++; }
            });
        });
        // No suffix pointed at a real certificate -- either the line never had one (a constant
        // value meant for every certificate) or its suffix(es) were all out of range (most
        // likely a trailing number that isn't a mapping index, e.g. "Industry 4.0"). Either way,
        // apply the item's TEXT with its suffix digits already stripped, never the raw line --
        // otherwise an out-of-range case like "...India3" with only 2 certificates would print
        // the literal "3" onto every certificate.
        if (mapped === 0) {
            const fallbackText = items.length > 0 ? items.map(it => it.val).join(', ') : line;
            records.forEach(record => { record[field] = fallbackText; });
        }
    }

    classification.Designation.forEach(line => applyMappedLine(line, 'Designation'));
    classification.PaperTitle.forEach(line => applyMappedLine(line, 'PaperTitle'));

    // A gap in the suffix sequence (e.g. "Alice1, Bob3" with no "2") leaves a record with no
    // NAME at all. Rather than silently exporting a blank certificate, drop it and warn instead.
    const warnings = [];
    const skippedIndexes = [];
    const filledRecords = [];
    records.forEach((record, i) => {
        if (record.NAME) {
            filledRecords.push(record);
        } else {
            skippedIndexes.push(i + 1);
        }
    });
    if (skippedIndexes.length > 0) {
        const label = skippedIndexes.length === 1 ? 'certificate' : 'certificates';
        warnings.push(`Skipped ${label} #${skippedIndexes.join(', #')} — no name is mapped to that index. Check the numeric suffixes in your name line(s).`);
    }
    if (maxIdx >= MAX_CERTIFICATES * 0.75) {
        warnings.push(`${maxIdx} certificates were requested — double-check the suffix ranges if that's more than you intended.`);
    }

    // Apply the publication details (DOI-derived or hand-entered) as constants to every certificate.
    const pub = getPublicationInfo();
    if (pub.error) warnings.push(pub.error);
    filledRecords.forEach(record => {
        record.DOI = pub.doi;
        record.Volume = pub.vol;
        record.Issue = pub.issue;
        record.Year = pub.year;
        record.Month = pub.month;
    });

    return { records: filledRecords, authorListLine, warnings };
}

function parseSuffix(suffix) {
    const indexes = [];
    // Supports comma-separated lists of numbers and/or ranges, e.g. "1,3-5,7"
    suffix.split(',').forEach(part => {
        part = part.trim();
        if (!part) return;
        if (part.includes('-')) {
            const parts = part.split('-');
            let start = parseInt(parts[0], 10);
            let end = parseInt(parts[1], 10);
            if (!isNaN(start) && !isNaN(end)) {
                if (start > end) { [start, end] = [end, start]; } // tolerate an inverted range like "5-2"
                for (let i = Math.max(start, 1); i <= end; i++) {
                    indexes.push(i);
                }
            }
        } else {
            const idx = parseInt(part, 10);
            if (!isNaN(idx) && idx >= 1) {
                indexes.push(idx);
            }
        }
    });
    return indexes;
}

// Extract the journal "file number" from a DOI, e.g. "10.17148/IJARCCE.2026.15817" -> "17"
// (the last 2 digits of the final numeric segment: VOLUME(15) + ISSUE(8) + FILE(17)). Purely
// cosmetic (used only in filenames), so it stays permissive even when parseDoiParts rejects the
// DOI for the certificate fields themselves.
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
//
// An empty DOI is a normal "not entered yet" state (no error). A non-empty DOI that doesn't
// decode cleanly returns an explicit `error` message instead of silently blank fields, so a
// malformed DOI doesn't ship a certificate with an empty volume/issue/month.
function parseDoiParts(doi) {
    const empty = { vol: '', issue: '', year: '', month: '', error: '' };
    if (!doi) return empty;

    const runs = doi.match(/\d+/g);
    if (!runs || runs.length < 2) {
        return { ...empty, error: `Couldn't find volume/issue/year numbers in the DOI "${doi}".` };
    }

    const last = runs[runs.length - 1];
    const year = runs[runs.length - 2];
    if (year.length !== 4) {
        return { ...empty, error: `Couldn't find a 4-digit year in the DOI "${doi}" (found "${year}").` };
    }
    if (last.length < 5) {
        return { vol: '', issue: '', year, month: '', error: `The DOI's final number ("${last}") is too short to contain a volume, issue and file number.` };
    }

    const vol = String(parseInt(last.slice(0, 2), 10));
    const issueNum = parseInt(last.slice(2, -2), 10);
    if (isNaN(issueNum) || issueNum < 1 || issueNum > 12) {
        return { vol, issue: '', year, month: '', error: `The DOI's issue digits don't resolve to a valid month (1-12).` };
    }

    return { vol, issue: String(issueNum), year, month: MONTH_NAMES[issueNum - 1], error: '' };
}

// The single source of truth for volume/issue/year/month/file number, resolved from whichever
// input mode is active. In "with DOI" mode everything is decoded from the DOI; in "without DOI"
// mode the user supplies volume, issue, file number and year by hand. The issue number sets the
// month in both modes, so that rule never has to be restated by the user. Either mode can return
// a non-empty `error` describing why a field came out blank/unreliable -- an empty field the
// user simply hasn't filled in yet is not an error.
function getPublicationInfo() {
    if (state.doiMode === 'with') {
        const doi = el.doiInput.value.trim();
        const parts = parseDoiParts(doi);
        return { doi, ...parts, fileNumber: extractPaperNumber(doi) };
    }

    const volRaw = el.volInput.value.trim();
    const issueRaw = el.issueInput.value.trim();
    const yearRaw = el.yearInput.value.trim();
    const fileRaw = el.fileNumInput.value.trim();
    const issueNum = parseInt(issueRaw, 10);

    let error = '';
    if (volRaw && !/^\d+$/.test(volRaw)) {
        error = `Volume "${volRaw}" isn't a number.`;
    } else if (issueRaw && (isNaN(issueNum) || issueNum < 1 || issueNum > 12)) {
        error = `Issue "${issueRaw}" must be a number from 1 to 12 (it sets the month).`;
    } else if (yearRaw && !/^\d{4}$/.test(yearRaw)) {
        error = `Year "${yearRaw}" must be a 4-digit number.`;
    }

    return {
        doi: '',
        vol: volRaw,
        issue: isNaN(issueNum) ? '' : String(issueNum),
        year: yearRaw,
        month: MONTH_NAMES[issueNum - 1] || '',
        fileNumber: fileRaw,
        error
    };
}

// Only strip characters that are actually invalid in filenames; keep spaces, commas, superscripts.
// Control characters are stripped too, and the result is capped well under the ~255-char
// filesystem limit so a long, multi-author list doesn't push "{base} {index}.docx" over it.
function sanitizeFilename(name) {
    const cleaned = name
        .replace(/[\\/:*?"<>|]/g, '_')
        .replace(/[\x00-\x1f\x7f]/g, '')
        .replace(/\s+/g, ' ')
        .trim()
        .replace(/[. ]+$/, '');
    return cleaned.length > 150 ? cleaned.slice(0, 150).trim() : cleaned;
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
    const first = state.records[0] || {};
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

// Render the parsed-data preview table shown between the input box and the publication-details
// card, so misclassification (a title landing in the wrong field, a blank certificate, an
// unresolved DOI) is visible before a download is generated rather than only after opening the
// .docx. Hidden entirely when there's nothing parsed yet.
function renderPreview(records, warnings) {
    el.parseWarnings.innerHTML = '';
    el.previewTableBody.innerHTML = '';

    if (!records || records.length === 0) {
        el.previewCard.hidden = true;
        return;
    }
    el.previewCard.hidden = false;

    (warnings || []).forEach(w => {
        const div = document.createElement('div');
        div.className = 'parse-warning';
        const icon = document.createElement('i');
        icon.className = 'fa-solid fa-triangle-exclamation';
        div.appendChild(icon);
        div.appendChild(document.createTextNode(' ' + w));
        el.parseWarnings.appendChild(div);
    });

    records.forEach((r, i) => {
        const tr = document.createElement('tr');
        [String(i + 1), r.NAME || '', r.Designation || '', r.PaperTitle || ''].forEach((val, ci) => {
            const td = document.createElement('td');
            td.textContent = val || (ci > 0 ? '(empty)' : '');
            if (ci > 0 && !val) td.classList.add('cell-empty');
            tr.appendChild(td);
        });
        el.previewTableBody.appendChild(tr);
    });

    const pub = records[0] || {};
    const parts = [];
    if (pub.DOI) parts.push(`DOI ${pub.DOI}`);
    parts.push(`Vol ${pub.Volume || '—'}`);
    parts.push(`Issue ${pub.Issue || '—'}`);
    parts.push(`${pub.Month || '—'}${pub.Year ? ' ' + pub.Year : ''}`);
    el.previewPubInfo.textContent = parts.join('  ·  ');
}

// Live Preview -- an actual render of the generated Word certificate (the first recipient's
// page), redrawn as section 2 is edited so the real output layout is visible before any
// download. The same docxtemplater pipeline that builds the downloads produces the .docx here;
// docx-preview then renders that .docx to HTML in the panel.
//
// It mirrors the parser's states: a prompt when there's no template or no data yet, the parse
// error when the text can't be read, otherwise the rendered certificate. If docx-preview or its
// JSZip dependency failed to load from the CDN, it degrades to renderLivePreviewMock() -- a
// plain-text summary of the parsed fields -- rather than showing nothing.
//
// The heavy .docx render is debounced and guarded by a generation counter: fast typing only
// pays for the final render, and a slow render that finishes after a newer edit is discarded
// instead of overwriting the current preview.
let livePreviewGen = 0;
let livePreviewTimer = null;

function showLivePreviewMsg(box, text, cls) {
    const p = document.createElement('p');
    p.className = cls;
    p.textContent = text;
    box.replaceChildren(p);
}

function docxPreviewAvailable() {
    return typeof window.JSZip === 'function'
        && window.docx && typeof window.docx.renderAsync === 'function';
}

function renderLivePreview({ records = [], error = '' } = {}) {
    const box = el.livePreview;
    if (!box) return;

    // Any newer call invalidates an in-flight or pending render.
    const gen = ++livePreviewGen;
    clearTimeout(livePreviewTimer);

    if (error) {
        showLivePreviewMsg(box, error, 'cert-preview-error');
        return;
    }
    if (!records || records.length === 0) {
        showLivePreviewMsg(box, 'Paste data in section 2 to preview the certificate.', 'cert-preview-empty');
        return;
    }
    if (!state.docxLoaded) {
        showLivePreviewMsg(box, 'Select a template in section 1 to preview the certificate.', 'cert-preview-empty');
        return;
    }
    if (!docxPreviewAvailable()) {
        renderLivePreviewMock(box, records);
        return;
    }

    // Keep the current preview on screen until the new render is ready, then swap it in.
    livePreviewTimer = setTimeout(() => runDocxPreview(box, records[0], records.length, gen), 300);
}

async function runDocxPreview(box, record, total, gen) {
    try {
        const blob = renderDocxBlobForRecord(record);
        if (gen !== livePreviewGen) return;

        const mount = document.createElement('div');
        mount.className = 'docx-render';
        // Pass the same node as body and style container so the injected <style> is removed
        // with the render on the next swap.
        await window.docx.renderAsync(blob, mount, mount, {
            className: 'docxpv',
            inWrapper: true,
            ignoreLastRenderedPageBreak: true,
        });
        if (gen !== livePreviewGen) return;

        // Scale the full-size Word page down to fit the panel width. A wrapping frame is
        // clipped to the scaled page box so there's no dead space beside a landscape page
        // and no horizontal scrollbar. This is a single measure-then-transform pass, not a
        // loop: `transform: scale()` is compositor-only so it doesn't relayout the
        // image-heavy render subtree the way a dynamic CSS `zoom` did (which used to freeze
        // the tab).
        const frame = document.createElement('div');
        frame.className = 'docx-scale-frame';
        frame.appendChild(mount);
        box.replaceChildren(frame);
        fitDocxPreview(box, frame, mount);

        if (total > 1) {
            const note = document.createElement('p');
            note.className = 'cert-preview-note';
            note.textContent = `Showing certificate 1 of ${total}.`;
            box.appendChild(note);
        }
    } catch (err) {
        if (gen !== livePreviewGen) return;
        showLivePreviewMsg(box, `Couldn't render the certificate preview: ${err.message}`, 'cert-preview-error');
        console.error(err);
    }
}

// Shrink the rendered Word page so it fits the preview panel, whatever the template's page
// size (portrait ~816px, landscape ~1123px at 96dpi). docx-preview does not lay these
// templates out cleanly -- it pushes the certificate block hundreds of px off the page
// origin, drops full-bleed background art at wild offsets, and for one template leaves an
// ~870px vertical gap between the header and the body. So rather than scaling the nominal
// page box (which would show a slab of blank page beside a clipped, off-centre certificate)
// this measures where the text actually landed, finds the dense cluster the certificate
// body forms, and fits a top-left `transform` to that. A single measure-then-transform
// pass, no loop: `transform: scale()` is compositor-only and doesn't relayout the render.
function fitDocxPreview(box, frame, mount) {
    const page = mount.querySelector('section.docxpv');
    if (!page) return;
    const pageW = page.offsetWidth;
    const pageH = page.offsetHeight;
    if (!pageW || !pageH) return;

    const pr = page.getBoundingClientRect();
    const sc0 = pr.width / pageW || 1;
    const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

    // Collect the on-page text rects (page-local, unscaled). Images are skipped entirely --
    // they're decorative here and docx-preview's placement of them is the least reliable.
    const rects = [];
    page.querySelectorAll('*').forEach(node => {
        if (![...node.childNodes].some(n => n.nodeType === 3 && n.textContent.trim())) return;
        const b = node.getBoundingClientRect();
        if (b.width === 0 || b.height === 0) return;
        rects.push({
            l: (b.left - pr.left) / sc0, r: (b.right - pr.left) / sc0,
            t: (b.top - pr.top) / sc0, b: (b.bottom - pr.top) / sc0,
        });
    });

    let originX = 0, originY = 0, contentW = pageW, contentH = pageH;
    if (rects.length) {
        // Vertical outlier trim: sort by top edge and find the tightest window holding ~75%
        // of the rects. That window is the certificate body; a stray header sitting far
        // above it (or spacing junk far below) falls outside and is dropped from the crop.
        const need = Math.max(1, Math.ceil(rects.length * 0.85));
        const byTop = [...rects].sort((p, q) => p.t - q.t);
        let bestI = 0, bestSpan = Infinity;
        for (let i = 0; i + need <= byTop.length; i++) {
            const span = byTop[i + need - 1].t - byTop[i].t;
            if (span < bestSpan) { bestSpan = span; bestI = i; }
        }
        const core = byTop.slice(bestI, bestI + need);
        const cl = Math.min(...core.map(x => x.l));
        const cr = Math.max(...core.map(x => x.r));
        const ct = Math.min(...core.map(x => x.t));
        const cb = Math.max(...core.map(x => x.b));

        // Pad a little, then clamp to a believable region so a rect that slipped the trim
        // can't stretch the crop past ~1.6x the page in either axis.
        const padX = pageW * 0.04, padY = pageH * 0.06;
        const left = clamp(cl - padX, -0.1 * pageW, 0.55 * pageW);
        const right = clamp(cr + padX, 0.55 * pageW, 1.6 * pageW);
        const top = clamp(ct - padY, -0.15 * pageH, 0.85 * pageH);
        const bottom = clamp(cb + padY, top + pageH * 0.2, top + pageH * 1.6);
        if (right - left > pageW * 0.3 && bottom - top > pageH * 0.2) {
            originX = Math.max(0, left);
            originY = Math.max(0, top);
            contentW = right - originX;
            contentH = bottom - originY;
        }
    }

    const cs = getComputedStyle(box);
    const avail = box.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
    const scale = Math.min(1, avail / contentW);

    mount.style.transformOrigin = 'top left';
    mount.style.transform = `translate(${-originX * scale}px, ${-originY * scale}px) scale(${scale})`;
    mount.style.width = `${pageW}px`;
    frame.style.width = `${Math.round(contentW * scale)}px`;
    frame.style.height = `${Math.round(contentH * scale)}px`;
}

// Fallback used only when docx-preview/JSZip didn't load: a plain-language summary of section
// 2's parsed content (paper title, author names upper-cased as the .docx stamps them, and each
// affiliation line) so the panel still says something useful.
function renderLivePreviewMock(box, records) {
    box.replaceChildren();

    const addLabel = text => {
        const d = document.createElement('div');
        d.className = 'cert-preview-label';
        d.textContent = text;
        box.appendChild(d);
    };

    const title = dedupeJoin(records.map(r => r.PaperTitle));
    if (title) {
        addLabel('Paper Title');
        const t = document.createElement('div');
        t.className = 'cert-preview-title';
        t.textContent = title;
        box.appendChild(t);
    }

    addLabel(records.length === 1 ? 'Author' : `Authors (${records.length})`);
    const authors = document.createElement('div');
    authors.className = 'cert-preview-authors';
    records.forEach(r => {
        const wrap = document.createElement('div');
        wrap.className = 'cert-preview-author';

        const name = document.createElement('span');
        name.className = 'cert-preview-name';
        name.textContent = (r.NAME || '').toUpperCase() || '(no name)';
        wrap.appendChild(name);

        if (r.Designation) {
            const affil = document.createElement('span');
            affil.className = 'cert-preview-affil';
            affil.textContent = r.Designation;
            wrap.appendChild(affil);
        }
        authors.appendChild(wrap);
    });
    box.appendChild(authors);
}

// Trigger a browser download of a blob. The object URL and the anchor are cleaned up on a
// timer rather than synchronously: revoking the URL (or removing the anchor) in the same tick
// as the click can abort the download in some browsers, which matters most for the larger
// combined files and for the rapid back-to-back downloads in "Multiple Downloads".
function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    setTimeout(() => {
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }, 1000);
}

// Reduce a placeholder tag to a canonical form so lookups tolerate case and spacing differences:
// "Paper Title", "paper_title" and "PAPERTITLE" all collapse to "papertitle".
function normalizeTag(tag) {
    return String(tag).toLowerCase().replace(/[\s_]+/g, '');
}

// Curly braces are docxtemplater's tag delimiters, so a literal "{" or "}" in user text
// (e.g. a paper title like "Study on {AI} Systems") would be parsed as a broken tag and the
// text between them silently dropped. Swap them for the full-width look-alikes so every other
// special character (&, <, >, %, quotes, dashes, ...) passes straight through untouched --
// docxtemplater XML-escapes those itself.
function neutralizeBraces(value) {
    return String(value == null ? '' : value).replace(/\{/g, '｛').replace(/\}/g, '｝');
}

// 3. Generate a filled DOCX package (PizZip instance) for one record
function renderDocxZipForRecord(record) {
    const docZip = new window.PizZip(state.docxBuffer);

    const name = neutralizeBraces(record.NAME ? record.NAME.toUpperCase() : '');
    const designation = neutralizeBraces(record.Designation || '');
    const paperTitle = neutralizeBraces(record.PaperTitle || '');
    const doi = neutralizeBraces(record.DOI || '');
    const vol = neutralizeBraces(record.Volume || '');
    const issue = neutralizeBraces(record.Issue || '');
    const year = neutralizeBraces(record.Year || '');
    const month = neutralizeBraces(record.Month || '');

    // Canonical values, keyed by normalized tag name. This backs the nullGetter below, which
    // catches placeholders whose spelling doesn't exactly match a key in setData -- notably the
    // RJR template, which spells its name tag "{NAMe}". Without this, docxtemplater's default
    // nullGetter would stamp the literal text "undefined" onto the certificate. Unknown tags
    // resolve to an empty string rather than failing the render.
    const canonical = {
        name: name,
        designation: designation,
        papertitle: paperTitle,
        doi: doi,
        vol: vol,
        volume: vol,
        issue: issue,
        year: year,
        month: month
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
        NAME: name,
        name: name,
        Designation: designation,
        designation: designation,
        "Paper Title": paperTitle,
        "paper title": paperTitle,
        PaperTitle: paperTitle,
        papertitle: paperTitle,
        DOI: doi,
        doi: doi,
        vol: vol,
        Volume: vol,
        issue: issue,
        Issue: issue,
        year: year,
        Year: year,
        month: month,
        Month: month
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

// Ensure every OOXML element that must carry a document-unique identifier -- drawing objects
// (wp:docPr), bookmarks, and paragraph/run revision ids (w14:paraId/textId) -- actually gets a
// unique one once multiple pre-rendered certificate bodies are concatenated into a single
// document. Each source template only has one page's worth of these, so naively concatenating
// several copies duplicates every id; OOXML requires them to be unique, and Word responds to a
// duplicate with an "unreadable content, repair?" prompt. `counters` is shared across every
// call (one call per certificate body) so ids never collide between certificates; `bookmarkMap`
// is per-call because a bookmarkStart/bookmarkEnd pair shares one original id that must still
// match after remapping.
function makeBodyIdsUnique(xml, counters) {
    const bookmarkMap = new Map();

    xml = xml.replace(/(<wp:docPr\s+id=")(\d+)(")/g, (m, a, id, b) => a + (counters.docPr++) + b);

    xml = xml.replace(/(<w:bookmark(?:Start|End)\s+w:id=")(\d+)(")/g, (m, a, id, b) => {
        if (!bookmarkMap.has(id)) bookmarkMap.set(id, counters.bookmark++);
        return a + bookmarkMap.get(id) + b;
    });
    // Bookmark NAMEs are also duplicated across copies (every page's is "page1"); make them
    // unique too so a "go to bookmark" doesn't always land on the first certificate.
    xml = xml.replace(/(<w:bookmarkStart\s+w:id="\d+"\s+w:name=")([^"]*)(")/g, (m, a, name, b) => `${a}${name}_${counters.bookmarkName++}${b}`);

    xml = xml.replace(/(w14:(?:para|text)Id=")[0-9A-Fa-f]{1,8}(")/g, (m, a, b) =>
        a + (counters.hex++).toString(16).toUpperCase().padStart(8, '0') + b);

    // Legacy VML drawing shapes (textboxes) carry their own id namespace.
    xml = xml.replace(/(o:spid="_x0000_s)\d+(")/g, (m, a, b) => a + (counters.vml++) + b);

    return xml;
}

// Combine every record's rendered DOCX into a single multi-page DOCX file, one certificate
// per page. Records are separated by a real section break (an empty paragraph carrying a
// copy of the template's own <w:sectPr>) rather than a manual page break: since the template
// uses <w:titlePg/> (a distinct "first page" header/layout), a manual page break would push
// every certificate after the first onto the section's "default" (non-first) page, which
// renders differently (visible as stray whitespace/misalignment). Giving each certificate its
// own one-page section means every certificate consistently gets the section's "first page"
// treatment, matching how page 1 renders. The separator paragraph is given explicit zero
// spacing/line-height/font-size: every template's page margins are 0, so even the default
// paragraph spacing on this synthetic paragraph would be enough to push a spurious blank page
// in between certificates.
function buildCombinedDocxBlob(records) {
    const baseZip = renderDocxZipForRecord(records[0]);
    const xml0 = baseZip.file('word/document.xml').asText();

    const bodyOpenIdx = xml0.indexOf('<w:body>') + '<w:body>'.length;
    const sectPrStart = xml0.lastIndexOf('<w:sectPr');
    const sectPrEnd = xml0.indexOf('</w:sectPr>', sectPrStart) + '</w:sectPr>'.length;
    const sectPrXml = xml0.slice(sectPrStart, sectPrEnd);
    const prefix = xml0.slice(0, bodyOpenIdx);
    const suffix = xml0.slice(sectPrEnd); // "</w:body></w:document>"

    const idCounters = { docPr: 1, bookmark: 0, bookmarkName: 1, hex: 1, vml: 1 };
    const bodies = [makeBodyIdsUnique(extractRecordBodyContent(xml0), idCounters)];
    for (let i = 1; i < records.length; i++) {
        const recordZip = renderDocxZipForRecord(records[i]);
        const recordXml = recordZip.file('word/document.xml').asText();
        bodies.push(makeBodyIdsUnique(extractRecordBodyContent(recordXml), idCounters));
    }

    const sectionBreakParagraph =
        `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/>` +
        `<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>${sectPrXml}</w:pPr></w:p>`;
    const combinedXml = prefix + bodies.join(sectionBreakParagraph) + sectPrXml + suffix;

    baseZip.file('word/document.xml', combinedXml);
    return baseZip.generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
}

// Log a docxtemplater render failure with the actual offending tag(s) rather than the bare
// "Multi error" message it throws by default.
function logRenderError(err) {
    const errors = err && err.properties && err.properties.errors;
    if (Array.isArray(errors) && errors.length > 0) {
        errors.forEach(e => {
            const props = e.properties || {};
            const tag = props.xtag || props.id || props.tag;
            const explanation = props.explanation || e.message || 'unknown error';
            log(`Template error${tag ? ` in {${tag}}` : ''}: ${explanation}`, 'error');
        });
    } else {
        log(`Failed generation: ${err.message}`, 'error');
    }
    console.error(err);
}

// Shared "disable every download button, show a spinner on the clicked one, restore afterward"
// wrapper so a user can't fire two exports at once, and errors from any mode are logged the
// same way.
async function runDownload(btnEl, task) {
    if (state.librariesMissing) {
        log('DOCX libraries failed to load — reload the page before generating certificates.', 'error');
        return;
    }
    if (!state.docxLoaded) {
        log('Select a journal template first.', 'error');
        return;
    }
    if (state.records.length === 0) {
        log('Enter recipient data first — at least one name with a numeric suffix (e.g. "Ravi Kumar1").', 'error');
        return;
    }

    const initialHTML = btnEl.innerHTML;
    toggleButtons(false);
    btnEl.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Generating...';

    try {
        await task();
    } catch (err) {
        logRenderError(err);
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

// DOM trigger. Guarded so init() runs exactly once regardless of whether the DOM was still
// parsing when this script executed: without the flag, a script that runs after DOMContentLoaded
// would fall through to the immediate call, while one that runs before it would fire on the
// event -- and any setup that registered both paths would bind every listener twice, doubling
// clicks (and downloads).
let initialized = false;
function bootstrap() {
    if (initialized) return;
    initialized = true;
    init();
}
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bootstrap);
} else {
    bootstrap();
}
