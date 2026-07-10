// Zero-Styling Certificate Generator Logic (PDF & DOCX Support)

const state = {
    fileType: null,          // 'pdf' or 'docx'
    templateBuffer: null,    // Raw ArrayBuffer of template
    pdfDoc: null,            // Loaded PDFJS Document
    viewport: null,          // Page Viewport
    page: null,              // PDFJS Page 1 Object
    placeholders: {},        // { NAME: {...}, Designation: {...}, PaperTitle: {...}, DOI: {...} }
    records: [],             // Parsed recipient row data
    canvasCache: null,       // Offscreen cached canvas with rendered PDF page
    currentPreviewIndex: 0
};

const el = {
    uploadZone: document.getElementById('upload-zone'),
    templateUpload: document.getElementById('template-upload'),
    fileName: document.getElementById('file-name'),
    dataInput: document.getElementById('data-input'),
    parserStatus: document.getElementById('parser-status'),
    btnLoadDemo: document.getElementById('btn-load-demo'),
    btnGenerateZip: document.getElementById('btn-generate-zip'),
    btnGeneratePdf: document.getElementById('btn-generate-pdf'),
    logsContainer: document.getElementById('logs-container'),
    previewCard: document.getElementById('preview-card'),
    previewCanvas: document.getElementById('preview-canvas'),
    previewIndex: document.getElementById('preview-index'),
    btnPrevPage: document.getElementById('btn-prev-page'),
    btnNextPage: document.getElementById('btn-next-page'),
    docxPreviewMsg: document.getElementById('docx-preview-msg')
};

// Logger utility
function log(msg, type = 'system') {
    const entry = document.createElement('div');
    entry.className = `log-entry ${type}`;
    
    let icon = 'fa-info-circle';
    if (type === 'success') icon = 'fa-circle-check';
    if (type === 'error') icon = 'fa-circle-exclamation';
    if (type === 'info') icon = 'fa-magnifying-glass';
    
    entry.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${msg}</span>`;
    el.logsContainer.appendChild(entry);
    el.logsContainer.scrollTop = el.logsContainer.scrollHeight;
}

// Initialize listeners
function init() {
    setupEventListeners();
}

function setupEventListeners() {
    el.uploadZone.addEventListener('click', () => el.templateUpload.click());
    el.templateUpload.addEventListener('change', handleFileSelect);
    
    el.uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        el.uploadZone.classList.add('dragover');
    });
    el.uploadZone.addEventListener('dragleave', () => {
        el.uploadZone.classList.remove('dragover');
    });
    el.uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        el.uploadZone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            el.templateUpload.files = e.dataTransfer.files;
            handleFileSelect();
        }
    });

    el.btnLoadDemo.addEventListener('click', loadDemoData);
    el.dataInput.addEventListener('input', handleDataInput);
    el.btnGenerateZip.addEventListener('click', handleZipExport);
    el.btnGeneratePdf.addEventListener('click', generateCombinedPDF);
    
    el.btnPrevPage.addEventListener('click', () => navigatePreview(-1));
    el.btnNextPage.addEventListener('click', () => navigatePreview(1));
}

// Load Example Data
function loadDemoData() {
    const demo = [
        'A STUDY ON RECRUITMENT METRICS AND THEIR IMPACT ON ORGANIZATIONAL EFFICIENCY',
        'Mr.AJAYRATHNA S1, Ms.GAYATHRI M2',
        'Assistant Professor, Department of Management Studies, EGS Pillay Engineering College, Nagapattinam, Tamilnadu, India1',
        'MBA Student, Department of Management Studies, EGS Pillay Engineering College, Nagapattinam, Tamilnadu, India2',
        '13706'
    ].join('\n');
    el.dataInput.value = demo;
    handleDataInput();
}

// 1. Load Template File (PDF or DOCX)
function handleFileSelect() {
    const file = el.templateUpload.files[0];
    if (!file) return;

    el.fileName.textContent = file.name;
    const extension = file.name.split('.').pop().toLowerCase();
    state.fileType = extension;
    
    log(`Loading ${extension.toUpperCase()} file: ${file.name}...`, 'system');

    const reader = new FileReader();
    reader.onload = async function(e) {
        state.templateBuffer = e.target.result;
        
        if (extension === 'docx') {
            log(`Word Document (DOCX) loaded successfully.`, 'success');
            setupDocxWorkflow();
        } else if (extension === 'pdf') {
            try {
                const loadingTask = pdfjsLib.getDocument({ data: state.templateBuffer });
                state.pdfDoc = await loadingTask.promise;
                log(`PDF loaded. Total Pages: ${state.pdfDoc.numPages}`, 'success');
                await analyzePDFTemplatePage(1);
            } catch (err) {
                log(`Error reading PDF: ${err.message}`, 'error');
                console.error(err);
            }
        } else {
            log(`Unsupported file type. Please upload a PDF or DOCX template.`, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// DOCX Flow Setup
function setupDocxWorkflow() {
    el.docxPreviewMsg.style.display = 'block';
    el.previewCanvas.style.display = 'none';
    el.previewCard.style.display = 'block';
    
    // DOCX outputs are ZIP archives of individual documents
    el.btnGeneratePdf.style.display = 'none'; // DOCX doesn't support combined PDF directly
    el.btnGenerateZip.innerHTML = '<i class="fa-solid fa-file-zipper"></i> Download ZIP (Individual DOCX Files)';
    
    log(`DOCX Mode Ready. Format mappings to replace {NAME}, {Designation}, {PaperTitle}, and {DOI} tags inside your Word document.`, 'info');
    handleDataInput();
}

// 2. Scan PDF Text Layer for Placeholders
async function analyzePDFTemplatePage(pageNum) {
    try {
        el.docxPreviewMsg.style.display = 'none';
        el.previewCanvas.style.display = 'block';
        el.btnGeneratePdf.style.display = 'inline-flex';
        el.btnGenerateZip.innerHTML = '<i class="fa-solid fa-file-zipper"></i> Download ZIP (Individual PDFs)';
        
        log(`Analyzing text layout on Page ${pageNum}...`, 'info');
        state.page = await state.pdfDoc.getPage(pageNum);
        state.viewport = state.page.getViewport({ scale: 2.5 });
        
        // Cache rendered base page
        state.canvasCache = document.createElement('canvas');
        state.canvasCache.width = state.viewport.width;
        state.canvasCache.height = state.viewport.height;
        const ctx = state.canvasCache.getContext('2d');
        
        log(`Rendering base template canvas...`, 'system');
        await state.page.render({ canvasContext: ctx, viewport: state.viewport }).promise;
        
        // Scan text items
        const textContent = await state.page.getTextContent();
        
        // Map transform coordinates to viewport coordinates
        const transformedItems = textContent.items.map(item => {
            const tx = pdfjsLib.Util.transform(state.viewport.transform, item.transform);
            const fontSize = Math.sqrt(tx[2]*tx[2] + tx[3]*tx[3]);
            return {
                str: item.str,
                fontName: item.fontName,
                x: tx[4],
                y: tx[5], // baseline
                width: item.width * (tx[0] / item.transform[0] || state.viewport.scale),
                height: fontSize,
                item: item
            };
        });

        // Filter items with braces
        const braceItems = transformedItems.filter(item => {
            const s = item.str;
            return s.includes('{') || s.includes('}') || s.includes('[') || s.includes(']');
        });

        // Explicitly sort brace items left-to-right (horizontal reading order)
        braceItems.sort((a, b) => {
            const yDiff = Math.abs(a.y - b.y);
            if (yDiff < Math.min(a.height, b.height) * 0.8) {
                return a.x - b.x; // same line, sort left-to-right
            }
            return a.y - b.y; // sort top-to-bottom
        });

        // Merge adjacent placeholders on the same horizontal line (e.g. {Paper and Title})
        const merged = [];
        braceItems.forEach(item => {
            if (merged.length > 0) {
                const last = merged[merged.length - 1];
                const yDiff = Math.abs(item.y - last.y);
                const xDiff = item.x - (last.x + last.width);
                if (yDiff < last.height * 0.8 && xDiff >= -10 && xDiff < 75) {
                    last.str += ' ' + item.str;
                    last.width = (item.x + item.width) - last.x;
                    last.height = Math.max(last.height, item.height);
                    return;
                }
            }
            merged.push({ ...item });
        });

        // Classify standard target placeholders
        state.placeholders = {
            NAME: null,
            Designation: null,
            PaperTitle: null,
            DOI: null
        };

        merged.forEach(ph => {
            const text = ph.str.toLowerCase();
            if (text.includes('name')) {
                state.placeholders.NAME = ph;
            } else if (text.includes('designation')) {
                state.placeholders.Designation = ph;
            } else if (text.includes('paper') || text.includes('title')) {
                state.placeholders.PaperTitle = ph;
            } else if (text.includes('doi')) {
                state.placeholders.DOI = ph;
            }
        });

        // Log results
        const detected = [];
        Object.keys(state.placeholders).forEach(key => {
            if (state.placeholders[key]) {
                detected.push(`{${key}}`);
            }
        });

        if (detected.length > 0) {
            log(`Placeholders detected: ${detected.join(', ')}`, 'success');
            
            // Sample styles and background colors
            sampleStylesAndColors(ctx);
            
            // Parse inputs & enable generate
            handleDataInput();
        } else {
            log(`No placeholders found containing '{...}' format.`, 'error');
            toggleButtons(false);
        }
    } catch (err) {
        log(`Failed template analysis: ${err.message}`, 'error');
        console.error(err);
    }
}

// Helper to disable/enable export buttons
function toggleButtons(enabled) {
    el.btnGenerateZip.disabled = !enabled;
    if (state.fileType === 'pdf') {
        el.btnGeneratePdf.disabled = !enabled;
    } else {
        el.btnGeneratePdf.disabled = true;
    }
}

// 3. Extract Styles (Original Fonts, Colors, Background Sample)
function sampleStylesAndColors(ctx) {
    Object.keys(state.placeholders).forEach(key => {
        const ph = state.placeholders[key];
        if (!ph) return;

        // A. Sample local background just above placeholder text
        const localX = Math.round(ph.x + ph.width / 2);
        const localY = Math.max(5, Math.round(ph.y - ph.height - 5));
        const localPixel = ctx.getImageData(localX, localY, 1, 1).data;
        ph.localBG = `rgb(${localPixel[0]}, ${localPixel[1]}, ${localPixel[2]})`;

        // B. Sample original text color (max deviation from local background)
        const pad = 1;
        const scanX = Math.round(ph.x) + pad;
        const scanY = Math.round(ph.y - ph.height) + pad;
        const scanW = Math.round(ph.width) - (pad * 2);
        const scanH = Math.round(ph.height) - (pad * 2);
        
        let textColor = 'rgb(0,0,0)';
        if (scanW > 0 && scanH > 0) {
            const imgData = ctx.getImageData(scanX, scanY, scanW, scanH);
            const pixels = imgData.data;
            let maxDist = -1;
            let bestRGB = [0, 0, 0];

            for (let i = 0; i < pixels.length; i += 4) {
                const r = pixels[i];
                const g = pixels[i+1];
                const b = pixels[i+2];
                const a = pixels[i+3];
                if (a < 100) continue;

                const dist = Math.abs(r - localPixel[0]) + Math.abs(g - localPixel[1]) + Math.abs(b - localPixel[2]);
                if (dist > maxDist) {
                    maxDist = dist;
                    bestRGB = [r, g, b];
                }
            }
            textColor = `rgb(${bestRGB[0]}, ${bestRGB[1]}, ${bestRGB[2]})`;
        }

        // C. Contrast Fallback: If text color matches the background color too closely, use a dark fallback
        const localBGStr = ph.localBG;
        const dev = getRGBColorDev(textColor, localBGStr);
        if (dev < 120) {
            textColor = 'rgb(6, 26, 56)'; // default dark blue
        }

        ph.textColor = textColor;
        log(`Extracted style for {${key}}: Size ${Math.round(ph.height)}px, Font: ${ph.fontName}, Color: ${ph.textColor}`, 'system');
    });

    // Make sure contrast fallbacks resolve if NAME itself is correct
    if (state.placeholders.NAME && state.placeholders.NAME.textColor) {
        Object.keys(state.placeholders).forEach(key => {
            const ph = state.placeholders[key];
            if (!ph) return;
            const dev = getRGBColorDev(ph.textColor, ph.localBG);
            if (dev < 120) {
                ph.textColor = state.placeholders.NAME.textColor;
            }
        });
    }
}

// Contrast helper
function getRGBColorDev(color1Str, color2Str) {
    const parse = c => {
        const m = c.match(/\d+/g);
        return m ? m.map(Number) : [0,0,0];
    };
    const c1 = parse(color1Str);
    const c2 = parse(color2Str);
    return Math.abs(c1[0] - c2[0]) + Math.abs(c1[1] - c2[1]) + Math.abs(c1[2] - c2[2]);
}

// 4. Data Inputs Parser
function handleDataInput() {
    const text = el.dataInput.value.trim();
    if (!text) {
        state.records = [];
        el.parserStatus.textContent = 'No data input';
        toggleButtons(false);
        el.previewCard.style.display = 'none';
        return;
    }

    try {
        state.records = parseDataInput(text);
        el.parserStatus.textContent = `${state.records.length} records parsed successfully`;
        
        const hasTemplate = state.fileType === 'docx' || (state.pdfDoc && Object.values(state.placeholders).some(p => p !== null));
        if (hasTemplate) {
            toggleButtons(true);
            state.currentPreviewIndex = 0;
            el.previewCard.style.display = 'block';
            updatePreview();
        }
    } catch (err) {
        el.parserStatus.textContent = `Parse error: ${err.message}`;
        toggleButtons(false);
    }
}

function parseMappingLine(line) {
    const items = [];
    let currentStart = 0;
    
    for (let i = 0; i < line.length; i++) {
        if (line[i] === ',') {
            // Check if this comma is immediately preceded by a numeric suffix (e.g., "...1,")
            const beforeComma = line.substring(currentStart, i).trim();
            const match = beforeComma.match(/\d+(?:-\d+)?$/);
            if (match) {
                const val = beforeComma.substring(0, match.index).trim();
                const suffix = match[0];
                items.push({ val, suffix });
                currentStart = i + 1;
            }
        }
    }
    
    // Handle last item
    const lastItem = line.substring(currentStart).trim();
    if (lastItem) {
        const match = lastItem.match(/\d+(?:-\d+)?$/);
        if (match) {
            const val = lastItem.substring(0, match.index).trim();
            const suffix = match[0];
            items.push({ val, suffix });
        }
    }
    return items;
}

function parseDataInput(text) {
    const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    
    // Classify lines dynamically
    const classification = {
        NAME: [],
        Designation: [],
        PaperTitle: [],
        DOI: []
    };

    lines.forEach(line => {
        const trimmed = line.trim();
        if (!trimmed) return;

        // 1. DOI Classify
        if (/^\d+$/.test(trimmed) || trimmed.startsWith('10.') || (trimmed.length < 25 && trimmed.includes('.'))) {
            classification.DOI.push(trimmed);
            return;
        }

        // 2. Name Classify (e.g. contains Mr/Ms prefix, or multiple comma items with numeric suffixes)
        const isNameList = trimmed.includes('Mr.') || trimmed.includes('Ms.') || trimmed.includes('Mrs.') || trimmed.includes('Dr.');
        const parsed = parseMappingLine(trimmed);
        const allShort = parsed.length > 0 && parsed.every(item => item.val.length < 25);
        if (isNameList || (parsed.length > 1 && allShort)) {
            classification.NAME.push(trimmed);
            return;
        }

        // 3. Designation Classify (contains academic/corporate keywords)
        const keywords = ['student', 'professor', 'lecturer', 'department', 'college', 'university', 'researcher', 'studies', 'faculty', 'nagapattinam', 'nagapatnam', 'tamilnadu', 'tamil nadu', 'india', 'scholar'];
        const lower = trimmed.toLowerCase();
        const hasKeyword = keywords.some(kw => lower.includes(kw));
        if (hasKeyword) {
            classification.Designation.push(trimmed);
            return;
        }

        // 4. Default to Paper Title
        classification.PaperTitle.push(trimmed);
    });

    // We must have at least one NAME line to determine the count
    if (classification.NAME.length === 0) {
        throw new Error("Could not detect any Name line containing certificate mapping suffixes (e.g., Name1, Name2).");
    }

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

    // Map DOI (constant applied directly to all certificates)
    if (classification.DOI.length > 0) {
        const doiValue = classification.DOI[0];
        records.forEach(record => {
            record.DOI = doiValue;
        });
    }

    return records;
}

function parseSuffix(suffix) {
    const indexes = [];
    if (suffix.includes('-')) {
        const parts = suffix.split('-');
        const start = parseInt(parts[0], 10);
        const end = parseInt(parts[1], 10);
        if (!isNaN(start) && !isNaN(end)) {
            for (let i = start; i <= end; i++) {
                indexes.push(i);
            }
        }
    } else {
        const idx = parseInt(suffix, 10);
        if (!isNaN(idx)) {
            indexes.push(idx);
        }
    }
    return indexes;
}

// 5. Render PDF Certificate Canvas
function renderCertificateCanvas(recordIdx, targetCanvas) {
    return new Promise((resolve) => {
        const record = state.records[recordIdx];
        
        // Match canvas dimensions
        targetCanvas.width = state.canvasCache.width;
        targetCanvas.height = state.canvasCache.height;
        const ctx = targetCanvas.getContext('2d');
        
        // 1. Draw cached original template drawing
        ctx.drawImage(state.canvasCache, 0, 0);

        // Wait for fonts to be ready in the document
        document.fonts.ready.then(() => {
            // 2. Perform style-perfect replacements
            Object.keys(state.placeholders).forEach(key => {
                const ph = state.placeholders[key];
                if (!ph) return;

                // A. Cover-up background color rect (extending below baseline for descenders/shadows)
                ctx.fillStyle = ph.localBG;
                const padX = Math.max(16, ph.width * 0.18);
                const padY = Math.max(6, ph.height * 0.15);
                ctx.fillRect(ph.x - padX, ph.y - ph.height - padY, ph.width + (padX * 2), ph.height + (padY * 3.5));

                // B. Prepare replacement text value
                let val = record[key] || '';
                let drawText = val;
                
                // Specific DOI prefix merge formatting (to avoid duplicate prefixes)
                if (key === 'DOI' && val) {
                    const originalStr = ph.str; // e.g. "{DOI}"
                    const cleanVal = val.replace(/^(doi\s*)?10\.17148\/iarjset\.2026\./i, '');
                    
                    if (originalStr.includes('{') || originalStr.includes('}')) {
                        drawText = originalStr.replace(/\{doi\}/i, cleanVal);
                    } else {
                        drawText = cleanVal;
                    }
                }

                // C. Name formatting: MUST be uppercase & styled in a bold serif with a subtle stroke & shadow
                ctx.save();
                ctx.textAlign = 'center';
                ctx.textBaseline = 'alphabetic';

                const centerX = ph.x + ph.width / 2;
                const baselineY = ph.y;

                if (key === 'NAME') {
                    drawText = drawText.toUpperCase();
                    
                    // Style matching for IARJSET journal expected style: Bold Serif with shadows
                    ctx.font = `bold ${ph.height}px "Playfair Display", "Cinzel", "Georgia", serif`;
                    
                    // Draw a subtle text outline shadow
                    ctx.shadowColor = 'rgba(0, 0, 0, 0.25)';
                    ctx.shadowBlur = 4;
                    ctx.shadowOffsetX = 2;
                    ctx.shadowOffsetY = 2;

                    // Draw filled text
                    ctx.fillStyle = '#061b38'; // Dark blue
                    ctx.fillText(drawText, centerX, baselineY);

                    // Add thin outline strokes for extra pop
                    ctx.shadowColor = 'transparent'; // turn off shadow for stroke
                    ctx.strokeStyle = 'rgba(255, 255, 255, 0.4)';
                    ctx.lineWidth = 1;
                    ctx.strokeText(drawText, centerX, baselineY);
                } else if (key === 'PaperTitle') {
                    // Green bold title
                    ctx.font = `bold ${ph.height}px "Montserrat", "Inter", sans-serif`;
                    ctx.fillStyle = '#065f46'; // Beautiful dark green
                    ctx.fillText(drawText, centerX, baselineY);
                } else if (key === 'DOI') {
                    // Brown DOI text
                    ctx.font = `${ph.height}px "Inter", sans-serif`;
                    ctx.fillStyle = '#c05621'; // Brown
                    ctx.fillText(drawText, centerX, baselineY);
                } else {
                    // General placeholder style fallback
                    const cleanFont = ph.fontName.replace(/['"]/g, '');
                    ctx.font = `${ph.height}px "${cleanFont}", sans-serif`;
                    ctx.fillStyle = ph.textColor;
                    ctx.fillText(drawText, centerX, baselineY);
                }
                ctx.restore();
            });
            
            resolve(targetCanvas);
        });
    });
}

// 6. Navigation Preview
async function navigatePreview(direction) {
    if (state.records.length === 0) return;
    state.currentPreviewIndex = (state.currentPreviewIndex + direction + state.records.length) % state.records.length;
    await updatePreview();
}

async function updatePreview() {
    el.previewIndex.textContent = `Certificate ${state.currentPreviewIndex + 1} of ${state.records.length}`;
    if (state.fileType === 'docx') {
        // Show text-based preview for DOCX
        const record = state.records[state.currentPreviewIndex];
        const previewText = `NAME: ${record.NAME ? record.NAME.toUpperCase() : ''}\nDesignation: ${record.Designation || ''}\nPaperTitle: ${record.PaperTitle || ''}\nDOI: ${record.DOI || ''}`;
        el.docxPreviewMsg.innerHTML = `
            <i class="fa-solid fa-file-word" style="font-size: 3rem; color: #2b579a; margin-bottom: 1rem; display: block;"></i>
            <h3>DOCX Record Preview (${state.currentPreviewIndex + 1} of ${state.records.length})</h3>
            <pre style="text-align: left; background-color: var(--bg-input); padding: 1rem; border-radius: 4px; margin-top: 1rem; font-family: monospace; font-size: 0.8rem; line-height: 1.5; color: var(--text);">${previewText}</pre>
        `;
    } else {
        await renderCertificateCanvas(state.currentPreviewIndex, el.previewCanvas);
    }
}

// 7. ZIP Export (Handles PDF ZIP and DOCX ZIP)
async function handleZipExport() {
    if (state.records.length === 0 || !state.templateBuffer) return;

    if (state.fileType === 'docx') {
        await generateDocxZIP();
    } else {
        await generatePdfZIP();
    }
}

// Generate ZIP of individual DOCX files (No hardcoded styles, editing tags dynamically)
async function generateDocxZIP() {
    const initialText = el.btnGenerateZip.innerHTML;
    el.btnGenerateZip.disabled = true;
    el.btnGenerateZip.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Zipping DOCX files...';
    log(`Compiling Word DOCX documents...`, 'info');

    try {
        const zipArchive = new JSZip();

        for (let i = 0; i < state.records.length; i++) {
            const record = state.records[i];
            log(`Generating DOCX page ${i + 1} for: ${record.NAME || `Record ${i+1}`}...`, 'system');

            // Load Pizzip binary buffer
            const docZip = new window.PizZip(state.templateBuffer);
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

            // Replace XML placeholders
            doc.render();

            // Export to blob
            const outBlob = doc.getZip().generate({
                type: "blob",
                mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            });

            const fileName = (record.NAME || `document_${i+1}`).trim().replace(/[^a-z0-9_-]/gi, '_');
            zipArchive.file(`${fileName}.docx`, outBlob);
        }

        log(`Bundling ZIP package...`, 'system');
        const zipContent = await zipArchive.generateAsync({ type: 'blob' });
        
        const link = document.createElement('a');
        link.href = URL.createObjectURL(zipContent);
        link.download = 'certificates_docx_archive.zip';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        log(`ZIP archive of Word documents downloaded successfully!`, 'success');
    } catch (err) {
        log(`DOCX compilation failed: ${err.message}`, 'error');
        console.error(err);
    } finally {
        el.btnGenerateZip.disabled = false;
        el.btnGenerateZip.innerHTML = initialText;
    }
}

// Generate ZIP of individual PDF files
async function generatePdfZIP() {
    const initialText = el.btnGenerateZip.innerHTML;
    el.btnGenerateZip.disabled = true;
    el.btnGenerateZip.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Bundling ZIP...';
    log(`Compiling individual certificates into ZIP file...`, 'info');

    try {
        const zip = new JSZip();
        const { jsPDF } = window.jspdf;
        const tempCanvas = document.createElement('canvas');

        for (let i = 0; i < state.records.length; i++) {
            const record = state.records[i];
            log(`Rendering individual PDF ${i + 1} of ${state.records.length}...`, 'system');
            
            await renderCertificateCanvas(i, tempCanvas);
            
            const w = tempCanvas.width;
            const h = tempCanvas.height;
            const orientation = w >= h ? 'landscape' : 'portrait';
            const imgData = tempCanvas.toDataURL('image/jpeg', 0.95);

            const singlePdf = new jsPDF({
                orientation: orientation,
                unit: 'px',
                format: [w, h]
            });
            singlePdf.addImage(imgData, 'JPEG', 0, 0, w, h);

            const pdfBlob = singlePdf.output('blob');
            const fileName = (record.NAME || `certificate_${i+1}`).trim().replace(/[^a-z0-9_-]/gi, '_');
            zip.file(`${fileName}.pdf`, pdfBlob);
        }

        log(`Creating ZIP package...`, 'system');
        const zipContent = await zip.generateAsync({ type: 'blob' });
        
        const link = document.createElement('a');
        link.href = URL.createObjectURL(zipContent);
        link.download = 'certificates_archive.zip';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        log(`Successfully generated and downloaded ZIP archive!`, 'success');
    } catch (err) {
        log(`Failed compilation: ${err.message}`, 'error');
        console.error(err);
    } finally {
        el.btnGenerateZip.disabled = false;
        el.btnGenerateZip.innerHTML = initialText;
    }
}

// 8. Generate Combined PDF
async function generateCombinedPDF() {
    if (state.records.length === 0 || !state.pdfDoc) return;

    const initialText = el.btnGeneratePdf.innerHTML;
    el.btnGeneratePdf.disabled = true;
    el.btnGeneratePdf.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Compiling PDF...';
    log(`Compiling combined bulk PDF...`, 'info');

    try {
        const { jsPDF } = window.jspdf;
        let pdf = null;
        const tempCanvas = document.createElement('canvas');

        for (let i = 0; i < state.records.length; i++) {
            await renderCertificateCanvas(i, tempCanvas);
            
            const w = tempCanvas.width;
            const h = tempCanvas.height;
            const orientation = w >= h ? 'landscape' : 'portrait';
            const imgData = tempCanvas.toDataURL('image/jpeg', 0.95);

            if (i === 0) {
                pdf = new jsPDF({
                    orientation: orientation,
                    unit: 'px',
                    format: [w, h]
                });
            } else {
                pdf.addPage([w, h], orientation);
            }

            pdf.addImage(imgData, 'JPEG', 0, 0, w, h);
        }

        if (pdf) {
            pdf.save('certificates_combined.pdf');
            log(`Successfully generated and downloaded combined PDF!`, 'success');
        }
    } catch (err) {
        log(`Failed compilation: ${err.message}`, 'error');
        console.error(err);
    } finally {
        el.btnGeneratePdf.disabled = false;
        el.btnGeneratePdf.innerHTML = initialText;
    }
}

// DOM trigger
document.addEventListener('DOMContentLoaded', init);
if (document.readyState === 'interactive' || document.readyState === 'complete') {
    init();
}
