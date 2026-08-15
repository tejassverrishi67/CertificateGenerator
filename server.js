// Local server for the Certificate Generator.
//
// Serves the static site and exposes POST /api/convert, which converts a batch of generated
// DOCX files to PDF using Microsoft Word itself (via COM automation in convert-docx-to-pdf.ps1).
// Word's own export engine is used so the PDFs are lossless, pixel-faithful renderings of the
// DOCX -- something no in-browser converter can achieve.
//
// No npm dependencies: only Node's standard library.

const http = require('http');
const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { execFile } = require('child_process');

const PORT = process.env.PORT || 8080;
const ROOT = __dirname;
const PS_SCRIPT = path.join(ROOT, 'convert-docx-to-pdf.ps1');
const MAX_BODY_BYTES = 256 * 1024 * 1024; // generous: base64 DOCX batches inflate ~33%

const MIME = {
    '.html': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.svg': 'image/svg+xml',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.ico': 'image/x-icon',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};

function send(res, status, body, headers = {}) {
    res.writeHead(status, { 'Cache-Control': 'no-store', ...headers });
    res.end(body);
}

function sendJson(res, status, obj) {
    send(res, status, JSON.stringify(obj), { 'Content-Type': 'application/json; charset=utf-8' });
}

function readBody(req) {
    return new Promise((resolve, reject) => {
        const chunks = [];
        let total = 0;
        req.on('data', (chunk) => {
            total += chunk.length;
            if (total > MAX_BODY_BYTES) {
                reject(new Error('Request body too large'));
                req.destroy();
                return;
            }
            chunks.push(chunk);
        });
        req.on('end', () => resolve(Buffer.concat(chunks)));
        req.on('error', reject);
    });
}

function runPowerShell(args) {
    return new Promise((resolve, reject) => {
        execFile(
            'powershell.exe',
            ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', ...args],
            { windowsHide: true, maxBuffer: 32 * 1024 * 1024 },
            (err, stdout, stderr) => {
                if (err) {
                    reject(new Error(`${err.message}\n${stderr || ''}`.trim()));
                    return;
                }
                resolve({ stdout, stderr });
            }
        );
    });
}

// Sanitize a client-supplied filename down to a safe basename with no path components.
function safeBaseName(name, fallback) {
    const base = path.basename(String(name || '')).replace(/\.docx$/i, '');
    const cleaned = base.replace(/[\\/:*?"<>|]/g, '_').trim();
    return cleaned || fallback;
}

// POST /api/convert
// Body: { files: [{ name: "cert 1", data: "<base64 docx>" }, ...] }
// Returns: { files: [{ name: "cert 1", data: "<base64 pdf>" }, ...] }
async function handleConvert(req, res) {
    let payload;
    try {
        const raw = await readBody(req);
        payload = JSON.parse(raw.toString('utf8'));
    } catch (err) {
        sendJson(res, 400, { error: `Invalid request body: ${err.message}` });
        return;
    }

    const files = Array.isArray(payload && payload.files) ? payload.files : null;
    if (!files || files.length === 0) {
        sendJson(res, 400, { error: 'Expected a non-empty "files" array.' });
        return;
    }

    const workDir = path.join(os.tmpdir(), `certgen-${crypto.randomUUID()}`);
    const inDir = path.join(workDir, 'in');
    const outDir = path.join(workDir, 'out');

    try {
        await fsp.mkdir(inDir, { recursive: true });
        await fsp.mkdir(outDir, { recursive: true });

        // Write each DOCX to disk under an index-prefixed name so we can map results back
        // deterministically even if two records share a display name.
        const order = [];
        for (let i = 0; i < files.length; i++) {
            const entry = files[i] || {};
            if (typeof entry.data !== 'string') {
                sendJson(res, 400, { error: `File at index ${i} is missing base64 "data".` });
                return;
            }
            const displayName = safeBaseName(entry.name, `certificate_${i + 1}`);
            const stem = String(i).padStart(5, '0');
            order.push({ stem, displayName });
            await fsp.writeFile(path.join(inDir, `${stem}.docx`), Buffer.from(entry.data, 'base64'));
        }

        console.log(`[convert] Converting ${files.length} document(s) via Microsoft Word...`);
        const { stdout } = await runPowerShell(['-File', PS_SCRIPT, '-InputDir', inDir, '-OutputDir', outDir]);
        console.log(`[convert] ${stdout.trim()}`);

        const results = [];
        for (const { stem, displayName } of order) {
            const pdfPath = path.join(outDir, `${stem}.pdf`);
            let buf;
            try {
                buf = await fsp.readFile(pdfPath);
            } catch {
                sendJson(res, 500, { error: `Word did not produce a PDF for "${displayName}".` });
                return;
            }
            results.push({ name: displayName, data: buf.toString('base64') });
        }

        sendJson(res, 200, { files: results });
    } catch (err) {
        console.error('[convert] failed:', err);
        sendJson(res, 500, { error: `Word conversion failed: ${err.message}` });
    } finally {
        fsp.rm(workDir, { recursive: true, force: true }).catch(() => {});
    }
}

async function serveStatic(req, res) {
    const urlPath = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
    const rel = urlPath === '/' ? 'index.html' : urlPath.replace(/^\/+/, '');
    const filePath = path.join(ROOT, rel);

    // Prevent path traversal outside the project directory
    if (!filePath.startsWith(ROOT + path.sep) && filePath !== path.join(ROOT, 'index.html')) {
        send(res, 403, 'Forbidden');
        return;
    }

    try {
        const data = await fsp.readFile(filePath);
        const type = MIME[path.extname(filePath).toLowerCase()] || 'application/octet-stream';
        send(res, 200, data, { 'Content-Type': type });
    } catch {
        send(res, 404, 'Not found');
    }
}

const server = http.createServer((req, res) => {
    if (req.method === 'POST' && req.url === '/api/convert') {
        handleConvert(req, res);
        return;
    }
    if (req.method === 'GET' || req.method === 'HEAD') {
        serveStatic(req, res);
        return;
    }
    send(res, 405, 'Method not allowed');
});

if (!fs.existsSync(PS_SCRIPT)) {
    console.error(`Missing ${PS_SCRIPT} - DOCX to PDF conversion will not work.`);
}

server.listen(PORT, () => {
    console.log(`Certificate Generator running at http://localhost:${PORT}`);
    console.log('DOCX to PDF conversion uses Microsoft Word (COM automation).');
});
