import 'dotenv/config';
import express from 'express';
import morgan from 'morgan';
import multer from 'multer';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { parse as csvParse } from 'csv-parse';
import { lookup as mimeLookup } from 'mime-types';

// Use the ESM build and wire Node fs (prevents readFile errors)
import * as XLSX from 'xlsx/xlsx.mjs';
XLSX.set_fs(fs);

import { SmppClient } from './smppClient.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ── Express setup
const app = express();
app.use(morgan('dev'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use('/static', express.static(path.join(__dirname, 'public')));

// ── Multer uploads
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir);
const upload = multer({ dest: uploadsDir });

// ── SMPP client
const client = new SmppClient({
    url: process.env.SMPP_URL,
    system_id: process.env.SMPP_SYSTEM_ID,
    password: process.env.SMPP_PASSWORD,
    senderId: process.env.SMPP_SENDER_ID,
    tps: process.env.SMPP_TPS,
    connectTimeout: process.env.SMPP_CONNECT_TIMEOUT_MS,
    enquireMs: process.env.SMPP_ENQUIRE_MS,
    reconnectMin: process.env.SMPP_RECONNECT_MIN_MS,
    reconnectMax: process.env.SMPP_RECONNECT_MAX_MS
});
client.connectAndBind().catch(e => console.error('Connect error:', e?.message || e));

// ── Helpers
const isExcel = (filename) => /\.(xlsx?|xlsm|xlsb|xls)$/i.test(filename);
const isCsv = (filename) => /\.csv$/i.test(filename);
const normalizeHeader = (s = '') =>
    String(s).trim().toLowerCase().replace(/\s+/g, ' ').replace(/[^\w ]/g, '').replace(/\s/g, '_');

function parseCsvFile(filePath) {
    return new Promise((resolve, reject) => {
        const rows = [];
        fs.createReadStream(filePath)
            .pipe(csvParse({ columns: true, trim: true }))
            .on('data', (r) => rows.push(r))
            .on('end', () => resolve(rows))
            .on('error', reject);
    });
}

// Auto-detect the row that actually contains the headers (skips title rows)
function parseExcelSmart(filePath) {
    const buf = fs.readFileSync(filePath);
    const wb = XLSX.read(buf, { type: 'buffer' });
    const sheet = wb.Sheets[wb.SheetNames[0]];

    // First read as array-of-arrays
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });

    // Heuristic to find the header row
    const wanted = new Set(['telephone', 'number', 'msisdn', 'phone', 'to', 'name', 'names', 'names_of_farmers']);
    let headerRow = 0;
    for (let i = 0; i < Math.min(50, aoa.length); i++) {
        const row = aoa[i].map(x => String(x).trim());
        const norm = row.map(normalizeHeader).filter(Boolean);
        const hasAny = norm.some(h => wanted.has(h));
        const filled = norm.length;
        if (hasAny && filled >= 2) { headerRow = i; break; }
    }

    // Re-read from the detected header row, returning objects
    const rangeRef = (() => {
        const ref = sheet['!ref'];
        if (!ref) return undefined;
        const R = XLSX.utils.decode_range(ref);
        R.s.r = headerRow;
        return XLSX.utils.encode_range(R);
    })();

    const rows = XLSX.utils.sheet_to_json(sheet, {
        raw: false,
        defval: '',
        range: rangeRef // start at header row
    });

    return rows;
}

// Phone normalization
function normalizeMsisdn(v) {
    if (v === null || v === undefined) return '';
    let s = String(v).trim();
    s = s.replace(/[^\d]/g, ''); // remove spaces, commas, dots, etc.
    return s;
}

function titleCaseName(name = '') {
    return String(name).trim().toLowerCase().replace(/\b([a-z])/g, (m, c) => c.toUpperCase());
}

function renderTemplate(template, rowObj, nameKey) {
    const tpl = String(template ?? '');
    let out = tpl.replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_, k) => {
        const kk = normalizeHeader(k);
        return rowObj[kk] !== undefined && rowObj[kk] !== null ? String(rowObj[kk]) : '';
    });
    if (nameKey && rowObj[nameKey]) {
        out = out.replace(/\{\{\s*name\s*\}\}/gi, String(rowObj[nameKey]));
    }
    return out;
}

function htmlEscape(s = '') {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
}

// ── In-memory pending uploads for Designer
const pendingUploads = new Map();
const UPLOAD_TTL_MS = 60 * 60 * 1000;
function newUploadId() { return (Date.now().toString(36) + Math.random().toString(36).slice(2, 10)).toUpperCase(); }
function cleanupPending() {
    const now = Date.now();
    for (const [id, info] of pendingUploads.entries()) {
        if (now - info.uploadedAt > UPLOAD_TTL_MS) {
            try { fs.unlinkSync(info.path); } catch { }
            pendingUploads.delete(id);
        }
    }
}

// ── Pages
app.get('/', (_req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.get('/designer', (_req, res) => res.sendFile(path.join(__dirname, 'public', 'designer.html')));

app.get('/template.csv', (_req, res) => {
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename="template.csv"');
    res.send('name,number,message\nJohn,256700000001,Hello John!\nMary,256700000002,Hello Mary!');
});

app.get('/health', (_req, res) => res.json({ status: 'ok', bound: client.bound, queue: client.queue.length }));

// ── Single send (kept)
app.post('/send-one', async (req, res) => {
    try {
        const { number, message, name } = req.body;
        if (!number || !message) return res.status(400).send('Missing number or message');

        const finalText = titleCaseName(name || '')
            ? renderTemplate('Hullo {{name}}, ' + message, { name }, 'name')
            : String(message);

        const r = await client.send(String(normalizeMsisdn(number)), String(finalText));
        res.send(`<pre>✅ Sent. Message ID: ${r.message_id}
Text: ${finalText}
<a href="/">Back</a></pre>`);
    } catch (e) {
        res.status(500).send(`<pre>❌ ${e?.message || e}
<a href="/">Back</a></pre>`);
    }
});

// ── Designer: upload & parse
app.post('/design-parse', upload.single('file'), async (req, res) => {
    const file = req.file;
    if (!file) return res.status(400).json({ error: 'No file uploaded' });
    cleanupPending();

    try {
        let rows = [];
        if (isCsv(file.originalname)) rows = await parseCsvFile(file.path);
        else if (isExcel(file.originalname)) rows = parseExcelSmart(file.path);
        else {
            const mt = mimeLookup(file.originalname) || '';
            if (String(mt).includes('csv')) rows = await parseCsvFile(file.path);
            else return res.status(400).json({ error: 'Unsupported file type. Use .csv or .xlsx' });
        }

        if (!rows.length) return res.status(400).json({ error: 'No rows found in the first sheet.' });

        // Normalize headers for uniform keys in UI
        const rawCols = Object.keys(rows[0] || {});
        const columns = rawCols.map(c => ({ label: c, key: normalizeHeader(c) }));
        const normRows = rows.map(r => {
            const o = {};
            for (const [k, v] of Object.entries(r)) o[normalizeHeader(k)] = v;
            return o;
        });

        const uploadId = newUploadId();
        pendingUploads.set(uploadId, { path: file.path, originalname: file.originalname, uploadedAt: Date.now() });

        const sample = normRows.slice(0, 25).map((r, i) => ({ idx: i, row: r }));
        res.json({ uploadId, columns, rowCount: normRows.length, sample });
    } catch (e) {
        res.status(500).json({ error: e?.message || String(e) });
    }
});

// ── Designer: send
app.post('/design-send', async (req, res) => {
    try {
        const { uploadId, numberCol, nameCol, messageTemplate, selectedIdxs, attachFields } = req.body || {};
        if (!uploadId) return res.status(400).json({ error: 'Missing uploadId' });
        if (!numberCol) return res.status(400).json({ error: 'Select the Number column' });
        if (!messageTemplate) return res.status(400).json({ error: 'Message template required' });

        const info = pendingUploads.get(uploadId);
        if (!info) return res.status(410).json({ error: 'Upload expired. Please re-upload.' });

        // Re-parse full file
        let rows = [];
        if (isCsv(info.originalname)) rows = await parseCsvFile(info.path);
        else rows = parseExcelSmart(info.path);

        const normRows = rows.map(r => {
            const o = {};
            for (const [k, v] of Object.entries(r)) o[normalizeHeader(k)] = v;
            return o;
        });

        const targets = Array.isArray(selectedIdxs) && selectedIdxs.length
            ? selectedIdxs.filter(n => Number.isInteger(n) && n >= 0 && n < normRows.length)
            : normRows.map((_, i) => i);

        const results = [];
        let ok = 0, fail = 0;

        for (const i of targets) {
            const row = normRows[i];
            const msisdn = normalizeMsisdn(row[numberCol]);
            if (!msisdn || !/^\d{10,15}$/.test(msisdn)) {
                results.push({ idx: i, status: 'error', detail: `Invalid number in ${numberCol}` });
                fail++; continue;
            }

            let text = renderTemplate(messageTemplate, row, nameCol);
            if (Array.isArray(attachFields) && attachFields.length) {
                const extras = attachFields
                    .filter(k => k && row[k] !== undefined && row[k] !== '')
                    .map(k => `${k.replace(/_/g, ' ')}: ${row[k]}`);
                if (extras.length) text = `${text}\n${extras.join('\n')}`;
            }

            try {
                const r = await client.send(msisdn, text);
                results.push({ idx: i, status: 'sent', message_id: r.message_id, number: msisdn, text });
                ok++;
            } catch (e) {
                results.push({ idx: i, status: 'error', number: msisdn, text, detail: e?.message || String(e) });
                fail++;
            }
        }

        try { fs.unlinkSync(info.path); } catch { }
        pendingUploads.delete(uploadId);

        res.json({ ok, fail, total: results.length, results });
    } catch (e) {
        res.status(500).json({ error: e?.message || String(e) });
    }
});

// ── Legacy bulk upload kept (unchanged behavior; uses Excel smart parser too)
function renderPersonalizedMessage(rawMsg, name, personalizeFlag) {
    const msg = String(rawMsg ?? '').trim();
    const cleanName = titleCaseName(name || '');
    let out = msg
        .replace(/\{\{\s*name\s*\}\}/gi, cleanName)
        .replace(/\{\s*name\s*\}/gi, cleanName);
    const replaced = out !== msg;
    if (!replaced && personalizeFlag && cleanName) out = `Hullo ${cleanName}, ${msg}`;
    return out;
}
function mapRow(row) {
    const entries = Object.fromEntries(Object.entries(row).map(([k, v]) => [normalizeHeader(k), v]));
    const name = entries['name'] ?? entries['full_name'] ?? entries['names'] ?? '';
    const message = entries['message'] ?? entries['text'] ?? '';
    const numberRaw = entries['number'] ?? entries['msisdn'] ?? entries['telephone'] ?? entries['phone'] ?? entries['to'] ?? '';
    const number = normalizeMsisdn(numberRaw);
    return { name: String(name || ''), message: String(message || ''), number };
}
function validateRow({ message, number }) {
    const errs = [];
    if (!number || !/^\d{10,15}$/.test(number)) errs.push('Invalid number');
    if (!message) errs.push('Missing message');
    return errs;
}
app.post('/upload', upload.single('file'), async (req, res) => {
    const file = req.file;
    if (!file) return res.status(400).send('No file uploaded');
    const personalize = !!req.body.personalize;

    try {
        let rows = [];
        if (isCsv(file.originalname)) rows = await parseCsvFile(file.path);
        else if (isExcel(file.originalname)) rows = parseExcelSmart(file.path);
        else {
            const mt = mimeLookup(file.originalname) || '';
            if (String(mt).includes('csv')) rows = await parseCsvFile(file.path);
            else return res.status(400).send('Unsupported file type. Use .csv or .xlsx');
        }

        const mapped = rows.map(mapRow);
        const results = [];
        let ok = 0, fail = 0;

        for (const row of mapped) {
            const errs = validateRow(row);
            if (errs.length) {
                results.push({ ...row, status: 'error', detail: errs.join('; ') });
                fail++;
                continue;
            }

            const finalText = renderPersonalizedMessage(row.message, row.name, personalize);
            try {
                const r = await client.send(row.number, finalText);
                results.push({ ...row, status: 'sent', message_id: r.message_id, finalText });
                ok++;
            } catch (e) {
                results.push({ ...row, status: 'error', detail: e?.message || String(e), finalText });
                fail++;
            }
        }

        res.send(`<!doctype html>
<html><head><meta charset="utf-8"><title>Upload report</title>
<style>
body{font-family:system-ui,Segoe UI,Roboto,Arial;margin:2rem}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ddd;padding:8px;font-size:14px}
th{background:#f3f4f6;text-align:left}
.badge{display:inline-block;padding:.2rem .5rem;border-radius:.5rem}
</style></head>
<body>
<h2>Bulk Send Report</h2>
<p><span class="badge" style="background:#e0ffe0">✅ OK: ${ok}</span>
   <span class="badge" style="background:#ffe0e0">❌ Failed: ${fail}</span>
   <span class="badge" style="background:#e0e7ff">Total: ${results.length}</span></p>
<p><a href="/">← Back</a> • <a href="/designer">Open Designer</a></p>
<table>
<thead><tr>
  <th>Name</th><th>Number</th><th>Message (final)</th><th>Status</th><th>Detail/MsgID</th>
</tr></thead>
<tbody>${results.map(r => `
<tr>
  <td>${r.name || ''}</td>
  <td>${r.number}</td>
  <td>${htmlEscape(r.finalText || r.message || '')}</td>
  <td>${r.status}</td>
  <td>${r.message_id || r.detail || ''}</td>
</tr>`).join('')}</tbody>
</table>
</body></html>`);
    } catch (e) {
        res.status(500).send(`Error: ${e?.message || e}`);
    } finally {
        try { fs.unlinkSync(file.path); } catch { }
    }
});

// ── Start server
const port = parseInt(process.env.HTTP_PORT || '8080', 10);
const host = process.env.HTTP_BIND || '127.0.0.1';
app.listen(port, host, () => console.log(`Web UI on http://${host}:${port}`));
