import 'dotenv/config';
import express from 'express';
import morgan from 'morgan';
import multer from 'multer';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { parse as csvParse } from 'csv-parse';
import * as XLSX from 'xlsx';
import { lookup as mimeLookup } from 'mime-types';
import { SmppClient } from './smppClient.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ── Express setup
const app = express();
app.use(morgan('dev'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use('/static', express.static(path.join(__dirname, 'public')));

// ── Ensure uploads folder exists, then Multer (uploads)
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
const normalizeHeader = (s = '') => String(s).trim().toLowerCase();
const isExcel = (filename) => /\.xlsx?$/i.test(filename);
const isCsv = (filename) => /\.csv$/i.test(filename);

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

// IMPORTANT: raw:false makes XLSX give us the *displayed text* (avoids 2.56E+11)
function parseExcelFile(filePath) {
    const wb = XLSX.readFile(filePath);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { raw: false, defval: '' }); // header row required
}

// Strip formatting (+, spaces, dashes, parentheses) and handle weird notations
function normalizeMsisdn(v) {
    if (v === null || v === undefined) return '';
    let s = String(v).trim();

    // If Excel scientific notation slipped in (shouldn't with raw:false), fall back to digits only
    if (/^\d+(\.\d+)?e\+\d+$/i.test(s)) {
        // We can’t accurately expand huge numbers safely here; keep only digits
        // (raw:false should already provide the full value as text)
    }

    // Remove non-digits
    s = s.replace(/[^\d]/g, '');
    return s;
}

function titleCaseName(name = '') {
    return String(name)
        .trim()
        .toLowerCase()
        .replace(/\b([a-z])/g, (m, c) => c.toUpperCase());
}

function renderPersonalizedMessage(rawMsg, name, personalizeFlag) {
    const msg = String(rawMsg ?? '').trim();
    const cleanName = titleCaseName(name || '');

    // Replace placeholders if present
    let out = msg
        .replace(/\{\{\s*name\s*\}\}/gi, cleanName)
        .replace(/\{\s*name\s*\}/gi, cleanName);

    // If no placeholder replacements happened and personalization requested, prefix it
    const replaced = out !== msg;
    if (!replaced && personalizeFlag && cleanName) {
        out = `Hullo ${cleanName}, ${msg}`;
    }
    return out;
}

function mapRow(row) {
    // Accept header variations
    const entries = Object.fromEntries(
        Object.entries(row).map(([k, v]) => [normalizeHeader(k), v])
    );
    const name = entries['name'] ?? entries['full name'] ?? '';
    const message = entries['message'] ?? entries['text'] ?? '';
    const numberRaw = entries['number'] ?? entries['msisdn'] ?? entries['phone'] ?? entries['to'] ?? '';

    const number = normalizeMsisdn(numberRaw);

    return { name: String(name || ''), message: String(message || ''), number };
}

function validateRow({ message, number }) {
    const errs = [];
    // E.164-ish: 10 to 15 digits
    if (!number || !/^\d{10,15}$/.test(number)) errs.push('Invalid number');
    if (!message) errs.push('Missing message');
    return errs;
}

// ── Routes
app.get('/', (_req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/template.csv', (_req, res) => {
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename="template.csv"');
    res.send('name,number,message\nJohn,256700000001,Hello John!\nMary,256700000002,Hello Mary!');
});

app.get('/health', (_req, res) => {
    res.json({ status: 'ok', bound: client.bound, queue: client.queue.length });
});

// Single send (form)
app.post('/send-one', async (req, res) => {
    try {
        const { number, message, name } = req.body;
        const personalize = !!req.body.personalize;

        if (!number || !message) return res.status(400).send('Missing number or message');

        const finalText = renderPersonalizedMessage(message, name, personalize);
        const r = await client.send(String(normalizeMsisdn(number)), String(finalText));

        res.send(`<pre>✅ Sent. Message ID: ${r.message_id}
Text: ${finalText}
<a href="/">Back</a></pre>`);
    } catch (e) {
        res.status(500).send(`<pre>❌ ${e?.message || e}
<a href="/">Back</a></pre>`);
    }
});

// Bulk upload
app.post('/upload', upload.single('file'), async (req, res) => {
    const file = req.file;
    if (!file) return res.status(400).send('No file uploaded');



    const personalize = !!req.body.personalize; // toggle from checkbox

    try {
        let rows = [];
        if (isCsv(file.originalname)) {
            rows = await parseCsvFile(file.path);
        } else if (isExcel(file.originalname)) {
            rows = parseExcelFile(file.path);
        } else {
            // try by MIME as a fallback
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

        // Simple HTML report
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
<p><a href="/">← Back</a></p>
<table>
<thead><tr>
  <th>Name</th><th>Number</th><th>Message (final)</th><th>Status</th><th>Detail/MsgID</th>
</tr></thead>
<tbody>${results.map(r => `
<tr>
  <td>${r.name || ''}</td>
  <td>${r.number}</td>
  <td>${(r.finalText || r.message || '').replace(/</g, '&lt;')}</td>
  <td>${r.status}</td>
  <td>${r.message_id || r.detail || ''}</td>
</tr>`).join('')}</tbody>
</table>
</body></html>`);
    } catch (e) {
        res.status(500).send(`Error: ${e?.message || e}`);
    } finally {
        // cleanup temp file
        try { fs.unlinkSync(file.path); } catch { }
    }
});

// ── Start server
const port = parseInt(process.env.HTTP_PORT || '8080', 10);
const host = process.env.HTTP_BIND || '127.0.0.1';
app.listen(port, host, () => {
    console.log(`Web UI on http://${host}:${port}`);
});
