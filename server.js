require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const { google } = require('googleapis');
const { Readable } = require('stream');

const app = express();
const PORT = process.env.PORT || 3000;
const SHEET_ID = process.env.SHEET_ID || '1K-uQpZn21dM0YzInjE2Lj-ZSuWOf-a-y4Vr1WMtDjWY';
const SHEET_NAME = process.env.SHEET_NAME || 'Job';
const CREDS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || './reflected-drake-427610-p8-13c0068b2a7a.json';
const SIGNATURE_FOLDER_ID = process.env.SIGNATURE_FOLDER_ID || '';
const POLL_MS = parseInt(process.env.SHEETS_POLL_MS || '2500', 10);

// ---- Google APIs ----
const auth = new google.auth.GoogleAuth({
  keyFile: path.resolve(CREDS_PATH),
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
  ],
});
const sheets = google.sheets({ version: 'v4', auth });
const drive = google.drive({ version: 'v3', auth });

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ---- State ----
let cached = { rows: null, ts: 0, sig: '' };
let headers = [];

// ---- SSE ----
const sseClients = new Set();
function broadcast(event, payload) {
  const chunk = `event: ${event}\ndata: ${JSON.stringify(payload)}\n\n`;
  for (const c of sseClients) { try { c.write(chunk); } catch {} }
}

function hashRows(rows) {
  let h = 0;
  const str = JSON.stringify(rows.map(r => [r.Key, r.Status, r.Cancel, r.Dropoff, r.Remark, r['Txt_01']]));
  for (let i = 0; i < str.length; i++) h = ((h << 5) - h + str.charCodeAt(i)) | 0;
  return String(h);
}

function colLetter(idx) {
  let n = idx + 1, s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ---- Core fetch ----
async function fetchRowsFromSheets() {
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: SHEET_NAME,
    valueRenderOption: 'FORMATTED_VALUE',
    dateTimeRenderOption: 'FORMATTED_STRING',
  });
  const values = resp.data.values || [];
  if (values.length === 0) return { headers: [], rows: [] };
  const hdrs = values[0].map(h => String(h || '').trim());
  const rows = values.slice(1).map(r => {
    const o = {};
    hdrs.forEach((h, i) => { if (h) o[h] = r[i] != null ? String(r[i]).trim() : ''; });
    return o;
  });
  return { headers: hdrs, rows };
}

async function refreshCache(force = false) {
  if (!force && cached.rows && (Date.now() - cached.ts) < 1000) return cached;
  const { headers: hdrs, rows } = await fetchRowsFromSheets();
  headers = hdrs;
  const sig = hashRows(rows);
  const changed = sig !== cached.sig;
  cached = { rows, ts: Date.now(), sig };
  if (changed) broadcast('jobs', { ts: cached.ts, count: rows.length });
  return cached;
}

function filterActive(rows) {
  return rows.filter(job => {
    if (!job.Key) return false;
    const s = String(job.Status || '').trim().toLowerCase();
    const c = String(job.Cancel || '').trim().toLowerCase();
    if (c === 'yes' || c === 'true') return false;
    if (s === 'cancelled') return false;
    return s !== 'finished' && s !== '3';
  });
}
function filterFinished(rows) {
  return rows.filter(job => {
    if (!job.Key) return false;
    const s = String(job.Status || '').trim().toLowerCase();
    const c = String(job.Cancel || '').trim().toLowerCase();
    if (c === 'yes' || c === 'true') return false;
    if (s === 'cancelled') return false;
    return s === 'finished' || s === '3';
  });
}

// ---- Presence ----
const presenceMap = new Map();
const PRESENCE_TTL = 25000;
function cleanPresence() {
  const now = Date.now();
  for (const [k, v] of presenceMap) if (now - v.ts > PRESENCE_TTL) presenceMap.delete(k);
}

// ---- Routes ----
app.get('/api/jobs', async (req, res) => {
  try {
    await refreshCache(req.query.refresh === '1');
    res.json({ success: true, data: filterActive(cached.rows), timestamp: cached.ts });
  } catch (err) {
    console.error('jobs error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/jobs/finished', async (req, res) => {
  try {
    await refreshCache(req.query.refresh === '1');
    let finishedJobs = filterFinished(cached.rows);
    finishedJobs.sort((a, b) => (b['Dropoff'] || b['เวลาทำรายการ'] || '').localeCompare(a['Dropoff'] || a['เวลาทำรายการ'] || ''));

    const { month, page = 1, limit = 50 } = req.query;
    if (month) {
      finishedJobs = finishedJobs.filter(item => {
        const dropoff = item['Dropoff'] || item['เวลาทำรายการ'] || '';
        const match = dropoff.match(/(\d{2})\/(\d{2})\/(\d{4})/);
        if (match) {
          let year = parseInt(match[3]);
          if (year > 2500) year -= 543;
          return `${year}-${match[2]}` === month;
        }
        return false;
      });
    }
    const startIdx = (parseInt(page) - 1) * parseInt(limit);
    const endIdx = parseInt(page) * parseInt(limit);
    res.json({
      success: true,
      data: finishedJobs.slice(startIdx, endIdx),
      hasMore: endIdx < finishedJobs.length,
      total: finishedJobs.length,
      timestamp: cached.ts,
    });
  } catch (err) {
    console.error('finished error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

async function findRowByKey(key) {
  await refreshCache(false);
  for (let i = 0; i < cached.rows.length; i++) {
    if (String(cached.rows[i].Key) === String(key)) return { rowNumber: i + 2, row: cached.rows[i] };
  }
  return null;
}

app.post('/api/jobs/update', async (req, res) => {
  try {
    cleanPresence();
    const { key, _userId, status, signature, dropoff, remark, action, note } = req.body || {};
    if (!key) return res.status(400).json({ success: false, error: 'ไม่ได้ระบุ Key' });

    if (_userId) {
      const cur = presenceMap.get(key);
      if (cur && cur.userId !== _userId) {
        return res.status(409).json({
          success: false, conflict: true, owner: cur.name,
          error: 'รายการ ' + key + ' กำลังถูกดำเนินการโดย ' + cur.name,
        });
      }
    }

    const found = await findRowByKey(key);
    if (!found) return res.status(404).json({ success: false, error: 'ไม่พบข้อมูล Key ที่ระบุ' });
    const { rowNumber } = found;
    const idxOf = name => headers.indexOf(name);
    const updates = [];

    let sigSaved = false;

    if (action === 'cancel') {
      const cancelIdx = idxOf('Cancel');
      if (cancelIdx !== -1) {
        updates.push({ range: `${SHEET_NAME}!${colLetter(cancelIdx)}${rowNumber}`, values: [['Yes']] });
      } else {
        const statusIdx = idxOf('Status');
        if (statusIdx !== -1) updates.push({ range: `${SHEET_NAME}!${colLetter(statusIdx)}${rowNumber}`, values: [['Cancelled']] });
      }
      updates.push({ range: `${SHEET_NAME}!T${rowNumber}`, values: [[note || '']] });
    } else {
      const statusIdx = idxOf('Status');
      if (statusIdx !== -1 && status !== undefined) {
        updates.push({ range: `${SHEET_NAME}!${colLetter(statusIdx)}${rowNumber}`, values: [[status]] });
      }
      if (status === 'Finished') {
        const dropoffIdx = idxOf('Dropoff');
        if (dropoffIdx !== -1 && dropoff) updates.push({ range: `${SHEET_NAME}!${colLetter(dropoffIdx)}${rowNumber}`, values: [[dropoff]] });
        const remarkIdx = idxOf('Remark');
        if (remarkIdx !== -1 && remark) updates.push({ range: `${SHEET_NAME}!${colLetter(remarkIdx)}${rowNumber}`, values: [[remark]] });

        // Store signature as data URL directly in the cell (no Drive needed)
        // Google Sheets cell limit = 50,000 chars; compressed JPEG signature ~4-11K chars (fits easily)
        if (signature && signature.startsWith('data:image')) {
          const CELL_LIMIT = 49500; // leave a little headroom
          if (signature.length > CELL_LIMIT) {
            console.warn(`signature too large (${signature.length} chars); skipped. Increase client compression.`);
          } else {
            const txt01Idx = idxOf('Txt_01');
            const dropoffSigIdx = idxOf('Dropoff Signature');
            if (txt01Idx !== -1) updates.push({ range: `${SHEET_NAME}!${colLetter(txt01Idx)}${rowNumber}`, values: [[signature]] });
            if (dropoffSigIdx !== -1) updates.push({ range: `${SHEET_NAME}!${colLetter(dropoffSigIdx)}${rowNumber}`, values: [[signature]] });
            sigSaved = true;
          }
        }
      }
    }

    if (updates.length === 0) {
      return res.json({ success: false, error: 'ไม่มีข้อมูลจะอัปเดต' });
    }

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { valueInputOption: 'RAW', data: updates },
    });

    cached = { rows: null, ts: 0, sig: '' };
    refreshCache(true).catch(() => {}); // fire-and-forget — SSE will catch up
    if (key) presenceMap.delete(key);

    res.json({
      success: true,
      message: action === 'cancel' ? 'ลบรายการเรียบร้อย' : 'อัปเดตข้อมูลสำเร็จ',
      sigSaved,
    });
  } catch (err) {
    console.error('update error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ---- Presence ----
app.post('/api/presence/heartbeat', (req, res) => {
  cleanPresence();
  const { userId, name, claims = [] } = req.body || {};
  if (!userId) return res.status(400).json({ success: false, error: 'userId required' });
  const now = Date.now();
  const claimSet = new Set(claims);
  for (const key of claimSet) {
    const cur = presenceMap.get(key);
    if (!cur || cur.userId === userId) presenceMap.set(key, { userId, name: name || 'ผู้ใช้', ts: now });
  }
  for (const [k, v] of presenceMap) {
    if (v.userId === userId && !claimSet.has(k)) presenceMap.delete(k);
  }
  const others = {};
  for (const [k, v] of presenceMap) if (v.userId !== userId) others[k] = { userId: v.userId, name: v.name };
  res.json({ success: true, others });
});

app.post('/api/presence/release', (req, res) => {
  const { userId, key } = req.body || {};
  if (!userId || !key) return res.status(400).json({ success: false });
  const cur = presenceMap.get(key);
  if (cur && cur.userId === userId) presenceMap.delete(key);
  res.json({ success: true });
});

// ---- SSE (real-time push) ----
app.get('/api/events', (req, res) => {
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache, no-transform',
    Connection: 'keep-alive',
    'X-Accel-Buffering': 'no',
  });
  res.write(`: connected\n\n`);
  res.write(`event: hello\ndata: ${JSON.stringify({ ts: Date.now() })}\n\n`);
  sseClients.add(res);
  const ping = setInterval(() => { try { res.write(`: ping\n\n`); } catch {} }, 20000);
  req.on('close', () => { clearInterval(ping); sseClients.delete(res); });
});

app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    cached: !!cached.rows,
    rowCount: cached.rows ? cached.rows.length : 0,
    cacheAge: Date.now() - cached.ts,
    sseClients: sseClients.size,
    pollMs: POLL_MS,
  });
});

// ---- Background poller for real-time ----
async function pollLoop() {
  try { await refreshCache(true); }
  catch (err) { console.error('poll error:', err.message); }
  finally { setTimeout(pollLoop, POLL_MS); }
}

app.listen(PORT, async () => {
  console.log(`\n🚀 DocDelivery App → http://localhost:${PORT}`);
  console.log(`   Poll interval: ${POLL_MS}ms | SSE stream: /api/events\n`);
  try {
    await refreshCache(true);
    console.log(`✅ Sheets API connected. ${cached.rows.length} rows cached.\n`);
  } catch (err) {
    console.error('❌ Sheets API error:', err.message);
    console.error('   ตรวจสอบ: (1) share Sheet ให้ service account, (2) path credentials JSON ถูกต้อง\n');
  }
  pollLoop();
});
