require('dotenv').config();
const express = require('express');
const axios = require('axios');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const SHEET_ID = process.env.SHEET_ID || '1K-uQpZn21dM0YzInjE2Lj-ZSuWOf-a-y4Vr1WMtDjWY';
const SHEET_NAME = 'Job';

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Cache
let dataCache = { rows: null, ts: 0 };
const CACHE_TTL = 15000;

// ====== Presence (in-memory) ======
// Map<key, { userId, name, ts }>
const presenceMap = new Map();
const PRESENCE_TTL = 25000;

function cleanPresence() {
  const now = Date.now();
  for (const [k, v] of presenceMap) {
    if (now - v.ts > PRESENCE_TTL) presenceMap.delete(k);
  }
}

function parseGvizResponse(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start === -1 || end === -1) throw new Error('Invalid gviz response');
  const json = JSON.parse(text.slice(start, end + 1));
  const cols = json.table.cols.map(c => c.label || c.id);
  const rows = (json.table.rows || []).map(row => {
    const obj = {};
    (row.c || []).forEach((cell, i) => {
      const key = cols[i];
      if (!key) return;
      if (!cell || cell.v === null || cell.v === undefined) {
        obj[key] = '';
      } else if (typeof cell.v === 'string' && cell.v.startsWith('Date(')) {
        obj[key] = cell.f || cell.v;
      } else {
        obj[key] = String(cell.v).trim();
      }
    });
    return obj;
  });
  return rows;
}

async function fetchJobs(force = false) {
  const now = Date.now();
  if (!force && dataCache.rows && (now - dataCache.ts) < CACHE_TTL) {
    return dataCache.rows;
  }
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(SHEET_NAME)}&_=${now}`;
  const { data } = await axios.get(url, { timeout: 10000 });
  const rows = parseGvizResponse(data);
  dataCache = { rows, ts: now };
  return rows;
}

app.get('/api/jobs', async (req, res) => {
  try {
    const jobs = await fetchJobs(req.query.refresh === '1');
    const activeJobs = jobs.filter(job => {
      if (!job.Key) return false;
      const statusRaw = String(job.Status || '').trim().toLowerCase();
      const cancel = String(job.Cancel || '').trim().toLowerCase();
      if (cancel === 'yes' || cancel === 'true') return false;
      if (statusRaw === 'cancelled') return false;
      return statusRaw !== 'finished' && statusRaw !== '3';
    });
    res.json({ success: true, data: activeJobs, timestamp: Date.now() });
  } catch (err) {
    console.error('fetchJobs error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/jobs/finished', async (req, res) => {
  try {
    const jobs = await fetchJobs(req.query.refresh === '1');
    let finishedJobs = jobs.filter(job => {
      if (!job.Key) return false;
      const statusRaw = String(job.Status || '').trim().toLowerCase();
      const cancel = String(job.Cancel || '').trim().toLowerCase();
      if (cancel === 'yes' || cancel === 'true') return false;
      if (statusRaw === 'cancelled') return false;
      return statusRaw === 'finished' || statusRaw === '3';
    });

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
    const paginated = finishedJobs.slice(startIdx, endIdx);

    res.json({
      success: true,
      data: paginated,
      hasMore: endIdx < finishedJobs.length,
      total: finishedJobs.length,
      timestamp: Date.now()
    });
  } catch (err) {
    console.error('fetchFinishedJobs error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/jobs/update', async (req, res) => {
  try {
    const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL;
    if (!APPS_SCRIPT_URL || APPS_SCRIPT_URL.includes('YOUR_SCRIPT_ID')) {
      return res.status(500).json({ success: false, error: 'APPS_SCRIPT_URL ยังไม่ได้ตั้งค่าใน .env' });
    }
    const response = await axios.post(APPS_SCRIPT_URL, req.body, {
      headers: { 'Content-Type': 'application/json' },
      timeout: 15000
    });
    dataCache = { rows: null, ts: 0 }; // Invalidate cache
    res.json(response.data);
  } catch (err) {
    console.error('update error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', cached: !!dataCache.rows, cacheAge: Date.now() - dataCache.ts });
});

app.listen(PORT, () => {
  console.log(`\n🚀 DocDelivery App → http://localhost:${PORT}\n`);
});
