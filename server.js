require('dotenv').config();
const express = require('express');
const path = require('path');
const nodemailer = require('nodemailer');
const { google } = require('googleapis');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
const REPO = 'frazeved/SAMANTHA';

const SHEET_BASE = 'https://docs.google.com/spreadsheets/d/1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q/edit';

const WORKFLOWS = [
  { id: 'po-detail.yml',   name: 'PO DETAIL',   tab: 'PO DETAIL',  sheetUrl: `${SHEET_BASE}?gid=2017761959#gid=2017761959` },
  { id: 'all-pos.yml',     name: 'ALL POs',      tab: 'PO TRADE',   sheetUrl: `${SHEET_BASE}?gid=890202899#gid=890202899`   },
  { id: 'po-headers.yml',  name: 'PO HEADERS',   tab: 'PO NEW',     sheetUrl: `${SHEET_BASE}?gid=406184613#gid=406184613`   },
  { id: 'po-invoiced.yml', name: 'PO INVOICED',  tab: 'PO INVOICE', sheetUrl: `${SHEET_BASE}?gid=1379989532#gid=1379989532` },
];

async function ghFetch(url, opts = {}) {
  return fetch(url, {
    ...opts,
    headers: {
      Authorization: `Bearer ${GITHUB_TOKEN}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json',
      ...opts.headers,
    },
  });
}

app.get('/api/status', async (req, res) => {
  try {
    const r = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/all-extractions.yml/runs?per_page=1`);
    const data = await r.json();
    const run = data.workflow_runs?.[0] || null;
    const duration =
      run && run.status === 'completed' && run.run_started_at && run.updated_at
        ? Math.round((new Date(run.updated_at) - new Date(run.run_started_at)) / 1000)
        : null;

    const results = WORKFLOWS.map(wf => ({
      id: wf.id,
      name: wf.name,
      tab: wf.tab,
      sheetUrl: wf.sheetUrl,
      status: run?.status || 'unknown',
      conclusion: run?.conclusion || null,
      started_at: run?.run_started_at || null,
      updated_at: run?.updated_at || null,
      duration,
    }));

    res.json(results);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/run-all', async (req, res) => {
  try {
    const r = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/all-extractions.yml/dispatches`, {
      method: 'POST',
      body: JSON.stringify({ ref: 'main' }),
    });
    if (r.status !== 204) {
      const body = await r.text();
      console.error(`GitHub dispatch failed: ${r.status} ${body}`);
      return res.status(500).json({ error: `GitHub returned ${r.status}: ${body}` });
    }
    res.json({ ok: true });
  } catch (e) {
    console.error('run-all error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/run-jhonny', async (req, res) => {
  try {
    const r = await ghFetch(`https://api.github.com/repos/frazeved/JHONNY/actions/workflows/update-fedex-status.yml/dispatches`, {
      method: 'POST',
      body: JSON.stringify({ ref: 'main' }),
    });
    if (r.status !== 204) {
      const body = await r.text();
      return res.status(500).json({ error: `GitHub returned ${r.status}: ${body}` });
    }
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── FedEx Status helpers ──────────────────────────────────────────────────────
const SHEET_ID  = '1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q';
const SHEET_GID = 99866922;
const TAB_NAME  = 'Warehouse Now Database';

function parseCSV(text) {
  const rows = []; let row = [], field = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (c === '"') inQ = false;
      else field += c;
    } else {
      if (c === '"') inQ = true;
      else if (c === ',') { row.push(field); field = ''; }
      else if (c === '\n' || (c === '\r' && text[i + 1] === '\n')) {
        if (c === '\r') i++;
        row.push(field); field = '';
        if (row.some(f => f.trim())) rows.push(row);
        row = [];
      } else field += c;
    }
  }
  if (row.length) { row.push(field); if (row.some(f => f.trim())) rows.push(row); }
  return rows;
}

function findCol(headers, ...kws) {
  return headers.findIndex(h => kws.some(k => h.trim().toLowerCase().includes(k.toLowerCase())));
}

function colLetter(idx) {
  let col = '', i = idx;
  while (i >= 0) { col = String.fromCharCode(65 + (i % 26)) + col; i = Math.floor(i / 26) - 1; }
  return col;
}

async function getFedExToken() {
  const base = process.env.FEDEX_ENV === 'production'
    ? 'https://apis.fedex.com' : 'https://apis-sandbox.fedex.com';
  const r = await fetch(`${base}/oauth/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: process.env.FEDEX_API_KEY,
      client_secret: process.env.FEDEX_API_SECRET,
    }),
  });
  if (!r.ok) throw new Error(`FedEx OAuth: ${r.status} ${await r.text()}`);
  const d = await r.json();
  if (!d.access_token) throw new Error('FedEx OAuth: no token');
  return { token: d.access_token, base };
}

async function trackBatch(base, token, numbers) {
  const r = await fetch(`${base}/track/v1/trackingnumbers`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', 'x-locale': 'en_US' },
    body: JSON.stringify({
      includeDetailedScans: false,
      trackingInfo: numbers.map(n => ({
        trackingNumberInfo: { trackingNumber: n },
        ...(process.env.FEDEX_ACCOUNT_NUMBER ? { associatedAccountNumber: { value: process.env.FEDEX_ACCOUNT_NUMBER } } : {}),
      })),
    }),
  });
  if (!r.ok) throw new Error(`FedEx Track: ${r.status} ${await r.text()}`);
  return r.json();
}

app.post('/api/fedex-status', async (req, res) => {
  try {
    // 1. Read sheet
    const csvRes = await fetch(`https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&gid=${SHEET_GID}`);
    if (!csvRes.ok) throw new Error(`CSV fetch failed: ${csvRes.status}`);
    const rows = parseCSV(await csvRes.text());
    if (rows.length < 2) throw new Error('Sheet has no data');

    const H = rows[0].map(h => h.trim().replace(/^﻿/, ''));
    const CI = {
      style:  findCol(H, 'style #', 'style#'),
      po:     findCol(H, 'po#', 'po #'),
      status: findCol(H, 'warehouse status', 'wh status', 'status'),
      ship:   findCol(H, 'ship date tradestone', 'ship date'),
      cancel: findCol(H, 'cancel date'),
      trk:    findCol(H, 'tracking number'),
      sts:    findCol(H, 'delivery status', 'fedex status'),
    };

    const get = (r, i) => i >= 0 ? (r[i] || '').trim() : '';
    const cutoff = new Date(Date.now() - 30 * 86400000);

    // 2. Filter: SHIPPED + last 30 days + has tracking number
    const toProcess = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const trk    = get(r, CI.trk);
      const sd     = new Date(get(r, CI.ship));
      const rowSts = get(r, CI.status).toLowerCase();
      if (!trk || isNaN(sd) || sd < cutoff || !rowSts.includes('ship')) continue;
      toProcess.push({
        rowIndex: i,
        tracking: trk,
        style:  get(r, CI.style),
        po:     get(r, CI.po),
        status: get(r, CI.status),
        ship:   get(r, CI.ship),
        cancel: get(r, CI.cancel),
        currentSts: get(r, CI.sts),
      });
    }

    if (!toProcess.length) return res.json({ rows: [], updated: 0 });

    // 3. FedEx tracking
    const { token, base } = await getFedExToken();
    const statusMap = {};
    for (let i = 0; i < toProcess.length; i += 30) {
      const batch = toProcess.slice(i, i + 30);
      const result = await trackBatch(base, token, batch.map(b => b.tracking));
      for (const item of (result.output?.completeTrackResults || [])) {
        const tr = item.trackResults?.[0];
        const s  = tr?.latestStatusDetail?.statusByLocale || tr?.latestStatusDetail?.description;
        if (s) statusMap[item.trackingNumber] = s;
      }
    }

    // 4. Write back to Google Sheet
    const updatedAt = new Date().toLocaleDateString('en-US', {
      month: 'short', day: 'numeric', year: '2-digit', hour: '2-digit', minute: '2-digit',
    });
    const sheetUpdates = toProcess.filter(r => statusMap[r.tracking]);
    if (sheetUpdates.length && process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
      try {
        const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
        if (creds.private_key) creds.private_key = creds.private_key.replace(/\\n/g, '\n');
        const auth = new google.auth.GoogleAuth({ credentials: creds, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
        const sheets = google.sheets({ version: 'v4', auth });
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: SHEET_ID,
          requestBody: {
            valueInputOption: 'RAW',
            data: sheetUpdates.map(u => ({
              range: `${TAB_NAME}!${colLetter(CI.sts)}${u.rowIndex + 1}`,
              values: [[`${statusMap[u.tracking]} · ${updatedAt}`]],
            })),
          },
        });
      } catch (err) { console.error('Sheet write error:', err.message); }
    }

    // 5. Return rows with updated statuses
    const resultRows = toProcess.map(r => ({
      ...r,
      deliveryStatus: statusMap[r.tracking]
        ? `${statusMap[r.tracking]} · ${updatedAt}`
        : r.currentSts,
    }));

    res.json({ rows: resultRows, updated: sheetUpdates.length });
  } catch (e) {
    console.error('/api/fedex-status:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/contact', async (req, res) => {
  const { name, email, message } = req.body;
  if (!name || !message) return res.status(400).json({ error: 'Name and message are required.' });

  try {
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
      },
    });

    await transporter.sendMail({
      from: `"305 Workspace" <${process.env.EMAIL_USER}>`,
      to: 'support@creativetwotwelve.com',
      replyTo: email || process.env.EMAIL_USER,
      subject: `Message from ${name} — 305 Workspace`,
      text: `From: ${name}\nEmail: ${email || 'not provided'}\n\n${message}`,
      html: `<p><strong>From:</strong> ${name}</p><p><strong>Email:</strong> ${email || 'not provided'}</p><br/><p>${message.replace(/\n/g, '<br/>')}</p>`,
    });

    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});


const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n  305 WORKSPACE TEAM`);
  console.log(`  ─────────────────────────────`);
  console.log(`  http://localhost:${PORT}\n`);
});
