require('dotenv').config();
const express = require('express');
const path = require('path');
const nodemailer = require('nodemailer');

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

async function triggerJhonnyUpdate(req, res) {
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
}

app.post('/api/run-jhonny', triggerJhonnyUpdate);
app.post('/api/fedex/update', triggerJhonnyUpdate);

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
