require('dotenv').config();
const express  = require('express');
const path     = require('path');
const fs       = require('fs');
const nodemailer = require('nodemailer');
const XLSX     = require('xlsx');
const { google } = require('googleapis');
const session  = require('express-session');

const suppliers  = require('./contacts/suppliers');
const team305    = require('./contacts/team305');
const TEAM_USERS = require('./contacts/users');

const app = express();
app.use(express.json());
app.use(session({
  secret: process.env.SESSION_SECRET || '305secret',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 24 * 60 * 60 * 1000 },
}));
app.use(express.static(path.join(__dirname, 'public')));

// ─── Constants ────────────────────────────────────────────────────────────────
const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
const REPO         = 'frazeved/SAMANTHA';
const SHEET_ID     = '1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q';
const SHEET_BASE   = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`;
const csvUrl       = (gid) => `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${gid}`;
const REDIRECT_URI = 'https://workspace305team.onrender.com/auth/callback';

const SIZE_MAP = {
  "3700":"XXS","3740":"XXS PETITE","4000":"XS","3701":"XS PETITE",
  "5000":"S","3702":"S PETITE","6000":"M","3703":"M PETITE",
  "7000":"L","3704":"L PETITE","8000":"XL","3705":"XL PETITE","9000":"XXL",
};

const WORKFLOWS = [
  { id: 'po-detail.yml',   name: 'PO DETAIL',  tab: 'PO DETAIL',  sheetUrl: `${SHEET_BASE}?gid=2017761959#gid=2017761959` },
  { id: 'all-pos.yml',     name: 'ALL POs',     tab: 'PO TRADE',   sheetUrl: `${SHEET_BASE}?gid=890202899#gid=890202899`   },
  { id: 'po-headers.yml',  name: 'PO HEADERS',  tab: 'PO NEW',     sheetUrl: `${SHEET_BASE}?gid=406184613#gid=406184613`   },
  { id: 'po-invoiced.yml', name: 'PO INVOICED', tab: 'PO INVOICE', sheetUrl: `${SHEET_BASE}?gid=1379989532#gid=1379989532` },
];

// ─── Token store ──────────────────────────────────────────────────────────────
const TOKEN_FILE = path.join(__dirname, 'tokens.json');
let userTokens = {};

(function loadTokens() {
  try { userTokens = JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8')); } catch (_) {}
  // Env vars override file — useful for persisting across Render deploys
  for (const [k, v] of Object.entries(process.env)) {
    if (k.startsWith('GMAIL_TOKEN_') && v) {
      const id = k.slice(12).toLowerCase();
      userTokens[id] = { ...(userTokens[id] || {}), refreshToken: v };
    }
  }
})();

function saveTokens() {
  try { fs.writeFileSync(TOKEN_FILE, JSON.stringify(userTokens, null, 2)); } catch (_) {}
}

function makeOAuth2Client() {
  return new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    REDIRECT_URI
  );
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
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

function parseCSV(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  const rows = []; let row = [], field = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"' && text[i+1] === '"') { field += '"'; i++; }
      else if (c === '"') inQ = false;
      else field += c;
    } else {
      if (c === '"') inQ = true;
      else if (c === ',') { row.push(field); field = ''; }
      else if (c === '\n' || (c === '\r' && text[i+1] === '\n')) {
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

// Build a raw MIME message buffer using nodemailer (no actual send)
function buildRawMime(mailOptions) {
  return new Promise((resolve, reject) => {
    const t = nodemailer.createTransport({ streamTransport: true, newline: 'unix', buffer: true });
    t.sendMail(mailOptions, (err, info) => err ? reject(err) : resolve(info.message));
  });
}

// ─── Auth routes ──────────────────────────────────────────────────────────────
app.get('/setup', (req, res) => res.sendFile(path.join(__dirname, 'public', 'setup.html')));

app.get('/api/users', (req, res) => {
  res.json(TEAM_USERS.map(u => ({
    ...u,
    connected: !!(userTokens[u.id]?.refreshToken),
  })));
});

app.get('/auth/google', (req, res) => {
  const { userId } = req.query;
  const user = TEAM_USERS.find(u => u.id === userId);
  if (!user) return res.status(400).send('Unknown user');
  req.session.pendingUserId = userId;
  const url = makeOAuth2Client().generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: ['https://www.googleapis.com/auth/gmail.compose'],
    login_hint: user.email,
  });
  res.redirect(url);
});

app.get('/auth/callback', async (req, res) => {
  const { code, error } = req.query;
  if (error) return res.redirect('/setup?error=denied');
  const userId = req.session.pendingUserId;
  if (!code || !userId) return res.redirect('/setup?error=session');
  try {
    const client = makeOAuth2Client();
    const { tokens } = await client.getToken(code);
    if (!tokens.refresh_token) return res.redirect('/setup?error=norefresh');
    userTokens[userId] = { refreshToken: tokens.refresh_token };
    saveTokens();
    res.redirect(`/setup?success=${userId}`);
  } catch (e) {
    console.error('auth callback error:', e.message);
    res.redirect('/setup?error=failed');
  }
});

// ─── Existing routes ──────────────────────────────────────────────────────────
app.get('/api/status', async (req, res) => {
  try {
    const r = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/all-extractions.yml/runs?per_page=1`);
    const data = await r.json();
    const run = data.workflow_runs?.[0] || null;
    const duration = run && run.status === 'completed' && run.run_started_at && run.updated_at
      ? Math.round((new Date(run.updated_at) - new Date(run.run_started_at)) / 1000) : null;
    res.json(WORKFLOWS.map(wf => ({
      id: wf.id, name: wf.name, tab: wf.tab, sheetUrl: wf.sheetUrl,
      status: run?.status || 'unknown', conclusion: run?.conclusion || null,
      started_at: run?.run_started_at || null, updated_at: run?.updated_at || null, duration,
    })));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/run-all', async (req, res) => {
  try {
    const r = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/all-extractions.yml/dispatches`, {
      method: 'POST', body: JSON.stringify({ ref: 'main' }),
    });
    if (r.status !== 204) { const b = await r.text(); return res.status(500).json({ error: `GitHub returned ${r.status}: ${b}` }); }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

async function triggerJhonnyUpdate(req, res) {
  try {
    const r = await ghFetch(`https://api.github.com/repos/frazeved/JHONNY/actions/workflows/update-fedex-status.yml/dispatches`, {
      method: 'POST', body: JSON.stringify({ ref: 'main' }),
    });
    if (r.status !== 204) { const b = await r.text(); return res.status(500).json({ error: `GitHub returned ${r.status}: ${b}` }); }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
}
app.post('/api/run-jhonny', triggerJhonnyUpdate);
app.post('/api/fedex/update', triggerJhonnyUpdate);

app.post('/api/po/search', async (req, res) => {
  try {
    const { style } = req.body;
    if (!style?.trim()) return res.status(400).json({ error: 'Style # is required' });
    const styleToSearch = style.toString().trim().toLowerCase();
    const response = await fetch(csvUrl(0));
    if (!response.ok) return res.status(500).json({ error: 'Failed to fetch sheet data' });
    const rows = parseCSV(await response.text());
    if (rows.length < 2) return res.status(500).json({ error: 'No data in sheet' });
    const headers = rows[0].map(h => h.trim().toLowerCase());
    const col = (...kws) => {
      for (const kw of kws) { const i = headers.findIndex(h => h === kw.toLowerCase()); if (i >= 0) return i; }
      for (const kw of kws) { const i = headers.findIndex(h => h.includes(kw.toLowerCase())); if (i >= 0) return i; }
      return -1;
    };
    const C = {
      style: col('style #','style#','style'), status: col('status'), supplier: col('supplier'),
      category: col('category'), subcategory: col('sub-category','subcategory'),
      piReceived: col('buyer presentation','pi received','pi date'),
      topSent: col('sms sent to anthro','top sent to anthro','sms sent'),
      topSampleStatus: col('top approval from anthro','top sample status','top approval'),
      basePo: col('po info anthro','base po','base po import farm'),
      finalNDC: col('final ndc','ndc'), topDeadline: col('sms deadline','top deadline'),
      exFactory: col('ex factory / flight date','ex factory','flight date'),
      cost: col('cost'), freight: col('freight'), duty: col('duty'), hts: col('hts code','hts'),
    };
    const get = (r, i) => i >= 0 ? (r[i] || '').trim() : '';
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (get(r, C.style).toLowerCase() !== styleToSearch) continue;
      if (!get(r, C.status).toLowerCase().includes("po'd")) continue;
      const finalNDC = get(r, C.finalNDC), topDeadline = get(r, C.topDeadline);
      let topDeadlineDays = '';
      if (finalNDC && topDeadline) { const a = new Date(finalNDC), b = new Date(topDeadline); if (!isNaN(a)&&!isNaN(b)) topDeadlineDays = Math.round((a-b)/86400000).toString(); }
      return res.json({ style: get(r,C.style), supplier: get(r,C.supplier), category: get(r,C.category), subcategory: get(r,C.subcategory), piReceived: get(r,C.piReceived), topSent: get(r,C.topSent), topSampleStatus: get(r,C.topSampleStatus), basePo: get(r,C.basePo), finalNDC, topDeadline, topDeadlineDays, exFactory: get(r,C.exFactory), cost: get(r,C.cost), freight: get(r,C.freight), duty: get(r,C.duty), hts: get(r,C.hts) });
    }
    res.status(404).json({ error: "Style not found or not in PO'd status" });
  } catch (e) { console.error('PO search error:', e.message); res.status(500).json({ error: e.message }); }
});

// ─── Breakdown Email — creates a Gmail draft ──────────────────────────────────
app.post('/api/po/breakdown-email', async (req, res) => {
  try {
    const { style, supplier, cost, hts, freight, exFactory, message, sendingAs } = req.body;
    if (!style?.trim())    return res.status(400).json({ error: 'Style # is required' });
    if (!supplier?.trim()) return res.status(400).json({ error: 'Supplier is required' });
    if (!sendingAs)        return res.status(400).json({ error: 'Please select who is sending this' });

    const token = userTokens[sendingAs];
    if (!token?.refreshToken) {
      const user = TEAM_USERS.find(u => u.id === sendingAs);
      return res.status(401).json({ error: `${user?.name || sendingAs} has not connected their Gmail yet. Please visit /setup first.` });
    }

    const sender = TEAM_USERS.find(u => u.id === sendingAs);

    // Normalize style
    const rawStyle    = style.trim();
    const displayStyle = /^[A-Za-z]-/.test(rawStyle) ? rawStyle.split('-').slice(1).join('-').trim() : rawStyle;
    const cleanStyle  = displayStyle.toUpperCase().replace(/\s+/g, '');
    const supplierKey = supplier.trim().toUpperCase();

    // Format dates
    const fmtDate = iso => { if (!iso) return 'xxxxxx'; const d = new Date(iso); return isNaN(d) ? 'xxxxxx' : d.toLocaleDateString('en-US',{month:'short',day:'numeric'}); };
    const handoverFormatted = fmtDate(exFactory);
    const invoiceFormatted  = exFactory ? fmtDate(new Date(new Date(exFactory).getTime()-2*86400000).toISOString().split('T')[0]) : 'xxxxxx';

    // Fetch sheets
    const [pdRes, ptRes] = await Promise.all([fetch(csvUrl(2017761959)), fetch(csvUrl(890202899))]);
    if (!pdRes.ok) throw new Error('Could not fetch PO DETAIL sheet');
    if (!ptRes.ok) throw new Error('Could not fetch PO TRADE sheet');
    const pdRows = parseCSV(await pdRes.text());
    const ptRows = parseCSV(await ptRes.text());

    const channelMap = {};
    for (let i = 1; i < ptRows.length; i++) { const po = (ptRows[i][0]||'').trim(); if (po) channelMap[po] = (ptRows[i][1]||'').trim(); }

    const H = pdRows[0] || [];
    const findCol = (...kws) => {
      for (const kw of kws) { const i = H.findIndex(h=>h.trim().toLowerCase()===kw); if(i>=0) return i; }
      for (const kw of kws) { const i = H.findIndex(h=>h.trim().toLowerCase().includes(kw)); if(i>=0) return i; }
      return -1;
    };
    const C = {
      po: findCol('po#','purchase order','po'), style: findCol('vendor style','style#','style #','style'),
      shortSku: findCol('short sku'), size: findCol('size desc','size'),
      packType: findCol('ship pack','pack type','pack'), sizeCode: findCol('size code'),
      qty: findCol('total qty','qty'), units: findCol('allocated','prepack','total units'),
      itemNum: findCol('long sku','item number','sku'), color: findCol('vendor color','color'),
    };
    const get = (r,i) => i>=0 ? (r[i]||'').trim() : '';

    const groupedByPO = {};
    for (let i=1; i<pdRows.length; i++) {
      const r = pdRows[i];
      if (get(r,C.style).toUpperCase().replace(/\s+/g,'') !== cleanStyle) continue;
      const po = get(r,C.po); if (!po) continue;
      (groupedByPO[po] = groupedByPO[po]||[]).push(r);
    }

    const poLines=[], poSubject=[];
    for (const po in groupedByPO) { const ch=channelMap[po]||''; poLines.push(`<b>PO# ${po} - ${ch}</b>`); poSubject.push(`#${po}`); }
    if (!poLines.length) { poLines.push('<span style="color:blue;"><b>PO# xxxxxx - xxxxx</b></span>'); poSubject.push('#xxxxxx'); }

    const gcd=(a,b)=>{a=Math.abs(a);b=Math.abs(b);while(b){const t=b;b=a%b;a=t;}return a;};
    const gcdArr=ns=>{let g=0;for(const n of ns){if(n)g=g?gcd(g,n):Math.abs(n);}return g||1;};

    const EXCEL_HEADERS=['PO#','Style#','Type','Pack Type','Item Number','Vendor Color','Size code','Size','Short SKU','Qty','RATIO','Total Units'];
    const allRows=[]; let grandQty=0,grandPack=0,grandUnits=0;
    for (const po in groupedByPO) {
      const rows=groupedByPO[po];
      const ppkUnits=rows.filter(r=>get(r,C.packType).toUpperCase()==='PPK').map(r=>Number(get(r,C.units))||0).filter(v=>v>0);
      const gcdUnits=gcdArr(ppkUnits);
      let poQty=0,poPack=0,poTotal=0;
      for (const r of rows) {
        const pack=get(r,C.packType).toUpperCase(), qty=Number(get(r,C.qty))||0, units=Number(get(r,C.units))||0;
        const ratio=pack==='PPK'?Math.round((units/(gcdUnits||1))*1e6)/1e6:1, sc=get(r,C.sizeCode);
        allRows.push([get(r,C.po),get(r,C.style),channelMap[po]||'',pack,get(r,C.itemNum),get(r,C.color),sc,SIZE_MAP[sc]||get(r,C.size),get(r,C.shortSku),qty,ratio,units]);
        poQty+=qty; poPack+=ratio; poTotal+=units;
      }
      allRows.push(['','','','','','','','TOTAL','',poQty,poPack,poTotal]);
      grandQty+=poQty; grandPack+=poPack; grandUnits+=poTotal;
    }
    allRows.push(Array(12).fill(''));
    allRows.push(['','','','','','','','GRAND TOTAL','',grandQty,grandPack,grandUnits]);

    const wb=XLSX.utils.book_new(), ws=XLSX.utils.aoa_to_sheet([EXCEL_HEADERS,...allRows]);
    XLSX.utils.book_append_sheet(wb,ws,'Breakdown');
    const excelBuf=XLSX.write(wb,{type:'buffer',bookType:'xlsx'});

    const toEmails  = suppliers.emails[supplierKey]      || ['logistics@creativetwotwelve.com'];
    const greetName = suppliers.mainContact[supplierKey] || supplierKey;
    const subject   = `BREAKDOWN STYLE# ${displayStyle} - PO ${poSubject.join(' ')} - ${supplierKey}`;

    const htmlBody = `
<p><b><span style="font-size:12pt;">Hi ${greetName} and ${supplierKey} team,</span></b></p>
<p>Please find attached the breakdown of<br><b>STYLE# ${displayStyle}</b></p>
<p><b>HTS# ${hts||'xxxxxx'}</b> | <b>${freight||'xxxxxx'} $${cost||'xxxxxx'}</b> | <b>INVOICE/PACKING LIST WITH FLAVIO: ${invoiceFormatted}</b> | <b>AGREED HANDOVER DATE: ${handoverFormatted}</b></p>
<p>${poLines.join('<br>')}</p>
<p>Could you please confirm this style fabric composition?</p>
<p><b><span style="color:red;">IMPORTANT:</span></b><br>Please, send the Invoice and Packing list before shipping for validation, also the custom description, HTS#, TAX ID on it. AWB when available.</p>
<p>Any delay or new agreed ship date on this Style#, please answer this email chain immediately!</p>
${message?`<p>${message}</p>`:''}
<p>Best,<br>${sender?.name||sendingAs}<br>Production Team<br>305 CONSULTING AND PRODUCTION</p>`;

    // Build raw MIME (nodemailer composes, does NOT send)
    const rawMime = await buildRawMime({
      from: `"${sender?.name||sendingAs}" <${sender?.email||''}>`,
      to:   toEmails.join(','),
      cc:   team305.breakdownCC.join(','),
      subject,
      html: htmlBody,
      attachments: [{ filename:`BREAKDOWN STYLE# ${cleanStyle}.xlsx`, content:excelBuf, contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }],
    });
    const encoded = rawMime.toString('base64').replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');

    // Create draft in the sender's Gmail
    const authClient = makeOAuth2Client();
    authClient.setCredentials({ refresh_token: token.refreshToken });
    const gmail = google.gmail({ version:'v1', auth:authClient });
    await gmail.users.drafts.create({ userId:'me', requestBody:{ message:{ raw:encoded } } });

    res.json({ ok:true, message:'Draft created in your Gmail Drafts folder' });
  } catch (e) {
    console.error('breakdown-email error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/contact', async (req, res) => {
  const { name, email, message } = req.body;
  if (!name || !message) return res.status(400).json({ error: 'Name and message are required.' });
  try {
    const transporter = nodemailer.createTransport({ service:'gmail', auth:{ user:process.env.EMAIL_USER, pass:process.env.EMAIL_PASS } });
    await transporter.sendMail({
      from: `"305 Workspace" <${process.env.EMAIL_USER}>`,
      to: 'support@creativetwotwelve.com',
      replyTo: email || process.env.EMAIL_USER,
      subject: `Message from ${name} — 305 Workspace`,
      text: `From: ${name}\nEmail: ${email||'not provided'}\n\n${message}`,
      html: `<p><strong>From:</strong> ${name}</p><p><strong>Email:</strong> ${email||'not provided'}</p><br/><p>${message.replace(/\n/g,'<br/>')}</p>`,
    });
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n  305 WORKSPACE TEAM`);
  console.log(`  ─────────────────────────────`);
  console.log(`  http://localhost:${PORT}\n`);
});
