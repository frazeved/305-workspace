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
const MAP_SHEET_ID = '1W88MKYr-q9g3F2fLFu2jjxvXzigK12PWohSVMsOQst4';
const SHEET_BASE   = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`;
const csvUrl       = (gid) => `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${gid}`;
const REDIRECT_URI = 'https://workspace305team.onrender.com/auth/callback';

const SIZE_MAP = {
  "3700":"XXS","3740":"XXS PETITE","4000":"XS","3701":"XS PETITE",
  "5000":"S","3702":"S PETITE","6000":"M","3703":"M PETITE",
  "7000":"L","3704":"L PETITE","8000":"XL","3705":"XL PETITE","9000":"XXL",
};

const FEDEX_API_URL      = 'https://apis.fedex.com';
const FEDEX_DRIVE_FOLDER = '1ufkdrO23m2C-MrmhR1iKN3QFSJFwuQPY';
const FEDEX_SERVICE_NAMES = {
  FIRST_OVERNIGHT:         'FedEx First Overnight',
  FEDEX_2_DAY_AM:          'FedEx 2Day AM',
  FEDEX_3_DAY_FREIGHT:     'FedEx 3Day Freight',
  FEDEX_EXPRESS_SAVER:     'FedEx Express Saver',
  PRIORITY_OVERNIGHT:      'FedEx Priority Overnight',
  FEDEX_2_DAY:             'FedEx 2Day',
  FEDEX_1_DAY_FREIGHT:     'FedEx 1Day Freight',
  FEDEX_2_DAY_FREIGHT:     'FedEx 2Day Freight',
  STANDARD_OVERNIGHT:      'FedEx Standard Overnight',
  FEDEX_GROUND:            'FedEx Ground',
  FIRST_OVERNIGHT_FREIGHT: 'FedEx First Overnight Freight',
};
const FEDEX_ADDRESS_BOOK = {
  SHIP_FROM: {
    company: '305 Consulting and Production Inc',
    contact: '305 Consulting and Production Inc',
    phone:   '9174992103',
    street:  ['1800 NW 15TH AVENUE, STE 110', 'UNIT 2 GROUND'],
    city:    'POMPANO BEACH', state: 'FL', zip: '33069', country: 'US',
  },
  BRO: { company: 'URBAN OUTFITTERS INC',   contact: 'BRISTOL RENTAL DC',           phone: '', street: ['2401 GREEN LN'],                             city: 'LEVITTOWN',    state: 'PA', zip: '19057', country: 'US' },
  GAP: { company: 'ANTHROPOLOGIE GAP',      contact: 'URBN GAP DC',                 phone: '', street: ['755 BRACKBILL ROAD'],                        city: 'GAP',          state: 'PA', zip: '17527', country: 'US' },
  GFC: { company: 'ANTHROPOLOGIE',          contact: 'URBN GAP FULFILLMENT CENTER', phone: '', street: ['766 BRACKBILL ROAD'],                        city: 'GAP',          state: 'PA', zip: '17527', country: 'US' },
  KC1: { company: 'URBAN OUTFITTERS INC',   contact: 'KANSAS CITY KANSAS FC',       phone: '', street: ['11681 STATE AVE'],                           city: 'KANSAS CITY',  state: 'KS', zip: '66111', country: 'US' },
  KC3: { company: 'URBAN OUTFITTERS INC',   contact: 'NUULY RAYMORE RENTAL DC',     phone: '', street: ['1300 S. DEAN AVE', 'BUILDING 3, SUITE 100'], city: 'RAYMORE',      state: 'MO', zip: '64083', country: 'US' },
  REN: { company: 'URBN RENO DC',           contact: 'URBN RENO DC',                phone: '', street: ['6640 ECHO AVE'],                             city: 'RENO',         state: 'NV', zip: '89506', country: 'US' },
  RNO: { company: 'ANTHROPOLOGIE',          contact: 'URBN WEST COAST FULFILLMENT', phone: '', street: ['12055 MOYA BLVD'],                           city: 'RENO',         state: 'NV', zip: '89506', country: 'US' },
  YRD: { company: 'NUULY NAVY YARD',        contact: 'NUULY NAVY YARD',             phone: '', street: ['5000 SOUTH BROAD ST'],                       city: 'PHILADELPHIA', state: 'PA', zip: '19112', country: 'US' },
};
const FEDEX_PACKAGING = {
  'ECICO':      { l: 20, w: 12, h: 12 },
  'GAIA':       { l: 20, w: 12, h: 12 },
  'H&F':        { l: 23, w: 12, h: 12 },
  'HS FASHION': { l: 20, w: 15, h: 11 },
  'KONCEPTION': { l: 20, w: 12, h: 12 },
  'MINI BOX':   { l: 12, w: 12, h: 12 },
  'PQSWIM':     { l: 20, w: 12, h: 12 },
  'S&S':        { l: 23, w: 16, h: 11 },
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

// ─── FedEx Validation Email Draft ────────────────────────────────────────────
app.post('/api/fedex/validation-draft', async (req, res) => {
  try {
    const token = userTokens['flavio'];
    if (!token?.refreshToken) return res.status(401).json({ error: 'flavio Gmail not connected. Visit /setup first.' });

    const htmlBody = `
<p>Hi FedEx Label Analysis Team,</p>
<p>Please find attached our test label for Ship API production validation.</p>
<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;">
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">Company</td><td>305 Consulting and Production Inc</td></tr>
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">FedEx Account Number</td><td>740561073</td></tr>
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">Production API Key</td><td>l766161b6e947e4d08b5284266e7afcee8</td></tr>
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">Contact</td><td>support@creativetwotwelve.com | 917-499-2103</td></tr>
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">Services Requested</td><td>FedEx Ground (domestic US)</td></tr>
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">Application</td><td>Internal warehouse shipping system — 305 Workspace</td></tr>
</table>
<br>
<p>Attached: scanned test label printed at 600 DPI.</p>
<p>Please advise on approval.</p>
<p>Thank you,<br>Flavio Azevedo<br>305 Consulting and Production Inc</p>`;

    const rawMime = await buildRawMime({
      from:    '"Flavio Azevedo" <support@creativetwotwelve.com>',
      to:      'label@fedex.com',
      subject: 'FedEx Ship API Label Validation — 305 Consulting and Production Inc',
      html:    htmlBody,
    });
    const encoded = rawMime.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

    const authClient = makeOAuth2Client();
    authClient.setCredentials({ refresh_token: token.refreshToken });
    const gmail = google.gmail({ version: 'v1', auth: authClient });
    await gmail.users.drafts.create({ userId: 'me', requestBody: { message: { raw: encoded } } });

    res.json({ ok: true, message: 'Draft saved to support@creativetwotwelve.com — check Gmail Drafts' });
  } catch (e) {
    console.error('validation-draft error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── FedEx Label Creation ─────────────────────────────────────────────────────
function colToLetter(col) {
  let letter = '', n = col + 1;
  while (n > 0) { const r = (n - 1) % 26; letter = String.fromCharCode(65 + r) + letter; n = Math.floor((n - 1) / 26); }
  return letter;
}

app.post('/api/fedex/create-label', async (req, res) => {
  try {
    const { po, boxQty, weight, serviceCode, deliverToCode, supplier } = req.body;
    if (!po || !boxQty || !serviceCode || !deliverToCode)
      return res.status(400).json({ error: 'Missing required fields: po, boxQty, serviceCode, deliverToCode' });

    const to = FEDEX_ADDRESS_BOOK[deliverToCode];
    if (!to) return res.status(400).json({ error: `Unknown deliver-to code: ${deliverToCode}` });
    const from = FEDEX_ADDRESS_BOOK.SHIP_FROM;

    const nk = s => (s || '').toLowerCase().trim();
    const dimsEntry = Object.entries(FEDEX_PACKAGING).find(([k]) =>
      nk(k) === nk(supplier) || nk(supplier).includes(nk(k)) || nk(k).includes(nk(supplier))
    );
    const dims = dimsEntry ? dimsEntry[1] : { l: 20, w: 12, h: 12 };
    const qty  = Math.max(1, parseInt(boxQty) || 1);
    const wt   = parseFloat(weight) || 20;

    // 1. FedEx OAuth token
    const tokenRes = await fetch(`${FEDEX_API_URL}/oauth/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: `grant_type=client_credentials&client_id=${encodeURIComponent(process.env.FEDEX_SHIP_API_KEY)}&client_secret=${encodeURIComponent(process.env.FEDEX_SHIP_API_SECRET)}`,
    });
    if (!tokenRes.ok) {
      const t = await tokenRes.text();
      return res.status(500).json({ error: `FedEx auth failed (${tokenRes.status}): ${t}` });
    }
    const { access_token } = await tokenRes.json();

    // 2. Build payload
    const packages = Array.from({ length: qty }, (_, i) => ({
      sequenceNumber: i + 1,
      weight:     { units: 'LB', value: wt },
      dimensions: { length: dims.l, width: dims.w, height: dims.h, units: 'IN' },
    }));

    const payload = {
      labelResponseOptions: 'URL_ONLY',
      requestedShipment: {
        shipper: {
          contact: { personName: from.contact, companyName: from.company, phoneNumber: from.phone },
          address: { streetLines: from.street, city: from.city, stateOrProvinceCode: from.state, postalCode: from.zip, countryCode: from.country },
        },
        recipients: [{
          contact: { personName: to.contact, companyName: to.company, phoneNumber: to.phone || '0000000000' },
          address: { streetLines: to.street, city: to.city, stateOrProvinceCode: to.state, postalCode: to.zip, countryCode: to.country },
        }],
        serviceType: serviceCode,
        packagingType: 'YOUR_PACKAGING',
        pickupType: 'USE_SCHEDULED_PICKUP',
        shippingChargesPayment: {
          paymentType: 'SENDER',
          payor: { responsibleParty: { accountNumber: { value: process.env.FEDEX_ACCOUNT_NUMBER } } },
        },
        labelSpecification: { labelFormatType: 'COMMON2D', imageType: 'PDF', labelStockType: 'PAPER_4X6' },
        customerReferences: [{ customerReferenceType: 'CUSTOMER_REFERENCE', value: po }],
        requestedPackageLineItems: packages,
      },
      accountNumber: { value: process.env.FEDEX_ACCOUNT_NUMBER },
    };

    // 3. Create shipment
    const shipRes = await fetch(`${FEDEX_API_URL}/ship/v1/shipments`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${access_token}`, 'Content-Type': 'application/json', 'x-locale': 'en_US' },
      body: JSON.stringify(payload),
    });
    const shipData = await shipRes.json();
    if (!shipRes.ok) {
      const errMsg = shipData?.errors?.[0]?.message || JSON.stringify(shipData).slice(0, 300);
      return res.status(500).json({ error: `FedEx Ship error: ${errMsg}`, raw: shipData });
    }

    const shipment    = shipData.output?.transactionShipments?.[0];
    if (!shipment) return res.status(500).json({ error: 'No shipment in FedEx response', raw: shipData });

    const tracking    = shipment.masterTrackingNumber || shipment.pieceResponses?.[0]?.trackingNumber || '';
    const allLabels   = (shipment.pieceResponses || []).flatMap(p => (p.packageDocuments || []).map(d => d.url).filter(Boolean));
    const rateDetail  = shipment.completedShipmentDetail?.shipmentRating?.shipmentRateDetails;
    const totalCharge = rateDetail?.[0]?.totalNetCharge?.amount ?? rateDetail?.[0]?.totalNetCharge ?? '';
    const serviceName = FEDEX_SERVICE_NAMES[serviceCode] || serviceCode;

    // 4. Upload labels to Google Drive (best-effort)
    let driveWebUrl = '';
    if (allLabels.length && process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
      try {
        const sa   = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
        const auth = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/drive'] });
        const drive = google.drive({ version: 'v3', auth });
        const MONTHS = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];
        let folderId = FEDEX_DRIVE_FOLDER;
        try {
          const listRes = await drive.files.list({
            q: `'${FEDEX_DRIVE_FOLDER}' in parents and mimeType='application/vnd.google-apps.folder' and name='${MONTHS[new Date().getMonth()]}' and trashed=false`,
            fields: 'files(id)', pageSize: 1,
          });
          if (listRes.data.files?.length) folderId = listRes.data.files[0].id;
        } catch (_) {}

        for (let i = 0; i < allLabels.length; i++) {
          const pdfRes = await fetch(allLabels[i]);
          if (!pdfRes.ok) continue;
          const buf = Buffer.from(await pdfRes.arrayBuffer());
          const { Readable } = require('stream');
          const fileName = allLabels.length === 1 ? `FEDEX ${po}` : `FEDEX ${po} - ${i + 1}`;
          const up = await drive.files.create({
            requestBody: { name: fileName, mimeType: 'application/pdf', parents: [folderId] },
            media: { mimeType: 'application/pdf', body: Readable.from(buf) },
            fields: 'id,webViewLink',
          });
          if (i === 0) driveWebUrl = up.data.webViewLink || '';
        }
      } catch (driveErr) { console.error('[Drive upload]', driveErr.message); }
    }

    // 5. Write back to Warehouse Now sheet (best-effort)
    if (tracking && process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
      try {
        const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
        const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
        const sheets = google.sheets({ version: 'v4', auth });
        const TAB    = 'Warehouse Now Database';

        const hdRes = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${TAB}'!1:1` });
        const hdrs  = (hdRes.data.values?.[0] || []).map(h => h.trim().toLowerCase());
        const fc = (...kws) => {
          for (const kw of kws) { const i = hdrs.findIndex(h => h === kw.toLowerCase()); if (i >= 0) return i; }
          for (const kw of kws) { const i = hdrs.findIndex(h => h.includes(kw.toLowerCase())); if (i >= 0) return i; }
          return -1;
        };
        const trackCol   = fc('tracking number', 'tracking #', 'tracking');
        const carrierCol = fc('carrier');
        const serviceCol = fc('shipping type', 'service type', 'ship type');
        const costCol    = fc('shipping cost', 'freight cost', 'ship cost');
        const poColIdx   = fc('po#', 'po number', 'po');

        if (poColIdx >= 0 && trackCol >= 0) {
          const colL   = colToLetter(poColIdx);
          const poVals = (await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${TAB}'!${colL}:${colL}` })).data.values || [];
          let sheetRow = -1;
          for (let i = 1; i < poVals.length; i++) {
            if ((poVals[i]?.[0] || '').toString().trim() === po.toString().trim()) { sheetRow = i + 1; break; }
          }
          if (sheetRow > 0) {
            const updates = [
              ...(trackCol   >= 0 ? [{ range: `'${TAB}'!${colToLetter(trackCol)}${sheetRow}`,   values: [[tracking]]          }] : []),
              ...(carrierCol >= 0 ? [{ range: `'${TAB}'!${colToLetter(carrierCol)}${sheetRow}`,  values: [['FEDEX']]            }] : []),
              ...(serviceCol >= 0 ? [{ range: `'${TAB}'!${colToLetter(serviceCol)}${sheetRow}`,  values: [[serviceName]]        }] : []),
              ...(costCol    >= 0 ? [{ range: `'${TAB}'!${colToLetter(costCol)}${sheetRow}`,     values: [[String(totalCharge)]]}] : []),
            ];
            if (updates.length) await sheets.spreadsheets.values.batchUpdate({
              spreadsheetId: SHEET_ID,
              requestBody: { valueInputOption: 'RAW', data: updates },
            });
          }
        }
      } catch (sheetErr) { console.error('[Sheet write]', sheetErr.message); }
    }

    res.json({ ok: true, tracking, cost: totalCharge ? `$${totalCharge}` : '', driveUrl: driveWebUrl });
  } catch (e) {
    console.error('[create-label]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Gabriel: MAP DATA (chart feed) ──────────────────────────────────────────
app.get('/api/gabriel/map-data', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    return res.status(500).json({ error: 'Google credentials not configured' });
  }
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: MAP_SHEET_ID,
      range: `'ANTHRO MAP 2026'`,
      valueRenderOption: 'FORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING',
    });
    const values = r.data.values || [];
    res.setHeader('Cache-Control', 'no-store');
    res.json({ values });
  } catch (e) {
    console.error('[map-data]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Samantha: PowerBI Database Sync ─────────────────────────────────────────
app.post('/api/samantha/powerbi-sync', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    return res.status(500).json({ error: 'Google credentials not configured' });
  }
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });

    const readTab = async (tab, formula = false) => {
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID, range: `'${tab}'`,
        valueRenderOption: formula ? 'FORMULA' : 'UNFORMATTED_VALUE',
        dateTimeRenderOption: 'FORMATTED_STRING',
      });
      return r.data.values || [];
    };
    const isFormulaPB = v => typeof v === 'string' && v.startsWith('=');

    const [tsData, poNewData, poInvData] = await Promise.all([
      readTab('TRADESTONE DATABASE', true), // FORMULA mode — preserves existing formulas
      readTab('PO NEW'), readTab('PO INVOICE'),
    ]);

    const tsH = tsData[0] || [], poNewH = poNewData[0] || [], poInvH = poInvData[0] || [];
    const hMap = H => { const m = {}; H.forEach((h,i) => { const k = String(h).trim().toLowerCase(); if (k) m[k]=i; }); return m; };
    const tsHM = hMap(tsH), pnHM = hMap(poNewH), piHM = hMap(poInvH);

    const tsPO  = tsHM['po#'] ?? -1;
    const pnPO  = pnHM['po#'] ?? -1;
    const piPO  = piHM['po#'] ?? -1;
    const piDate  = piHM['invoice date']  ?? -1;
    const piValue = piHM['invoice value'] ?? piHM['invoice amount'] ?? -1;
    const tsShip  = tsHM['ship date'] ?? -1;
    const pnShip  = pnHM['ship date'] ?? -1;
    const pnCancel = pnHM['cancel date'] ?? -1;
    const tsIPcol  = tsHM['ip class'] ?? tsHM['ipclass'] ?? -1;

    // Columns that must not be overwritten from PO NEW (V=21, AL=37, AM=38)
    const SKIP = new Set([21, 37, 38]);
    const COL_AJ = 35, COL_AK = 36, COL_H = 7, COL_V = 21;

    const IP_CAT = {
      1501:'Blouses',1502:'SWTRS & SWTSHRTS',1504:'Pants',1505:'Skirts',1506:'SHORTS',
      1508:'Dresses',1509:'Dresses',1515:'Dresses',4110:'Blouses',4114:'Sweaters',
      4120:'Skirts',4123:'Pants',4124:'Jumpsuit',4125:'Shorts',4130:'Dresses',
      4134:'Dresses',4141:'LOUNGE',4142:'Swimwear',4148:'Blouses',4152:'Hats',
      4153:'Accessories',4157:'Wraps',4428:'Dresses',4915:'Waters Edge',
      4920:'Blouses',4925:'Skirts',4927:'Pants',4928:'Dresses',4942:'Swimwear',
    };

    // Column mapping: TRADESTONE col → PO NEW col (matched by header name, skip protected)
    const colMap = [];
    for (let ti = 0; ti < tsH.length; ti++) {
      if (SKIP.has(ti)) continue;
      const k = String(tsH[ti]).trim().toLowerCase();
      if (k && pnHM[k] !== undefined) colMap.push({ tsIdx: ti, poIdx: pnHM[k] });
    }

    // Existing TRADESTONE rows by PO
    const existByPO = new Map();
    for (let i = 1; i < tsData.length; i++) {
      const po = String(tsData[i][tsPO] || '').trim().toUpperCase();
      if (po) existByPO.set(po, i);
    }

    // Invoice data by PO
    const invByPO = new Map();
    for (let i = 1; i < poInvData.length; i++) {
      const po = String(poInvData[i][piPO] || '').trim();
      if (po) invByPO.set(po, {
        AJ: piValue >= 0 ? poInvData[i][piValue] : '',
        AK: piDate  >= 0 ? poInvData[i][piDate]  : '',
      });
    }

    const colH = row => { const ak = row[COL_AK]; return (ak != null && ak !== '') ? ak : (row[6] || ''); };

    // Step 1 — refresh AJ/AK invoice data on ALL existing rows
    for (const [po, inv] of invByPO) {
      const ei = existByPO.get(po.toUpperCase());
      if (ei === undefined) continue;
      const r = tsData[ei];
      if (inv.AJ != null && inv.AJ !== '') r[COL_AJ] = inv.AJ;
      if (inv.AK != null && inv.AK !== '') r[COL_AK] = inv.AK;
      r[COL_H] = colH(r);
    }

    // Step 2 — add only NEW POs (not already in TRADESTONE DATABASE)
    // Find formula column indices by header name
    const cat2Idx  = tsHM['category 2']  ?? tsHM['category2']  ?? -1;
    const cat3Idx  = tsHM['category 3']  ?? tsHM['category3']  ?? -1;
    const boxesIdx = tsHM['boxes']       ?? -1;
    const yr2Idx   = tsHM['year 2']      ?? tsHM['year2']       ?? -1;
    const mo2Idx   = tsHM['month 2']     ?? tsHM['month2']      ?? -1;

    // First data row = 2, last existing = tsData.length (since tsData[0]=header)
    const firstNewSheetRow = tsData.length + 1; // sheet row of first new row to be appended

    const newRows = [];
    for (let si = 1; si < poNewData.length; si++) {
      const pnRow = poNewData[si];
      const po  = String(pnRow[pnPO]    || '').trim();
      const st  = String(pnRow[pnHM['vendor style #'] ?? -1] || '').trim();
      if (!po || !st) continue;
      if (existByPO.has(po.toUpperCase())) continue; // already in TRADESTONE — skip

      const nr = Array(tsH.length).fill('');
      for (const {tsIdx, poIdx} of colMap) nr[tsIdx] = pnRow[poIdx];
      if (tsShip >= 0 && pnShip >= 0) nr[tsShip] = pnRow[pnShip];
      if (pnCancel >= 0) nr[COL_H] = pnRow[pnCancel];
      const inv = invByPO.get(po);
      if (inv) { nr[COL_AJ] = inv.AJ; nr[COL_AK] = inv.AK; }
      nr[COL_H] = colH(nr);
      if (tsIPcol >= 0) nr[COL_V] = IP_CAT[parseInt(nr[tsIPcol])] || 'OTHER';

      // Inject formulas with the exact sheet row number for this new row
      const r = firstNewSheetRow + newRows.length;
      if (cat2Idx  >= 0) nr[cat2Idx]  = `=IFNA(VLOOKUP(V${r},{"Dresses","Dresses";"Rompers","Dresses";"JUMPERS & ROMPERS","Dresses";"Blouses","Blouses";"BLOUSES & SHIRTS","Blouses";"SLEEP","Lounge";"Fine Gauge","Sweaters";"Sweaters","Sweaters";"SWTRS & SWTSHRTS","Sweaters";"Heavyweight","Knit";"Knit","Knit";"Pants","Bottoms";"PANTS & LEGGINGS","Bottoms";"Jumpsuit","Bottoms";"Swimwear","Swimwear";"Water''s Edge","Swimwear";"Wraps","Accessories";"Shorts","Shorts";"Skirts","Skirts"},2,0),"No Match")`;
      if (cat3Idx  >= 0) nr[cat3Idx]  = `=IFNA(VLOOKUP(V${r},{"Sleep","Lounge";"Blouses","Blouses";"BLOUSES & SHIRTS","Blouses";"Dresses","Dresses";"Fine Gauge","Sweaters";"Heavyweight","Knit";"JUMPERS & ROMPERS","Bottoms";"Jumpsuit","Bottoms";"Pants","Bottoms";"PANTS & LEGGINGS","Bottoms";"Rompers","Dresses";"Shorts","Skirts";"Skirts","Skirts";"Sweaters","Sweaters";"SWTRS & SWTSHRTS","Sweaters";"Swimwear","Swimwear";"Water''s Edge","Swimwear";"Wraps","Accessories"},2,0),"No Match")`;
      if (boxesIdx >= 0) nr[boxesIdx] = `=AF${r}/30`;
      if (yr2Idx   >= 0) nr[yr2Idx]   = `=YEAR(H${r})`;
      if (mo2Idx   >= 0) nr[mo2Idx]   = `=TEXT(H${r},"MM") & " - " & TEXT(H${r},"MMM")`;

      newRows.push(nr);
    }

    // Write AJ/AK/H updates — preserve any existing formula cells
    const ajVals = [], akVals = [], hVals = [];
    for (let i = 1; i < tsData.length; i++) {
      const orig = tsData[i];
      // If cell is a formula, write the formula back unchanged; otherwise write computed value
      ajVals.push([isFormulaPB(orig[COL_AJ]) ? orig[COL_AJ] : (orig[COL_AJ] ?? '')]);
      akVals.push([isFormulaPB(orig[COL_AK]) ? orig[COL_AK] : (orig[COL_AK] ?? '')]);
      hVals.push([isFormulaPB(orig[COL_H])  ? orig[COL_H]  : (orig[COL_H]  ?? '')]);
    }

    // Deduplicate newRows by PO# (in case PO NEW has duplicate entries)
    const seenPOs = new Set();
    const dedupedRows = newRows.filter(row => {
      const po = String(row[tsPO] || '').trim().toUpperCase();
      if (!po || seenPOs.has(po)) return false;
      seenPOs.add(po);
      return true;
    });

    const writes = [];
    if (tsData.length > 1) {
      writes.push(
        sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: SHEET_ID,
          requestBody: {
            valueInputOption: 'USER_ENTERED',
            data: [
              { range: `'TRADESTONE DATABASE'!AJ2`, values: ajVals },
              { range: `'TRADESTONE DATABASE'!AK2`, values: akVals },
              { range: `'TRADESTONE DATABASE'!H2`,  values: hVals  },
            ],
          },
        })
      );
    }
    if (dedupedRows.length) {
      writes.push(sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID, range: `'TRADESTONE DATABASE'!A1`,
        valueInputOption: 'USER_ENTERED', insertDataOption: 'INSERT_ROWS',
        requestBody: { values: dedupedRows },
      }));
    }
    await Promise.all(writes);

    // Full dedup pass — read fresh, remove duplicate PO# rows, rewrite with corrected formula row refs
    const freshData = await readTab('TRADESTONE DATABASE', true);
    const freshH = freshData[0] || [];
    const freshPOcol = hMap(freshH)['po#'] ?? tsPO;

    // Build deduped list keeping track of each row's original sheet position
    const seenAll = new Set();
    const dedupedWithPos = [];
    for (let i = 1; i < freshData.length; i++) {
      const row = freshData[i];
      const po = String(row[freshPOcol] || '').trim().toUpperCase();
      if (po && seenAll.has(po)) continue;
      if (po) seenAll.add(po);
      dedupedWithPos.push({ row, origRow: i + 1 }); // origRow = 1-indexed sheet row
    }

    const removedCount = (freshData.length - 1) - dedupedWithPos.length;
    if (removedCount > 0) {
      // Adjust formula row references to match each row's new sheet position
      // e.g. =YEAR(H2939) in a row that moves to row 100 becomes =YEAR(H100)
      const adjustFormulas = (row, origRow, newRow) => {
        if (origRow === newRow) return row;
        return row.map(cell => {
          if (!isFormulaPB(cell)) return cell;
          return cell.replace(/([A-Z]+)(\d+)/g, (match, col, num) =>
            parseInt(num) === origRow ? col + newRow : match
          );
        });
      };

      const writeRows = dedupedWithPos.map(({ row, origRow }, j) =>
        adjustFormulas(row, origRow, j + 2) // new sheet row = j+2 (data starts at row 2)
      );

      await sheets.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: `'TRADESTONE DATABASE'!A2:ZZ` });
      if (writeRows.length) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID, range: `'TRADESTONE DATABASE'!A2`,
          valueInputOption: 'USER_ENTERED', requestBody: { values: writeRows },
        });
      }
    }

    res.json({ ok: true, newRows: dedupedRows.length, duplicatesRemoved: removedCount, invoiceRefreshed: ajVals.length });
  } catch (e) {
    console.error('[powerbi-sync]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Gabriel: MAP DATA SYNC ───────────────────────────────────────────────────
app.post('/api/gabriel/map-sync', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    return res.status(500).json({ error: 'Google credentials not configured' });
  }
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });

    // Read all 5 sheets in parallel — MAP is a separate spreadsheet
    // Target sheet (MAP) is read with FORMULA mode so we can detect and preserve formula cells
    const readSheet = async (sheetId, tab, formula = false) => {
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range: `'${tab}'`,
        valueRenderOption: formula ? 'FORMULA' : 'UNFORMATTED_VALUE',
        dateTimeRenderOption: 'FORMATTED_STRING',
      });
      return r.data.values || [];
    };
    const isFormula = v => typeof v === 'string' && v.startsWith('=');
    const [targetData, tsData, pdData, ptData, sourceData] = await Promise.all([
      readSheet(MAP_SHEET_ID, 'ANTHRO MAP 2026', true), // FORMULA mode — preserves formulas on write-back
      readSheet(SHEET_ID,     'TRADESTONE DATABASE'),
      readSheet(SHEET_ID,     'PO DETAIL'),
      readSheet(SHEET_ID,     'PO TRADE'),
      readSheet(SHEET_ID,     'Production & PO DataBase'),
    ]);

    const srcH  = sourceData[0] || [];
    const tgtH  = targetData[0] || [];
    const tsH   = tsData[0]    || [];
    const pdH   = pdData[0]    || [];

    const idx  = (H, name) => H.indexOf(name);
    const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const smartIdx = (H, ...names) => {
      const nH = H.map(h => norm(h));
      for (const name of names) {
        const t = norm(name);
        let i = nH.indexOf(t); if (i !== -1) return i;
        i = nH.findIndex(h => h.includes(t) || t.includes(h)); if (i !== -1) return i;
      }
      return -1;
    };

    // ── PASS 1: refreshManu ─────────────────────────────────────────────────
    const statusAllowed  = ["PO'd", "Waiting PO", "PO'd + production ok"];
    const alwaysWrite    = ["COST", "EX FACTORY / FLIGHT DATE"];
    const onlyIfBlank    = ["SUPPLIER", "HTS CODE", "DUTY", "FREIGHT"];
    const allCols        = [...alwaysWrite, ...onlyIfBlank];

    const srcStyleCol  = smartIdx(srcH, 'ORIGINAL STYLE#', 'STYLE #', 'STYLE#', 'Style Number', 'Style');
    const srcStatusCol = smartIdx(srcH, 'STATUS');
    const tgtStyleCol  = smartIdx(tgtH, 'STYLE #', 'STYLE#', 'Style Number');
    const srcSubCat    = smartIdx(srcH, 'SUB-CATEGORY', 'SubCategory', 'Sub Category');
    const tgtSubCat    = smartIdx(tgtH, 'SUB-CATEGORY', 'SubCategory', 'Sub Category');

    const srcCI = {}, tgtCI = {};
    for (const c of allCols) {
      srcCI[c] = smartIdx(srcH, c);
      tgtCI[c] = smartIdx(tgtH, c);
    }

    const styleMap = {};
    for (let i = 1; i < sourceData.length; i++) {
      const row    = sourceData[i];
      const status = srcStatusCol >= 0 ? String(row[srcStatusCol] || '') : '';
      const style  = srcStyleCol  >= 0 ? String(row[srcStyleCol]  || '').trim() : '';
      if (!statusAllowed.includes(status) || !style) continue;
      const entry  = {};
      for (const c of allCols) { const v = srcCI[c] >= 0 ? row[srcCI[c]] : undefined; if (v != null && v !== '') entry[c] = v; }
      if (srcSubCat >= 0 && row[srcSubCat] != null && row[srcSubCat] !== '') entry['SUB-CATEGORY'] = row[srcSubCat];
      styleMap[style] = entry;
    }

    // ── All column indices pre-computed once ───────────────────────────────
    const poCol = smartIdx(tgtH, 'Purchase Order', 'PO#', 'PO');
    const tsPO  = smartIdx(tsH,  'Purchase Order', 'PO#', 'PO');
    const pdPO  = smartIdx(pdH,  'Purchase Order', 'PO#', 'PO');
    const tsNorm = tsH.map(h => h.toString().trim().toLowerCase());

    // IP CLASS → CATEGORY mapping
    const IP_TO_CATEGORY = {
      1501:'Blouses',1502:'SWTRS & SWTSHRTS',1504:'Pants',1505:'Skirts',
      1506:'SHORTS',1508:'Dresses',1509:'Dresses',1515:'Dresses',
      4110:'Blouses',4114:'Sweaters',4120:'Skirts',4123:'Pants',
      4124:'Jumpsuit',4125:'Shorts',4130:'Dresses',4134:'Dresses',
      4141:'LOUNGE',4142:'Swimwear',4148:'Blouses',4152:'Hats',
      4153:'Accessories',4157:'Wraps',4428:'Dresses',4915:'Waters Edge',
      4920:'Blouses',4925:'Skirts',4927:'Pants',4928:'Dresses',4942:'Swimwear',
    };
    const ipToCategory = ip => { const v = parseInt(ip); return IP_TO_CATEGORY[v] || 'OTHER'; };

    // Convert 0-based column index to spreadsheet letter (A, B, ... Z, AA, AB ...)
    const colLetter = n => {
      let s = '', m = n + 1;
      while (m > 0) { s = String.fromCharCode(64 + ((m - 1) % 26 + 1)) + s; m = Math.floor((m - 1) / 26); }
      return s;
    };

    const tgt = {
      period:        idx(tgtH, 'PERIOD'),
      invoiceDate:   idx(tgtH, 'URBN INVOICE DATE'),
      invoiceTotal:  idx(tgtH, 'URBN INVOICE TOTAL'),
      totalQty:      idx(tgtH, 'Total Qty'),
      wholesale:     idx(tgtH, 'PO WHOLESALE'),
      shipDate:      idx(tgtH, 'Ship Date'),
      cancelDate:    idx(tgtH, 'Cancel Date'),
      originCountry: idx(tgtH, 'Origin Country'),
      styleDesc:     idx(tgtH, 'Style Description'),
      vendorColor:   idx(tgtH, 'Vendor Color'),
      ipClass:       idx(tgtH, 'IP CLASS'),
      category:      idx(tgtH, 'CATEGORY'),
      customsDesc:   idx(tgtH, 'Customs Description'),
      brand:         idx(tgtH, 'BRAND'),
      deliverTo:     idx(tgtH, 'Deliver To'),
      fobPrice:      idx(tgtH, 'FOB Price'),
      channel:       idx(tgtH, 'CHANNEL'),
    };

    // ── STEP 1: new PO discovery ────────────────────────────────────────────
    const ts_style  = tsNorm.indexOf('vendor style #');
    const ts_po     = tsNorm.indexOf('po#');
    const ts_qty    = tsNorm.indexOf('total qty');
    const ts_wsp    = tsNorm.indexOf('po wholesale');
    const ts_fob    = tsNorm.indexOf('total fob');
    const ts_ship   = tsNorm.indexOf('ship date');
    const ts_cancel = tsNorm.indexOf('cancel date');
    const ts_status = tsNorm.indexOf('urbn status');

    const existingPOs = new Set(
      targetData.slice(1).map(r => String(r[poCol] || '').trim()).filter(Boolean)
    );

    const emptyRow = Array(tgtH.length).fill('');
    const newRows  = [];

    for (let i = 1; i < tsData.length; i++) {
      const r      = tsData[i];
      const status = String(r[ts_status] || '').toUpperCase();
      if (status.includes('CANCEL')) continue;

      const po = String(r[ts_po] || '').trim();
      if (!po || existingPOs.has(po)) continue;

      const cancelStr  = typeof r[ts_cancel] === 'string' ? r[ts_cancel] : '';
      const cancelYear = cancelStr ? new Date(cancelStr).getFullYear() : 0;
      if (cancelYear !== 2026) continue;

      const row = [...emptyRow];
      if (tgtStyleCol    >= 0) row[tgtStyleCol]    = r[ts_style]  || '';
      if (poCol          >= 0) row[poCol]           = po;
      if (tgt.totalQty   >= 0) row[tgt.totalQty]   = r[ts_qty]   || '';
      if (tgt.wholesale  >= 0) row[tgt.wholesale]   = r[ts_wsp]  || r[ts_fob] || '';
      if (tgt.shipDate   >= 0) row[tgt.shipDate]    = r[ts_ship]  || '';
      if (tgt.cancelDate >= 0) row[tgt.cancelDate]  = r[ts_cancel]|| '';
      newRows.push(row);
      existingPOs.add(po);
    }

    newRows.sort((a, b) =>
      String(a[tgtStyleCol]).localeCompare(String(b[tgtStyleCol])) ||
      String(a[poCol]).localeCompare(String(b[poCol]))
    );

    // ── Build mapRows (existing rows — passes applied below) ───────────────
    let pass1 = 0;
    const mapRows = targetData.slice(1).map(rawRow => [...rawRow]);

    const ts_ = {
      period:       smartIdx(tsH, 'PERIOD'),
      invoiceDate:  smartIdx(tsH, 'URBN INVOICE DATE', 'INVOICE DATE', 'Invoice Dt', 'Invoice Date'),
      invoiceTotal: smartIdx(tsH, 'URBN INVOICE TOTAL', 'INVOICE TOTAL', 'Invoice Amount', 'Invoice Total'),
      totalQty:     smartIdx(tsH, 'TOTAL QTY', 'Total Qty', 'Total Units', 'QTY'),
      wholesale:    smartIdx(tsH, 'TOTAL FOB', 'Total FOB', 'PO WHOLESALE', 'WHOLESALE', 'Unit Cost', 'Wholesale'),
      shipDate:     smartIdx(tsH, 'SHIP DATE', 'Ship Date', 'Shipment Date'),
      cancelDate:   smartIdx(tsH, 'CANCEL DATE', 'Cancel Date'),
    };

    const pd_ = {
      originCountry: smartIdx(pdH, 'Origin Country', 'ORIGIN COUNTRY', 'Country of Origin'),
      styleDesc:     smartIdx(pdH, 'Style Description', 'STYLE DESCRIPTION', 'Description'),
      vendorColor:   smartIdx(pdH, 'Vendor Color', 'VENDOR COLOR', 'Color'),
      ipClass:       smartIdx(pdH, 'IP CLASS', 'IP Class'),
      customsDesc:   smartIdx(pdH, 'Customs Description', 'CUSTOMS DESCRIPTION', 'Customs Desc'),
      brand:         smartIdx(pdH, 'Brand', 'BRAND'),
      deliverTo:     smartIdx(pdH, 'Deliver To', 'DELIVER TO', 'Destination'),
      fobPrice:      smartIdx(pdH, 'FOB Price', 'FOB PRICE', 'Unit FOB'),
    };

    const tsMap = new Map();
    if (tsPO >= 0) for (let i = 1; i < tsData.length; i++) { const po = norm(tsData[i][tsPO]); if (po) tsMap.set(po, tsData[i]); }
    const pdMap = new Map();
    if (pdPO >= 0) for (let i = 1; i < pdData.length; i++) { const po = norm(pdData[i][pdPO]); if (po) pdMap.set(po, pdData[i]); }

    // ptMap: find PO and CHANNEL columns by header, not hardcoded index
    const ptH       = ptData[0] || [];
    const ptPO      = smartIdx(ptH, 'Purchase Order', 'PO#', 'PO');
    const ptChannel = smartIdx(ptH, 'CHANNEL', 'Channel');
    const ptMap = new Map();
    if (ptPO >= 0 && ptChannel >= 0) {
      for (let i = 1; i < ptData.length; i++) {
        const po = norm(ptData[i][ptPO]);
        if (po) ptMap.set(po, ptData[i][ptChannel]);
      }
    }

    // helper: only write if src column found, value is non-empty, and target cell is NOT a formula
    const setSrc = (row, tgtIdx, srcRow, srcIdx) => {
      if (tgtIdx < 0 || srcIdx < 0) return;
      if (isFormula(row[tgtIdx])) return;
      const v = srcRow[srcIdx];
      if (v != null && v !== '') row[tgtIdx] = v;
    };

    const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    const monthName = v => {
      if (!v && v !== 0) return '';
      const d = new Date(v);
      return isNaN(d) ? '' : MONTH_NAMES[d.getMonth()];
    };

    // Enrich a single row with Tradestone, PO DETAIL, PO TRADE, PERIOD, CATEGORY
    const enrichRow = row => {
      const po = norm(row[poCol]);
      if (!po) return;
      const ts = tsMap.get(po);
      const pd = pdMap.get(po);
      if (ts) {
        setSrc(row, tgt.invoiceDate,  ts, ts_.invoiceDate);
        setSrc(row, tgt.invoiceTotal, ts, ts_.invoiceTotal);
        setSrc(row, tgt.totalQty,     ts, ts_.totalQty);
        setSrc(row, tgt.wholesale,    ts, ts_.wholesale);
        setSrc(row, tgt.shipDate,     ts, ts_.shipDate);
        setSrc(row, tgt.cancelDate,   ts, ts_.cancelDate);
      }
      // PERIOD rule: month name of URBN INVOICE DATE if populated, else Cancel Date
      if (tgt.period >= 0 && !isFormula(row[tgt.period])) {
        const invDateVal    = tgt.invoiceDate >= 0 ? row[tgt.invoiceDate] : '';
        const cancelDateVal = tgt.cancelDate  >= 0 ? row[tgt.cancelDate]  : '';
        const src = (invDateVal !== '' && invDateVal != null) ? invDateVal : cancelDateVal;
        const mn = monthName(src);
        if (mn) row[tgt.period] = mn;
      }
      if (pd) {
        setSrc(row, tgt.originCountry, pd, pd_.originCountry);
        setSrc(row, tgt.styleDesc,     pd, pd_.styleDesc);
        setSrc(row, tgt.vendorColor,   pd, pd_.vendorColor);
        setSrc(row, tgt.ipClass,       pd, pd_.ipClass);
        setSrc(row, tgt.customsDesc,   pd, pd_.customsDesc);
        setSrc(row, tgt.brand,         pd, pd_.brand);
        setSrc(row, tgt.deliverTo,     pd, pd_.deliverTo);
        setSrc(row, tgt.fobPrice,      pd, pd_.fobPrice);
        if (tgt.category >= 0 && pd_.ipClass >= 0 && !isFormula(row[tgt.category])) {
          const ip = pd[pd_.ipClass];
          if (ip != null && ip !== '') row[tgt.category] = ipToCategory(ip);
        }
      }
      if (ptMap.has(po) && tgt.channel >= 0) row[tgt.channel] = ptMap.get(po);
    };

    // Apply styleMap (Pass 1 / Production DB) to a row by STYLE#
    const applyStyleMap = row => {
      const style = String(row[tgtStyleCol] || '').trim();
      if (!style) return;
      const match = styleMap[style];
      if (!match) return;
      for (const c of allCols) {
        const ti = tgtCI[c];
        if (ti < 0) continue;
        if (isFormula(row[ti])) continue;
        const nv = match[c];
        if (nv === undefined) continue;
        const cur = row[ti];
        const blank = cur === '' || cur === null || cur === undefined;
        if (alwaysWrite.includes(c) || (onlyIfBlank.includes(c) && blank)) row[ti] = nv;
      }
      if (!isFormula(row[tgtSubCat])) {
        const subBlank = row[tgtSubCat] === '' || row[tgtSubCat] === null || row[tgtSubCat] === undefined;
        if (subBlank && match['SUB-CATEGORY'] !== undefined) row[tgtSubCat] = match['SUB-CATEGORY'];
      }
    };

    // Apply to existing rows (Pass 1 + Pass 2/3)
    let pass2 = 0;
    for (const row of mapRows) {
      const style = String(row[tgtStyleCol] || '').trim();
      if (style && styleMap[style]) { applyStyleMap(row); pass1++; }
      const po = norm(row[poCol]);
      if (po) { enrichRow(row); pass2++; }
    }

    // Apply ALL passes to new rows in the same run — no second sync needed
    for (const row of newRows) {
      applyStyleMap(row);
      enrichRow(row);
    }

    // Write only the specific columns we touched (preserves all formulas in other columns)
    const writeCols = [
      ...Object.values(tgtCI),                           // Pass 1 cols (COST, EX FACTORY, SUPPLIER, HTS CODE, DUTY, FREIGHT)
      tgtSubCat,                                          // SUB-CATEGORY
      tgt.period, tgt.invoiceDate, tgt.invoiceTotal,     // Tradestone
      tgt.totalQty, tgt.wholesale, tgt.shipDate, tgt.cancelDate,
      tgt.originCountry, tgt.styleDesc, tgt.vendorColor, // PO DETAIL
      tgt.ipClass, tgt.category, tgt.customsDesc,
      tgt.brand, tgt.deliverTo, tgt.fobPrice,
      tgt.channel,                                        // PO TRADE
    ].filter(ci => ci != null && ci >= 0);

    const uniqueCols = [...new Set(writeCols)];

    const batchData = uniqueCols.map(ci => ({
      range: `'ANTHRO MAP 2026'!${colLetter(ci)}2`,
      values: mapRows.map(r => [r[ci] ?? '']),
    }));

    const writes = [
      sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: MAP_SHEET_ID,
        requestBody: { valueInputOption: 'USER_ENTERED', data: batchData },
      }),
    ];

    if (newRows.length) {
      writes.push(sheets.spreadsheets.values.append({
        spreadsheetId: MAP_SHEET_ID,
        range: `'ANTHRO MAP 2026'!A1`,
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: { values: newRows },
      }));
    }
    await Promise.all(writes);

    res.json({
      ok: true,
      newPOs: newRows.length,
      pass1Updated: pass1,
      pass2Updated: pass2,
      totalRows: mapRows.length + newRows.length,
      debug: {
        srcStyleCol, srcStatusCol, styleMapSize: Object.keys(styleMap).length,
        tsPO, pdPO, ptPO, ptChannel,
        tsMapSize: tsMap.size, pdMapSize: pdMap.size, ptMapSize: ptMap.size,
        ts_, pd_,
      }
    });

  } catch (e) {
    console.error('[map-sync]', e.message);
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n  305 WORKSPACE TEAM`);
  console.log(`  ─────────────────────────────`);
  console.log(`  http://localhost:${PORT}\n`);
});
