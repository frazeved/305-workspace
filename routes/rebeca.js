const { Router }   = require('express');
const { google }   = require('googleapis');
const { Readable } = require('stream');

const router   = Router();
const SHEET_ID = '1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q';

function sheetsClient(readonly = true) {
  const sa    = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth  = new google.auth.GoogleAuth({
    credentials: sa,
    scopes: [readonly
      ? 'https://www.googleapis.com/auth/spreadsheets.readonly'
      : 'https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

function requireCreds(res) {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    res.status(500).json({ error: 'Google credentials not configured' });
    return false;
  }
  return true;
}

// Search Style
router.get('/search-style', async (req, res) => {
  const styleNum = String(req.query.style || '').trim().toUpperCase();
  if (!styleNum) return res.status(400).json({ error: 'Style # required' });
  if (!requireCreds(res)) return;
  try {
    const sheets = sheetsClient(true);
    const [rFmt, rFormula] = await Promise.all([
      sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `'Design DataBase'`,
        valueRenderOption: 'FORMATTED_VALUE',
        dateTimeRenderOption: 'FORMATTED_STRING',
      }),
      sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `'Design DataBase'`,
        valueRenderOption: 'FORMULA',
      }),
    ]);
    const rows = rFmt.data.values || [];
    if (rows.length < 2) return res.status(404).json({ error: 'Sheet is empty' });
    const headers  = rows[0].map(h => String(h).trim());
    const styleCol = headers.findIndex(h => h.toUpperCase() === 'STYLE #');
    if (styleCol < 0) return res.status(500).json({ error: 'STYLE # column not found' });
    const rowIdx = rows.slice(1).findIndex(r => String(r[styleCol] || '').trim().toUpperCase() === styleNum);
    if (rowIdx < 0) return res.status(404).json({ error: `Style "${styleNum}" not found` });
    const row        = rows.slice(1)[rowIdx];
    const formulaRow = (rFormula.data.values || []).slice(1)[rowIdx] || [];
    const result = {};
    headers.forEach((h, i) => {
      const fVal     = String(formulaRow[i] ?? '');
      const imgMatch = fVal.match(/^=IMAGE\("([^"]+)"/i);
      result[h] = imgMatch ? imgMatch[1] : (row[i] ?? '');
    });
    res.json({ style: result });
  } catch (e) {
    console.error('[rebeca/search-style]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// Save / Create Style
router.post('/save-style', async (req, res) => {
  const formData = req.body || {};
  const styleNum = String(formData['STYLE #'] || '').trim().toUpperCase();
  if (!styleNum) return res.status(400).json({ error: 'STYLE # required' });
  if (!requireCreds(res)) return;
  try {
    const sheets = sheetsClient(false);

    const readTab = async (tab) => {
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID, range: `'${tab}'`,
        valueRenderOption: 'FORMATTED_VALUE', dateTimeRenderOption: 'FORMATTED_STRING',
      });
      return r.data.values || [];
    };

    const designRows = await readTab('Design DataBase');
    const headers    = (designRows[0] || []).map(h => String(h).trim());
    const styleCol   = headers.findIndex(h => h.toUpperCase() === 'STYLE #');
    const toImageFormula = (url) => {
      if (!url) return '';
      const m = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
      if (m) return `=IMAGE("https://drive.google.com/uc?id=${m[1]}")`;
      const m2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
      if (m2) return `=IMAGE("https://drive.google.com/uc?id=${m2[1]}")`;
      return url;
    };
    const buildRow = (hdrs) => hdrs.map(h => {
      const val = formData[h] ?? '';
      return h === 'CAD IMAGE' ? toImageFormula(val) : val;
    });

    const existIdx = styleCol >= 0
      ? designRows.slice(1).findIndex(r => String(r[styleCol] || '').trim().toUpperCase() === styleNum)
      : -1;

    if (existIdx >= 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `'Design DataBase'!A${existIdx + 2}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [buildRow(headers)] },
      });
      return res.json({ ok: true, action: 'updated' });
    }

    const tabs = ['Design DataBase', 'Production & PO DataBase', 'Print DataBase', 'PowerBI database Process'];
    for (const tab of tabs) {
      const tabRows = await readTab(tab);
      const tabHdrs = (tabRows[0] || []).map(h => String(h).trim());
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID, range: `'${tab}'!A1`,
        valueInputOption: 'USER_ENTERED', insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [buildRow(tabHdrs)] },
      });
    }
    res.json({ ok: true, action: 'created' });
  } catch (e) {
    console.error('[rebeca/save-style]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// Delete Style from all sheets
router.delete('/delete-style', async (req, res) => {
  const styleNum = String(req.query.style || '').trim().toUpperCase();
  if (!styleNum) return res.status(400).json({ error: 'STYLE # required' });
  if (!requireCreds(res)) return;

  const sheets      = sheetsClient(false);
  const tabs        = ['Design DataBase', 'Production & PO DataBase', 'Print DataBase', 'PowerBI database Process'];
  const deletedFrom = [];
  const skipped     = [];

  for (const tab of tabs) {
    let rows;
    try {
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID, range: `'${tab}'`,
        valueRenderOption: 'FORMATTED_VALUE',
      });
      rows = r.data.values || [];
    } catch (e) {
      skipped.push(`${tab} (read: ${e.response?.data?.error?.message || e.message})`);
      continue;
    }

    const headers  = (rows[0] || []).map(h => String(h).trim());
    const styleCol = headers.findIndex(h => h.toUpperCase() === 'STYLE #');
    if (styleCol < 0) continue;

    const rowIdx = rows.slice(1).findIndex(r => String(r[styleCol] || '').trim().toUpperCase() === styleNum);
    if (rowIdx < 0) continue;

    const deletedSheetRow = rowIdx + 2;
    const lastSheetRow    = rows.length;
    const rowsToShift     = rows.slice(rowIdx + 2);

    try {
      if (rowsToShift.length > 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID,
          range: `'${tab}'!A${deletedSheetRow}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: rowsToShift },
        });
      }
      await sheets.spreadsheets.values.clear({
        spreadsheetId: SHEET_ID,
        range: `'${tab}'!A${lastSheetRow}:ZZ${lastSheetRow}`,
      });
      deletedFrom.push(tab);
    } catch (e) {
      const detail = e.response?.data?.error?.message || e.message;
      console.error(`[rebeca/delete-style] ${tab}:`, detail);
      skipped.push(`${tab} (${detail})`);
    }
  }

  if (deletedFrom.length === 0) {
    const reason = skipped.length ? skipped.join('; ') : `Style "${styleNum}" not found`;
    return res.status(404).json({ error: reason });
  }
  res.json({ ok: true, deletedFrom, skipped });
});

// Dashboard stats
router.get('/dashboard', async (req, res) => {
  if (!requireCreds(res)) return;
  try {
    const sheets = sheetsClient(true);
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `'Design DataBase'`,
      valueRenderOption: 'FORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING',
    });
    const rows = r.data.values || [];
    if (rows.length < 2) return res.json({ total: 0, missingTp: 0, missingPrint: 0, months: [] });

    const headers  = rows[0].map(h => String(h).trim());
    const colIdx   = name => headers.findIndex(h => h.toUpperCase() === name.toUpperCase());
    const styleCol = colIdx('STYLE #');
    const tpCol    = colIdx('TP SENT TO SUPPLIER');
    const printCol = colIdx('PRINT SENT TO SUPPLIER');
    const ndcCol   = headers.findIndex(h => h.toUpperCase().replace(/\s+/g,'').includes('NDC'));
    const yearCol  = headers.findIndex(h => {
      const k = h.toUpperCase().replace(/\s+/g,'');
      return k === 'YEAR' || k === 'NDCYEAR';
    });

    const MON_MAP = {
      JAN:1,JANUARY:1, FEB:2,FEBRUARY:2, MAR:3,MARCH:3, APR:4,APRIL:4,
      MAY:5, JUN:6,JUNE:6, JUL:7,JULY:7, AUG:8,AUGUST:8,
      SEP:9,SEPT:9,SEPTEMBER:9, OCT:10,OCTOBER:10, NOV:11,NOVEMBER:11, DEC:12,DECEMBER:12,
    };

    const parseNdc = (ndcRaw, yearRaw) => {
      const s = String(ndcRaw || '').trim();
      if (!s) return null;

      // Full date: "6/1/2026" "06/01/2026"
      const fullDate = s.match(/^(\d{1,2})\/\d{1,2}\/(\d{4})$/);
      if (fullDate) return { mon: parseInt(fullDate[1]), yr: parseInt(fullDate[2]) };

      // ISO: "2026-06-01"
      const isoDate = s.match(/^(\d{4})-(\d{2})-\d{2}$/);
      if (isoDate) return { mon: parseInt(isoDate[2]), yr: parseInt(isoDate[1]) };

      // Combined: "05/2026" "May/2026" "May 2026" "May-2026" "May/26"
      const withYear = s.match(/^([A-Za-z]+|\d{1,2})[\s\/\-](\d{2,4})$/);
      if (withYear) {
        const monPart = withYear[1];
        let yr = parseInt(withYear[2]);
        if (yr < 100) yr += 2000;
        const mon = /^\d+$/.test(monPart) ? parseInt(monPart) : (MON_MAP[monPart.toUpperCase()] || 0);
        return mon ? { mon, yr } : null;
      }

      // Month name or number only
      const mon = /^\d{1,2}$/.test(s) ? parseInt(s) : (MON_MAP[s.toUpperCase()] || 0);
      if (!mon) return null;
      let yr = yearRaw ? parseInt(String(yearRaw).trim()) : NaN;
      if (yr < 100) yr += 2000;
      if (isNaN(yr)) yr = new Date().getFullYear();
      return { mon, yr };
    };

    const MONTH_NAMES = ['','JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    const styleRows  = rows.slice(1).filter(r => String(r[styleCol] || '').trim());

    const monthMap = {};
    let total = 0, missingTp = 0, missingPrint = 0;

    styleRows.forEach(r => {
      const ndcRaw  = ndcCol  >= 0 ? r[ndcCol]  : '';
      const yearRaw = yearCol >= 0 ? r[yearCol] : '';
      const parsed  = parseNdc(ndcRaw, yearRaw);
      if (!parsed || parsed.yr !== 2026) return;

      total++;
      const noTp    = !String(r[tpCol]    || '').trim();
      const noPrint = !String(r[printCol] || '').trim();
      if (noTp)    missingTp++;
      if (noPrint) missingPrint++;

      if (!monthMap[parsed.mon]) monthMap[parsed.mon] = { count: 0, missingTp: 0, missingPrint: 0 };
      monthMap[parsed.mon].count++;
      if (noTp)    monthMap[parsed.mon].missingTp++;
      if (noPrint) monthMap[parsed.mon].missingPrint++;
    });

    // 4 months starting 2 months ahead (e.g. in May → Jul Aug Sep Oct)
    const nowMonth = new Date().getMonth() + 1;
    const months = [2, 3, 4, 5].map(i => {
      let mm = nowMonth + i;
      if (mm > 12) mm -= 12;
      return {
        label:        MONTH_NAMES[mm],
        sortKey:      mm,
        count:        monthMap[mm]?.count        || 0,
        missingTp:    monthMap[mm]?.missingTp    || 0,
        missingPrint: monthMap[mm]?.missingPrint || 0,
      };
    });

    res.json({ total, missingTp, missingPrint, months });
  } catch (e) {
    console.error('[rebeca/dashboard]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// CAD Image proxy
router.get('/cad-image', async (req, res) => {
  const fileId = String(req.query.id || '').trim();
  if (!fileId) return res.status(400).end();
  try {
    const driveUrl = `https://drive.google.com/uc?export=download&id=${encodeURIComponent(fileId)}`;
    const r = await fetch(driveUrl, {
      redirect: 'follow',
      headers: { 'User-Agent': 'Mozilla/5.0' },
    });
    if (!r.ok) return res.status(502).end();
    const ct = r.headers.get('content-type') || 'image/jpeg';
    if (ct.includes('text/html')) return res.status(404).end();
    res.setHeader('Content-Type', ct);
    res.setHeader('Cache-Control', 'public, max-age=86400');
    Readable.fromWeb(r.body).pipe(res);
  } catch (e) {
    console.error('[rebeca/cad-image]', e.message);
    res.status(500).end();
  }
});

module.exports = router;
