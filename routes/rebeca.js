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

// ─── Search Style ─────────────────────────────────────────────────────────────
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
      const fVal    = String(formulaRow[i] ?? '');
      const imgMatch = fVal.match(/^=IMAGE\("([^"]+)"/i);
      result[h] = imgMatch ? imgMatch[1] : (row[i] ?? '');
    });
    res.json({ style: result });
  } catch (e) {
    console.error('[rebeca/search-style]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Save / Create Style ──────────────────────────────────────────────────────
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

    // New style — append to all 4 tabs
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

// ─── Delete Style from all sheets ─────────────────────────────────────────────
router.delete('/delete-style', async (req, res) => {
  const styleNum = String(req.query.style || '').trim().toUpperCase();
  if (!styleNum) return res.status(400).json({ error: 'STYLE # required' });
  if (!requireCreds(res)) return;
  try {
    const sheets = sheetsClient(false);
    const tabs   = ['Design DataBase', 'Production & PO DataBase', 'Print DataBase', 'PowerBI database Process'];

    // Fetch sheet metadata + all tab data in parallel
    const [meta, ...tabResults] = await Promise.all([
      sheets.spreadsheets.get({ spreadsheetId: SHEET_ID }),
      ...tabs.map(tab => sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID, range: `'${tab}'`,
        valueRenderOption: 'FORMATTED_VALUE',
      }).catch(() => null)),
    ]);

    const sheetIdMap = {};
    for (const s of meta.data.sheets) sheetIdMap[s.properties.title] = s.properties.sheetId;

    const deletedFrom = [];
    const skipped     = [];

    for (let i = 0; i < tabs.length; i++) {
      const tab    = tabs[i];
      const result = tabResults[i];
      if (!result) { skipped.push(tab); continue; }

      const rows     = result.data.values || [];
      const headers  = (rows[0] || []).map(h => String(h).trim());
      const styleCol = headers.findIndex(h => h.toUpperCase() === 'STYLE #');
      if (styleCol < 0) continue;

      const rowIdx = rows.slice(1).findIndex(r => String(r[styleCol] || '').trim().toUpperCase() === styleNum);
      if (rowIdx < 0) continue;

      const sheetId = sheetIdMap[tab];
      if (sheetId === undefined) continue;

      try {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SHEET_ID,
          requestBody: { requests: [{ deleteDimension: {
            range: { sheetId, dimension: 'ROWS', startIndex: rowIdx + 1, endIndex: rowIdx + 2 },
          }}]},
        });
        deletedFrom.push(tab);
      } catch (tabErr) {
        const detail = tabErr.response?.data?.error?.message || tabErr.message;
        console.error(`[rebeca/delete-style] ${tab} row ${rowIdx + 1} sheetId ${sheetId}:`, detail);
        skipped.push(`${tab} (${detail})`);
      }
    }

    if (deletedFrom.length === 0) {
      const reason = skipped.length ? skipped.join('; ') : `Style "${styleNum}" not found`;
      return res.status(404).json({ error: reason });
    }
    res.json({ ok: true, deletedFrom, skipped });
  } catch (e) {
    const detail = e.response?.data?.error?.message || e.message;
    console.error('[rebeca/delete-style]', detail);
    res.status(500).json({ error: detail });
  }
});

// ─── Dashboard stats ──────────────────────────────────────────────────────────
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
    const col      = (name) => headers.findIndex(h => h.toUpperCase() === name.toUpperCase());
    const styleCol = col('STYLE #');
    const tpCol    = col('TP SENT TO SUPPLIER');
    const printCol = col('PRINT SENT TO SUPPLIER');
    const ndcCol   = col('NDC MONTH/YEAR');

    // Only 2026 styles with a style number
    const styleRows = rows.slice(1).filter(r => {
      if (!String(r[styleCol] || '').trim()) return false;
      const ndc = String(r[ndcCol] || '').trim();
      return ndc.endsWith('/2026');
    });

    const total        = styleRows.length;
    const missingTp    = styleRows.filter(r => !String(r[tpCol]    || '').trim()).length;
    const missingPrint = styleRows.filter(r => !String(r[printCol] || '').trim()).length;

    // Group by month — last 4 months with data, each with its own TP/Print counts
    const monthMap = {};
    styleRows.forEach(r => {
      const ndc = String(r[ndcCol] || '').trim();
      if (!ndc) return;
      if (!monthMap[ndc]) monthMap[ndc] = { count: 0, missingTp: 0, missingPrint: 0 };
      monthMap[ndc].count++;
      if (!String(r[tpCol]    || '').trim()) monthMap[ndc].missingTp++;
      if (!String(r[printCol] || '').trim()) monthMap[ndc].missingPrint++;
    });

    const MONTH_NAMES = ['', 'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
                              'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    const months = Object.entries(monthMap)
      .map(([key, stats]) => {
        const [mm] = key.split('/');
        const m = parseInt(mm);
        return { label: MONTH_NAMES[m] || mm, sortKey: m, ...stats };
      })
      .filter(m => !isNaN(m.sortKey))
      .sort((a, b) => b.sortKey - a.sortKey)
      .slice(0, 4);

    res.json({ total, missingTp, missingPrint, months });
  } catch (e) {
    console.error('[rebeca/dashboard]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── CAD Image proxy (bypasses Google Drive embed restrictions) ───────────────
router.get('/cad-image', async (req, res) => {
  const fileId = String(req.query.id || '').trim();
  if (!fileId) return res.status(400).end();
  try {
    // Google Drive direct download URL
    const driveUrl = `https://drive.google.com/uc?export=download&id=${encodeURIComponent(fileId)}`;
    const r = await fetch(driveUrl, {
      redirect: 'follow',
      headers: { 'User-Agent': 'Mozilla/5.0' },
    });
    if (!r.ok) return res.status(502).end();
    const ct = r.headers.get('content-type') || 'image/jpeg';
    // If Google returns HTML (virus/size warning page) treat as not found
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
