const { Router } = require('express');
const { google }  = require('googleapis');

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
    const buildRow   = (hdrs) => hdrs.map(h => formData[h] ?? '');

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

    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const sheetIdMap = {};
    for (const s of meta.data.sheets) sheetIdMap[s.properties.title] = s.properties.sheetId;

    const tabs        = ['Design DataBase', 'Production & PO DataBase', 'Print DataBase', 'PowerBI database Process'];
    const deletedFrom = [];

    for (const tab of tabs) {
      const sheetId = sheetIdMap[tab];
      if (sheetId === undefined) continue;
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID, range: `'${tab}'`,
        valueRenderOption: 'FORMATTED_VALUE',
      });
      const rows     = r.data.values || [];
      const headers  = (rows[0] || []).map(h => String(h).trim());
      const styleCol = headers.findIndex(h => h.toUpperCase() === 'STYLE #');
      if (styleCol < 0) continue;
      const rowIdx = rows.slice(1).findIndex(r => String(r[styleCol] || '').trim().toUpperCase() === styleNum);
      if (rowIdx < 0) continue;
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { requests: [{ deleteDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex: rowIdx + 1, endIndex: rowIdx + 2 },
        }}]},
      });
      deletedFrom.push(tab);
    }

    if (deletedFrom.length === 0) return res.status(404).json({ error: `Style "${styleNum}" not found in any sheet` });
    res.json({ ok: true, deletedFrom });
  } catch (e) {
    console.error('[rebeca/delete-style]', e.message);
    res.status(500).json({ error: e.message });
  }
});

module.exports = router;
