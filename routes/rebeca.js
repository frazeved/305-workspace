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
// Uses values.update + values.clear (same API as save) — avoids batchUpdate/deleteDimension failures.
// Strategy: shift all rows above the deleted row down by one, then clear the now-empty last row.
router.delete('/delete-style', async (req, res) => {
  const styleNum = String(req.query.style || '').trim().toUpperCase();
  if (!styleNum) return res.status(400).json({ error: 'STYLE # required' });
  if (!requireCreds(res)) return;

  const sheets = sheetsClient(false);
  const tabs   = ['Design DataBase', 'Production & PO DataBase', 'Print DataBase', 'PowerBI database Process'];

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
      const detail = e.response?.data?.error?.message || e.message;
      skipped.push(`${tab} (read: ${detail})`);
      continue;
    }

    const headers  = (rows[0] || []).map(h => String(h).trim());
    const styleCol = headers.findIndex(h => h.toUpperCase() === 'STYLE #');
    if (styleCol < 0) continue;

    const rowIdx = rows.slice(1).findIndex(r => String(r[styleCol] || '').trim().toUpperCase() === styleNum);
    if (rowIdx < 0) continue;

    // Sheet row of the deleted style (1-based): header=1, first data=2
    const deletedSheetRow = rowIdx + 2;
    const lastSheetRow    = rows.length;
    // Rows that need to shift up: everything after the deleted row
    const rowsToShift = rows.slice(rowIdx + 2);

    try {
      // Write shifted rows starting at the deleted row's position (never touches header row 1)
      if (rowsToShift.length > 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID,
          range: `'${tab}'!A${deletedSheetRow}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: rowsToShift },
        });
      }
      // Clear the now-empty last row
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

// ─── TP Generator ─────────────────────────────────────────────────────────────
router.post('/generate-tp', async (req, res) => {
  if (!requireCreds(res)) return;

  const {
    style, model, category, subCategory, fabric,
    printType, printName, supplier, cadUrl,
    printSentToSupplier, tpSentToSupplier, linkStyleFolder,
  } = req.body;

  if (!style)           return res.status(400).json({ error: 'STYLE # required' });
  if (!model)           return res.status(400).json({ error: 'TECH PACK MODEL required' });
  if (!linkStyleFolder) return res.status(400).json({ error: 'LINK STYLE FOLDER is required — fill in the Drive folder URL for this style first' });

  const norm = s => (s || '').replace(/ /g, ' ').replace(/\s+/g, ' ').toUpperCase().trim();

  const MODEL_TEMPLATES = {
    'TP_FARM ANTHRO NEW BODY':    '1Zvg1gdMMzE3mgdDuKHTnqhgJJKrskeM3MIw2ZVpSRsY',
    'TP_FARM ANTHRO REPEAT BODY': '1bMFvktTY0zb2rRmr1Fm4M3fqFYIdlAOXtFG9GsXL-00',
  };
  const templateId = MODEL_TEMPLATES[norm(model)];
  if (!templateId) return res.status(400).json({ error: `Unknown model: "${model}"` });

  // Extract folder ID from LINK STYLE FOLDER Drive URL
  const folderIdMatch = (linkStyleFolder || '').match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (!folderIdMatch) return res.status(400).json({ error: 'LINK STYLE FOLDER must be a valid Google Drive folder URL' });
  const folderId = folderIdMatch[1];

  // Extract Drive file ID from CAD URL (Drive sharing link or uc?id= URL)
  const extractFileId = url => {
    const m1 = (url || '').match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (m1) return m1[1];
    const m2 = (url || '').match(/[?&]id=([a-zA-Z0-9_-]+)/);
    return m2 ? m2[1] : null;
  };
  const cadFileId = extractFileId(cadUrl);

  // SAMPLE_DUE_DATE = tpSentToSupplier + 14 days
  const padZ = n => String(n).padStart(2, '0');
  const fmtDate = d => (!d || isNaN(d)) ? '' : `${padZ(d.getMonth()+1)}/${padZ(d.getDate())}/${d.getFullYear()}`;
  const parseSheetDate = s => { // YYYY-MM-DD → display MM/DD/YYYY
    if (!s) return '';
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? `${m[2]}/${m[3]}/${m[1]}` : s;
  };
  let sampleDueDate = '';
  if (tpSentToSupplier) {
    const d = new Date(tpSentToSupplier + 'T12:00:00');
    if (!isNaN(d)) { d.setDate(d.getDate() + 14); sampleDueDate = fmtDate(d); }
  }

  try {
    const sa = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth = new google.auth.GoogleAuth({
      credentials: sa,
      scopes: ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/presentations'],
    });
    const authClient = await auth.getClient();
    const drive  = google.drive({ version: 'v3', auth: authClient });
    const slides = google.slides({ version: 'v1', auth: authClient });

    // Build presentation file name
    const safe = s => (s || '').replace(/[\\/:*?"<>|#\[\]\r\n]/g, '').trim();
    const cleanStyle = (style || '').replace(/^[A-Za-z]+-?/, '');
    const copyName = `${safe(norm(model))} – ${safe(supplier || '')} – ${safe(cleanStyle)}`;

    // Copy template into LINK STYLE FOLDER
    const copyRes = await drive.files.copy({
      fileId: templateId,
      supportsAllDrives: true,
      requestBody: { name: copyName, parents: [folderId] },
    });
    const copyId = copyRes.data.id;

    // Get presentation structure to find {{CAD_IMAGE}} placeholder shapes
    const presData = await slides.presentations.get({ presentationId: copyId });

    // Text token replacement requests
    const tokens = {
      '{{STYLE}}':                  style,
      '{{STYLE #}}':                style,
      '{{DESIGN_STYLE}}':           style,
      '{{CATEGORY}}':               category || '',
      '{{SUB_CATEGORY}}':           subCategory || '',
      '{{FABRIC}}':                 fabric || '',
      '{{PRINT_TYPE}}':             printType || '',
      '{{PRINT_NAME}}':             printName || '',
      '{{SUPPLIER}}':               supplier || '',
      '{{PRINT_SENT_TO_SUPPLIER}}': parseSheetDate(printSentToSupplier),
      '{{PRINT_SENT_TO_SUP}}':      parseSheetDate(printSentToSupplier),
      '{{TP_SENT_TO_SUPPLIER}}':    parseSheetDate(tpSentToSupplier),
      '{{SAMPLE_DUE_DATE}}':        sampleDueDate,
      '{{SAMPLE_SIZE}}':            'SMALL 4/6',
      '{{TECHNICAL_DESIGNER}}':     'bertha@creativetwotwelve.com',
      '{{CAD_URL}}':                cadUrl || '',
    };
    const requests = Object.entries(tokens).map(([ph, val]) => ({
      replaceAllText: { containsText: { text: ph, matchCase: true }, replaceText: String(val ?? '') },
    }));

    // Find {{CAD_IMAGE}} placeholder shapes and replace with actual image
    if (cadFileId || cadUrl) {
      const cadImgUrl = cadFileId
        ? `https://drive.google.com/uc?export=view&id=${cadFileId}`
        : cadUrl;
      const INCH = 914400; // EMU per inch
      for (const slide of (presData.data.slides || [])) {
        for (const el of (slide.pageElements || [])) {
          if (!el.shape?.text) continue;
          const text = (el.shape.text.textElements || [])
            .map(te => te.textRun?.content || '').join('').trim();
          if (text === '{{CAD_IMAGE}}') {
            requests.push(
              { deleteObject: { objectId: el.objectId } },
              {
                createImage: {
                  url: cadImgUrl,
                  elementProperties: {
                    pageObjectId: slide.objectId,
                    size: {
                      width:  { magnitude: Math.round(2.59 * INCH), unit: 'EMU' },
                      height: { magnitude: Math.round(7.00 * INCH), unit: 'EMU' },
                    },
                    transform: {
                      scaleX: 1, scaleY: 1, shearX: 0, shearY: 0,
                      translateX: Math.round(0.97 * INCH),
                      translateY: Math.round(1.61 * INCH),
                      unit: 'EMU',
                    },
                  },
                },
              }
            );
          }
        }
      }
    }

    await slides.presentations.batchUpdate({ presentationId: copyId, requestBody: { requests } });

    res.json({
      ok: true,
      presentationUrl: `https://docs.google.com/presentation/d/${copyId}/edit`,
      folderUrl: `https://drive.google.com/drive/folders/${folderId}`,
    });
  } catch (e) {
    const detail = e.response?.data?.error?.message || e.message;
    console.error('[rebeca/generate-tp]', detail);
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
      .sort((a, b) => a.sortKey - b.sortKey)
      .slice(-4);

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
