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

// Delete Style from all sheets
// Uses values.update + values.clear — avoids batchUpdate/deleteDimension failures.
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

    const deletedSheetRow = rowIdx + 2; // 1-based; header=1, first data=2
    const lastSheetRow    = rows.length;
    const rowsToShift     = rows.slice(rowIdx + 2); // rows after the deleted one

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

// TP Generator
// Exports the template as PPTX, edits the ZIP directly (tokens + CAD image),
// then re-uploads as Google Slides. No Slides API required.
router.post('/generate-tp', async (req, res) => {
  if (!requireCreds(res)) return;

  const {
    style, model, category, subCategory, fabric,
    printType, printName, supplier, cadUrl,
    printSentToSupplier, tpSentToSupplier,
  } = req.body;

  if (!style) return res.status(400).json({ error: 'STYLE # required' });
  if (!model) return res.status(400).json({ error: 'TECH PACK MODEL required' });

  const norm = s => (s || '').replace(/ /g, ' ').replace(/\s+/g, ' ').toUpperCase().trim();

  const MODEL_TEMPLATES = {
    'TP_FARM ANTHRO NEW BODY':    '1Zvg1gdMMzE3mgdDuKHTnqhgJJKrskeM3MIw2ZVpSRsY',
    'TP_FARM ANTHRO REPEAT BODY': '1bMFvktTY0zb2rRmr1Fm4M3fqFYIdlAOXtFG9GsXL-00',
  };
  const templateId = MODEL_TEMPLATES[norm(model)];
  if (!templateId) return res.status(400).json({ error: `Unknown model: "${model}"` });

  const folderId = '1CTH7KZzFJLd-a4bN51MEMmSF7LiOc1f6';

  const extractFileId = url => {
    const m1 = (url || '').match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (m1) return m1[1];
    const m2 = (url || '').match(/[?&]id=([a-zA-Z0-9_-]+)/);
    return m2 ? m2[1] : null;
  };
  const cadFileId = extractFileId(cadUrl);

  const padZ = n => String(n).padStart(2, '0');
  const fmtDate = d => (!d || isNaN(d)) ? '' : `${padZ(d.getMonth()+1)}/${padZ(d.getDate())}/${d.getFullYear()}`;
  const parseSheetDate = s => {
    if (!s) return '';
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? `${m[2]}/${m[3]}/${m[1]}` : s;
  };
  let sampleDueDate = '';
  if (tpSentToSupplier) {
    const d = new Date(tpSentToSupplier + 'T12:00:00');
    if (!isNaN(d)) { d.setDate(d.getDate() + 14); sampleDueDate = fmtDate(d); }
  }

  const escapeXml = s => String(s ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');

  const tokens = {
    '{{STYLE}}': style, '{{STYLE #}}': style, '{{DESIGN_STYLE}}': style,
    '{{CATEGORY}}': category || '', '{{SUB_CATEGORY}}': subCategory || '',
    '{{FABRIC}}': fabric || '', '{{PRINT_TYPE}}': printType || '',
    '{{PRINT_NAME}}': printName || '', '{{SUPPLIER}}': supplier || '',
    '{{PRINT_SENT_TO_SUPPLIER}}': parseSheetDate(printSentToSupplier),
    '{{PRINT_SENT_TO_SUP}}':      parseSheetDate(printSentToSupplier),
    '{{TP_SENT_TO_SUPPLIER}}':    parseSheetDate(tpSentToSupplier),
    '{{SAMPLE_DUE_DATE}}':        sampleDueDate,
    '{{SAMPLE_SIZE}}':            'SMALL 4/6',
    '{{TECHNICAL_DESIGNER}}':     'bertha@creativetwotwelve.com',
    '{{CAD_URL}}':                cadUrl || '',
    '{{CAD_IMAGE}}':              '', // cleared if no actual image is inserted
  };

  // Merge text runs within a paragraph so split tokens get joined before replacement.
  const mergeRunsInPara = paraXml => {
    const runMatches = [];
    const runRe = /<a:r>([\s\S]*?)<\/a:r>/g;
    let m;
    while ((m = runRe.exec(paraXml)) !== null)
      runMatches.push({ start: m.index, end: m.index + m[0].length, inner: m[1] });
    if (runMatches.length < 2) return paraXml;

    const texts = runMatches.map(r => {
      const tm = r.inner.match(/<a:t[^>]*>([\s\S]*?)<\/a:t>/);
      return tm ? tm[1] : '';
    });
    const combined = texts.join('');
    if (!Object.keys(tokens).some(k => combined.includes(k))) return paraXml;

    const rPrM = runMatches[0].inner.match(/<a:rPr[\s\S]*?(?:\/>|<\/a:rPr>)/);
    const rPr  = rPrM ? rPrM[0] : '';
    const merged = `<a:r>${rPr}<a:t>${combined}</a:t></a:r>`;
    return (
      paraXml.slice(0, runMatches[0].start) +
      merged +
      paraXml.slice(runMatches[runMatches.length - 1].end)
    );
  };

  const replaceTokens = xml => {
    let out = xml.replace(/<a:p(?:\s[^>]*)?>[\s\S]*?<\/a:p>/g, mergeRunsInPara);
    for (const [ph, val] of Object.entries(tokens)) {
      const esc = ph.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      out = out.replace(new RegExp(esc, 'g'), escapeXml(String(val ?? '')));
    }
    return out;
  };

  try {
    const sa = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth = new google.auth.GoogleAuth({
      credentials: sa,
      scopes: ['https://www.googleapis.com/auth/drive'],
    });
    const authClient = await auth.getClient();
    const drive = google.drive({ version: 'v3', auth: authClient });

    // 1. Export the template as a PPTX binary
    let exportResp;
    try {
      exportResp = await drive.files.export(
        { fileId: templateId, mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
        { responseType: 'arraybuffer' }
      );
    } catch (e) {
      const detail = e.response?.data?.error?.message || e.message;
      console.error('[rebeca/generate-tp] export failed:', detail);
      return res.status(500).json({ error: `Export failed: ${detail}` });
    }

    // 2. Open the ZIP (PPTX is a ZIP archive)
    const JSZip = require('jszip');
    let zip;
    try {
      const rawExport = exportResp.data;
      zip = await JSZip.loadAsync(Buffer.isBuffer(rawExport) ? rawExport : Buffer.from(rawExport));
    } catch (e) {
      console.error('[rebeca/generate-tp] zip load failed:', e.message);
      return res.status(500).json({ error: `ZIP load failed: ${e.message}` });
    }

    // 3. Download the CAD image via Drive API
    let cadImgBuf = null, cadImgMime = 'image/png', cadImgExt = 'png';
    if (cadFileId) {
      try {
        const imgResp = await drive.files.get(
          { fileId: cadFileId, alt: 'media' },
          { responseType: 'arraybuffer' }
        );
        cadImgBuf = Buffer.from(imgResp.data);
        const ct  = (imgResp.headers?.['content-type'] || 'image/png').split(';')[0].trim();
        cadImgMime = ct;
        cadImgExt  = (ct.includes('jpg') || ct.includes('jpeg')) ? 'jpg' : 'png';
      } catch (e) {
        console.warn('[rebeca/generate-tp] CAD image download failed:', e.message);
      }
    }

    // 4. Process each slide XML
    const INCH = 914400; // EMU per inch
    let imgIdx = 500;
    let ctNeedsUpdate = false;

    const slideKeys = Object.keys(zip.files)
      .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
      .sort((a, b) => {
        const na = parseInt(a.match(/slide(\d+)/)[1]);
        const nb = parseInt(b.match(/slide(\d+)/)[1]);
        return na - nb;
      });

    for (const slidePath of slideKeys) {
      let xml = await zip.files[slidePath].async('string');
      const slideNum = slidePath.match(/slide(\d+)\.xml$/)[1];
      const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
      let relsXml = zip.files[relsPath]
        ? await zip.files[relsPath].async('string')
        : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

      // Replace {{CAD_IMAGE}} placeholder shape with a real image element
      let cadInserted = false, cadRId = '', cadName = '';
      if (cadImgBuf) {
        xml = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, shapeXml => {
          if (cadInserted) return shapeXml;
          const texts = [];
          const aRe = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g;
          let am;
          while ((am = aRe.exec(shapeXml)) !== null) texts.push(am[1]);
          if (texts.join('').trim() !== '{{CAD_IMAGE}}') return shapeXml;

          cadInserted = true;
          imgIdx++;
          cadName = `image${imgIdx}.${cadImgExt}`;
          cadRId  = `rIdImg${imgIdx}`;
          zip.file(`ppt/media/${cadName}`, cadImgBuf);
          ctNeedsUpdate = true;

          const x = Math.round(0.97 * INCH);
          const y = Math.round(1.61 * INCH);
          const w = Math.round(2.59 * INCH);
          const h = Math.round(7.00 * INCH);
          return (
            `<p:pic>` +
            `<p:nvPicPr>` +
              `<p:cNvPr id="${9900 + imgIdx}" name="${cadName}"/>` +
              `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
              `<p:nvPr/>` +
            `</p:nvPicPr>` +
            `<p:blipFill>` +
              `<a:blip r:embed="${cadRId}"/>` +
              `<a:stretch><a:fillRect/></a:stretch>` +
            `</p:blipFill>` +
            `<p:spPr>` +
              `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
              `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
            `</p:spPr>` +
            `</p:pic>`
          );
        });

        if (cadInserted) {
          relsXml = relsXml.replace(
            '</Relationships>',
            `<Relationship Id="${cadRId}" ` +
            `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ` +
            `Target="../media/${cadName}"/></Relationships>`
          );
          zip.file(relsPath, relsXml);
        }
      }

      // Text token replacement
      xml = replaceTokens(xml);
      zip.file(slidePath, xml);
    }

    // Update [Content_Types].xml for the new image extension if needed
    if (ctNeedsUpdate) {
      const ctPath = '[Content_Types].xml';
      if (zip.files[ctPath]) {
        let ctXml = await zip.files[ctPath].async('string');
        if (!ctXml.includes(`Extension="${cadImgExt}"`)) {
          ctXml = ctXml.replace('</Types>', `<Default Extension="${cadImgExt}" ContentType="${cadImgMime}"/></Types>`);
          zip.file(ctPath, ctXml);
        }
      }
    }

    // 5. Generate the modified PPTX as a Buffer
    const pptxBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });

    const safe = s => (s || '').replace(/[\\/:*?"<>|#\[\]\r\n]/g, '').trim();
    const cleanStyle = (style || '').replace(/^[A-Za-z]+-?/, '');
    const fileName = `${safe(norm(model))} - ${safe(supplier || '')} - ${safe(cleanStyle)}`;
    const pptxMime = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

    // 6. Copy the template into the folder first (copy always works; avoids
    //    Shared Drive create-with-parents issues).
    let copyId;
    try {
      const copyRes = await drive.files.copy({
        fileId: templateId,
        supportsAllDrives: true,
        requestBody: { name: fileName, parents: [folderId] },
      });
      copyId = copyRes.data.id;
    } catch (e) {
      const detail = e.response?.data?.error?.message || e.message;
      return res.status(500).json({ error: `Copy failed: ${detail}` });
    }

    // 7. Update the copy's content with the processed PPTX (resumable PATCH).
    //    No parents needed — file already exists in the right folder.
    const { token: accessToken } = await authClient.getAccessToken();

    const initResp = await fetch(
      `https://www.googleapis.com/upload/drive/v3/files/${copyId}?uploadType=resumable&supportsAllDrives=true&fields=id`,
      {
        method: 'PATCH',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json; charset=UTF-8',
          'X-Upload-Content-Type': pptxMime,
          'X-Upload-Content-Length': String(pptxBuf.length),
        },
        body: '{}',
      }
    );
    if (!initResp.ok) {
      const d = await initResp.json().catch(() => ({}));
      return res.status(500).json({ error: `Update init failed: ${d?.error?.message || initResp.statusText}` });
    }
    const uploadUri = initResp.headers.get('location');
    if (!uploadUri) return res.status(500).json({ error: 'No upload URI from Drive' });

    const putResp = await fetch(uploadUri, {
      method: 'PUT',
      headers: { 'Content-Type': pptxMime, 'Content-Length': String(pptxBuf.length) },
      body: pptxBuf,
    });
    const putData = await putResp.json().catch(() => ({}));
    if (!putResp.ok) {
      const detail = putData?.error?.message || putResp.statusText;
      return res.status(500).json({ error: `Content update failed: ${detail}` });
    }

    const newId = putData.id || copyId;
    res.json({
      ok: true,
      presentationUrl: `https://docs.google.com/presentation/d/${newId}/edit`,
      folderUrl: `https://drive.google.com/drive/folders/${folderId}`,
    });
  } catch (e) {
    const detail = e.response?.data?.error?.message || e.message;
    console.error('[rebeca/generate-tp]', detail);
    res.status(500).json({ error: detail });
  }
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
    const colIdx   = (name) => headers.findIndex(h => h.toUpperCase() === name.toUpperCase());
    const styleCol = colIdx('STYLE #');
    const tpCol    = colIdx('TP SENT TO SUPPLIER');
    const printCol = colIdx('PRINT SENT TO SUPPLIER');

    // Fuzzy-find NDC column and optional separate YEAR column
    const ndcCol  = headers.findIndex(h => h.toUpperCase().replace(/\s+/g,'').includes('NDC'));
    const yearCol = headers.findIndex(h => h.toUpperCase().replace(/\s+/g,'') === 'YEAR' ||
                                          h.toUpperCase().replace(/\s+/g,'') === 'NDCYEAR');

    // Parse month number (1-12) and year from an NDC cell value
    const MON_MAP = {
      JAN:1,JANUARY:1, FEB:2,FEBRUARY:2, MAR:3,MARCH:3, APR:4,APRIL:4,
      MAY:5, JUN:6,JUNE:6, JUL:7,JULY:7, AUG:8,AUGUST:8,
      SEP:9,SEPT:9,SEPTEMBER:9, OCT:10,OCTOBER:10, NOV:11,NOVEMBER:11, DEC:12,DECEMBER:12,
    };
    const parseNdc = (ndcRaw, yearRaw) => {
      const s = String(ndcRaw || '').trim();
      if (!s) return null;
      // Combined "05/2026" or "May/2026" or "May 2026" or "May-2026"
      const withYear = s.match(/^([A-Za-z]+|\d{1,2})[\s\/\-](\d{4})$/);
      if (withYear) {
        const monPart = withYear[1], yr = parseInt(withYear[2]);
        const mon = /^\d+$/.test(monPart) ? parseInt(monPart) : (MON_MAP[monPart.toUpperCase()] || 0);
        return mon ? { mon, yr } : null;
      }
      // Month only: "May", "MAY", "05", "5"
      const mon = /^\d{1,2}$/.test(s)
        ? parseInt(s)
        : (MON_MAP[s.toUpperCase()] || 0);
      if (!mon) return null;
      const yr = yearRaw ? parseInt(String(yearRaw).trim()) : NaN;
      return { mon, yr };
    };

    const MONTH_NAMES = ['','JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

    const styleRows = rows.slice(1).filter(r => String(r[styleCol] || '').trim());

    // Build month map — include all 2026 rows
    const monthMap = {};
    let total = 0, missingTp = 0, missingPrint = 0;
    styleRows.forEach(r => {
      const ndcRaw  = ndcCol  >= 0 ? r[ndcCol]  : '';
      const yearRaw = yearCol >= 0 ? r[yearCol] : '';
      const parsed  = parseNdc(ndcRaw, yearRaw);
      if (!parsed || parsed.yr !== 2026) return; // only 2026

      total++;
      const noTp    = !String(r[tpCol]    || '').trim();
      const noPrint = !String(r[printCol] || '').trim();
      if (noTp)    missingTp++;
      if (noPrint) missingPrint++;

      const key = parsed.mon; // group by month number
      if (!monthMap[key]) monthMap[key] = { count: 0, missingTp: 0, missingPrint: 0 };
      monthMap[key].count++;
      if (noTp)    monthMap[key].missingTp++;
      if (noPrint) monthMap[key].missingPrint++;
    });

    // Always show the 4 calendar months starting from next month
    const nowMonth = new Date().getMonth() + 1; // 1-12
    const months = [1, 2, 3, 4].map(i => {
      const m = (nowMonth % 12) + i; // next month, then +1, +2, +3
      const mm = m > 12 ? m - 12 : m;  // wrap Dec→Jan
      return {
        label: MONTH_NAMES[mm],
        sortKey: mm,
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

// CAD Image proxy (bypasses Google Drive embed restrictions)
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
