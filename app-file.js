/* ===== File Processing ===== */

async function processFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  APP.fileContent = new Uint8Array(arrayBuffer);
  APP.uploadedFileName = file.name;
  const ext = file.name.split('.').pop().toLowerCase();

  if (ext === 'xlsx' || ext === 'xls') {
    return processExcel(arrayBuffer, file.name);
  } else if (ext === 'pdf') {
    return processPDF(arrayBuffer, file.name);
  } else if (['png', 'jpg', 'jpeg', 'tiff'].includes(ext)) {
    APP.fileType = 'pdf-scan';
    const base64 = arrayBufferToBase64(arrayBuffer);
    return analyzeWithClaude(base64, file.name, 'pdf-scan', '');
  }
  showToast('Format de fichier non supporté.', 'error');
  return null;
}

async function processExcel(arrayBuffer, fileName) {
  APP.fileType = fileName.split('.').pop().toLowerCase() === 'xls' ? 'xls' : 'xlsx';
  const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array', cellStyles: true, cellFormula: true, cellDates: true });
  APP.workbook = workbook;

  let textContent = '';
  workbook.SheetNames.forEach(name => {
    if (name.toLowerCase() === 'menus') return;
    const sheet = workbook.Sheets[name];
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    textContent += `=== Feuille: ${name} ===\n`;
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[addr];
        row.push(cell ? String(cell.v ?? '') : '');
      }
      const rowStr = row.join(' | ');
      if (rowStr.trim()) {
        textContent += `Ligne ${r + 1}: ${rowStr}\n`;
      }
    }
    textContent += '\n';
  });

  return analyzeWithClaude('', fileName, 'xlsx', textContent);
}

async function processPDF(arrayBuffer, fileName) {
  const uint8 = new Uint8Array(arrayBuffer);
  try {
    const pdf = await pdfjsLib.getDocument({ data: uint8 }).promise;
    let textContent = '';
    let hasFormFields = false;
    APP.pdfPageImages = [];

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const pageText = content.items.map(item => item.str).join(' ');
      textContent += `--- Page ${i} ---\n${pageText}\n\n`;
      const viewport = page.getViewport({ scale: 1.5 });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport }).promise;
      APP.pdfPageImages.push(canvas.toDataURL('image/png').split(',')[1]);
    }

    try {
      const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
      const form = pdfDoc.getForm();
      const fields = form.getFields();
      if (fields.length > 0) {
        hasFormFields = true;
        textContent += '\n=== Champs de formulaire PDF ===\n';
        fields.forEach(field => {
          const name = field.getName();
          const type = field.constructor.name;
          let value = '';
          try { value = field.getText?.() || ''; } catch(e) {}
          textContent += `[${name}] (type: ${type}) = "${value}"\n`;
        });
      }
    } catch(e) {}

    APP.fileType = hasFormFields ? 'pdf-form' : 'pdf-scan';
    return analyzeWithClaude('', fileName, APP.fileType, textContent);
  } catch(e) {
    console.error('Erreur traitement PDF:', e);
    showToast('Erreur lors du traitement du PDF: ' + e.message, 'error');
    return null;
  }
}

function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

/* ===== Collect cell updates from validated fields ===== */

function collectCellUpdates() {
  const updates = [];
  (APP.analysisResult.champs || []).forEach((champ, idx) => {
    let raw = (champ.cellule_ou_position || '').trim();
    if (!raw) return;
    const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
    if (!value) return;

    let sheetName = null;
    let cellRef = raw;
    if (raw.includes('!')) {
      const parts = raw.split('!');
      sheetName = parts[0].replace(/^'|'$/g, '');
      cellRef = parts[1];
    }
    cellRef = cellRef.split(/[-:]/)[0].trim().toUpperCase();
    if (!/^[A-Z]+\d+$/.test(cellRef)) return;

    if (sheetName) {
      updates.push({ sheetName, cellRef, value });
    } else {
      APP.workbook.SheetNames.forEach(n => {
        if (!['menus', 'menu', 'listes', 'lists'].includes(n.toLowerCase())) {
          updates.push({ sheetName: n, cellRef, value });
        }
      });
    }
  });
  return updates;
}

/* ====================================================================
   APPROCHE 1 — ZIP/XML directe (.xlsx uniquement)
   Ouvre le fichier original comme ZIP, modifie le XML des cellules,
   ne touche à rien d'autre → formatage 100% préservé
   ==================================================================== */

async function generateCompletedExcel() {
  if (!APP.analysisResult) return;

  const updates = collectCellUpdates();
  console.log('Updates to apply:', updates.length);

  // Pour .xlsx : approche ZIP/XML directe
  if (APP.fileType === 'xlsx' && APP.fileContent) {
    try {
      const blob = await fillXlsxViaZip(APP.fileContent, updates);
      downloadBlob(blob, APP.uploadedFileName.replace(/\.xlsx$/i, '_complété.xlsx'));
      showToast('Document Excel complété téléchargé !');
      return;
    } catch (e) {
      console.error('fillXlsxViaZip failed:', e);
      showToast('Erreur Excel ZIP, génération du PDF récapitulatif...', 'warning');
    }
  }

  // Fallback pour .xls ou si ZIP échoue : PDF récapitulatif
  await generateRecapPDF();
}

async function fillXlsxViaZip(originalBytes, updates) {
  const zip = await JSZip.loadAsync(originalBytes);

  // 1. Map sheet names to XML paths via workbook.xml + rels
  const sheetMap = await mapSheetNames(zip);
  console.log('Sheet map:', sheetMap);

  // 2. Group updates by XML path
  const byPath = {};
  for (const u of updates) {
    const path = sheetMap[u.sheetName.toLowerCase()];
    if (!path) { console.warn('No XML path for sheet:', u.sheetName); continue; }
    if (!byPath[path]) byPath[path] = [];
    byPath[path].push(u);
  }

  // 3. For each sheet XML, inject cell values
  for (const [xmlPath, cellUpdates] of Object.entries(byPath)) {
    const file = zip.file(xmlPath);
    if (!file) { console.warn('File not found in ZIP:', xmlPath); continue; }

    let xml = await file.async('string');

    // Detect namespace prefix (some files use <x:row> instead of <row>)
    const nsMatch = xml.match(/<(\w+):sheetData[\s>\/]/);
    const ns = nsMatch ? nsMatch[1] + ':' : '';

    console.log(`Processing ${xmlPath}, ns prefix="${ns}", updates=${cellUpdates.length}`);

    for (const { cellRef, value } of cellUpdates) {
      xml = injectCellValue(xml, cellRef, value, ns);
    }

    zip.file(xmlPath, xml);
  }

  // 4. Generate output
  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    compression: 'DEFLATE'
  });
}

async function mapSheetNames(zip) {
  const map = {};
  try {
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');

    // Parse rels: rId → target path
    const rels = {};
    const relRe = /Id="([^"]+)"[^>]*Target="([^"]+)"/g;
    let m;
    while ((m = relRe.exec(relsXml)) !== null) {
      rels[m[1]] = m[2].startsWith('/') ? m[2].substring(1) : 'xl/' + m[2];
    }

    // Parse sheets: name + rId
    const sheetRe = /name="([^"]+)"[^>]*r:id="([^"]+)"/gi;
    while ((m = sheetRe.exec(wbXml)) !== null) {
      const name = m[1];
      const rId = m[2];
      if (rels[rId]) map[name.toLowerCase()] = rels[rId];
    }
    // Try reverse attribute order too
    const sheetRe2 = /r:id="([^"]+)"[^>]*name="([^"]+)"/gi;
    while ((m = sheetRe2.exec(wbXml)) !== null) {
      const rId = m[1];
      const name = m[2];
      if (rels[rId] && !map[name.toLowerCase()]) map[name.toLowerCase()] = rels[rId];
    }
  } catch (e) {
    console.warn('mapSheetNames fallback:', e.message);
  }

  // Fallback: enumerate sheet files
  if (Object.keys(map).length === 0) {
    let idx = 1;
    while (zip.file(`xl/worksheets/sheet${idx}.xml`)) {
      map[`sheet${idx}`] = `xl/worksheets/sheet${idx}.xml`;
      idx++;
    }
  }
  return map;
}

/**
 * Inject a single cell value into sheet XML string.
 * Strategy: find the row, find or create the cell, write inlineStr value.
 * Uses indexOf-based scanning — no regex for row/cell finding.
 */
function injectCellValue(xml, cellRef, value, ns) {
  const rowNum = cellRef.replace(/[A-Z]+/g, '');
  const escaped = value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

  const rowTag = `<${ns}row`;
  const rowClose = `</${ns}row>`;
  const cellTag = `<${ns}c`;
  const cellClose = `</${ns}c>`;
  const sdClose = `</${ns}sheetData>`;

  const newCellStr = `<${ns}c r="${cellRef}" t="inlineStr"><${ns}is><${ns}t>${escaped}</${ns}t></${ns}is></${ns}c>`;

  // Find the row with r="rowNum"
  let pos = 0;
  let rowStart = -1;
  let rowEnd = -1;

  while (true) {
    const idx = xml.indexOf(rowTag, pos);
    if (idx === -1) break;

    // Extract the opening tag to check r attribute
    const tagEnd = xml.indexOf('>', idx);
    if (tagEnd === -1) break;
    const openTag = xml.substring(idx, tagEnd + 1);

    // Check if this row has r="rowNum"
    const rMatch = openTag.match(/\br="(\d+)"/);
    if (rMatch && rMatch[1] === rowNum) {
      rowStart = idx;
      // Find end of this row
      if (openTag.endsWith('/>')) {
        rowEnd = tagEnd + 1;
      } else {
        rowEnd = xml.indexOf(rowClose, tagEnd);
        if (rowEnd !== -1) rowEnd += rowClose.length;
      }
      break;
    }
    pos = tagEnd + 1;
  }

  if (rowStart === -1) {
    // Row doesn't exist — insert before </sheetData>
    const sdPos = xml.lastIndexOf(sdClose);
    if (sdPos === -1) {
      console.warn('Cannot find sheetData close tag for', cellRef);
      return xml;
    }
    const newRow = `<${ns}row r="${rowNum}">${newCellStr}</${ns}row>`;
    console.log(`  [NEW ROW] ${cellRef} = "${value}"`);
    return xml.substring(0, sdPos) + newRow + xml.substring(sdPos);
  }

  // Row exists — check if the row is self-closing
  const rowContent = xml.substring(rowStart, rowEnd);
  if (rowContent.trimEnd().endsWith('/>')) {
    // Self-closing row: <row r="11" ... /> → open it
    const selfClose = xml.substring(rowStart, rowEnd);
    const opened = selfClose.replace(/\s*\/>$/, `>${newCellStr}${rowClose}`);
    console.log(`  [OPEN ROW] ${cellRef} = "${value}"`);
    return xml.substring(0, rowStart) + opened + xml.substring(rowEnd);
  }

  // Row is open — look for existing cell with this ref
  const rowClosePos = xml.indexOf(rowClose, rowStart);
  const rowXml = xml.substring(rowStart, rowClosePos + rowClose.length);

  // Try to find existing cell: look for r="B11" in a <c tag
  const cellRefPattern = `r="${cellRef}"`;
  const cellIdx = rowXml.indexOf(cellRefPattern);

  if (cellIdx !== -1) {
    // Cell exists — find its boundaries and replace
    // Go backward to find the <c or <ns:c before this r="..."
    let cStart = rowXml.lastIndexOf(cellTag, cellIdx);
    if (cStart === -1) {
      console.warn('Found r= but no cell tag for', cellRef);
      return xml;
    }

    // Find end: either </c> or /> from the cell start
    let cEnd;
    const selfCloseIdx = rowXml.indexOf('/>', cStart);
    const contentCloseIdx = rowXml.indexOf(cellClose, cStart);

    if (selfCloseIdx !== -1 && (contentCloseIdx === -1 || selfCloseIdx < contentCloseIdx)) {
      // Check if the /> is BEFORE any > (meaning truly self-closing)
      const firstGt = rowXml.indexOf('>', cStart);
      if (firstGt === selfCloseIdx + 1 || rowXml[firstGt - 1] === '/') {
        cEnd = selfCloseIdx + 2;
      } else {
        cEnd = contentCloseIdx + cellClose.length;
      }
    } else if (contentCloseIdx !== -1) {
      cEnd = contentCloseIdx + cellClose.length;
    } else {
      console.warn('Cannot find cell end for', cellRef);
      return xml;
    }

    // Extract existing cell to preserve style
    const existingCell = rowXml.substring(cStart, cEnd);
    const styleMatch = existingCell.match(/\bs="(\d+)"/);
    const styleAttr = styleMatch ? ` s="${styleMatch[1]}"` : '';

    const replacement = `<${ns}c r="${cellRef}"${styleAttr} t="inlineStr"><${ns}is><${ns}t>${escaped}</${ns}t></${ns}is></${ns}c>`;

    const globalStart = rowStart + cStart;
    const globalEnd = rowStart + cEnd;
    console.log(`  [REPLACE] ${cellRef} = "${value}"`);
    return xml.substring(0, globalStart) + replacement + xml.substring(globalEnd);
  }

  // Cell doesn't exist in this row — insert before </row>
  console.log(`  [INSERT] ${cellRef} = "${value}"`);
  return xml.substring(0, rowClosePos) + newCellStr + xml.substring(rowClosePos);
}

/* ===== PDF Form Fill ===== */

async function generateCompletedPDFForm() {
  try {
    const pdfDoc = await PDFLib.PDFDocument.load(APP.fileContent);
    const form = pdfDoc.getForm();
    APP.analysisResult.champs.forEach((champ, idx) => {
      const fieldName = champ.cellule_ou_position;
      const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
      if (fieldName && value) {
        try { const f = form.getTextField(fieldName); if (f) f.setText(value); } catch(e) {}
      }
    });
    const pdfBytes = await pdfDoc.save();
    downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }),
      APP.uploadedFileName.replace(/\.pdf$/i, '_complété.pdf'));
  } catch(e) {
    showToast('Erreur génération PDF: ' + e.message, 'error');
  }
}

/* ===== PDF Récapitulatif ===== */

async function generateRecapPDF() {
  try {
    const pdfDoc = await PDFLib.PDFDocument.create();
    const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    const bold = await pdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);
    const orange = PDFLib.rgb(0.957, 0.455, 0.345);
    const gray = PDFLib.rgb(0.55, 0.57, 0.65);
    const dark = PDFLib.rgb(0.93, 0.94, 0.96);

    let page = pdfDoc.addPage([595, 842]);
    let y = 790;
    const m = 45;
    const lh = 16;
    const colVal = 240;

    const addPage = () => { page = pdfDoc.addPage([595, 842]); y = 790; };
    const checkY = () => { if (y < 50) addPage(); };

    // Header
    page.drawText('FICHE FOURNISSEUR — RÉCAPITULATIF', { x: m, y, size: 14, font: bold, color: orange });
    y -= 24;
    const fr = APP.companyData?.entites?.france;
    page.drawText(`${fr?.nom_juridique || ''} (${fr?.nom_commercial || ''})`, { x: m, y, size: 10, font: bold, color: dark });
    y -= 14;
    page.drawText(`Fichier traité : ${APP.uploadedFileName}  —  ${new Date().toLocaleDateString('fr-FR')}`, { x: m, y, size: 8, font, color: gray });
    y -= 6;
    page.drawRectangle({ x: m, y, width: 505, height: 1, color: orange });
    y -= 20;

    // Fields
    const allChamps = APP.analysisResult.champs || [];
    allChamps.forEach((champ, idx) => {
      checkY();
      const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
      const label = champ.label_original || champ.categorie_identifiee || '?';

      // Label
      const labelTrunc = label.length > 45 ? label.substring(0, 42) + '...' : label;
      page.drawText(labelTrunc, { x: m, y, size: 8.5, font: bold, color: dark });

      // Value
      if (value) {
        const valTrunc = value.length > 55 ? value.substring(0, 52) + '...' : value;
        page.drawText(valTrunc, { x: m + colVal, y, size: 8.5, font, color: dark });
      }
      y -= lh;
    });

    // Unknown fields
    const unknowns = APP.analysisResult.champs_inconnus || [];
    if (unknowns.length > 0) {
      checkY();
      y -= 8;
      page.drawText('CHAMPS SUPPLÉMENTAIRES', { x: m, y, size: 9, font: bold, color: orange });
      y -= lh;
      unknowns.forEach((u, idx) => {
        checkY();
        const val = APP.fieldValues[`unknown_${idx}`] || '';
        page.drawText(u.label_original || '?', { x: m, y, size: 8.5, font: bold, color: dark });
        if (val) page.drawText(val, { x: m + colVal, y, size: 8.5, font, color: dark });
        y -= lh;
      });
    }

    // Signature block
    checkY();
    y -= 12;
    page.drawRectangle({ x: m, y, width: 505, height: 1, color: gray });
    y -= 16;
    const ct = fr?.contacts?.service_comptabilite_facturation;
    page.drawText(`${ct?.prenom || ''} ${ct?.nom || ''} — ${ct?.intitule_poste || ''}`, { x: m, y, size: 8.5, font: bold, color: dark });
    y -= 12;
    page.drawText(`${ct?.email || ''}  |  ${ct?.telephone || ''}`, { x: m, y, size: 8, font, color: gray });

    const pdfBytes = await pdfDoc.save();
    downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }),
      APP.uploadedFileName.replace(/\.[^.]+$/, '_recap.pdf'));
  } catch(e) {
    console.error('Erreur PDF récap:', e);
    showToast('Erreur génération PDF: ' + e.message, 'error');
  }
}

/* ===== Email Recap HTML ===== */

function generateEmailRecap() {
  if (!APP.analysisResult) return '';
  const fr = APP.companyData?.entites?.france;
  const ct = fr?.contacts?.service_comptabilite_facturation;

  let html = `<div style="font-family:Arial,sans-serif;max-width:600px;">`;
  html += `<h2 style="color:#F47458;border-bottom:2px solid #F47458;padding-bottom:8px;">Fiche Fournisseur — ${fr?.nom_juridique || ''}</h2>`;
  html += `<p style="color:#666;font-size:12px;">Fichier traité : ${APP.uploadedFileName} — ${new Date().toLocaleDateString('fr-FR')}</p>`;
  html += `<table style="width:100%;border-collapse:collapse;font-size:13px;">`;

  (APP.analysisResult.champs || []).forEach((champ, idx) => {
    const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
    const label = champ.label_original || champ.categorie_identifiee;
    html += `<tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;font-weight:bold;color:#333;width:40%;">${label}</td><td style="padding:6px 8px;color:#111;">${value || '<em style="color:#999;">—</em>'}</td></tr>`;
  });

  html += `</table>`;

  const unknowns = APP.analysisResult.champs_inconnus || [];
  if (unknowns.length > 0) {
    html += `<h3 style="color:#F47458;margin-top:16px;">Champs supplémentaires</h3><table style="width:100%;border-collapse:collapse;font-size:13px;">`;
    unknowns.forEach((u, idx) => {
      const val = APP.fieldValues[`unknown_${idx}`] || '';
      html += `<tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;font-weight:bold;color:#333;width:40%;">${u.label_original}</td><td style="padding:6px 8px;">${val || '—'}</td></tr>`;
    });
    html += `</table>`;
  }

  html += `<hr style="margin-top:16px;border:none;border-top:1px solid #ddd;">`;
  html += `<p style="font-size:12px;color:#666;">Cordialement,<br><strong>${ct?.prenom || ''} ${ct?.nom || ''}</strong> — ${ct?.intitule_poste || ''}<br>${ct?.email || ''} | ${ct?.telephone || ''}</p>`;
  html += `</div>`;
  return html;
}

/* ===== Download helper ===== */

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
