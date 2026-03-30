/* ===== File Processing ===== */

// Server URL — derived from how the page was loaded
let SERVER_URL = '';

async function detectServer() {
  // If the page is served by Flask, window.location.origin IS the server
  if (window.location.protocol === 'http:' || window.location.protocol === 'https:') {
    const origin = window.location.origin; // e.g. http://127.0.0.1:5002
    try {
      const resp = await fetch(`${origin}/api/health`, { signal: AbortSignal.timeout(2000) });
      if (resp.ok) {
        SERVER_URL = origin;
        console.log(`Server detected at ${SERVER_URL} (from window.location)`);
        return true;
      }
    } catch (e) {}
  }

  // Fallback: scan common ports (for file:// or dev scenarios)
  for (const port of [5000, 5001, 5002, 5003, 5004, 5005]) {
    try {
      const resp = await fetch(`http://127.0.0.1:${port}/api/health`, { signal: AbortSignal.timeout(1000) });
      if (resp.ok) {
        SERVER_URL = `http://127.0.0.1:${port}`;
        console.log(`Server detected at ${SERVER_URL} (port scan)`);
        return true;
      }
    } catch (e) {}
  }
  console.warn('No local server detected');
  return false;
}

async function processFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  APP.fileContent = new Uint8Array(arrayBuffer);
  APP.uploadedFile = file;
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
  const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
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
      if (rowStr.trim()) textContent += `Ligne ${r + 1}: ${rowStr}\n`;
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
      textContent += `--- Page ${i} ---\n${content.items.map(item => item.str).join(' ')}\n\n`;
      const viewport = page.getViewport({ scale: 1.5 });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
      APP.pdfPageImages.push(canvas.toDataURL('image/png').split(',')[1]);
    }

    try {
      const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
      const fields = pdfDoc.getForm().getFields();
      if (fields.length > 0) {
        hasFormFields = true;
        textContent += '\n=== Champs de formulaire PDF ===\n';
        fields.forEach(f => {
          let val = '';
          try { val = f.getText?.() || ''; } catch(e) {}
          textContent += `[${f.getName()}] = "${val}"\n`;
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
  for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

/* ===== Collect cell updates ===== */

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
   GÉNÉRATION EXCEL — via serveur Flask + openpyxl
   ==================================================================== */

async function generateCompletedExcel() {
  if (!APP.analysisResult || !APP.uploadedFile) return;

  const updates = collectCellUpdates();
  console.log('Updates to send:', updates.length);

  if (!SERVER_URL) {
    showToast('Serveur local non détecté. Génération du récap PDF...', 'warning');
    await generateRecapPDF();
    return;
  }

  try {
    showLoading('Génération du fichier complété...');

    const formData = new FormData();
    formData.append('file', APP.uploadedFile);
    formData.append('updates', JSON.stringify(updates));

    const resp = await fetch(`${SERVER_URL}/api/fill-excel`, {
      method: 'POST',
      body: formData,
    });

    hideLoading();

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err.error || `Server error ${resp.status}`);
    }

    const blob = await resp.blob();
    const outName = APP.uploadedFileName.replace(/\.xlsx?$/i, '_complété.xlsx');
    downloadBlob(blob, outName);
    showToast('Document Excel complété téléchargé !');

  } catch (e) {
    hideLoading();
    console.error('fill-excel error:', e);
    showToast('Erreur: ' + e.message + '. Génération du récap PDF...', 'error');
    await generateRecapPDF();
  }
}

/* ===== PDF Form Fill via server ===== */

async function generateCompletedPDFForm() {
  if (!APP.analysisResult || !APP.uploadedFile) return;

  const updates = [];
  (APP.analysisResult.champs || []).forEach((champ, idx) => {
    const fieldName = champ.cellule_ou_position;
    const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
    if (fieldName && value) updates.push({ fieldName, value });
  });

  if (SERVER_URL) {
    try {
      const formData = new FormData();
      formData.append('file', APP.uploadedFile);
      formData.append('updates', JSON.stringify(updates));

      const resp = await fetch(`${SERVER_URL}/api/fill-pdf`, {
        method: 'POST',
        body: formData,
      });

      if (resp.ok) {
        const blob = await resp.blob();
        downloadBlob(blob, APP.uploadedFileName.replace(/\.pdf$/i, '_complété.pdf'));
        showToast('PDF complété téléchargé !');
        return;
      }
    } catch (e) {
      console.warn('Server PDF fill failed, using client-side:', e);
    }
  }

  // Fallback client-side via pdf-lib
  try {
    const pdfDoc = await PDFLib.PDFDocument.load(APP.fileContent);
    const form = pdfDoc.getForm();
    for (const u of updates) {
      try { const f = form.getTextField(u.fieldName); if (f) f.setText(u.value); } catch(e) {}
    }
    const pdfBytes = await pdfDoc.save();
    downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }),
      APP.uploadedFileName.replace(/\.pdf$/i, '_complété.pdf'));
  } catch(e) {
    showToast('Erreur PDF: ' + e.message, 'error');
  }
}

/* ===== PDF Récapitulatif via server or client ===== */

async function generateRecapPDF() {
  const fr = APP.companyData?.entites?.france;
  const ct = fr?.contacts?.service_comptabilite_facturation;

  const fields = [];
  (APP.analysisResult.champs || []).forEach((champ, idx) => {
    const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
    fields.push({ label: champ.label_original || champ.categorie_identifiee || '?', value });
  });
  (APP.analysisResult.champs_inconnus || []).forEach((u, idx) => {
    const val = APP.fieldValues[`unknown_${idx}`] || '';
    fields.push({ label: u.label_original || '?', value: val });
  });

  // Try server-side generation (nicer output)
  if (SERVER_URL) {
    try {
      const resp = await fetch(`${SERVER_URL}/api/recap-pdf`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: APP.uploadedFileName,
          companyName: `${fr?.nom_juridique || ''} (${fr?.nom_commercial || ''})`,
          fields,
          contact: {
            name: `${ct?.prenom || ''} ${ct?.nom || ''}`,
            role: ct?.intitule_poste || '',
            email: ct?.email || '',
            phone: ct?.telephone || '',
          }
        })
      });
      if (resp.ok) {
        const blob = await resp.blob();
        downloadBlob(blob, APP.uploadedFileName.replace(/\.[^.]+$/, '_recap.pdf'));
        showToast('Récapitulatif PDF téléchargé !');
        return;
      }
    } catch (e) {
      console.warn('Server recap PDF failed, using client-side:', e);
    }
  }

  // Fallback: client-side PDF via pdf-lib
  try {
    const pdfDoc = await PDFLib.PDFDocument.create();
    const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    const bold = await pdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);
    const orange = PDFLib.rgb(0.957, 0.455, 0.345);
    const gray = PDFLib.rgb(0.55, 0.57, 0.65);
    const white = PDFLib.rgb(0.93, 0.94, 0.96);

    let page = pdfDoc.addPage([595, 842]);
    let y = 790;
    const m = 45;
    const lh = 16;

    page.drawText('FICHE FOURNISSEUR — RÉCAPITULATIF', { x: m, y, size: 14, font: bold, color: orange });
    y -= 22;
    page.drawText(`${fr?.nom_juridique || ''} (${fr?.nom_commercial || ''})`, { x: m, y, size: 10, font: bold, color: white });
    y -= 14;
    page.drawText(`Fichier : ${APP.uploadedFileName} — ${new Date().toLocaleDateString('fr-FR')}`, { x: m, y, size: 8, font, color: gray });
    y -= 20;

    for (const f of fields) {
      if (y < 50) { page = pdfDoc.addPage([595, 842]); y = 790; }
      const lbl = (f.label || '').substring(0, 45);
      const val = (f.value || '—').substring(0, 55);
      page.drawText(lbl, { x: m, y, size: 8.5, font: bold, color: white });
      page.drawText(val, { x: m + 230, y, size: 8.5, font, color: white });
      y -= lh;
    }

    y -= 10;
    if (y < 50) { page = pdfDoc.addPage([595, 842]); y = 790; }
    page.drawText(`${ct?.prenom || ''} ${ct?.nom || ''} — ${ct?.intitule_poste || ''}`, { x: m, y, size: 8.5, font: bold, color: white });
    y -= 12;
    page.drawText(`${ct?.email || ''} | ${ct?.telephone || ''}`, { x: m, y, size: 8, font, color: gray });

    const pdfBytes = await pdfDoc.save();
    downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }),
      APP.uploadedFileName.replace(/\.[^.]+$/, '_recap.pdf'));
    showToast('Récapitulatif PDF téléchargé !');
  } catch(e) {
    console.error('Recap PDF error:', e);
    showToast('Erreur PDF: ' + e.message, 'error');
  }
}

/* ===== Email Recap HTML ===== */

function generateEmailRecap() {
  if (!APP.analysisResult) return '';
  const fr = APP.companyData?.entites?.france;
  const ct = fr?.contacts?.service_comptabilite_facturation;

  let html = `<div style="font-family:Arial,sans-serif;max-width:600px;">`;
  html += `<h2 style="color:#F47458;border-bottom:2px solid #F47458;padding-bottom:8px;">Fiche Fournisseur — ${fr?.nom_juridique || ''}</h2>`;
  html += `<p style="color:#666;font-size:12px;">Fichier : ${APP.uploadedFileName} — ${new Date().toLocaleDateString('fr-FR')}</p>`;
  html += `<table style="width:100%;border-collapse:collapse;font-size:13px;">`;

  (APP.analysisResult.champs || []).forEach((champ, idx) => {
    const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
    const label = champ.label_original || champ.categorie_identifiee;
    html += `<tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;font-weight:bold;color:#333;width:40%;">${escapeHtml(label)}</td><td style="padding:6px 8px;color:#111;">${escapeHtml(value) || '<em style="color:#999;">—</em>'}</td></tr>`;
  });
  html += `</table>`;

  html += `<hr style="margin-top:16px;border:none;border-top:1px solid #ddd;">`;
  html += `<p style="font-size:12px;color:#666;">Cordialement,<br><strong>${ct?.prenom || ''} ${ct?.nom || ''}</strong> — ${ct?.intitule_poste || ''}<br>${ct?.email || ''} | ${ct?.telephone || ''}</p>`;
  html += `</div>`;
  return html;
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
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
