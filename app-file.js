/* ===== File Processing ===== */

async function processFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  // Store a Uint8Array copy — ArrayBuffer can be detached by XLSX.read()
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

  // Read the workbook
  const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array', cellStyles: true, cellFormula: true, cellDates: true });
  APP.workbook = workbook;

  // Extract data validation (dropdown) options per cell
  APP.dataValidations = {};
  workbook.SheetNames.forEach(name => {
    const sheet = workbook.Sheets[name];
    // XLSX.js stores data validations in sheet['!dataValidation']
    const dvs = sheet['!dataValidation'] || [];
    dvs.forEach(dv => {
      if (dv.type === 'list' && dv.sqref && dv.formula1) {
        // Parse options from comma-separated formula or range
        let options = [];
        const f = dv.formula1;
        if (f.startsWith('"') && f.endsWith('"')) {
          // Inline list: "Option1,Option2,Option3"
          options = f.slice(1, -1).split(',').map(s => s.trim());
        } else if (f.includes('!')) {
          // Range reference — try to resolve from sheet
          try {
            const refSheet = f.split('!')[0].replace(/'/g, '');
            const refRange = f.split('!')[1];
            const ws = workbook.Sheets[refSheet] || sheet;
            const range = XLSX.utils.decode_range(refRange);
            for (let r = range.s.r; r <= range.e.r; r++) {
              for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                if (cell && cell.v) options.push(String(cell.v));
              }
            }
          } catch (e) {}
        }
        if (options.length > 0) {
          // Parse sqref (can be "B25" or "B25:B30")
          const refs = dv.sqref.split(/\s+/);
          refs.forEach(ref => {
            if (ref.includes(':')) {
              try {
                const range = XLSX.utils.decode_range(ref);
                for (let r = range.s.r; r <= range.e.r; r++) {
                  for (let c = range.s.c; c <= range.e.c; c++) {
                    APP.dataValidations[XLSX.utils.encode_cell({ r, c })] = options;
                  }
                }
              } catch (e) {}
            } else {
              APP.dataValidations[ref.toUpperCase()] = options;
            }
          });
        }
      }
    });

    // Also check for "Menus" sheet referenced in data validations
    if (workbook.Sheets['Menus']) {
      const menuSheet = workbook.Sheets['Menus'];
      const menuRange = XLSX.utils.decode_range(menuSheet['!ref'] || 'A1');
      // Each column in the Menus sheet is a dropdown list
      for (let c = menuRange.s.c; c <= menuRange.e.c; c++) {
        const headerCell = menuSheet[XLSX.utils.encode_cell({ r: 0, c })];
        if (!headerCell) continue;
        const options = [];
        for (let r = 1; r <= menuRange.e.r; r++) {
          const cell = menuSheet[XLSX.utils.encode_cell({ r, c })];
          if (cell && cell.v) options.push(String(cell.v));
        }
        if (options.length > 0) {
          // Store by header name for later matching
          APP.dataValidations[`_menu_${String(headerCell.v).trim()}`] = options;
        }
      }
    }
  });

  // Build text representation including dropdown info
  let textContent = '';
  workbook.SheetNames.forEach(name => {
    if (name === 'Menus') return; // Skip menu reference sheet
    const sheet = workbook.Sheets[name];
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    textContent += `=== Feuille: ${name} ===\n`;
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[addr];
        let cellText = cell ? String(cell.v ?? '') : '';
        // Add dropdown info
        if (APP.dataValidations[addr]) {
          cellText += ` [MENU DÉROULANT: ${APP.dataValidations[addr].join(', ')}]`;
        }
        row.push(cellText);
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
      const imgData = canvas.toDataURL('image/png').split(',')[1];
      APP.pdfPageImages.push(imgData);
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

    if (hasFormFields) {
      APP.fileType = 'pdf-form';
      return analyzeWithClaude('', fileName, 'pdf-form', textContent);
    } else if (textContent.trim().length > 100) {
      APP.fileType = 'pdf-scan';
      return analyzeWithClaude('', fileName, 'pdf-scan', textContent);
    } else {
      APP.fileType = 'pdf-scan';
      return analyzeWithClaude('', fileName, 'pdf-scan', textContent || '');
    }
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

/* ===== Document Generation ===== */

async function generateCompletedExcel() {
  if (!APP.workbook || !APP.analysisResult) return;

  try {
    // Build a fresh copy of the workbook to avoid any mutation issues
    const freshData = XLSX.write(APP.workbook, { bookType: 'xlsx', type: 'array' });
    const wb = XLSX.read(freshData, { type: 'array' });

    // Collect all updates
    const updates = collectCellUpdates();
    console.log('=== Cell updates to apply ===', updates);

    // Apply updates to the fresh workbook
    for (const { sheetName, cellRef, value } of updates) {
      const sheet = wb.Sheets[sheetName];
      if (!sheet) { console.warn('Sheet not found:', sheetName); continue; }
      XLSX.utils.sheet_add_aoa(sheet, [[value]], { origin: cellRef });
      console.log(`  Written: ${sheetName}!${cellRef} = "${sheet[cellRef]?.v}"`);
    }

    // Write the modified fresh workbook
    const bookType = APP.fileType === 'xls' ? 'xls' : 'xlsx';
    const mimeType = bookType === 'xls'
      ? 'application/vnd.ms-excel'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const wbout = XLSX.write(wb, { bookType, type: 'array' });
    const outputBlob = new Blob([wbout], { type: mimeType });

    downloadBlob(outputBlob, APP.uploadedFileName.replace(/\.xlsx?$/i, `_complété.${bookType}`));
    showToast('Document Excel complété téléchargé !');
  } catch (e) {
    console.error('Erreur génération Excel:', e);
    showToast('Erreur génération Excel: ' + e.message, 'error');
  }
}

// Collect all cell updates as a flat array
function collectCellUpdates() {
  const updates = [];
  const champs = APP.analysisResult.champs || [];

  champs.forEach((champ, idx) => {
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
      // Write to all data sheets
      APP.workbook.SheetNames.forEach(n => {
        if (!['menus', 'menu', 'listes', 'lists'].includes(n.toLowerCase())) {
          updates.push({ sheetName: n, cellRef, value });
        }
      });
    }
  });

  return updates;
}

// Restore formatting: copy styles/themes/images from original xlsx into new xlsx
async function restoreFormatting(newXlsxData, originalData) {
  try {
    const origZip = await JSZip.loadAsync(originalData);
    const newZip = await JSZip.loadAsync(newXlsxData);

    // Copy styles and theme
    for (const f of ['xl/styles.xml', 'xl/theme/theme1.xml']) {
      const origFile = origZip.file(f);
      if (origFile) {
        newZip.file(f, await origFile.async('uint8array'));
      }
    }

    // Copy media files (images, logos) — must await each one
    const mediaFiles = [];
    origZip.folder('xl/media')?.forEach((path) => { mediaFiles.push('xl/media/' + path); });
    for (const path of mediaFiles) {
      const f = origZip.file(path);
      if (f) newZip.file(path, await f.async('uint8array'));
    }

    // Copy drawings
    const drawingFiles = [];
    origZip.folder('xl/drawings')?.forEach((path) => { drawingFiles.push('xl/drawings/' + path); });
    for (const path of drawingFiles) {
      const f = origZip.file(path);
      if (f) newZip.file(path, await f.async('uint8array'));
    }

    // Copy drawing rels
    const drawingRels = [];
    origZip.folder('xl/drawings/_rels')?.forEach((path) => { drawingRels.push('xl/drawings/_rels/' + path); });
    for (const path of drawingRels) {
      const f = origZip.file(path);
      if (f) newZip.file(path, await f.async('uint8array'));
    }

    // Copy worksheet rels (for images linked to sheets)
    const wsRels = [];
    origZip.folder('xl/worksheets/_rels')?.forEach((path) => { wsRels.push('xl/worksheets/_rels/' + path); });
    for (const path of wsRels) {
      const f = origZip.file(path);
      if (f) newZip.file(path, await f.async('uint8array'));
    }

    return await newZip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE'
    });
  } catch (e) {
    console.warn('Could not restore formatting, using XLSX.js output:', e.message);
    return new Blob([newXlsxData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
}

async function generateCompletedPDFForm() {
  try {
    const pdfDoc = await PDFLib.PDFDocument.load(APP.fileContent);
    const form = pdfDoc.getForm();

    APP.analysisResult.champs.forEach((champ, idx) => {
      const fieldName = champ.cellule_ou_position;
      const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
      if (fieldName && value) {
        try {
          const field = form.getTextField(fieldName);
          if (field) field.setText(value);
        } catch(e) {}
      }
    });

    const pdfBytes = await pdfDoc.save();
    downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }),
      APP.uploadedFileName.replace(/\.pdf$/i, '_complété.pdf'));
  } catch(e) {
    showToast('Erreur génération PDF: ' + e.message, 'error');
  }
}

async function generateRecapPDF() {
  try {
    const pdfDoc = await PDFLib.PDFDocument.create();
    const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    const boldFont = await pdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);

    let page = pdfDoc.addPage([595, 842]);
    let y = 800;
    const margin = 50;
    const lineHeight = 18;

    page.drawText('Récapitulatif - Fiche Fournisseur', {
      x: margin, y, size: 16, font: boldFont, color: PDFLib.rgb(0.91, 0.32, 0.1)
    });
    y -= 30;

    page.drawText(`Fichier: ${APP.uploadedFileName}`, {
      x: margin, y, size: 10, font, color: PDFLib.rgb(0.4, 0.4, 0.4)
    });
    y -= 10;
    page.drawText(`Date: ${new Date().toLocaleDateString('fr-FR')}`, {
      x: margin, y, size: 10, font, color: PDFLib.rgb(0.4, 0.4, 0.4)
    });
    y -= 30;

    const allChamps = APP.analysisResult.champs || [];
    allChamps.forEach((champ, idx) => {
      if (y < 60) {
        page = pdfDoc.addPage([595, 842]);
        y = 800;
      }
      const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
      const label = champ.label_original || champ.categorie_identifiee || 'Champ inconnu';

      page.drawText(`${label}:`, { x: margin, y, size: 10, font: boldFont });
      page.drawText(value, { x: margin + 200, y, size: 10, font });
      y -= lineHeight;
    });

    const pdfBytes = await pdfDoc.save();
    downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }),
      APP.uploadedFileName.replace(/\.[^.]+$/, '_recap.pdf'));
  } catch(e) {
    console.error('Erreur génération récap PDF:', e);
    showToast('Erreur génération récap PDF: ' + e.message, 'error');
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function generateEmailRecap() {
  if (!APP.analysisResult) return '';
  const fr = APP.companyData?.entites?.france;
  const ct = fr?.contacts?.service_comptabilite_facturation;
  let text = `FICHE FOURNISSEUR - RÉCAPITULATIF\n`;
  text += `${'='.repeat(50)}\n\n`;
  text += `Entreprise : ${fr?.nom_juridique || ''} (${fr?.nom_commercial || ''})\n`;
  text += `Date : ${new Date().toLocaleDateString('fr-FR')}\n`;
  text += `Fichier traité : ${APP.uploadedFileName}\n\n`;
  text += `${'\u2500'.repeat(50)}\n\n`;

  const allChamps = APP.analysisResult.champs || [];
  allChamps.forEach((champ, idx) => {
    const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
    const label = champ.label_original || champ.categorie_identifiee;
    text += `${label} : ${value}\n`;
  });

  const unknown = APP.analysisResult.champs_inconnus || [];
  if (unknown.length > 0) {
    text += `\n${'\u2500'.repeat(50)}\nChamps supplémentaires :\n\n`;
    unknown.forEach((u, idx) => {
      const val = APP.fieldValues[`unknown_${idx}`] || '';
      text += `${u.label_original} : ${val}\n`;
    });
  }

  text += `\n${'\u2500'.repeat(50)}\n`;
  text += `Cordialement,\n${ct?.prenom || ''} ${ct?.nom || ''}`;
  text += `\n${ct?.email || ''}`;
  text += `\n${ct?.telephone || ''}`;

  return text;
}
