/* ===== File Processing ===== */

async function processFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  APP.fileContent = arrayBuffer;
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
  APP.fileType = 'xlsx';
  // Read with full style/format preservation
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true, cellFormula: true, cellDates: true });
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
  if (!APP.fileContent || !APP.analysisResult) return;

  try {
    // Load the original xlsx as a ZIP to preserve ALL formatting
    const zip = await JSZip.loadAsync(APP.fileContent);

    // Collect all values to write: { sheetIndex: { cellRef: value } }
    const cellUpdates = {};
    APP.analysisResult.champs.forEach((champ, idx) => {
      const cellRef = (champ.cellule_ou_position || '').trim().toUpperCase();
      if (!cellRef || !/^[A-Z]+\d+$/.test(cellRef)) return;
      const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
      if (!value) return;
      // Default to sheet 1 (index 0) — most forms use the first sheet
      if (!cellUpdates[0]) cellUpdates[0] = {};
      cellUpdates[0][cellRef] = value;
    });

    // Find all sheet XML files in the zip
    const sheetFiles = [];
    zip.folder('xl/worksheets').forEach((path, file) => {
      if (path.match(/^sheet\d+\.xml$/)) {
        sheetFiles.push({ path: 'xl/worksheets/' + path, file });
      }
    });
    sheetFiles.sort((a, b) => a.path.localeCompare(b.path));

    // Also load shared strings table
    let sstXml = null;
    let sstEntries = [];
    const sstFile = zip.file('xl/sharedStrings.xml');
    if (sstFile) {
      sstXml = await sstFile.async('string');
      // Parse existing shared strings
      const siMatches = sstXml.match(/<si>([\s\S]*?)<\/si>/g) || [];
      sstEntries = siMatches.map(si => si);
    }

    // Process each sheet that has updates
    for (const [sheetIdx, updates] of Object.entries(cellUpdates)) {
      const idx = parseInt(sheetIdx);
      if (idx >= sheetFiles.length) continue;

      let sheetXml = await sheetFiles[idx].file.async('string');

      for (const [cellRef, value] of Object.entries(updates)) {
        const col = cellRef.replace(/\d+/g, '');
        const row = parseInt(cellRef.replace(/[A-Z]+/g, ''));

        // Try to find existing cell and update it
        // Match <c r="B11" ...>...</c> or <c r="B11" ... />
        const cellRegex = new RegExp(`(<c\\s[^>]*r="${cellRef}"[^>]*)(/>|>([\\s\\S]*?)</c>)`, 'i');
        const cellMatch = sheetXml.match(cellRegex);

        if (cellMatch) {
          // Cell exists — update its value, change type to inlineStr
          let cellTag = cellMatch[1];
          // Remove existing type attribute and add t="inlineStr"
          cellTag = cellTag.replace(/\s+t="[^"]*"/, '');
          cellTag = cellTag.replace(/\s+s="(\d+)"/, ' s="$1"'); // keep style
          const newCell = `${cellTag} t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
          sheetXml = sheetXml.replace(cellRegex, newCell);
        } else {
          // Cell doesn't exist — need to insert it in the correct row
          const rowRegex = new RegExp(`(<row\\s[^>]*r="${row}"[^>]*)(/>|>([\\s\\S]*?)</row>)`, 'i');
          const rowMatch = sheetXml.match(rowRegex);

          if (rowMatch) {
            const newCellXml = `<c r="${cellRef}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
            if (rowMatch[2] === '/>') {
              // Empty row tag — open it and add cell
              sheetXml = sheetXml.replace(rowRegex, `${rowMatch[1]}>${newCellXml}</row>`);
            } else {
              // Row has children — append cell before </row>
              sheetXml = sheetXml.replace(rowRegex, `${rowMatch[1]}>${rowMatch[3]}${newCellXml}</row>`);
            }
          } else {
            // Row doesn't exist — insert a new row in sheetData
            const newRowXml = `<row r="${row}"><c r="${cellRef}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c></row>`;
            sheetXml = sheetXml.replace('</sheetData>', `${newRowXml}</sheetData>`);
          }
        }
      }

      zip.file(sheetFiles[idx].path, sheetXml);
    }

    // Generate the modified zip
    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE'
    });

    downloadBlob(blob, APP.uploadedFileName.replace(/\.xlsx?$/i, '_complété.xlsx'));
    showToast('Document Excel complété téléchargé !');
  } catch (e) {
    console.error('Erreur génération Excel:', e);
    showToast('Erreur génération Excel: ' + e.message, 'error');
  }
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
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
