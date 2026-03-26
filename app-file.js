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
  if (!APP.analysisResult) return;

  // For .xls (old binary format), use XLSX.js workbook modification
  // For .xlsx, use JSZip to preserve full formatting
  if (APP.fileType === 'xls') {
    return generateCompletedXls();
  }

  if (!APP.fileContent) return;
  try {
    const zip = await JSZip.loadAsync(APP.fileContent);

    // 1. Build sheet name → XML file path mapping from workbook.xml.rels + workbook.xml
    const sheetMap = await buildSheetMap(zip);

    // 2. Collect all values to write, grouped by sheet
    // Format: { sheetXmlPath: { cellRef: value } }
    const updatesBySheet = {};

    APP.analysisResult.champs.forEach((champ, idx) => {
      let raw = (champ.cellule_ou_position || '').trim();
      if (!raw) return;
      const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
      if (!value) return;

      let sheetName = null;
      let cellRef = raw;

      // Handle "SheetName!B11" or "'Sheet Name'!B11" format
      if (raw.includes('!')) {
        const parts = raw.split('!');
        sheetName = parts[0].replace(/^'|'$/g, '');
        cellRef = parts[1];
      }

      // Handle range refs like "A33-A34" or "A33:A34" — use first cell
      cellRef = cellRef.split(/[-:]/)[0].trim().toUpperCase();

      if (!/^[A-Z]+\d+$/.test(cellRef)) return;

      // Determine target XML path
      let targetPath = null;
      if (sheetName) {
        targetPath = sheetMap.byName[sheetName.toLowerCase()];
      }
      if (!targetPath) {
        // Default: write to ALL data sheets (skip known non-data sheets like "Menus")
        // This ensures the cell is found wherever it is
        for (const [name, path] of Object.entries(sheetMap.byName)) {
          if (['menus', 'menu', 'listes', 'lists', 'paramètres', 'config'].includes(name)) continue;
          if (!updatesBySheet[path]) updatesBySheet[path] = {};
          updatesBySheet[path][cellRef] = value;
        }
        return;
      }

      if (!updatesBySheet[targetPath]) updatesBySheet[targetPath] = {};
      updatesBySheet[targetPath][cellRef] = value;
    });

    // Debug: log what we're about to write
    console.log('=== Excel Generation Debug ===');
    for (const [xmlPath, updates] of Object.entries(updatesBySheet)) {
      console.log(`Sheet: ${xmlPath}`, updates);
    }

    // 3. Apply updates to each sheet XML using DOM parsing
    for (const [xmlPath, updates] of Object.entries(updatesBySheet)) {
      const file = zip.file(xmlPath);
      if (!file) { console.warn('Sheet file not found:', xmlPath); continue; }

      const sheetXml = await file.async('string');
      const updatedXml = updateSheetXmlDOM(sheetXml, updates);

      if (updatedXml) {
        zip.file(xmlPath, updatedXml);
      }
    }

    // 4. Generate the modified zip
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

// Build mapping: sheet name (lowercase) → xl/worksheets/sheetN.xml path
async function buildSheetMap(zip) {
  const map = { byName: {}, byIndex: [] };

  try {
    // Parse workbook.xml to get sheet names and rIds
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    // Match <sheet> tags — both self-closing and not, with any attributes
    const sheetTags = wbXml.match(/<sheet\s[^>]*>/gi) || [];

    // Parse relationships to map rId → file path
    const relsFile = zip.file('xl/_rels/workbook.xml.rels');
    const relsXml = relsFile ? await relsFile.async('string') : '';
    const relMap = {};
    const relMatches = relsXml.match(/<Relationship\s[^>]*>/gi) || [];
    relMatches.forEach(rel => {
      const id = (rel.match(/Id="([^"]+)"/) || [])[1];
      const target = (rel.match(/Target="([^"]+)"/) || [])[1];
      if (id && target) {
        relMap[id] = target.startsWith('/') ? target.substring(1) : 'xl/' + target;
      }
    });

    sheetTags.forEach(tag => {
      const name = (tag.match(/name="([^"]+)"/) || [])[1];
      // r:id can also appear as r:Id or just id in some files
      const rId = (tag.match(/r:id="([^"]+)"/i) || tag.match(/\bId="(rId\d+)"/i) || [])[1];
      if (name && rId && relMap[rId]) {
        map.byName[name.toLowerCase()] = relMap[rId];
        map.byIndex.push({ name, path: relMap[rId] });
      }
    });
  } catch (e) {
    console.warn('buildSheetMap: parsing failed, using fallback', e);
  }

  // Fallback: if no sheets found, enumerate files directly
  if (map.byIndex.length === 0) {
    zip.folder('xl/worksheets').forEach((path) => {
      if (path.match(/^sheet\d+\.xml$/)) {
        const fullPath = 'xl/worksheets/' + path;
        const idx = map.byIndex.length;
        const name = `Sheet${idx + 1}`;
        map.byName[name.toLowerCase()] = fullPath;
        map.byIndex.push({ name, path: fullPath });
      }
    });
  }

  console.log('Sheet map:', JSON.stringify(map, null, 2));
  return map;
}

// Pure string manipulation — no DOMParser, no XMLSerializer, no namespace issues
function updateSheetXmlDOM(sheetXml, updates) {
  try {
    let xml = sheetXml;
    let changeCount = 0;

    for (const [cellRef, value] of Object.entries(updates)) {
      const rowNum = cellRef.replace(/[A-Z]+/g, '');
      const escaped = value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      const newCellXml = `<c r="${cellRef}" t="inlineStr"><is><t>${escaped}</t></is></c>`;

      // Step 1: Try to find and replace existing cell
      // Match <c ...r="B11"...>...</c> or <c ...r="B11".../>
      const cellWithContent = new RegExp(`<c\\b([^>]*\\br="${cellRef}"[^>]*)>([\\s\\S]*?)</c>`);
      const cellSelfClose = new RegExp(`<c\\b([^>]*\\br="${cellRef}"[^>]*)/>`);

      let matched = false;

      if (cellWithContent.test(xml)) {
        const m = xml.match(cellWithContent);
        const styleMatch = m[1].match(/\bs="(\d+)"/);
        const style = styleMatch ? ` s="${styleMatch[1]}"` : '';
        xml = xml.replace(cellWithContent, `<c r="${cellRef}"${style} t="inlineStr"><is><t>${escaped}</t></is></c>`);
        matched = true;
      } else if (cellSelfClose.test(xml)) {
        const m = xml.match(cellSelfClose);
        const styleMatch = m[1].match(/\bs="(\d+)"/);
        const style = styleMatch ? ` s="${styleMatch[1]}"` : '';
        xml = xml.replace(cellSelfClose, `<c r="${cellRef}"${style} t="inlineStr"><is><t>${escaped}</t></is></c>`);
        matched = true;
      }

      if (!matched) {
        // Step 2: Cell doesn't exist — find row and append cell before </row>
        const rowWithContent = new RegExp(`(<row\\b[^>]*\\br="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`);
        const rowSelfClose = new RegExp(`(<row\\b[^>]*\\br="${rowNum}"[^>]*)(/\\s*>)`);

        if (rowWithContent.test(xml)) {
          xml = xml.replace(rowWithContent, `$1$2${newCellXml}$3`);
          matched = true;
        } else if (rowSelfClose.test(xml)) {
          xml = xml.replace(rowSelfClose, `$1>${newCellXml}</row>`);
          matched = true;
        }
      }

      if (!matched) {
        // Step 3: Row doesn't exist — add before </sheetData>
        xml = xml.replace('</sheetData>', `<row r="${rowNum}">${newCellXml}</row></sheetData>`);
        matched = true;
      }

      if (matched) {
        changeCount++;
        console.log(`  Written: ${cellRef} = "${value}"`);
      } else {
        console.warn(`  FAILED: ${cellRef}`);
      }
    }

    console.log(`Total cells written: ${changeCount}`);
    return xml;
  } catch (e) {
    console.error('updateSheetXmlDOM error:', e);
    return null;
  }
}

// Generate completed .xls using XLSX.js (modifies workbook in memory, writes back as .xls)
function generateCompletedXls() {
  if (!APP.workbook || !APP.analysisResult) return;

  try {
    // Modify the workbook cells directly
    APP.analysisResult.champs.forEach((champ, idx) => {
      let raw = (champ.cellule_ou_position || '').trim();
      if (!raw) return;
      const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
      if (!value) return;

      // Handle "SheetName!B11" format
      let sheetName = null;
      let cellRef = raw;
      if (raw.includes('!')) {
        const parts = raw.split('!');
        sheetName = parts[0].replace(/^'|'$/g, '');
        cellRef = parts[1];
      }
      cellRef = cellRef.split(/[-:]/)[0].trim().toUpperCase();
      if (!/^[A-Z]+\d+$/.test(cellRef)) return;

      // Find target sheet(s)
      const targetSheets = [];
      if (sheetName) {
        const found = APP.workbook.SheetNames.find(n => n.toLowerCase() === sheetName.toLowerCase());
        if (found) targetSheets.push(found);
      }
      if (targetSheets.length === 0) {
        // Write to all non-menu sheets
        APP.workbook.SheetNames.forEach(n => {
          if (!['menus', 'menu', 'listes', 'lists'].includes(n.toLowerCase())) {
            targetSheets.push(n);
          }
        });
      }

      // Write value to each target sheet
      targetSheets.forEach(name => {
        const sheet = APP.workbook.Sheets[name];
        if (!sheet) return;
        const existing = sheet[cellRef];
        if (existing) {
          existing.v = value;
          existing.t = 's';
          delete existing.w;
        } else {
          sheet[cellRef] = { t: 's', v: value };
        }
      });
    });

    // Debug: verify cells were written
    console.log('=== XLS Generation Debug ===');
    APP.workbook.SheetNames.forEach(name => {
      const sheet = APP.workbook.Sheets[name];
      APP.analysisResult.champs.forEach((champ, idx) => {
        const ref = (champ.cellule_ou_position || '').split(/[-:!]/).pop().trim().toUpperCase();
        if (ref && sheet[ref]) {
          console.log(`  ${name}!${ref} = "${sheet[ref].v}"`);
        }
      });
    });

    // Write back as .xls
    const wbout = XLSX.write(APP.workbook, { bookType: 'xls', type: 'array' });
    downloadBlob(
      new Blob([wbout], { type: 'application/vnd.ms-excel' }),
      APP.uploadedFileName.replace(/\.xlsx?$/i, '_complété.xls')
    );
    showToast('Document Excel complété téléchargé !');
  } catch (e) {
    console.error('Erreur génération XLS:', e);
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
