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
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  APP.workbook = workbook;

  let textContent = '';
  workbook.SheetNames.forEach(name => {
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

function generateCompletedExcel() {
  if (!APP.workbook || !APP.analysisResult) return;
  const wb = XLSX.utils.book_new();

  APP.workbook.SheetNames.forEach(name => {
    const origSheet = APP.workbook.Sheets[name];
    const newSheet = Object.assign({}, origSheet);

    APP.analysisResult.champs.forEach((champ, idx) => {
      const cellRef = champ.cellule_ou_position;
      if (cellRef && /^[A-Z]+\d+$/i.test(cellRef)) {
        const value = APP.fieldValues[`field_${idx}`] ?? champ.valeur_a_inserer ?? '';
        if (value) {
          newSheet[cellRef.toUpperCase()] = { t: 's', v: value };
        }
      }
    });

    XLSX.utils.book_append_sheet(wb, newSheet, name);
  });

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  downloadBlob(new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    APP.uploadedFileName.replace(/\.xlsx?$/i, '_complété.xlsx'));
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
  let text = `FICHE FOURNISSEUR - RÉCAPITULATIF\n`;
  text += `${'='.repeat(50)}\n\n`;
  text += `Entreprise : ${APP.companyData?.entreprise?.raison_sociale || ''}\n`;
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
  text += `Cordialement,\n${APP.companyData?.contacts?.contact_commercial?.prenom || ''} ${APP.companyData?.contacts?.contact_commercial?.nom || ''}`;
  text += `\n${APP.companyData?.contacts?.contact_commercial?.email || ''}`;
  text += `\n${APP.companyData?.contacts?.contact_commercial?.telephone || ''}`;

  return text;
}
