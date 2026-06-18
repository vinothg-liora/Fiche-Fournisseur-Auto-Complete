/* ===== Main App Controller ===== */
/* Modules loaded via index.html: app-utils.js, app-data.js, app-api.js, app-file.js, app-ui.js */

document.addEventListener('DOMContentLoaded', async () => {
  await loadCompanyData();
  await detectServer();
  renderDocuments();
  renderHistory();
  renderSettings();
  initTabs();
  initDropZone();
  initButtons();

  // Show server status
  if (SERVER_URL) {
    document.getElementById('connectionStatus').textContent = '🟢 Serveur connecté';
  }
});

/* ===== Tab Navigation ===== */
function initTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(`tab-${btn.dataset.tab}`).classList.add('active');

      if (btn.dataset.tab === 'documents') renderDocuments();
      if (btn.dataset.tab === 'historique') renderHistory();
      if (btn.dataset.tab === 'parametres') renderSettings();
    });
  });
}

/* ===== Drop Zone ===== */
function initDropZone() {
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');

  dropZone.addEventListener('click', () => fileInput.click());

  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
      handleFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFile(e.target.files[0]);
    }
  });
}

function handleFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  const validExts = ['xlsx', 'xls', 'pdf', 'png', 'jpg', 'jpeg', 'tiff'];
  if (!validExts.includes(ext)) {
    showToast('Format non supporté. Utilisez Excel (.xlsx), PDF ou image.', 'error');
    return;
  }

  APP.uploadedFile = file;
  const fileInfo = document.getElementById('fileInfo');
  const sizeMB = (file.size / 1024 / 1024).toFixed(2);
  fileInfo.textContent = `📎 ${file.name} (${sizeMB} Mo)`;
  fileInfo.style.display = 'block';
  document.getElementById('btnComplete').disabled = false;
}

/* ===== Buttons ===== */
function initButtons() {
  document.getElementById('btnComplete').addEventListener('click', async () => {
    if (!APP.uploadedFile) return;
    if (!localStorage.getItem('anthropicApiKey')) {
      showToast('Configurez votre clé API dans les Paramètres.', 'error');
      return;
    }

    showLoading('Analyse du document par Claude AI...');
    try {
      const result = await processFile(APP.uploadedFile);
      hideLoading();

      if (result) {
        (result.champs || []).forEach(champ => {
          const val = getCompanyValue(champ.categorie_identifiee);
          if (val && !champ.valeur_a_inserer) {
            champ.valeur_a_inserer = val;
          }
        });

        document.getElementById('uploadSection').style.display = 'none';
        document.getElementById('verificationSection').classList.add('active');
        renderVerificationTable(result);
      }
    } catch (e) {
      hideLoading();
      showToast('Erreur: ' + e.message, 'error');
    }
  });

  document.getElementById('btnBackToUpload').addEventListener('click', resetToUpload);
  document.getElementById('btnNewFile').addEventListener('click', resetToUpload);

  document.getElementById('btnValidate').addEventListener('click', async () => {
    const newAliases = [];

    // Learn rules from manually validated known fields (checkbox "Mémoriser")
    document.querySelectorAll('.learn-rule-cb:checked').forEach(cb => {
      const label = cb.dataset.label;
      const value = cb.dataset.value;
      const category = cb.dataset.category;
      if (label && value) {
        if (!APP.companyData.champs_appris) APP.companyData.champs_appris = {};
        APP.companyData.champs_appris[label] = value;
        newAliases.push(label);
      }
    });

    // Learn from unknown fields
    document.querySelectorAll('.alias-cb:checked').forEach(cb => {
      const idx = cb.dataset.unknownIdx;
      const label = APP.analysisResult.champs_inconnus?.[idx]?.label_original;
      const catInput = document.querySelector(`.category-input[data-unknown-idx="${idx}"]`);
      const category = catInput?.value || label;
      if (label) {
        if (!APP.companyData.champs_appris) APP.companyData.champs_appris = {};
        APP.companyData.champs_appris[label] = category;
        newAliases.push(label);
      }
    });

    document.querySelectorAll('.memorize-cb:checked').forEach(cb => {
      const idx = cb.dataset.unknownIdx;
      const catInput = document.querySelector(`.category-input[data-unknown-idx="${idx}"]`);
      const category = catInput?.value;
      const value = APP.fieldValues[`unknown_${idx}`];
      if (category && value) {
        if (!APP.companyData.champs_appris) APP.companyData.champs_appris = {};
        APP.companyData.champs_appris[category] = value;
      }
    });

    let autoCount = 0, manualCount = 0;
    (APP.analysisResult.champs || []).forEach((champ) => {
      if (champ.confiance === 'haute' && champ.valeur_a_inserer) autoCount++;
      else manualCount++;
    });

    addToHistory({
      date: new Date().toLocaleDateString('fr-FR') + ' ' + new Date().toLocaleTimeString('fr-FR'),
      fichier: APP.uploadedFileName,
      champs_auto: autoCount,
      champs_manuels: manualCount,
      alias_appris: newAliases
    });

    document.getElementById('verificationSection').classList.remove('active');
    document.getElementById('resultSection').classList.add('active');
    document.getElementById('resultSummary').textContent =
      `${autoCount + manualCount} champs traités (${autoCount} automatiques, ${manualCount} manuels). ${newAliases.length} alias appris.`;

    if (newAliases.length > 0 || document.querySelectorAll('.memorize-cb:checked').length > 0) {
      showToast(`${newAliases.length} alias mémorisés. Pensez à sauvegarder company_data.json dans les paramètres.`);
    }
  });

  document.getElementById('btnDownload').addEventListener('click', async () => {
    if (APP.fileType === 'xlsx' || APP.fileType === 'xls') await generateCompletedExcel();
    else if (APP.fileType === 'pdf-form') await generateCompletedPDFForm();
    else await generateRecapPDF();
  });

  document.getElementById('btnRecapPDF').addEventListener('click', async () => {
    await generateRecapPDF();
  });

  document.getElementById('btnEmailRecap').addEventListener('click', () => {
    const recap = generateEmailRecap();
    document.getElementById('emailRecapContent').innerHTML = recap;
    document.getElementById('emailModal').classList.add('active');
  });
  document.getElementById('btnCopyEmail').addEventListener('click', () => {
    // Copy as rich HTML for email paste
    const el = document.getElementById('emailRecapContent');
    const range = document.createRange();
    range.selectNodeContents(el);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand('copy');
    sel.removeAllRanges();
    showToast('Copié dans le presse-papiers (formaté) !');
  });
  document.getElementById('btnCloseEmailModal').addEventListener('click', () => {
    document.getElementById('emailModal').classList.remove('active');
  });

  document.getElementById('btnSharePoint').addEventListener('click', () => {
    const url = localStorage.getItem('sharePointUrl');
    if (url) window.open(url, '_blank');
    else showToast('URL SharePoint non configurée. Allez dans Paramètres.', 'warning');
  });

  document.getElementById('btnSaveApiKey').addEventListener('click', () => {
    // Strip non-ASCII chars that copy-paste may add (BOM, smart quotes, etc.)
    const key = document.getElementById('apiKeyInput').value.replace(/[^\x20-\x7E]/g, '').trim();
    if (key) {
      localStorage.setItem('anthropicApiKey', key);
      showToast('Cle API sauvegardee.');
    } else {
      showToast('Cle API vide ou invalide.', 'error');
    }
  });

  document.getElementById('btnSaveSharePoint').addEventListener('click', () => {
    const url = document.getElementById('sharePointUrl').value.trim();
    localStorage.setItem('sharePointUrl', url);
    showToast('URL SharePoint sauvegardée.');
  });

  document.getElementById('btnSaveNetworkPath').addEventListener('click', () => {
    const path = document.getElementById('networkPath').value.trim();
    if (path) {
      localStorage.setItem('networkPath', path);
      APP.networkPath = path;
      showToast('Chemin réseau sauvegardé.');
    }
  });

  document.getElementById('btnTestConnection').addEventListener('click', async () => {
    const ok = await loadCompanyData();
    if (ok) showToast('Connexion réussie ! Données chargées.');
  });

  document.getElementById('btnSaveCompany').addEventListener('click', () => {
    document.querySelectorAll('.company-field').forEach(input => {
      setNestedValue(APP.companyData, input.dataset.path, input.value);
    });
    saveCompanyData();
  });

  document.getElementById('btnSaveDocDates').addEventListener('click', () => {
    const docs = APP.companyData?.entites?.france?.documents_a_joindre;
    if (!docs) return;
    document.querySelectorAll('.doc-date-field').forEach(input => {
      const key = input.dataset.docKey;
      if (docs[key]) {
        docs[key].date_obtention = input.value;
      }
    });
    document.querySelectorAll('.doc-expiry-field').forEach(input => {
      const key = input.dataset.docKey;
      if (docs[key]) {
        docs[key].date_validite = input.value;
      }
    });
    saveCompanyData();
  });

  document.getElementById('btnExportHistory').addEventListener('click', () => {
    const history = APP.companyData?.historique || [];
    if (history.length === 0) {
      showToast('Aucun historique à exporter.', 'warning');
      return;
    }
    const wsData = [['Date', 'Fichier', 'Champs auto', 'Champs manuels', 'Alias appris']];
    history.forEach(entry => {
      wsData.push([
        entry.date,
        entry.fichier,
        entry.champs_auto || 0,
        entry.champs_manuels || 0,
        (entry.alias_appris || []).join(', ')
      ]);
    });
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Historique');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    downloadBlob(
      new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      'historique_fiches_fournisseur.xlsx'
    );
    showToast('Historique exporté.');
  });
}

function resetToUpload() {
  document.getElementById('uploadSection').style.display = 'block';
  document.getElementById('verificationSection').classList.remove('active');
  document.getElementById('resultSection').classList.remove('active');
  document.getElementById('fileInfo').style.display = 'none';
  document.getElementById('btnComplete').disabled = true;
  document.getElementById('fileInput').value = '';
  APP.uploadedFile = null;
  APP.analysisResult = null;
  APP.fieldValues = {};
}
