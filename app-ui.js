/* ===== UI Rendering ===== */

function renderVerificationTable(result) {
  APP.analysisResult = result;
  APP.fieldValues = {};

  document.getElementById('structureDesc').textContent = result.structure_document || 'Non détectée';
  document.getElementById('langueDoc').textContent = result.langue_document || 'Non détectée';

  const tbody = document.getElementById('fieldsBody');
  tbody.innerHTML = '';

  (result.champs || []).forEach((champ, idx) => {
    const knownValue = getCompanyValue(champ.categorie_identifiee);
    const prefilledValue = knownValue || champ.valeur_a_inserer || '';
    APP.fieldValues[`field_${idx}`] = prefilledValue;

    const tr = document.createElement('tr');
    const confClass = `confidence-${champ.confiance}`;
    const needsValidation = champ.confiance === 'basse' || champ.confiance === 'moyenne';

    tr.innerHTML = `
      <td title="${escapeHtml(champ.justification || '')}">${escapeHtml(champ.label_original)}</td>
      <td><code>${escapeHtml(champ.cellule_ou_position)}</code></td>
      <td>${escapeHtml(champ.categorie_identifiee)}</td>
      <td><span class="confidence-badge ${confClass}">${champ.confiance}</span></td>
      <td>
        <input type="text" class="field-input ${needsValidation ? '' : 'validated'}"
               id="input_field_${idx}"
               value="${escapeHtml(prefilledValue)}"
               data-field-idx="${idx}"
               data-needs-validation="${needsValidation}"
               ${!needsValidation && prefilledValue ? 'readonly' : ''}>
      </td>
      <td>
        ${needsValidation || !prefilledValue ?
          `<button class="btn btn-sm btn-secondary" onclick="validateField(${idx})">✓ Valider</button>` :
          `<span style="color:var(--green)">✓</span>`
        }
      </td>
    `;
    tbody.appendChild(tr);
  });

  const unknownFields = result.champs_inconnus || [];
  const unknownSection = document.getElementById('unknownFieldsSection');
  const unknownBody = document.getElementById('unknownFieldsBody');
  unknownBody.innerHTML = '';

  if (unknownFields.length > 0) {
    unknownSection.style.display = 'block';
    unknownFields.forEach((u, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(u.label_original)}</td>
        <td><code>${escapeHtml(u.position)}</code></td>
        <td>${escapeHtml(u.description || '—')}</td>
        <td>
          <input type="text" class="field-input" id="input_unknown_${idx}"
                 data-unknown-idx="${idx}" data-needs-validation="true"
                 placeholder="Valeur à insérer">
        </td>
        <td>
          <div class="learn-section">
            <label>
              <input type="checkbox" class="memorize-cb" data-unknown-idx="${idx}">
              Mémoriser cette valeur
            </label>
            <input type="text" class="category-input" data-unknown-idx="${idx}"
                   placeholder="Nom de la catégorie" style="display:none;">
            <label style="margin-top:6px;">
              <input type="checkbox" class="alias-cb" data-unknown-idx="${idx}">
              Ajouter comme alias connu
            </label>
          </div>
        </td>
      `;
      unknownBody.appendChild(tr);
    });
  } else {
    unknownSection.style.display = 'none';
  }

  document.querySelectorAll('.memorize-cb').forEach(cb => {
    cb.addEventListener('change', (e) => {
      const idx = e.target.dataset.unknownIdx;
      const catInput = document.querySelector(`.category-input[data-unknown-idx="${idx}"]`);
      catInput.style.display = e.target.checked ? 'block' : 'none';
    });
  });

  document.querySelectorAll('.field-input').forEach(input => {
    input.addEventListener('input', (e) => {
      const fieldIdx = e.target.dataset.fieldIdx;
      const unknownIdx = e.target.dataset.unknownIdx;
      if (fieldIdx !== undefined) APP.fieldValues[`field_${fieldIdx}`] = e.target.value;
      if (unknownIdx !== undefined) APP.fieldValues[`unknown_${unknownIdx}`] = e.target.value;
      updateValidateButton();
    });
  });

  updateValidateButton();
}

function validateField(idx) {
  const input = document.getElementById(`input_field_${idx}`);
  if (!input.value.trim()) {
    showToast('Veuillez saisir une valeur avant de valider.', 'warning');
    return;
  }
  input.classList.add('validated');
  input.dataset.needsValidation = 'false';
  input.readOnly = true;
  APP.fieldValues[`field_${idx}`] = input.value;

  const btn = input.closest('tr').querySelector('button');
  if (btn) btn.outerHTML = '<span style="color:var(--green)">✓</span>';

  updateValidateButton();
}

function updateValidateButton() {
  const allInputs = document.querySelectorAll('.field-input[data-needs-validation="true"]');
  let allValidated = true;
  allInputs.forEach(input => {
    if (input.dataset.needsValidation === 'true' && !input.classList.contains('validated')) {
      allValidated = false;
    }
  });

  document.getElementById('btnValidate').disabled = !allValidated;
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/* ===== Documents Tab ===== */
function renderDocuments() {
  if (!APP.companyData) return;
  const fr = APP.companyData.entites?.france;
  if (!fr) return;

  const tbody = document.getElementById('documentsBody');
  tbody.innerHTML = '';

  // Build a unified docs list from the new structure
  const docsList = [
    { nom: 'Extrait Kbis', date: fr.documents_a_joindre?.kbis?.date_obtention, freq: 3 },
    { nom: 'Attestation URSSAF', date: fr.documents_a_joindre?.attestation_vigilance_urssaf?.date_obtention, freq: 3 },
    { nom: `Attestation ${fr.assurance?.assureur || 'Assurance'}`, date: '01/01/2026', freq: 12, expiry: '31/12/2026' },
    { nom: 'Certificat Qualiopi', date: '30/12/2024', expiry: fr.documents_a_joindre?.attestation_qualiopi?.date_validite },
  ];

  docsList.forEach(doc => {
    const tr = document.createElement('tr');
    let expiryDate, validityLabel;

    // Parse dates (handle "Mars 2026" style or DD/MM/YYYY)
    const parseDate = (str) => {
      if (!str) return null;
      // DD/MM/YYYY
      const dmy = str.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (dmy) return new Date(dmy[3], dmy[2] - 1, dmy[1]);
      // "Mars 2026" style
      const months = { janvier:0, février:1, mars:2, avril:3, mai:4, juin:5, juillet:6, août:7, septembre:8, octobre:9, novembre:10, décembre:11 };
      const myr = str.match(/(\w+)\s+(\d{4})/i);
      if (myr && months[myr[1].toLowerCase()] !== undefined) return new Date(myr[2], months[myr[1].toLowerCase()], 1);
      return new Date(str);
    };

    const lastUpdate = parseDate(doc.date);

    if (doc.expiry) {
      expiryDate = parseDate(doc.expiry);
      validityLabel = `Jusqu'au ${doc.expiry}`;
    } else if (lastUpdate && doc.freq) {
      expiryDate = new Date(lastUpdate);
      expiryDate.setMonth(expiryDate.getMonth() + doc.freq);
      validityLabel = `${doc.freq} mois`;
    } else {
      expiryDate = new Date();
      validityLabel = '—';
    }

    const now = new Date();
    const daysLeft = expiryDate ? daysBetween(now, expiryDate) : 0;
    let statusIcon, statusClass;
    if (daysLeft < 0) {
      statusIcon = '🔴'; statusClass = 'Expiré';
    } else if (daysLeft <= 30) {
      statusIcon = '🟠'; statusClass = 'Expire bientôt';
    } else {
      statusIcon = '🟢'; statusClass = 'Valide';
    }

    tr.innerHTML = `
      <td>${escapeHtml(doc.nom)}</td>
      <td>${doc.date || '—'}</td>
      <td>${validityLabel}</td>
      <td><span class="status-dot">${statusIcon}</span> ${statusClass}${daysLeft >= 0 ? ` (${daysLeft}j)` : ''}</td>
    `;
    tbody.appendChild(tr);
  });
}

/* ===== History Tab ===== */
function renderHistory() {
  const container = document.getElementById('historyContent');
  const history = APP.companyData?.historique || [];

  if (history.length === 0) {
    container.innerHTML = `
      <div class="history-empty">
        <div class="icon">📋</div>
        <h3>Aucun traitement enregistré</h3>
        <p>L'historique apparaîtra ici après le premier traitement de fichier.</p>
      </div>
    `;
    return;
  }

  let html = `<table><thead><tr>
    <th>Date</th><th>Fichier</th><th>Champs auto</th><th>Champs manuels</th><th>Alias appris</th>
  </tr></thead><tbody>`;

  history.forEach(entry => {
    html += `<tr>
      <td>${escapeHtml(entry.date)}</td>
      <td>${escapeHtml(entry.fichier)}</td>
      <td>${entry.champs_auto || 0}</td>
      <td>${entry.champs_manuels || 0}</td>
      <td>${(entry.alias_appris || []).map(a => escapeHtml(a)).join(', ') || '—'}</td>
    </tr>`;
  });

  html += '</tbody></table>';
  container.innerHTML = html;
}

/* ===== Settings Tab ===== */
function renderSettings() {
  const savedKey = localStorage.getItem('anthropicApiKey') || '';
  document.getElementById('apiKeyInput').value = savedKey;
  document.getElementById('sharePointUrl').value = localStorage.getItem('sharePointUrl') || '';
  document.getElementById('networkPath').value = APP.networkPath;
  renderCompanyForm();
  renderAliases();
  renderDocValidityForm();
}

function renderCompanyForm() {
  if (!APP.companyData) return;
  const container = document.getElementById('companyForm');
  const fields = [
    ['Nom commercial', 'entites.france.nom_commercial'],
    ['Nom juridique', 'entites.france.nom_juridique'],
    ['Forme juridique', 'entites.france.forme_juridique.type'],
    ['Capital social', 'entites.france.forme_juridique.capital_social'],
    ['Date de création', 'entites.france.forme_juridique.date_creation'],
    ['SIRET siège', 'entites.france.identifiants.SIRET_siege'],
    ['SIREN', 'entites.france.identifiants.SIREN'],
    ['Code NAF/APE', 'entites.france.identifiants.code_NAF_APE'],
    ['TVA Intracommunautaire', 'entites.france.identifiants.TVA_intracommunautaire'],
    ['RCS', 'entites.france.identifiants.RCS'],
    ['NDA', 'entites.france.identifiants.NDA'],
    ['UAI', 'entites.france.identifiants.UAI'],
    ['N° Qualiopi', 'entites.france.identifiants.numero_Qualiopi'],
    ['Adresse', 'entites.france.siege_social.adresse_ligne1'],
    ['Code postal', 'entites.france.siege_social.code_postal'],
    ['Ville', 'entites.france.siege_social.ville'],
    ['Pays', 'entites.france.siege_social.pays'],
    ['Téléphone', 'entites.france.coordonnees.telephone'],
    ['Email facturation', 'entites.france.coordonnees.email_facturation'],
    ['Email contact', 'entites.france.coordonnees.email_contact'],
    ['Site web', 'entites.france.coordonnees.site_web'],
    ['Président', 'entites.france.dirigeants.president'],
    ['Nom contact ADV', 'entites.france.contacts.service_comptabilite_facturation.nom'],
    ['Prénom contact ADV', 'entites.france.contacts.service_comptabilite_facturation.prenom'],
    ['Email contact ADV', 'entites.france.contacts.service_comptabilite_facturation.email'],
    ['Tél contact ADV', 'entites.france.contacts.service_comptabilite_facturation.telephone'],
    ['Nom contact commercial', 'entites.france.contacts.responsable_commercial.nom'],
    ['Prénom contact commercial', 'entites.france.contacts.responsable_commercial.prenom'],
    ['Email contact commercial', 'entites.france.contacts.responsable_commercial.email'],
    ['IBAN', 'entites.france.bancaire.IBAN'],
    ['BIC', 'entites.france.bancaire.BIC'],
    ['Banque', 'entites.france.bancaire.banque'],
    ['Domiciliation', 'entites.france.bancaire.domiciliation'],
    ['CA N-1', 'entites.france.donnees_financieres.chiffre_affaires_N1'],
    ['Effectif actuel', 'entites.france.donnees_financieres.effectif_actuel'],
    ['Assureur', 'entites.france.assurance.assureur'],
    ['N° police assurance', 'entites.france.assurance.numero_police'],
  ];

  let html = '<div class="company-form-grid">';
  fields.forEach(([label, path]) => {
    const value = getNestedValue(APP.companyData, path) || '';
    html += `
      <div class="form-group">
        <label>${label}</label>
        <input type="text" data-path="${path}" class="company-field" value="${escapeHtml(String(value))}">
      </div>
    `;
  });
  html += '</div>';
  container.innerHTML = html;
}

function renderAliases() {
  const container = document.getElementById('aliasList');
  const aliases = APP.companyData?.entites?.france?.aliases_champs_connus || {};
  const appris = APP.companyData?.champs_appris || {};
  const all = { ...aliases, ...appris };

  if (Object.keys(all).length === 0) {
    container.innerHTML = '<p style="color:var(--text-muted);text-align:center;padding:20px;">Aucun alias appris pour le moment.</p>';
    return;
  }

  let html = '';
  Object.entries(all).forEach(([label, target]) => {
    if (label.startsWith('_')) return;
    const cat = typeof target === 'object' ? JSON.stringify(target) : target;
    html += `
      <div class="alias-item">
        <span><span class="alias-label">${escapeHtml(label)}</span> → <span class="alias-category">${escapeHtml(cat)}</span></span>
        <button class="btn btn-danger btn-sm" onclick="deleteAlias('${escapeHtml(label)}')">Supprimer</button>
      </div>
    `;
  });
  container.innerHTML = html;
}

function renderDocValidityForm() {
  if (!APP.companyData) return;
  const container = document.getElementById('docValidityForm');
  const docs = APP.companyData?.entites?.france?.documents_a_joindre || {};

  let html = '';
  Object.entries(docs).forEach(([key, doc]) => {
    const freq = doc.frequence_renouvellement_mois
      ? `Renouvellement : tous les ${doc.frequence_renouvellement_mois} mois`
      : (doc.date_expiration ? `Expire le : ${formatDate(doc.date_expiration)}` : '');

    html += `
      <div class="form-group">
        <label>${doc.nom} <small style="color:var(--text-muted);">(${freq})</small></label>
        <input type="date" data-doc-key="${key}" class="doc-date-field"
               value="${doc.derniere_mise_a_jour || ''}">
        ${doc.date_expiration !== undefined ? `
          <label style="margin-top:4px;font-size:.8rem;">Date d'expiration</label>
          <input type="date" data-doc-key="${key}" data-field="expiration" class="doc-expiry-field"
                 value="${doc.date_expiration || ''}">
        ` : ''}
      </div>
    `;
  });
  container.innerHTML = html;
}

function deleteAlias(label) {
  // Try both locations
  if (APP.companyData?.entites?.france?.aliases_champs_connus?.[label]) {
    delete APP.companyData.entites.france.aliases_champs_connus[label];
  }
  if (APP.companyData?.champs_appris?.[label]) {
    delete APP.companyData.champs_appris[label];
  }
  renderAliases();
  showToast(`Alias "${label}" supprimé. Pensez à sauvegarder.`);
}

/* ===== Helpers ===== */
function getNestedValue(obj, path) {
  return path.split('.').reduce((o, k) => (o && o[k] !== undefined) ? o[k] : '', obj);
}

function setNestedValue(obj, path, value) {
  const keys = path.split('.');
  const last = keys.pop();
  const target = keys.reduce((o, k) => {
    if (!o[k]) o[k] = {};
    return o[k];
  }, obj);
  target[last] = value;
}
