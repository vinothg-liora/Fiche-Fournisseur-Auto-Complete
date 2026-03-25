/* ===== Data Layer ===== */
async function loadCompanyData() {
  try {
    const path = APP.networkPath;
    const resp = await fetch(path);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    APP.companyData = await resp.json();
    document.getElementById('connectionStatus').textContent = '🟢 Connecté';
    return true;
  } catch (e) {
    console.error('Erreur chargement données:', e);
    document.getElementById('connectionStatus').textContent = '🔴 Déconnecté';
    showToast('Impossible de charger company_data.json : ' + e.message, 'error');
    return false;
  }
}

async function saveCompanyData() {
  const blob = new Blob([JSON.stringify(APP.companyData, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'company_data.json';
  a.click();
  URL.revokeObjectURL(url);
  showToast('Fichier company_data.json téléchargé. Remplacez le fichier existant sur le dossier partagé.');
}

function getCompanyValue(category) {
  if (!APP.companyData) return '';
  const cat = category.toLowerCase().trim();
  const data = APP.companyData;

  // Check learned aliases first
  if (data.alias_appris) {
    for (const [alias, info] of Object.entries(data.alias_appris)) {
      if (cat.includes(alias.toLowerCase())) {
        if (data.valeurs_memorisees && data.valeurs_memorisees[info.categorie]) {
          return data.valeurs_memorisees[info.categorie];
        }
      }
    }
  }

  const mappings = {
    'raison_sociale': data.entreprise?.raison_sociale,
    'raison sociale': data.entreprise?.raison_sociale,
    'company name': data.entreprise?.raison_sociale,
    'nom entreprise': data.entreprise?.raison_sociale,
    'denomination': data.entreprise?.raison_sociale,
    'forme_juridique': data.entreprise?.forme_juridique,
    'forme juridique': data.entreprise?.forme_juridique,
    'legal form': data.entreprise?.forme_juridique,
    'capital': data.entreprise?.capital_social,
    'capital social': data.entreprise?.capital_social,
    'siret': data.entreprise?.siret,
    'n° siret': data.entreprise?.siret,
    'numero siret': data.entreprise?.siret,
    'siret fournisseur': data.entreprise?.siret,
    'siren': data.entreprise?.siren,
    'n° siren': data.entreprise?.siren,
    'code naf': data.entreprise?.code_naf_ape,
    'code ape': data.entreprise?.code_naf_ape,
    'naf': data.entreprise?.code_naf_ape,
    'ape': data.entreprise?.code_naf_ape,
    'tva intracommunautaire': data.entreprise?.tva_intracommunautaire,
    'tva intracom': data.entreprise?.tva_intracommunautaire,
    'n° tva': data.entreprise?.tva_intracommunautaire,
    'vat number': data.entreprise?.tva_intracommunautaire,
    'tax id': data.entreprise?.tva_intracommunautaire,
    'numero identification fiscale': data.entreprise?.tva_intracommunautaire,
    'steuernummer': data.entreprise?.tva_intracommunautaire,
    'fiscal number': data.entreprise?.tva_intracommunautaire,
    'rcs': data.entreprise?.rcs,
    'nda': data.entreprise?.nda,
    'numero declaration activite': data.entreprise?.nda,
    'uai': data.entreprise?.uai,
    'qualiopi': data.entreprise?.certifications?.qualiopi?.numero,
    'certificat qualiopi': data.entreprise?.certifications?.qualiopi?.numero,
    'adresse': data.adresse?.siege_social?.adresse,
    'adresse siege': data.adresse?.siege_social?.adresse,
    'adresse siege social': data.adresse?.siege_social?.adresse,
    'address': data.adresse?.siege_social?.adresse,
    'rue': data.adresse?.siege_social?.adresse,
    'street': data.adresse?.siege_social?.adresse,
    'complement adresse': data.adresse?.siege_social?.complement,
    'code postal': data.adresse?.siege_social?.code_postal,
    'zip code': data.adresse?.siege_social?.code_postal,
    'postal code': data.adresse?.siege_social?.code_postal,
    'plz': data.adresse?.siege_social?.code_postal,
    'ville': data.adresse?.siege_social?.ville,
    'city': data.adresse?.siege_social?.ville,
    'ort': data.adresse?.siege_social?.ville,
    'pays': data.adresse?.siege_social?.pays,
    'country': data.adresse?.siege_social?.pays,
    'land': data.adresse?.siege_social?.pays,
    'dirigeant nom': data.contacts?.dirigeant ? `${data.contacts.dirigeant.prenom} ${data.contacts.dirigeant.nom}` : '',
    'representant legal': data.contacts?.dirigeant ? `${data.contacts.dirigeant.prenom} ${data.contacts.dirigeant.nom}` : '',
    'nom dirigeant': data.contacts?.dirigeant?.nom,
    'prenom dirigeant': data.contacts?.dirigeant?.prenom,
    'fonction dirigeant': data.contacts?.dirigeant?.fonction,
    'email dirigeant': data.contacts?.dirigeant?.email,
    'telephone dirigeant': data.contacts?.dirigeant?.telephone,
    'contact commercial': data.contacts?.contact_commercial ? `${data.contacts.contact_commercial.prenom} ${data.contacts.contact_commercial.nom}` : '',
    'email commercial': data.contacts?.contact_commercial?.email,
    'email contact': data.contacts?.contact_commercial?.email,
    'telephone commercial': data.contacts?.contact_commercial?.telephone,
    'telephone contact': data.contacts?.contact_commercial?.telephone,
    'email': data.contacts?.contact_commercial?.email,
    'telephone': data.contacts?.contact_commercial?.telephone,
    'phone': data.contacts?.contact_commercial?.telephone,
    'tel': data.contacts?.contact_commercial?.telephone,
    'contact comptabilite': data.contacts?.contact_comptabilite ? `${data.contacts.contact_comptabilite.prenom} ${data.contacts.contact_comptabilite.nom}` : '',
    'email comptabilite': data.contacts?.contact_comptabilite?.email,
    'banque': data.coordonnees_bancaires?.banque,
    'iban': data.coordonnees_bancaires?.iban,
    'bic': data.coordonnees_bancaires?.bic_swift,
    'bic swift': data.coordonnees_bancaires?.bic_swift,
    'swift': data.coordonnees_bancaires?.bic_swift,
    'titulaire compte': data.coordonnees_bancaires?.titulaire_compte,
    'regime tva': data.informations_fiscales?.regime_tva,
    'assureur': data.assurances?.responsabilite_civile?.assureur,
    'numero police assurance': data.assurances?.responsabilite_civile?.numero_police,
    'rc pro': data.assurances?.responsabilite_civile_professionnelle?.assureur,
    'effectif': String(data.effectifs?.effectif_total || ''),
    'effectif total': String(data.effectifs?.effectif_total || ''),
    'convention collective': data.effectifs?.convention_collective,
    'date creation': data.entreprise?.date_creation,
  };

  for (const [key, value] of Object.entries(mappings)) {
    if (cat.includes(key) || key.includes(cat)) {
      return value || '';
    }
  }
  return '';
}

function addToHistory(entry) {
  if (!APP.companyData) return;
  if (!APP.companyData.historique) APP.companyData.historique = [];
  APP.companyData.historique.unshift(entry);
}
