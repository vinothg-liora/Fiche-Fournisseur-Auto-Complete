/* ===== Default Company Data — DataScientest / Liora (embedded for file:// usage) ===== */
const COMPANY_DATA = {
  "_notice": "Fichier de référence fournisseur — DataScientest / Liora.",
  "entites": {
    "france": {
      "nom_commercial": "LIORA",
      "nom_juridique": "DATASCIENTEST",
      "nom_alternatif": "Liora",
      "identifiants": {
        "SIREN": "831450069",
        "SIRET_siege": "83145006900024",
        "SIRET_etablissements_secondaires": {
          "Courbevoie_92400": "83145006900040",
          "Puteaux_92800": "83145006900032",
          "Paris_75016_Flandrin": "83145006900016",
          "Paris_75016_Barcelone_2": "83145006900057"
        },
        "TVA_intracommunautaire": "FR69831450069",
        "code_NAF_APE": "8559A",
        "libelle_NAF": "Formation continue d'adultes",
        "RCS": "831 450 069 RCS Paris",
        "UAI": "0923015C",
        "NDA": "11755665975",
        "numero_Qualiopi": "B01084",
        "DUNS": ""
      },
      "forme_juridique": {
        "type": "Société par actions simplifiée (SAS)",
        "code_forme_juridique": "5710",
        "capital_social": "2 513 €",
        "date_creation": "11/08/2017",
        "exercice_social": "Du 1er juillet au 30 juin"
      },
      "siege_social": {
        "adresse_ligne1": "2 Place de Barcelone",
        "adresse_ligne2": "",
        "code_postal": "75016",
        "ville": "Paris",
        "pays": "France",
        "adresse_complete": "2 Place de Barcelone, 75016 Paris, France"
      },
      "coordonnees": {
        "telephone": "09 80 80 79 49",
        "email_facturation": "billing@liora.io",
        "email_contact": "contact@liora.io",
        "site_web": "liora.io"
      },
      "dirigeants": {
        "president": "EDUCIN TOPCO",
        "directeur_general": "EDUCIN TOPCO",
        "representant_legal_nom": "EDUCIN TOPCO",
        "representant_legal_prenom": "",
        "representant_legal_qualite": "Président"
      },
      "bancaire": {
        "banque": "BNP PARIBAS",
        "domiciliation": "83 bd Sebastopol, 75002 Paris",
        "IBAN": "FR76 3000 4028 3700 0113 0486 194",
        "BIC": "BNPAFRPPXXX",
        "RIB_titulaire": "DATASCIENTEST",
        "RIB_banque": "30004",
        "RIB_guichet": "02837",
        "RIB_numero_compte": "00011304861",
        "RIB_cle": "94"
      },
      "donnees_financieres": {
        "chiffre_affaires_N1": "21 029 000 €",
        "chiffre_affaires_N1_libelle": "Exercice clos au 30/06/2024",
        "chiffre_affaires_N": "28 048 000 €",
        "chiffre_affaires_N_libelle": "Exercice clos au 30/06/2025",
        "effectif_actuel": "120"
      },
      "contacts": {
        "service_comptabilite_facturation": {
          "nom": "NGAN",
          "prenom": "Cédric",
          "civilite": "M.",
          "email": "billing@liora.io",
          "telephone": "07 55 52 08 49",
          "intitule_poste": "Responsable ADV"
        },
        "responsable_commercial": {
          "nom": "LOTH",
          "prenom": "Nathan",
          "civilite": "M.",
          "email": "nathan@liora.io",
          "telephone": "",
          "intitule_poste": "Directeur B2B"
        }
      },
      "documents_a_joindre": {
        "kbis": {
          "description": "Extrait Kbis de moins de 6 mois",
          "date_obtention": "Mars 2026",
          "frequence_renouvellement_mois": 3
        },
        "attestation_vigilance_urssaf": {
          "description": "Attestation de vigilance URSSAF",
          "date_obtention": "Mars 2026",
          "frequence_renouvellement_mois": 3
        },
        "attestation_qualiopi": {
          "description": "Certificat Qualiopi",
          "date_validite": "29/12/2027"
        }
      },
      "assurance": {
        "assureur": "AXA",
        "numero_police": "6055097904",
        "attestation_validite": "Valide du 01/01/2026 au 31/12/2026"
      },
      "certifications": {
        "Qualiopi": "Oui",
        "date_validite_Qualiopi": "Valide du 30/12/2024 au 29/12/2027",
        "certificateur_Qualiopi": "ICPF",
        "numero_Qualiopi": "B01084"
      },
      "aliases_champs_connus": {}
    }
  },
  "champs_appris": {},
  "historique": []
};

/* ===== Data Layer ===== */

// Shortcut to France entity
function getFR() {
  return APP.companyData?.entites?.france;
}

async function loadCompanyData() {
  APP.companyData = JSON.parse(JSON.stringify(COMPANY_DATA));
  document.getElementById('connectionStatus').textContent = '🟢 Connecté';
  return true;
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
  const fr = getFR();
  if (!fr) return '';
  const cat = category.toLowerCase().trim();

  // Check learned aliases first
  const aliases = fr.aliases_champs_connus || {};
  for (const [alias, target] of Object.entries(aliases)) {
    if (cat.includes(alias.toLowerCase())) {
      const resolved = resolveAlias(target, fr);
      if (resolved) return resolved;
    }
  }

  // Check champs_appris
  const appris = APP.companyData?.champs_appris || {};
  for (const [key, val] of Object.entries(appris)) {
    if (cat.includes(key.toLowerCase()) && val && typeof val === 'string' && !val.startsWith('[')) {
      return val;
    }
  }

  // Direct mappings — all known variants
  const id = fr.identifiants || {};
  const fj = fr.forme_juridique || {};
  const siege = fr.siege_social || {};
  const coord = fr.coordonnees || {};
  const dir = fr.dirigeants || {};
  const bank = fr.bancaire || {};
  const fin = fr.donnees_financieres || {};
  const ctFactu = fr.contacts?.service_comptabilite_facturation || {};
  const ctComm = fr.contacts?.responsable_commercial || {};
  const assur = fr.assurance || {};
  const cert = fr.certifications || {};

  const mappings = {
    // Identité
    'raison sociale': fr.nom_juridique,
    'raison_sociale': fr.nom_juridique,
    'company name': fr.nom_juridique,
    'nom entreprise': fr.nom_juridique,
    'denomination': fr.nom_juridique,
    'nom commercial': fr.nom_commercial,
    'nom juridique': fr.nom_juridique,
    // Forme juridique
    'forme juridique': fj.type,
    'forme_juridique': fj.type,
    'legal form': fj.type,
    'capital': fj.capital_social,
    'capital social': fj.capital_social,
    'date creation': fj.date_creation,
    'date de creation': fj.date_creation,
    // Identifiants
    'siret': id.SIRET_siege,
    'n° siret': id.SIRET_siege,
    'numero siret': id.SIRET_siege,
    'siret fournisseur': id.SIRET_siege,
    'identification fiscale': id.SIRET_siege,
    'siren': id.SIREN,
    'n° siren': id.SIREN,
    'code naf': id.code_NAF_APE,
    'code ape': id.code_NAF_APE,
    'naf': id.code_NAF_APE,
    'ape': id.code_NAF_APE,
    'tva intracommunautaire': id.TVA_intracommunautaire,
    'tva intracom': id.TVA_intracommunautaire,
    'n° tva': id.TVA_intracommunautaire,
    'numero de tva': id.TVA_intracommunautaire,
    'tva': id.TVA_intracommunautaire,
    'vat number': id.TVA_intracommunautaire,
    'vat': id.TVA_intracommunautaire,
    'tax id': id.TVA_intracommunautaire,
    'numero identification fiscale': id.TVA_intracommunautaire,
    'identification légale - tva': id.TVA_intracommunautaire,
    'identification legale - tva': id.TVA_intracommunautaire,
    'steuernummer': id.TVA_intracommunautaire,
    'fiscal number': id.TVA_intracommunautaire,
    'ust-idnr': id.TVA_intracommunautaire,
    'rcs': id.RCS,
    'nda': id.NDA,
    'numero declaration activite': id.NDA,
    'n° declaration': id.NDA,
    'uai': id.UAI,
    'n° uai': id.UAI,
    'code uai': id.UAI,
    'qualiopi': cert.numero_Qualiopi,
    'numero qualiopi': cert.numero_Qualiopi,
    'certificat qualiopi': cert.numero_Qualiopi,
    'duns': id.DUNS,
    // Adresse
    'adresse': siege.adresse_complete,
    'adresse siege': siege.adresse_complete,
    'adresse siege social': siege.adresse_complete,
    'adresse facturation': siege.adresse_complete,
    'address': siege.adresse_complete,
    'rue': siege.adresse_ligne1,
    'street': siege.adresse_ligne1,
    'code postal': siege.code_postal,
    'zip code': siege.code_postal,
    'postal code': siege.code_postal,
    'plz': siege.code_postal,
    'ville': siege.ville,
    'city': siege.ville,
    'ort': siege.ville,
    'pays': siege.pays,
    'country': siege.pays,
    'land': siege.pays,
    // Coordonnées
    'telephone': coord.telephone,
    'phone': coord.telephone,
    'tel': coord.telephone,
    'email': coord.email_contact,
    'email contact': coord.email_contact,
    'email facturation': coord.email_facturation,
    'site web': coord.site_web,
    'website': coord.site_web,
    // Dirigeants
    'representant legal': dir.representant_legal_nom,
    'dirigeant': dir.president,
    'president': dir.president,
    'directeur general': dir.directeur_general,
    // Contact facturation / ADV
    'contact facturation': `${ctFactu.prenom} ${ctFactu.nom}`,
    'service comptabilite': `${ctFactu.prenom} ${ctFactu.nom}`,
    'responsable facturation': `${ctFactu.prenom} ${ctFactu.nom}`,
    'nom contact facturation': ctFactu.nom,
    'prenom contact facturation': ctFactu.prenom,
    'email comptabilite': ctFactu.email,
    'telephone facturation': ctFactu.telephone,
    // Contact commercial
    'responsable commercial': `${ctComm.prenom} ${ctComm.nom}`,
    'contact commercial': `${ctComm.prenom} ${ctComm.nom}`,
    'email commercial': ctComm.email,
    'telephone commercial': ctComm.telephone,
    // Banque
    'banque': bank.banque,
    'iban': bank.IBAN,
    'bic': bank.BIC,
    'bic swift': bank.BIC,
    'swift': bank.BIC,
    'domiciliation bancaire': bank.domiciliation,
    'titulaire compte': bank.RIB_titulaire,
    // Finance
    'chiffre affaires': fin.chiffre_affaires_N1,
    'ca n-1': fin.chiffre_affaires_N1,
    'ca n': fin.chiffre_affaires_N,
    'effectif': fin.effectif_actuel,
    'effectif total': fin.effectif_actuel,
    'nombre de salaries': fin.effectif_actuel,
    // Assurance
    'assureur': assur.assureur,
    'numero police assurance': assur.numero_police,
    'numero police': assur.numero_police,
  };

  // Fuzzy match
  for (const [key, value] of Object.entries(mappings)) {
    if (value && (cat.includes(key) || key.includes(cat))) {
      return value;
    }
  }
  return '';
}

function resolveAlias(target, fr) {
  // Simple dot-path resolution on the france entity
  if (!target || typeof target !== 'string') return '';
  const parts = target.split('.');
  let obj = fr;
  for (const p of parts) {
    if (!obj || typeof obj !== 'object') return '';
    obj = obj[p];
  }
  return (typeof obj === 'string' && !obj.startsWith('[')) ? obj : '';
}

function addToHistory(entry) {
  if (!APP.companyData) return;
  if (!APP.companyData.historique) APP.companyData.historique = [];
  APP.companyData.historique.unshift(entry);
}
