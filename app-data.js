/* ===== Default Company Data (embedded for file:// usage) ===== */
const COMPANY_DATA = {
  "entreprise": {
    "raison_sociale": "Liora Formation SAS",
    "forme_juridique": "SAS",
    "capital_social": "50 000 €",
    "date_creation": "2018-03-15",
    "siret": "123 456 789 00012",
    "siren": "123 456 789",
    "code_naf_ape": "8559A",
    "tva_intracommunautaire": "FR12 123456789",
    "rcs": "RCS Paris B 123 456 789",
    "nda": "11 75 12345 75",
    "uai": "0753456A",
    "certifications": {
      "qualiopi": {
        "numero": "QUA-2024-001234",
        "date_obtention": "30/12/2024",
        "date_expiration": "29/12/2027",
        "categories": ["Actions de formation", "Bilans de compétences"]
      }
    }
  },
  "adresse": {
    "siege_social": {
      "adresse": "15 Rue de la Formation",
      "complement": "Bâtiment A, 3ème étage",
      "code_postal": "75008",
      "ville": "Paris",
      "pays": "France"
    },
    "adresse_facturation": {
      "identique_siege": true,
      "adresse": "",
      "complement": "",
      "code_postal": "",
      "ville": "",
      "pays": ""
    }
  },
  "contacts": {
    "dirigeant": {
      "civilite": "M.",
      "nom": "Dupont",
      "prenom": "Jean",
      "fonction": "Directeur Général",
      "email": "j.dupont@liora-formation.fr",
      "telephone": "+33 1 23 45 67 89"
    },
    "contact_commercial": {
      "civilite": "Mme",
      "nom": "Martin",
      "prenom": "Sophie",
      "fonction": "Responsable ADV",
      "email": "s.martin@liora-formation.fr",
      "telephone": "+33 1 23 45 67 90"
    },
    "contact_comptabilite": {
      "civilite": "M.",
      "nom": "Bernard",
      "prenom": "Pierre",
      "fonction": "Responsable Comptabilité",
      "email": "p.bernard@liora-formation.fr",
      "telephone": "+33 1 23 45 67 91"
    }
  },
  "coordonnees_bancaires": {
    "banque": "BNP Paribas",
    "iban": "FR76 1234 5678 9012 3456 7890 123",
    "bic_swift": "BNPAFRPPXXX",
    "titulaire_compte": "Liora Formation SAS"
  },
  "informations_fiscales": {
    "regime_tva": "Normal",
    "assujetti_tva": true,
    "exoneration_tva_formation": true,
    "reference_exoneration": "Article 261-4-4°a du CGI"
  },
  "assurances": {
    "responsabilite_civile": {
      "assureur": "AXA France",
      "numero_police": "POL-2024-789456",
      "date_debut": "01/01/2026",
      "date_fin": "31/12/2026",
      "montant_couverture": "5 000 000 €"
    },
    "responsabilite_civile_professionnelle": {
      "assureur": "AXA France",
      "numero_police": "POL-2024-789457",
      "date_debut": "01/01/2026",
      "date_fin": "31/12/2026",
      "montant_couverture": "2 000 000 €"
    }
  },
  "effectifs": {
    "effectif_total": 45,
    "effectif_formateurs": 30,
    "convention_collective": "IDCC 1516 - Organismes de formation"
  },
  "references_clients": [
    "SNCF",
    "EDF",
    "Orange",
    "BNP Paribas",
    "Société Générale"
  ],
  "documents_officiels": {
    "kbis": {
      "nom": "Extrait Kbis",
      "derniere_mise_a_jour": "2026-03-01",
      "frequence_renouvellement_mois": 3
    },
    "attestation_urssaf": {
      "nom": "Attestation URSSAF",
      "derniere_mise_a_jour": "2026-03-01",
      "frequence_renouvellement_mois": 3
    },
    "attestation_assurance": {
      "nom": "Attestation AXA",
      "derniere_mise_a_jour": "2026-01-15",
      "frequence_renouvellement_mois": 12
    },
    "certificat_qualiopi": {
      "nom": "Certificat Qualiopi",
      "derniere_mise_a_jour": "2024-12-30",
      "date_expiration": "2027-12-29"
    }
  },
  "alias_appris": {},
  "valeurs_memorisees": {},
  "historique": []
};

/* ===== Data Layer ===== */
async function loadCompanyData() {
  // Use embedded data directly (works with file:// protocol)
  APP.companyData = JSON.parse(JSON.stringify(COMPANY_DATA));
  document.getElementById('connectionStatus').textContent = '🟢 Connecté';
  return true;
}

async function saveCompanyData() {
  // For local file:// usage, we download the updated JSON.
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

  // Direct mappings
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

  // Fuzzy match
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
