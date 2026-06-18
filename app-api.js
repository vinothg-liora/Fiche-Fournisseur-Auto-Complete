/* ===== Claude API Integration ===== */

const SYSTEM_PROMPT = `Tu es un expert en analyse de formulaires. On te donne un fichier (Excel ou PDF) a completer au nom d un fournisseur. Ta mission est d identifier TOUTES les zones a remplir, quelle que soit la structure.

Methode d analyse en 3 etapes :

Etape 1 - Cartographie : analyse l integralite du document. Identifie toutes les zones contenant du texte (labels) et toutes les zones vides ou a completer.
  - Pour les Excel : identifie les cellules vides, les cellules avec validation de donnees, les zones fusionnees. Indique pour chaque cellule vide si la cellule a gauche ou au-dessus contient un label.
  - Pour les PDF : analyse visuellement les champs de formulaire, les lignes pointillees, les cases, les rectangles vides.

Etape 2 - Association label/valeur : pour chaque zone vide, trouve le label qui lui correspond en cherchant dans cet ordre :
  (1) texte immediatement a gauche sur la meme ligne
  (2) texte immediatement au-dessus dans la meme colonne
  (3) header de section ou zone fusionnee englobante
  (4) contexte visuel global
  Indique ton niveau de confiance pour chaque association.

Etape 3 - Detection des choix possibles : pour chaque zone vide, verifie s il existe des options predefinies :
  - Listes deroulantes Excel (mentions [MENU DEROULANT: ...] dans le contenu ou feuille "Menus")
  - Cases a cocher PDF, boutons radio PDF
  - Valeurs entre parentheses dans le label
  Si oui, liste toutes les options et choisis la plus appropriee parmi ces options uniquement.

REGLES CRITIQUES :
- cellule_a_remplir doit TOUJOURS etre une zone VIDE. Verifie dans le contenu que cette cellule/zone ne contient pas de texte.
- cellule_label est la zone qui contient le texte descriptif (le label/question).
- Ne retourne JAMAIS une zone contenant du texte comme cellule_a_remplir.
- Si tu ne trouves pas de zone vide adjacente a un label, ne retourne pas ce champ.

Retourne UNIQUEMENT ce JSON :
{
  "champs": [
    {
      "label": "Raison Sociale / Company Registered Name",
      "position": "cellule B11",
      "type_fichier": "excel",
      "cellule_label": "A11",
      "cellule_a_remplir": "B11",
      "categorie": "raison sociale",
      "valeur": "",
      "confiance": "haute/moyenne/basse",
      "justification": "Label en A11, cellule vide en B11 sur la meme ligne",
      "est_choix_multiple": false,
      "options_disponibles": null,
      "valeur_choisie": null
    }
  ],
  "structure_document": "Labels en colonne A, valeurs en colonne B/C",
  "langue_document": "francais",
  "champs_inconnus": [
    {
      "label": "...",
      "position": "...",
      "description": "..."
    }
  ]
}`;

async function analyzeWithClaude(fileContentBase64, fileName, fileType, textContent) {
  let apiKey = localStorage.getItem('anthropicApiKey');
  if (!apiKey) {
    showToast('Cle API Anthropic non configuree. Allez dans Parametres.', 'error');
    return null;
  }

  apiKey = apiKey.replace(/[^\x20-\x7E]/g, '').trim();
  if (!apiKey) {
    showToast('Cle API invalide (caracteres non-ASCII). Re-saisissez dans Parametres.', 'error');
    return null;
  }

  const messages = [];
  const userContent = [];

  if (fileType === 'xlsx') {
    userContent.push({
      type: 'text',
      text: `Voici le contenu d un fichier Excel nomme "${fileName}". Analyse toutes les cellules vides a remplir. Pour chaque cellule vide, indique la cellule du label ET la cellule vide cible.\n\nContenu du fichier :\n${textContent}`
    });
  } else if (fileType === 'pdf-form') {
    userContent.push({
      type: 'text',
      text: `Voici le contenu d un PDF formulaire nomme "${fileName}". Les noms de champs PDF sont entre crochets. Pour chaque champ, indique le nom du champ PDF dans cellule_a_remplir et le type_fichier "pdf".\n\nContenu :\n${textContent}`
    });
  } else {
    // PDF scan / image
    const mediaType = fileName.endsWith('.pdf') ? 'application/pdf' :
                      fileName.endsWith('.png') ? 'image/png' : 'image/jpeg';

    if (mediaType === 'application/pdf') {
      if (textContent) {
        userContent.push({
          type: 'text',
          text: `Voici un PDF scanne nomme "${fileName}". Texte extrait :\n${textContent}\n\nAnalyse visuellement toutes les zones a remplir. Indique type_fichier "pdf" et la position visuelle dans cellule_a_remplir.`
        });
      }
      if (APP.pdfPageImages && APP.pdfPageImages.length > 0) {
        for (const img of APP.pdfPageImages) {
          userContent.push({
            type: 'image',
            source: { type: 'base64', media_type: 'image/png', data: img }
          });
        }
      }
    } else {
      userContent.push({
        type: 'image',
        source: { type: 'base64', media_type: mediaType, data: fileContentBase64 }
      });
      userContent.push({
        type: 'text',
        text: `Analyse cette image de document nomme "${fileName}" et identifie toutes les zones a remplir. Indique type_fichier "pdf".`
      });
    }
  }

  messages.push({ role: 'user', content: userContent });

  const maxRetries = 3;
  const retryDelays = [3000, 8000, 15000];

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true'
        },
        body: JSON.stringify({
          model: 'claude-sonnet-4-6',
          max_tokens: 8000,
          system: SYSTEM_PROMPT,
          messages: messages
        })
      });

      if ((response.status === 529 || response.status === 429) && attempt < maxRetries) {
        const delay = retryDelays[attempt];
        showLoading(`API surchargee - tentative ${attempt + 1}/${maxRetries} dans ${delay / 1000}s...`);
        await new Promise(r => setTimeout(r, delay));
        continue;
      }

      if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        const msg = err.error?.message || `API error ${response.status}`;
        if (response.status === 529) throw new Error('API surchargee. Reessayez dans quelques minutes.');
        if (response.status === 429) throw new Error('Limite de requetes atteinte. Reessayez.');
        if (response.status === 401) throw new Error('Cle API invalide. Verifiez dans Parametres.');
        throw new Error(msg);
      }

      const result = await response.json();
      const text = result?.content?.[0]?.text;
      if (!text) throw new Error('Reponse API vide');

      const jsonMatch = text.match(/\{[\s\S]*\}/);
      if (!jsonMatch) throw new Error('Pas de JSON dans la reponse API');

      return JSON.parse(jsonMatch[0]);
    } catch (e) {
      if (attempt < maxRetries && (e.message.includes('fetch') || e.message.includes('network'))) {
        const delay = retryDelays[attempt];
        showLoading(`Erreur reseau - tentative ${attempt + 1}/${maxRetries}...`);
        await new Promise(r => setTimeout(r, delay));
        continue;
      }
      console.error('Erreur API Claude:', e);
      showToast('Erreur API : ' + e.message, 'error');
      return null;
    }
  }
  return null;
}
