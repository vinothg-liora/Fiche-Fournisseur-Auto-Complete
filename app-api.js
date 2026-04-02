/* ===== Claude API Integration ===== */

const SYSTEM_PROMPT = `Tu es un expert en analyse de formulaires Excel. On te donne un fichier Excel a completer au nom d un fournisseur. Ta mission est d identifier TOUTES les cellules vides a remplir, quelle que soit la structure du fichier.

Methode d analyse en 3 etapes :

Etape 1 - Cartographie : analyse l integralite du fichier. Identifie toutes les cellules non vides (labels) et toutes les cellules vides (a remplir). Repere les zones fusionnees.

Etape 2 - Association label/valeur : pour chaque cellule vide, trouve le label qui lui correspond en cherchant dans cet ordre de priorite :
  (1) cellule immediatement a gauche sur la meme ligne
  (2) cellule immediatement au-dessus dans la meme colonne
  (3) cellule fusionnee englobante
  (4) header de colonne (premiere ligne non vide de la colonne)
  (5) header de ligne (premiere cellule non vide de la ligne)
  Indique pour chaque association ton niveau de confiance.

Etape 3 - Detection des listes deroulantes : pour chaque cellule vide, verifie si le contenu du fichier mentionne [MENU DEROULANT: ...] ou si une feuille "Menus" contient les options. Si oui, liste toutes les options disponibles et choisis la valeur la plus appropriee parmi ces options uniquement.

REGLES CRITIQUES :
- cellule_a_remplir doit TOUJOURS etre une cellule VIDE. Verifie dans le contenu du fichier.
- cellule_label est la cellule qui contient le texte descriptif (le label/question).
- Ne retourne JAMAIS une cellule contenant du texte comme cellule_a_remplir.
- Si tu ne trouves pas de cellule vide adjacente a un label, ne retourne pas ce champ.

Pour les PDF formulaires : indique le nom du champ PDF dans cellule_a_remplir.
Pour les PDF scannes : indique la position visuelle (ex: ‘ligne 3, colonne droite’).

Retourne UNIQUEMENT ce JSON :
{
  "champs": [
    {
      "label": "Raison Sociale / Company Registered Name",
      "cellule_label": "A11",
      "cellule_a_remplir": "B11",
      "categorie": "raison sociale",
      "valeur": "",
      "confiance": "haute/moyenne/basse",
      "est_liste_deroulante": false,
      "options_liste": null,
      "valeur_choisie_dans_liste": null
    }
  ],
  "structure_document": "Labels en colonne A, valeurs en colonne B",
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

  // Strip any non-ASCII chars (BOM, invisible spaces, smart quotes from copy-paste)
  apiKey = apiKey.replace(/[^\x20-\x7E]/g, '').trim();
  if (!apiKey) {
    showToast('Cle API invalide (caracteres non-ASCII detectes). Re-saisissez-la dans Parametres.', 'error');
    return null;
  }

  const messages = [];
  const userContent = [];

  if (fileType === 'xlsx') {
    userContent.push({
      type: 'text',
      text: `Voici le contenu d'un fichier Excel nommé "${fileName}". Analyse tous les champs à remplir.\n\nContenu du fichier :\n${textContent}`
    });
  } else if (fileType === 'pdf-form') {
    userContent.push({
      type: 'text',
      text: `Voici le contenu textuel extrait d'un PDF formulaire nommé "${fileName}". Les noms de champs PDF sont indiqués entre crochets.\n\nContenu :\n${textContent}`
    });
  } else {
    const mediaType = fileName.endsWith('.pdf') ? 'application/pdf' :
                      fileName.endsWith('.png') ? 'image/png' : 'image/jpeg';

    if (mediaType === 'application/pdf') {
      if (textContent) {
        userContent.push({
          type: 'text',
          text: `Voici un PDF scanné nommé "${fileName}". Texte extrait par OCR :\n${textContent}\n\nAnalyse tous les champs à remplir.`
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
        text: `Analyse cette image de document nommé "${fileName}" et identifie tous les champs à remplir.`
      });
    }
  }

  messages.push({ role: 'user', content: userContent });

  const maxRetries = 3;
  const retryDelays = [3000, 8000, 15000]; // 3s, 8s, 15s

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
          model: 'claude-sonnet-4-20250514',
          max_tokens: 8000,
          system: SYSTEM_PROMPT,
          messages: messages
        })
      });

      // Retry on overload (529) or rate limit (429)
      if ((response.status === 529 || response.status === 429) && attempt < maxRetries) {
        const delay = retryDelays[attempt];
        showLoading(`API surchargee - nouvelle tentative dans ${delay / 1000}s (${attempt + 1}/${maxRetries})...`);
        await new Promise(r => setTimeout(r, delay));
        continue;
      }

      if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        const msg = err.error?.message || `API error ${response.status}`;
        if (response.status === 529) throw new Error('API Claude temporairement surchargée. Réessayez dans quelques minutes.');
        if (response.status === 429) throw new Error('Limite de requêtes atteinte. Réessayez dans quelques minutes.');
        if (response.status === 401) throw new Error('Clé API invalide. Vérifiez dans Paramètres.');
        throw new Error(msg);
      }

      const result = await response.json();
      const text = result?.content?.[0]?.text;
      if (!text) throw new Error('Réponse API vide ou format inattendu');

      const jsonMatch = text.match(/\{[\s\S]*\}/);
      if (!jsonMatch) throw new Error('Réponse API invalide - pas de JSON trouvé');

      return JSON.parse(jsonMatch[0]);
    } catch (e) {
      if (attempt < maxRetries && (e.message.includes('fetch') || e.message.includes('network'))) {
        const delay = retryDelays[attempt];
        showLoading(`Erreur réseau — nouvelle tentative dans ${delay / 1000}s...`);
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
