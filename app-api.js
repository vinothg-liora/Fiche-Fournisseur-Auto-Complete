/* ===== Claude API Integration ===== */

const SYSTEM_PROMPT = `Tu es un expert en référencement fournisseur. On te donne un document (Excel ou PDF) qu'un client a envoyé à son fournisseur pour le référencer dans sa base de données. Ta mission est d'identifier TOUS les champs à remplir dans ce document, quelle que soit sa structure.

Règles de reconnaissance :
1. Analyse la structure complète du document avant de conclure : certains champs sont en ligne, d'autres en colonne, d'autres dans des tableaux imbriqués
2. Reconnais les variantes linguistiques et orthographiques : 'N° SIRET', 'Siret', 'SIRET fournisseur', 'Numéro d’identification', 'Tax ID', 'Steuernummer', 'fiscal number' peuvent tous désigner le SIRET ou équivalent
3. Reconnais les abréviations métier françaises spécifiques aux organismes de formation : UAI, NDA, OPCO, Qualiopi, BPF
4. Si un champ est ambigu, propose la catégorie la plus probable avec un niveau de confiance
5. Ne saute aucun champ, même s'il te semble inhabituel — liste-le en 'champ inconnu' plutôt que de l'ignorer
6. Pour les fichiers Excel : indique la référence exacte de la cellule à remplir (ex: B3, C12)
7. Pour les PDF formulaires : indique le nom du champ PDF
8. Pour les PDF scannés : indique la position visuelle du champ (ex: 'ligne 3, colonne droite')

Retourne UNIQUEMENT un JSON structuré ainsi :
{
  "champs": [
    {
      "label_original": "...",
      "cellule_ou_position": "...",
      "categorie_identifiee": "...",
      "valeur_a_inserer": "...",
      "confiance": "haute/moyenne/basse",
      "justification": "..."
    }
  ],
  "structure_document": "description courte de la structure détectée",
  "langue_document": "...",
  "champs_inconnus": [
    {
      "label_original": "...",
      "position": "...",
      "description": "..."
    }
  ]
}`;

async function analyzeWithClaude(fileContentBase64, fileName, fileType, textContent) {
  const apiKey = localStorage.getItem('anthropicApiKey');
  if (!apiKey) {
    showToast('Clé API Anthropic non configurée. Allez dans Paramètres.', 'error');
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

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error?.message || `API error ${response.status}`);
    }

    const result = await response.json();
    const text = result.content[0].text;

    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error('Réponse API invalide - pas de JSON trouvé');

    return JSON.parse(jsonMatch[0]);
  } catch (e) {
    console.error('Erreur API Claude:', e);
    showToast('Erreur API : ' + e.message, 'error');
    return null;
  }
}
