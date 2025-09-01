// =================================================================================
// == FICHIER : UtilsV3.gs
// == PROJET [MOTEUR] - FONCTIONS UTILITAIRES
// == VERSION : 8.1 - Correction du bug setGoToPage en mode multi-langues.
// == RÔLE    : Contient toutes les fonctions de support, appelées par les
// ==           autres scripts du projet.
// =================================-===============================================

// ⚙️ ID de la feuille de configuration centrale (CONFIG)
const ID_FEUILLE_CONFIGURATION = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";

// ------------------------------------
// IDs système (CONFIG → onglet sys_ID_Fichiers)
// ------------------------------------
function getSystemIds() {
  const configSS = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION);
  const idSheet = configSS.getSheetByName('sys_ID_Fichiers');
  if (!idSheet) throw new Error("L'onglet 'sys_ID_Fichiers' est introuvable dans CONFIG.");
  const data = idSheet.getDataRange().getValues();
  const ids = {};
  data.slice(1).forEach(row => {
    if (row[0] && row[1]) ids[row[0]] = row[1];
  });
  return ids;
}

// ------------------------------------
// Lecture d'une ligne de CONFIG (format horizontal)
// ------------------------------------
function getConfigurationFromRow(rowIndex) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("Paramètres Généraux");
  if (!sheet) throw new Error("L'onglet 'Paramètres Généraux' est introuvable.");

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  if (!rowIndex || isNaN(rowIndex) || rowIndex < 2) {
      throw new Error('getConfigurationFromRow: rowIndex invalide (' + rowIndex + ')');
  }

  const values = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  const cfg = {};
  headers.forEach((h, i) => { if (h) cfg[h] = values[i]; });
  cfg._rowIndex = rowIndex;
  return cfg;
}

// ------------------------------------
// Fonctions liées aux questions du formulaire
// ------------------------------------

/**
 * Identifie les langues disponibles pour un type de test donné.
 */
function _identifierLangues(bdd, typeTest) {
    const toutesLesFeuillesBDD = bdd.getSheets();
    const regexLangues = new RegExp('^Questions_' + typeTest + '_([A-Z]{2})$', 'i');
    const languesAInclure = [];
    toutesLesFeuillesBDD.forEach(feuille => {
        const match = feuille.getName().match(regexLangues);
        if (match && match[1]) {
            languesAInclure.push({ 
                code: match[1].toUpperCase(), 
                nomComplet: getLangueFullName(match[1]), 
                feuille: feuille 
            });
        }
    });

    // --- AJOUT : On trie les langues pour garantir un ordre constant (FR, EN...) ---
    const ordreLangues = ['FR', 'EN', 'ES', 'DE']; // Ordre de priorité
    languesAInclure.sort((a, b) => {
        const indexA = ordreLangues.indexOf(a.code);
        const indexB = ordreLangues.indexOf(b.code);
        if (indexA === -1) return 1; // Mettre les langues non listées à la fin
        if (indexB === -1) return -1;
        return indexA - indexB;
    });
    Logger.log(`Langues triées pour la génération : ${languesAInclure.map(l => l.code).join(', ')}`);
    // --- FIN DE L'AJOUT ---

    if (languesAInclure.length === 0) {
        throw new Error("Aucune feuille de questions trouvée pour le type '" + typeTest + "'.");
    }
    return languesAInclure;
}


/**
 * Construit les questions dans le formulaire, en gérant le multi-langues.
 */
function _construireQuestionsFormulaire(form, languesAInclure, nbQuestionsConfig) {
    if (languesAInclure.length > 1) {
        Logger.log(`Mode multi-langues détecté (${languesAInclure.length} langues).`);
        const itemLangue = form.addMultipleChoiceItem().setTitle("Langue / Language").setRequired(true);
        const choices = [];
        languesAInclure.forEach(langue => {
            const page = form.addPageBreakItem().setTitle("Questions (" + langue.nomComplet + ")");
            choices.push(itemLangue.createChoice(langue.nomComplet, page));
            
            _ajouterQuestionsDepuisFeuille(form, langue.feuille, nbQuestionsConfig);
            
            // --- DÉBUT DE LA CORRECTION v8.1 ---
        
            // On s'assure que la redirection vers la page de soumission est bien appliquée
            // à l'objet 'page' que nous venons de créer, qui est un PageBreakItem.
            if (page && typeof page.setGoToPage === 'function') {
                page.setGoToPage(FormApp.PageNavigationType.SUBMIT);
            }
            
            // --- FIN DE LA CORRECTION v8.1 ---
        });
        itemLangue.setChoices(choices);
    } else {
        Logger.log(`Mode langue unique détecté. Insertion directe des questions.`);
        const uniqueLangue = languesAInclure[0];
        _ajouterQuestionsDepuisFeuille(form, uniqueLangue.feuille, nbQuestionsConfig);
    }
}

/**
 * Ajoute une série de questions à un formulaire à partir d'une feuille de calcul.
 */
function _ajouterQuestionsDepuisFeuille(form, feuilleQuestions, nbQuestionsConfig) {
    const nbQuestionsDisponibles = feuilleQuestions.getLastRow() - 1;
    let nbQuestionsAUtiliser = (nbQuestionsConfig && nbQuestionsConfig > 0) 
        ?
        Math.min(nbQuestionsConfig, nbQuestionsDisponibles) 
        : nbQuestionsDisponibles;

    if (nbQuestionsAUtiliser <= 0) return;
    const questionsData = feuilleQuestions.getRange(2, 1, nbQuestionsAUtiliser, 7).getValues();
    questionsData.forEach(q_data => {
        const [id, type_old, titre, options, logique, description, params_json] = q_data;
        let final_type = type_old;
        if (params_json) {
            try {
                const p = JSON.parse(params_json);
                if (p.mode) final_type = p.mode;
            } catch (e) { /* Ignorer les erreurs de parsing JSON */ }
        }
        creerItemFormulaire(form, final_type, id + ': ' + titre, options, description, params_json);
    });
}


/**
 * Crée un item (question) dans le formulaire en fonction de ses spécifications.
 */
function creerItemFormulaire(form, type, titre, optionsString, description, paramsJSONString) {
  let isRequired = !titre.toLowerCase().includes('(optionnel)');

  let params = null;
  if (paramsJSONString) {
    try { params = JSON.parse(paramsJSONString); } catch (e) { params = null; }
  }

  let resolvedType = (params && params.mode) ? String(params.mode).trim().toUpperCase() : String(type || '').trim().toUpperCase();
  if (resolvedType === 'TEXTE_EMAIL') resolvedType = 'EMAIL';

  let item;
  const choices = (params && params.options) ?
    params.options.map(o => o.libelle) : (optionsString || '').split(';').map(s => s.trim()).filter(Boolean);

  switch (resolvedType) {
    case 'QRM_CAT':
    case 'QCU_CAT':
      if (choices.length > 0) {
        item = (resolvedType.startsWith('QRM')) 
          ?
          form.addCheckboxItem() 
          : form.addMultipleChoiceItem();
        item.setTitle(titre).setChoiceValues(choices).setRequired(isRequired);
      } else {
        item = form.addParagraphTextItem().setTitle(`[Erreur ${resolvedType}: Options manquantes] ${titre}`);
      }
      break;

    case 'ECHELLE_NOTE':
    case 'LIKERT_5':
      const min = params ? (params.echelle_min ?? params.min) : 1;
      const max = params ? (params.echelle_max ?? params.max) : 5;
      if (min != null && max != null) {
        item = form.addScaleItem().setTitle(titre).setBounds(Number(min), Number(max)).setRequired(isRequired);
        const lmin = params ? (params.label_min ?? params.libelle_min) : null;
        const lmax = params ? (params.label_max ?? params.libelle_max) : null;
        if (lmin && lmax) item.setLabels(String(lmin), String(lmax));
      } else {
        item = form.addParagraphTextItem().setTitle(`[Erreur ${resolvedType}: bornes min/max manquantes] ${titre}`);
      }
      break;

    case 'EMAIL':
      item = form.addTextItem().setTitle(titre).setRequired(isRequired);
      item.setValidation(FormApp.createTextValidation().requireTextIsEmail().build());
      break;
    case 'TEXTE_COURT':
      item = form.addTextItem().setTitle(titre).setRequired(isRequired);
      break;
    default:
      item = form.addParagraphTextItem().setTitle(`[Type Inconnu: ${resolvedType}] ${titre}`);
  }

  if (description && typeof item.setHelpText === 'function') {
    item.setHelpText(description);
  }
}

// ------------------------------------
// Fonctions utilitaires diverses
// ------------------------------------
function getLangueFullName(code) {
  const map = { FR: 'Français', EN: 'English', ES: 'Español', DE: 'Deutsch' };
  return map[String(code || '').toUpperCase()] || code;
}