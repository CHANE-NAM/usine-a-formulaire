// =================================================================================
// FICHIER : Utils V2.gs (Projet MOTEUR)
// RÔLE : Fonctions utilitaires pour le Moteur.
// VERSION : 2.7 - Ajout de la gestion du type de question ECHELLE_NOTE
// =================================================================================

// ID de la feuille de calcul centrale qui pilote toute l'usine.
const ID_FEUILLE_CONFIGURATION = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";

/**
 * Lit l'onglet 'sys_ID_Fichiers' de la feuille de configuration
 * et retourne un objet avec tous les ID système.
 * @returns {Object} Un objet où les clés sont les noms des ID et les valeurs sont les ID.
 */
function getSystemIds() {
  try {
    const configSS = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION);
    const idSheet = configSS.getSheetByName('sys_ID_Fichiers');
    if (!idSheet) {
      throw new Error("L'onglet de configuration 'sys_ID_Fichiers' est introuvable.");
    }
    
    const data = idSheet.getDataRange().getValues();
    const ids = {};
    
    data.slice(1).forEach(row => {
      const key = row[0];
      const value = row[1];
      if (key && value) {
        ids[key] = value;
      }
    });
    
    return ids;
  } catch (e) {
    Logger.log("Impossible de charger les ID système : " + e.toString());
    throw new Error("Impossible de charger les ID système. Vérifiez l'onglet 'sys_ID_Fichiers'. Erreur: " + e.message);
  }
}

/**
 * Récupère les données de configuration d'une ligne spécifique de manière robuste.
 * @param {number} rowIndex Le numéro de la ligne à lire.
 * @returns {Object} Un objet contenant la configuration, avec des clés correspondant exactement aux en-têtes.
 */
function getConfigurationFromRow(rowIndex) {
    const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION);
    const sheet = ss.getSheetByName('Paramètres Généraux');
    if (!sheet) {
      throw new Error("L'onglet 'Paramètres Généraux' est introuvable.");
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const config = {};
    headers.forEach((header, i) => {
        if (header) {
            config[header] = rowValues[i];
        }
    });

    Logger.log("Configuration lue depuis la ligne " + rowIndex + " : " + JSON.stringify(config, null, 2));
    return config;
}

/**
 * Mappe un code langue à son nom complet.
 * @param {string} code Le code de la langue (ex: 'FR').
 * @returns {string} Le nom complet de la langue.
 */
function getLangueFullName(code) {
  const map = { 'FR': 'Français', 'EN': 'English', 'ES': 'Español', 'DE': 'Deutsch' };
  return map[code.toUpperCase()] || code.toUpperCase();
}

/**
 * Crée un item dans le formulaire Google en fonction de son type.
 * Gère QCU (radio), QRM (checkbox) et les différentes sources d'options (JSON/V1).
 * @param {GoogleAppsScript.Forms.Form} form L'objet formulaire auquel ajouter l'item.
 * @param {string} type Le type de question (ex: 'QCU_CAT', 'ECHELLE_NOTE').
 * @param {string} titre Le titre de la question.
 * @param {string} optionsString Une chaîne de caractères contenant les options, séparées par ';'.
 * @param {string} description Une chaîne de caractères pour la description / texte d'aide.
 * @param {string} paramsJSONString Une chaîne de caractères contenant les paramètres V2 au format JSON.
 */
function creerItemFormulaire(form, type, titre, optionsString, description, paramsJSONString) {

    // --- DÉBUT MODIFICATION V2.6 ---
    let finalDescription = description;
    const placeholderRegex = /\[LIEN_FICHIER:(.*?)\]/;
    const match = description ? description.match(placeholderRegex) : null;

    if (match && match[1]) {
        const nomFichier = match[1].trim();
        try {
            const systemIds = getSystemIds();
            const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
            const listeFichiersSheet = bdd.getSheetByName('Liste_Fichiers_Drive');

            if (listeFichiersSheet) {
                const data = listeFichiersSheet.getDataRange().getValues();
                const fileRow = data.find(row => row[0].toString().trim() === nomFichier);

                if (fileRow && fileRow[1]) {
                    const fileId = fileRow[1].toString().trim();
                    const fileUrl = `https://drive.google.com/file/d/${fileId}/view`;
                    finalDescription = description.replace(placeholderRegex, fileUrl);
                } else {
                    finalDescription = description.replace(placeholderRegex, `[ERREUR: Fichier '${nomFichier}' introuvable dans la BDD]`);
                }
            } else {
                finalDescription = description.replace(placeholderRegex, `[ERREUR: Onglet 'Liste_Fichiers_Drive' introuvable]`);
            }
        } catch (e) {
            Logger.log("Erreur lors de la recherche du lien de fichier : " + e.message);
            finalDescription = description.replace(placeholderRegex, `[ERREUR SCRIPT: ${e.message}]`);
        }
    }
    // --- FIN MODIFICATION V2.6 ---

    let params = null;
    let choices = [];
    let item; 

    if (paramsJSONString && typeof paramsJSONString === 'string' && paramsJSONString.trim().startsWith('{')) {
        try {
            params = JSON.parse(paramsJSONString);
        } catch (e) {
            item = form.addParagraphTextItem().setTitle("[Erreur V2: JSON invalide] " + titre);
            if(finalDescription) item.setHelpText(finalDescription);
            return;
        }
    }

    if (params && params.options && Array.isArray(params.options) && params.options.length > 0) {
        choices = params.options.map(opt => (typeof opt === 'object' && opt !== null) ? opt.libelle : opt);
    } else if (optionsString) {
        choices = optionsString.split(';').map(String);
    }

    const formItemType = type ? type.toUpperCase() : '';

    if (formItemType.startsWith('QRM')) {
        if (choices.length > 0) {
            item = form.addCheckboxItem().setTitle(titre).setChoiceValues(choices).setRequired(true);
        } else {
            item = form.addParagraphTextItem().setTitle("[Erreur QRM: Options manquantes] " + titre);
        }
    } else if (formItemType.startsWith('QCU')) {
        if (choices.length > 0) {
            item = form.addMultipleChoiceItem().setTitle(titre).setChoiceValues(choices).setRequired(true);
        } else {
            item = form.addParagraphTextItem().setTitle("[Erreur QCU: Options manquantes] " + titre);
        }
    } else if (formItemType === 'ECHELLE_NOTE') { // =================== DÉBUT DE LA CORRECTION ===================
        if (params && params.echelle_min !== undefined && params.echelle_max !== undefined) {
            const scaleItem = form.addScaleItem()
                .setTitle(titre)
                .setBounds(params.echelle_min, params.echelle_max)
                .setRequired(true);
            if (params.label_min && params.label_max) {
                scaleItem.setLabels(params.label_min, params.label_max);
            }
            item = scaleItem;
        } else {
            item = form.addParagraphTextItem().setTitle("[Erreur ECHELLE_NOTE: Paramètres JSON manquants] " + titre);
        }
    } else if (formItemType === 'ECHELLE') { // Ancien mode conservé pour compatibilité
        const libelles = description ? description.split(';') : [];
        if (optionsString && optionsString.split(';').length >= 2) {
            const bounds = optionsString.split(';').map(Number);
            const scaleItem = form.addScaleItem().setTitle(titre).setBounds(bounds[0], bounds[bounds.length - 1]).setRequired(true);
            if (libelles.length === 2) {
                scaleItem.setLabels(libelles[0], libelles[1]);
            }
            item = scaleItem;
        } else {
            item = form.addParagraphTextItem().setTitle("[Erreur ECHELLE: Bornes manquantes] " + titre);
        }
    } else if (formItemType === 'TEXTE_EMAIL') {
        const textItem = form.addTextItem().setTitle(titre).setRequired(true);
        const emailValidation = FormApp.createTextValidation()
            .setHelpText("Veuillez entrer une adresse e-mail valide.")
            .requireTextIsEmail()
            .build();
        item = textItem.setValidation(emailValidation);
    } else if (formItemType === 'TEXTE_COURT') {
        item = form.addTextItem().setTitle(titre).setRequired(true);
    } else {
        item = form.addParagraphTextItem().setTitle("[Type Inconnu: " + type + "] " + titre);
    } // =================== FIN DE LA CORRECTION ===================

    if (finalDescription && item && typeof item.setHelpText === 'function') {
        // Pour les échelles, la description est utilisée pour les labels, ne pas l'écraser.
        if(formItemType !== 'ECHELLE' && formItemType !== 'ECHELLE_NOTE') {
             item.setHelpText(finalDescription);
        }
    }
}