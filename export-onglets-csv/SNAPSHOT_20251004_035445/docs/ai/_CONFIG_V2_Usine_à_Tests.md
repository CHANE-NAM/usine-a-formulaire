# _CONFIG_V2_Usine_à_Tests

> Généré automatiquement depuis **scripts__CONFIG_V2_Usine_à_Tests.txt** — snapshot: **SNAPSHOT_20251004_035445**.

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\02_configuration\appsscript.json

```json

{
  "timeZone": "Indian/Mauritius",
  "dependencies": {
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\02_configuration\Menu.js

```javascript

// =================================================================================
// == FICHIER : Menu.js
// == VERSION : 4.7 - Finalisation de la sauvegarde de tous les champs.
// == RÃ”LE  : Logique cÃ´tÃ© serveur pour l'application web de configuration.
// =================================================================================

const ID_FEUILLE_CONFIG = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";
const ID_MODELE_FICHE_TEST = "1W_amKwp5kyyGWmg5LTaIQe5K8Gzxf_qvcjGskRy1Sq8";

/**
 * CrÃ©e le menu de l'application Ã  l'ouverture de la feuille de calcul.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ðŸ­ Usine')
    .addSubMenu(ui.createMenu('Configuration')
      .addItem('Configurer un nouveau test...', 'showConfigurationSidebar')
      .addItem('Modifier un test existant...', 'showEditSidebar_UI')
      .addItem('Dupliquer un test existant...', 'showDuplicateUI'))
    .addSubMenu(ui.createMenu('Validation')
      .addItem('VÃ©rifier les en-tÃªtes (CONFIG, BDD, TEMPLATE)', 'validateAllHeaders'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Documents')
      .addItem('GÃ©nÃ©rer la fiche de test (imprimable)...', 'showPrintableSheetUI'))
    .addToUi();
}

/**
 * Affiche la barre latÃ©rale de configuration (FormulaireUI.html).
 */
function showConfigurationSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('FormulaireUI').setTitle('Configuration Usine Ã  Tests').setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * RÃ©cupÃ¨re les donnÃ©es initiales pour peupler les listes dÃ©roulantes de la barre latÃ©rale.
 * @returns {Object} Un objet contenant les listes d'options.
 */
function getInitialData() {
  try {
    const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
    const optionsSheet = ss.getSheetByName("sys_Options_Parametres");
    if (!optionsSheet) throw new Error("L'onglet 'sys_Options_Parametres' est introuvable.");
    
    const optionsData = optionsSheet.getDataRange().getValues();
    const headers = optionsData.shift().map(h => String(h || '').trim());
    const optionsMap = {};
    headers.forEach((header, i) => {
      if (header) optionsMap[header] = optionsData.map(row => row[i]).filter(String);
    });

    let availableMetaBlocks = [];
    try {
      const systemIds = getSystemIds();
      const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
      const metaSheet = bdd.getSheetByName('Questions_META_FR');
      if (metaSheet) {
          const metaData = metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, 3).getValues();
          availableMetaBlocks = metaData.map(row => ({ id: row[0], title: row[2] })).filter(block => block.id && block.title);
      }
    } catch(e) {
      console.error("Impossible de charger les blocs mÃ©ta : " + e.message);
    }
    optionsMap.availableMetaBlocks = availableMetaBlocks;
    return optionsMap;
  } catch (err) {
    Logger.log("ERREUR FATALE dans getInitialData: " + err.stack);
    throw new Error("Erreur cÃ´tÃ© serveur : " + err.message);
  }
}

/**
 * RÃ©cupÃ¨re le nombre de questions disponibles pour un type de test donnÃ© depuis la BDD.
 * @param {string} typeTest Le type de test (ex: 'CouleursV6').
 * @returns {number} Le nombre de questions.
 */
function getQuestionCountForTestType(typeTest) {
  if (!typeTest) return 0;
  try {
    const systemIds = getSystemIds();
    if (systemIds && systemIds.ID_BDD) {
      const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
      const questionSheet = bdd.getSheets().find(s => s.getName().startsWith('Questions_' + typeTest));
      return questionSheet ? questionSheet.getLastRow() - 1 : 0;
    }
    return 0;
  } catch (err) {
    Logger.log('Erreur getQuestionCountForTestType pour ' + typeTest + ': ' + err.message);
    return 0;
  }
}

/**
 * Traite les donnÃ©es soumises depuis la barre latÃ©rale et crÃ©e une nouvelle ligne de configuration.
 * @param {Object} formObject L'objet contenant les donnÃ©es du formulaire.
 * @returns {string} Un message de succÃ¨s.
 */
function processNewTestConfiguration(formObject) {
  try {
    const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
    const paramsSheet = ss.getSheetByName("ParamÃ¨tres GÃ©nÃ©raux");
    if (!paramsSheet) throw new Error("L'onglet 'ParamÃ¨tres GÃ©nÃ©raux' est introuvable.");

    const headers = paramsSheet.getRange(1, 1, 1, paramsSheet.getLastColumn()).getValues()[0];
    
    const dataRow = {};
    
    // Cartographie complÃ¨te des donnÃ©es du formulaire vers les colonnes de la feuille
    dataRow['Titre_Formulaire_Utilisateur'] = formObject.titre;
    dataRow['Sous-Titre_Formulaire'] = formObject.sousTitre;
    dataRow['Type_Test'] = formObject.type;
    dataRow['nbQuestions'] = formObject.nbQuestions;
    dataRow['Limite_Lignes_A_Traiter'] = getQuestionCountForTestType(formObject.type);
    dataRow['Blocs_Meta_A_Inclure'] = formObject.blocsMeta ? formObject.blocsMeta.join(',') : '';
    
    if (String(formObject.repondantContenu).includes('Niveau1')) dataRow['ID_Gabarit_Email_Repondant'] = 'RESULTATS_N1';
    else if (String(formObject.repondantContenu).includes('Niveau2')) dataRow['ID_Gabarit_Email_Repondant'] = 'RESULTATS_N2';
    else if (String(formObject.repondantContenu).includes('Niveau3')) dataRow['ID_Gabarit_Email_Repondant'] = 'RESULTATS_N3';

    // Emails
    dataRow['Repondant_Email_Actif'] = formObject.repondantActif ? "Oui" : "Non";
    dataRow['Repondant_Quand'] = formObject.repondantQuand;
    dataRow['Repondant_Contenu'] = formObject.repondantContenu;
    dataRow['Patron_Email_Mode'] = formObject.patronActif ? "Oui" : "Non";
    dataRow['Patron_Email'] = formObject.patronEmail;
    dataRow['Patron_Quand'] = formObject.patronQuand;
    dataRow['Patron_Contenu'] = formObject.patronContenu;
    dataRow['Formateur_Email_Actif'] = formObject.formateurActif ? "Oui" : "Non";
    dataRow['Formateur_Nom'] = formObject.formateurNom;
    dataRow['Formateur_Email'] = formObject.formateurEmail;
    dataRow['Formateur_Quand'] = formObject.formateurQuand;
    dataRow['Formateur_Contenu'] = formObject.formateurContenu;
    dataRow['Developpeur_Email'] = formObject.devEmail || "chanenam@gmail.com";
    dataRow['Email_Alias'] = formObject.Email_Alias;

    // Moteur & AccÃ¨s
    dataRow['Moteur_Calcul'] = formObject.Moteur_Calcul;
    dataRow['Mode_Acces_Test'] = formObject.Mode_Acces_Test;
    
    // Paiement
    dataRow['PAYMENT_PROVIDER'] = formObject.PAYMENT_PROVIDER;
    dataRow['BYPASS_PAYMENT'] = formObject.BYPASS_PAYMENT;
    dataRow['REQUIRE_PASSWORD'] = formObject.REQUIRE_PASSWORD;
    dataRow['FORM2_PASSWORD'] = formObject.FORM2_PASSWORD;
    
    // Livrables
    dataRow['DELIVERABLE_TYPE'] = formObject.DELIVERABLE_TYPE;
    dataRow['DELIVERABLE_TTL_MIN'] = formObject.DELIVERABLE_TTL_MIN;

    // Contexte & RGPD
    dataRow['CTX_ASK_ROLE'] = formObject.CTX_ASK_ROLE;
    dataRow['CTX_ASK_DEPARTMENT'] = formObject.CTX_ASK_DEPARTMENT;
    dataRow['CTX_ASK_RGPD'] = formObject.CTX_ASK_RGPD;
    dataRow['CTX_COUNTRY_SOURCE'] = formObject.CTX_COUNTRY_SOURCE;

    dataRow['Statut'] = 'En construction';
    
    const nouvelleLigne = headers.map(header => dataRow[header] !== undefined ? dataRow[header] : '');
    
    paramsSheet.appendRow(nouvelleLigne);
    return "Configuration enregistrÃ©e avec succÃ¨s !";
  } catch (e) {
    Logger.log("ERREUR processNewTestConfiguration: " + e.stack);
    throw new Error("Erreur interne lors de la sauvegarde : " + e.message);
  }
}

/**
 * Affiche une UI pour demander la ligne Ã  modifier.
 */
function showEditSidebar_UI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Modifier un Test', 'Veuillez entrer le numÃ©ro de la ligne Ã  modifier :', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowIndex = parseInt(response.getResponseText());
    if (!isNaN(rowIndex) && rowIndex > 1) {
      showEditSidebar(rowIndex);
    } else {
      ui.alert('NumÃ©ro de ligne invalide.');
    }
  }
}

/**
 * Affiche la barre latÃ©rale d'Ã©dition pour une ligne donnÃ©e.
 * @param {number} rowIndex Le numÃ©ro de la ligne Ã  Ã©diter.
 */
function showEditSidebar(rowIndex) {
  const template = HtmlService.createTemplateFromFile('ModifierTestUI');
  template.rowIndex = rowIndex;
  const html = template.evaluate().setTitle('Ã‰dition du Test (Ligne ' + rowIndex + ')').setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * RÃ©cupÃ¨re les donnÃ©es d'une ligne de configuration pour l'Ã©dition.
 * @param {number} rowIndex Le numÃ©ro de la ligne.
 * @returns {Object} Un objet avec les en-tÃªtes et les valeurs de la ligne.
 */
function getTestDataForEdit(rowIndex) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG).getSheetByName("ParamÃ¨tres GÃ©nÃ©raux");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { headers: headers, values: values };
}

/**
 * Met Ã  jour les donnÃ©es d'une ligne de configuration.
 * @param {number} rowIndex Le numÃ©ro de la ligne Ã  mettre Ã  jour.
 * @param {Object} updatedData Un objet contenant les nouvelles donnÃ©es.
 * @returns {string} Un message de succÃ¨s.
 */
function updateTestData(rowIndex, updatedData) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG).getSheetByName("ParamÃ¨tres GÃ©nÃ©raux");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndexMap = new Map(headers.map((h, i) => [h, i]));
  for (const header in updatedData) {
    if (colIndexMap.has(header)) {
      const colIndex = colIndexMap.get(header);
      sheet.getRange(rowIndex, colIndex + 1).setValue(updatedData[header]);
    }
  }
  return "Modifications enregistrÃ©es avec succÃ¨s !";
}

/**
 * Affiche une UI pour demander la ligne Ã  dupliquer.
 */
function showDuplicateUI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Dupliquer une Configuration de Test', 'Veuillez entrer le numÃ©ro de la ligne Ã  dupliquer :', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowIndex = parseInt(response.getResponseText());
    if (!isNaN(rowIndex) && rowIndex > 1) {
      try {
        const newRowIndex = duplicateTestConfiguration(rowIndex);
        ui.alert('SuccÃ¨s', `La ligne ${rowIndex} a Ã©tÃ© dupliquÃ©e Ã  la fin de la feuille (nouvelle ligne : ${newRowIndex}).`, ui.ButtonSet.OK);
      } catch (e) {
        ui.alert('Erreur', e.message, ui.ButtonSet.OK);
      }
    } else {
      ui.alert('NumÃ©ro de ligne invalide.');
    }
  }
}

/**
 * Duplique une ligne de configuration.
 * @param {number} rowIndex Le numÃ©ro de la ligne Ã  dupliquer.
 * @returns {number} Le numÃ©ro de la nouvelle ligne crÃ©Ã©e.
 */
function duplicateTestConfiguration(rowIndex) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG).getSheetByName("ParamÃ¨tres GÃ©nÃ©raux");
  if (rowIndex > sheet.getLastRow()) {
    throw new Error("La ligne spÃ©cifiÃ©e n'existe pas.");
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const sourceValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fieldsToClear = ['Id_Unique', 'Nom_Fichier_Complet', 'Lien_Formulaire_Public', 'ID_Formulaire_Cible', 'ID_Sheet_Cible', 'AccÃ¨s Direct Formulaire'];
  const shortUrlHeader = headers.find(h => h && h.toLowerCase().includes('raccourci'));
  if (shortUrlHeader) fieldsToClear.push(shortUrlHeader);
  
  const newValues = headers.map(header => {
    const colIndex = headers.indexOf(header);
    const sourceValue = sourceValues[colIndex];
    if (fieldsToClear.includes(header)) return '';
    if (header === 'Statut') return 'En construction';
    if (header === 'Titre_Formulaire_Utilisateur') return sourceValue + ' (Copie)';
    return sourceValue;
  });
  
  sheet.appendRow(newValues);
  const newRowIndex = sheet.getLastRow();
  sheet.getRange(newRowIndex, 1).activate();
  return newRowIndex;
}

/**
 * Affiche une UI pour gÃ©nÃ©rer une fiche de test imprimable.
 */
function showPrintableSheetUI() {
  if (ID_MODELE_FICHE_TEST === "METTEZ_ICI_L_ID_DE_VOTRE_MODELE_GOOGLE_DOC") {
    SpreadsheetApp.getUi().alert("Configuration requise", "Veuillez d'abord renseigner l'ID de votre modÃ¨le Google Doc.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('GÃ©nÃ©rer une Fiche de Test', 'Veuillez entrer le numÃ©ro de la ligne :', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowIndex = parseInt(response.getResponseText());
    if (!isNaN(rowIndex) && rowIndex > 1) {
      try {
        const fileUrl = generatePrintableSheet(rowIndex);
        const htmlOutput = HtmlService.createHtmlOutput(`<p>La fiche de test a Ã©tÃ© gÃ©nÃ©rÃ©e.</p><a href="${fileUrl}" target="_blank">Ouvrir le document</a>`).setWidth(300).setHeight(100);
        ui.showModalDialog(htmlOutput, 'Document CrÃ©Ã©');
      } catch (e) {
        ui.alert('Erreur', e.message, ui.ButtonSet.OK);
      }
    } else {
      ui.alert('NumÃ©ro de ligne invalide.');
    }
  }
}

/**
 * GÃ©nÃ¨re une fiche de test Ã  partir d'un modÃ¨le Google Doc.
 * @param {number} rowIndex Le numÃ©ro de la ligne de configuration Ã  utiliser.
 * @returns {string} L'URL du document gÃ©nÃ©rÃ©.
 */
function generatePrintableSheet(rowIndex) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG).getSheetByName("ParamÃ¨tres GÃ©nÃ©raux");
  if (rowIndex > sheet.getLastRow()) throw new Error("La ligne n'existe pas.");
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dataForFusion = {};
  headers.forEach((header, i) => {
    dataForFusion[header] = values[i];
  });

  const templateFile = DriveApp.getFileById(ID_MODELE_FICHE_TEST);
  const destinationFolder = DriveApp.getRootFolder();
  const testTitle = dataForFusion['Titre_Formulaire_Utilisateur'] || 'Fiche de Test';
  const newFileName = `Fiche - ${testTitle}`;
  const newFile = templateFile.makeCopy(newFileName, destinationFolder);
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();
  
  for (const key in dataForFusion) {
    body.replaceText(`{{${key}}}`, dataForFusion[key]);
  }
  
  doc.saveAndClose();
  Logger.log(`Document gÃ©nÃ©rÃ© : ${newFile.getName()} (ID: ${newFile.getId()})`);
  return newFile.getUrl();
}

/**
 * RÃ©cupÃ¨re les IDs des fichiers systÃ¨me (BDD, etc.) depuis l'onglet sys_ID_Fichiers.
 * @returns {Object} Un dictionnaire des IDs.
 */
function getSystemIds() {
  const configSS = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
  const idSheet = configSS.getSheetByName('sys_ID_Fichiers');
  if (!idSheet) { throw new Error("L'onglet 'sys_ID_Fichiers' est introuvable."); }
  
  const data = idSheet.getDataRange().getValues();
  const ids = {};
  data.slice(1).forEach(row => {
    if (row[0] && row[1]) ids[row[0]] = row[1];
  });
  return ids;
}
```

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\02_configuration\UtilitaireConversion.js

```javascript

// Remplacez cette variable par l'ID de votre feuille de calcul [CONFIG]V2 Usine Ã  Tests.
// const ID_FEUILLE_CONFIG = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";

/**
 * Fonction Ã  usage unique pour convertir toutes les URLs de formulaires existantes
 * dans l'onglet 'ParamÃ¨tres GÃ©nÃ©raux' en leurs versions courtes (forms.gle).
 */
function convertirLiensExistantsEnCourts() {
  const nomOnglet = "ParamÃ¨tres GÃ©nÃ©raux";
  
  try {
    const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
    const sheet = ss.getSheetByName(nomOnglet);
    
    if (!sheet) {
      throw new Error(`L'onglet "${nomOnglet}" est introuvable.`);
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    
    // Trouve automatiquement la colonne contenant les liens
    const linkColumnIndex = headers.indexOf("Lien_Formulaire_Public");
    if (linkColumnIndex === -1) {
      throw new Error("La colonne 'Lien_Formulaire_Public' est introuvable.");
    }

    // Boucle sur chaque ligne (en sautant l'en-tÃªte)
    for (let i = 1; i < values.length; i++) {
      const longUrl = values[i][linkColumnIndex];
      
      // Ne traite que les URLs longues et non vides
      if (longUrl && typeof longUrl === 'string' && longUrl.includes("docs.google.com/forms")) {
        // Extrait l'ID du formulaire Ã  partir de l'URL longue
        const formId = longUrl.split('/d/')[1].split('/')[0];
        
        if (formId) {
          // Ouvre le formulaire par son ID et obtient l'URL courte
          const form = FormApp.openById(formId);
          const shortUrl = form.getShortUrl();
          
          // Met Ã  jour la cellule avec la nouvelle URL courte
          // Les indices de range commencent Ã  1, donc i+1 et linkColumnIndex+1
          sheet.getRange(i + 1, linkColumnIndex + 1).setValue(shortUrl);
          Logger.log(`Ligne ${i + 1}: URL convertie pour le formulaire ${formId}`);
        }
      }
    }
    
    SpreadsheetApp.getUi().alert("Conversion terminÃ©e avec succÃ¨s !");
    
  } catch (e) {
    Logger.log(`Erreur lors de la conversion : ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`Une erreur est survenue : ${e.message}`);
  }
}
```

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\02_configuration\ValidationRunner.js

```javascript

/** ValidationRunner.gs â€” Runner de validation des en-tÃªtes (CONFIG â†’ BDD â†’ TEMPLATE)
 * Ajoute un menu "Validation" dans le classeur CONFIG pour vÃ©rifier les onglets requis
 * et les en-tÃªtes attendues, en se basant STRICTEMENT sur les noms dâ€™en-tÃªtes (jamais dâ€™indices).
 * Rapport affichÃ© en sidebar (HTML).
 *
 * PRÃ‰REQUIS :
 * - Avoir les IDs des 3 classeurs :
 *    - ID_CONFIG : celui du classeur courant (dÃ©tectÃ© automatiquement)
 *    - ID_BDD, ID_TEMPLATE : soit lus dâ€™un onglet "sys_ID_Fichiers" (si prÃ©sent),
 *      soit saisis en dur ci-dessous dans FALLBACK_IDS.
 */

/** ================ CONFIGURATION MINIMALE ================ **/
const FALLBACK_IDS = {
  ID_BDD:      '1m2MGBd0nyiAl3qw032B6Nfj7zQL27bRSBexiOPaRZd8', // â† remplace si besoin
  ID_TEMPLATE: '1XwyTt9hcFLd-_IrCYuKY4_E6Dw9aUrls-AGQp65dzDU'  // â† remplace si besoin
};
// Variantes de noms acceptÃ©es pour certains onglets (tolÃ©rance orthographe/accents)
const SHEET_NAME_VARIANTS = {
  'ParamÃ¨tres GÃ©nÃ©raux': ['ParamÃ¨tres GÃ©nÃ©raux','Parametres Generaux','Parameters','ParamÃ¨tres Generaux','Parametres GÃ©nÃ©raux']
};

/** ================ MENU ================== **/
function addValidationMenu_(ui) {
  ui.createMenu('Validation')
    .addItem('VÃ©rifier les en-tÃªtes (CONFIG, BDD, TEMPLATE)', 'validateAllHeaders')
    .addToUi();
}
/** ================ HELPERS ================== **/
function normalizeHeader_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // retire accents
    .replace(/\s+/g, ' ') // espaces multiples â†’ un espace
    .trim()
    .toLowerCase();
}

function getHeaderRow_(sheet) {
  if (!sheet) return [];
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1,1,1,lastCol).getValues()[0] || [];
}

function findSheetByVariants_(ss, canonicalName) {
  const variants = SHEET_NAME_VARIANTS[canonicalName] || [canonicalName];
  for (const name of variants) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  return null;
}

function assertHeaders_(sheet, requiredNames, report, context) {
  const headers = getHeaderRow_(sheet);
  const normalized = headers.map(normalizeHeader_);
  const missing = requiredNames
    .map(normalizeHeader_)
    .filter(req => !normalized.includes(req));

  if (missing.length) {
    report.push({
      classeur: context.classeur,
      onglet: sheet ? sheet.getName() : context.ongletAttendu,
      type: 'En-tÃªtes manquantes',
      details: 'Manquantes: ' + missing.join(', '),
      headersTrouves: headers.join(' | ')
    });
  }
}

function getSystemIdsFromConfig_() {
  // Essaie de lire dans un onglet "sys_ID_Fichiers" (2 colonnes : ClÃ©, Valeur)
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('sys_ID_Fichiers');
  const ids = { ID_CONFIG: ss.getId(), ID_BDD: FALLBACK_IDS.ID_BDD, ID_TEMPLATE: FALLBACK_IDS.ID_TEMPLATE };

  if (!sh) return ids;

  const values = sh.getDataRange().getValues();
  const map = {};
  values.forEach(row => {
    const k = String(row[0] || '').trim();
    const v = String(row[1] || '').trim();
    if (k && v) map[k] = v;
  });

  if (map.ID_CONFIG)   ids.ID_CONFIG = map.ID_CONFIG;
  if (map.ID_BDD)      ids.ID_BDD = map.ID_BDD;
  if (map.ID_TEMPLATE) ids.ID_TEMPLATE = map.ID_TEMPLATE;

  return ids;
}

function htmlReport_(rows) {
  const esc = s => String(s||'').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
  const head = `
    <style>
      body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;padding:12px}
      h2{margin:0 0 10px 0}
      table{border-collapse:collapse;width:100%}
      th,td{border:1px solid #ddd;padding:6px;vertical-align:top}
      th{background:#fafafa;text-align:left}
      tr:nth-child(even){background:#fcfcfc}
      .ok{color:#2e7d32}
      .err{color:#b71c1c}
    </style>`;
  if (!rows.length) {
    return HtmlService.createHtmlOutput(head + `<h2>Validation des en-tÃªtes</h2><p class="ok">Aucune anomalie dÃ©tectÃ©e ðŸ‘</p>`)
      .setTitle('Validation en-tÃªtes');
  }
  const rowsHtml = rows.map(r => `
    <tr>
      <td>${esc(r.classeur)}</td>
      <td>${esc(r.onglet)}</td>
      <td class="err">${esc(r.type)}</td>
      <td>${esc(r.details)}</td>
      <td>${esc(r.headersTrouves || '')}</td>
    </tr>`).join('');
  const html = `
    ${head}
    <h2>Validation des en-tÃªtes</h2>
    <table>
      <thead><tr>
        <th>Classeur</th><th>Onglet</th><th>Type</th><th>DÃ©tails</th><th>En-tÃªtes trouvÃ©es</th>
      </tr></thead>
      <tbody>${rowsHtml}</tbody>
    </table>`;
  return HtmlService.createHtmlOutput(html).setTitle('Validation en-tÃªtes');
}

/** ================ RUNNER PRINCIPAL ================== **/
function validateAllHeaders() {
  const report = [];
  const { ID_CONFIG, ID_BDD, ID_TEMPLATE } = getSystemIdsFromConfig_();

  // --- CONFIG ---
  try {
    const ssCfg = SpreadsheetApp.openById(ID_CONFIG);
    const shParams = findSheetByVariants_(ssCfg, 'ParamÃ¨tres GÃ©nÃ©raux');
    if (!shParams) {
      report.push({ classeur:'CONFIG', onglet:'ParamÃ¨tres GÃ©nÃ©raux', type:'Onglet manquant', details:'Aucune variante trouvÃ©e (ParamÃ¨tres GÃ©nÃ©raux/Parametres Generaux/Parameters)' });
    } else {
      assertHeaders_(shParams, [
        'Type_Test','Repondant_Quand','Repondant_Contenu','Patron_Quand','Patron_Contenu','Formateur_Quand','Formateur_Contenu'
      ], report, { classeur:'CONFIG', ongletAttendu:'ParamÃ¨tres GÃ©nÃ©raux' });
    }
  } catch (e) {
    report.push({ classeur:'CONFIG', onglet:'*', type:'Erreur', details:String(e) });
  }

  // --- BDD ---
  try {
    const ssBdd = SpreadsheetApp.openById(ID_BDD);
    // sys_Composition_Emails
    const shCompo = ssBdd.getSheetByName('sys_Composition_Emails');
    if (!shCompo) {
      report.push({ classeur:'BDD', onglet:'sys_Composition_Emails', type:'Onglet manquant', details:'Non trouvÃ©' });
    } else {
      assertHeaders_(shCompo, [
        'Type_Test','Code_Langue','Code_Niveau_Email','Code_Profil','Element','Ordre','Contenu / ID_Document'
      ], report, { classeur:'BDD', ongletAttendu:'sys_Composition_Emails' });
    }
    // sys_PiecesJointes
    const shPJ = ssBdd.getSheetByName('sys_PiecesJointes');
    if (!shPJ) {
      report.push({ classeur:'BDD', onglet:'sys_PiecesJointes', type:'Onglet manquant', details:'Non trouvÃ©' });
    } else {
      assertHeaders_(shPJ, [
        'Type_Test','Code_Langue','Code_Niveau_Email','Code_Profil','ID_Document','Nom_Fichier'
      ], report, { classeur:'BDD', ongletAttendu:'sys_PiecesJointes' });
    }
    // Questions_META_FR (optionnel mais frÃ©quent)
    const shMeta = ssBdd.getSheetByName('Questions_META_FR');
    if (shMeta) {
      assertHeaders_(shMeta, [
        'ID_Question','Libelle','Type','Obligatoire','Bloc'
      ], report, { classeur:'BDD', ongletAttendu:'Questions_META_FR' });
    }
  } catch (e) {
    report.push({ classeur:'BDD', onglet:'*', type:'Erreur', details:String(e) });
  }

  // --- TEMPLATE ---
  try {
    const ssTpl = SpreadsheetApp.openById(ID_TEMPLATE);
    // Onglet de config attendu cÃ´tÃ© TEMPLATE (Ã  adapter si besoin)
    const shTplCfg = ssTpl.getSheetByName('sys_Template_Config');
    if (shTplCfg) {
      assertHeaders_(shTplCfg, [
        'Cle','Valeur'
      ], report, { classeur:'TEMPLATE', ongletAttendu:'sys_Template_Config' });
    } else {
      // pas bloquant : beaucoup de logique template est dans la BDD (compo emails / PJ)
      // On signale juste l'absence d'onglet de config local si on s'y attendait :
      // report.push({ classeur:'TEMPLATE', onglet:'sys_Template_Config', type:'Onglet manquant', details:'Optionnel mais recommandÃ©' });
    }
  } catch (e) {
    report.push({ classeur:'TEMPLATE', onglet:'*', type:'Erreur', details:String(e) });
  }

  // Affiche le rapport
  const out = htmlReport_(report);
  SpreadsheetApp.getUi().showSidebar(out);
}


```

