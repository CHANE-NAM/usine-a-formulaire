# _CONFIG_V2_Usine_à_Tests

> Généré automatiquement depuis **scripts__CONFIG_V2_Usine_à_Tests.txt** — snapshot: **SNAPSHOT_20250925_210002**.

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
// == FICHIER : Menu.gs
// == VERSION : 4.3 - Ajout de la gestion du Type de Moteur dans le formulaire.
// == RÃ”LE  : Logique cÃ´tÃ© serveur pour l'application web de configuration.
// =================================================================================

const ID_FEUILLE_CONFIG = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";
// ID du modÃ¨le pour la fiche de test (catalogue)
const ID_MODELE_FICHE_TEST = "1W_amKwp5kyyGWmg5LTaIQe5K8Gzxf_qvcjGskRy1Sq8";


// --- SECTION 1 : INTERFACE UTILISATEUR (MENU) ---

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const main = ui.createMenu('ðŸ­ Usine');
  const conf = ui.createMenu('Configuration')
    .addItem('Configurer un nouveau test...', 'showConfigurationSidebar')
    .addItem('Modifier un test existant...', 'showEditSidebar_UI')
    .addItem('Dupliquer un test existant...', 'showDuplicateUI');

  const val = ui.createMenu('Validation')
    .addItem('VÃ©rifier les en-tÃªtes (CONFIG, BDD, TEMPLATE)', 'validateAllHeaders');
    
  // --- NOUVEAU MENU ---
  const docs = ui.createMenu('Documents')
    .addItem('GÃ©nÃ©rer la fiche de test (imprimable)...', 'showPrintableSheetUI');

  main.addSubMenu(conf);
  main.addSubMenu(val);
  main.addSeparator();
  main.addSubMenu(docs);
  main.addToUi();
}


// --- SECTION 2 : FONCTIONS POUR LA CRÃ‰ATION D'UN NOUVEAU TEST ---

function showConfigurationSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('FormulaireUI')
      .setTitle('Configuration Usine Ã  Tests')
      .setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getInitialData() {
  const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
  const optionsSheet = ss.getSheetByName("sys_Options_Parametres");
  if (!optionsSheet) {
    throw new Error("L'onglet 'sys_Options_Parametres' est introuvable.");
  }

  const optionsData = optionsSheet.getDataRange().getValues();
  const headers = optionsData.shift().map(h => String(h || '').trim());
  const optionsMap = {};

  headers.forEach((header, i) => {
    const options = optionsData.map(row => row[i]).filter(String);
    optionsMap[header] = options;
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
    console.error("Impossible de charger les blocs mÃ©ta depuis la BDD : " + e.message);
  }

  return {
    typesDeTest: optionsMap['Type_Test'] || [],
    // MODIFICATION : On ajoute la liste des types de moteur
    typesDeMoteur: optionsMap['Type_Moteur'] || [],
    availableMetaBlocks: availableMetaBlocks,
    options: {
      Repondant_Quand: optionsMap['Repondant_Quand'] || [],
      Repondant_Contenu: optionsMap['Repondant_Contenu'] || [],
      Patron_Quand: optionsMap['Patron_Quand'] || [],
      Patron_Contenu: optionsMap['Patron_Contenu'] || [],
      Formateur_Quand: optionsMap['Formateur_Quand'] || [],
      Formateur_Contenu: optionsMap['Formateur_Contenu'] || []
    }
  };
}

function getQuestionCountForTestType(typeTest) {
  if (!typeTest) return 0;
  try {
    const systemIds = getSystemIds();
    if (systemIds && systemIds.ID_BDD) {
      const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
      const questionSheet = bdd.getSheets().find(s => s.getName().startsWith('Questions_' + typeTest));
      if (questionSheet) {
        return questionSheet.getLastRow() - 1;
      }
    }
    return 0;
  } catch (err) {
    Logger.log('Erreur lors du calcul du nombre de questions pour ' + typeTest + ': ' + err.message);
    return 0;
  }
}

function processNewTestConfiguration(formObject) {
  try {
    const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
    const paramsSheet = ss.getSheetByName("ParamÃ¨tres GÃ©nÃ©raux");
    if (!paramsSheet) { throw new Error("L'onglet 'ParamÃ¨tres GÃ©nÃ©raux' est introuvable."); }
    
    let headers = paramsSheet.getRange(1, 1, 1, paramsSheet.getLastColumn()).getValues()[0];
    const requiredHeaders = ['Blocs_Meta_A_Inclure', 'ID_Gabarit_Email_Repondant', 'Email_Alias', 'Moteur_Calcul'];

    requiredHeaders.forEach(headerName => {
        if (headers.indexOf(headerName) === -1) {
            paramsSheet.getRange(1, paramsSheet.getLastColumn() + 1).setValue(headerName);
        }
    });
    headers = paramsSheet.getRange(1, 1, 1, paramsSheet.getLastColumn()).getValues()[0];
    
    let emailDev = formObject.devEmail;
    if (!emailDev || emailDev.trim() === "") { emailDev = "chanenam@gmail.com"; }

    const limiteLignes = getQuestionCountForTestType(formObject.type);
    const blocsMetaString = formObject.blocsMeta.join(',');

    let idGabaritRepondant = '';
    if (formObject.repondantContenu && formObject.repondantContenu.includes('Niveau1')) {
        idGabaritRepondant = 'RESULTATS_N1';
    } else if (formObject.repondantContenu && formObject.repondantContenu.includes('Niveau2')) {
        idGabaritRepondant = 'RESULTATS_N2';
    } else if (formObject.repondantContenu && formObject.repondantContenu.includes('Niveau3')) {
        idGabaritRepondant = 'RESULTATS_N3';
    }

    const dataRow = {
      'Id_Unique': '', 'Titre_Formulaire_Utilisateur': formObject.titre, 'Nom_Fichier_Complet': '',
      'Statut': 'En construction', 'Type_Test': formObject.type, 
      // MODIFICATION : On utilise la valeur du formulaire au lieu de "Universel" en dur
      'Moteur_Calcul': formObject.moteur,
      'Blocs_Meta_A_Inclure': blocsMetaString, 'ID_Gabarit_Email_Repondant': idGabaritRepondant,
      'ID_Dossier_Cible': '', 'Limite_Lignes_A_Traiter': limiteLignes, 'nbQuestions': formObject.nbQuestions,
      'Repondant_Email_Actif': formObject.repondantActif ? "Oui" : "Non", 'Repondant_Quand': formObject.repondantQuand,
      'Repondant_Contenu': formObject.repondantContenu, 'Patron_Email_Mode': formObject.patronActif ? "Oui" : "Non",
      'Patron_Quand': formObject.patronQuand, 'Patron_Contenu': formObject.patronContenu, 'Patron_Email': formObject.patronEmail,
      'Formateur_Email_Actif': formObject.formateurActif ? "Oui" : "Non", 'Formateur_Quand': formObject.formateurQuand,
      'Formateur_Contenu': formObject.formateurContenu, 'Formateur_Email': formObject.formateurEmail,
      'Developpeur_Email': emailDev, 'ID_Formulaire_Cible': '', 'ID_Sheet_Cible': '', 'Email_Alias': formObject.emailAlias
    };

    const nouvelleLigne = headers.map(header => dataRow[header] !== undefined ? dataRow[header] : '');
    paramsSheet.appendRow(nouvelleLigne);
    return "Configuration enregistrÃ©e avec succÃ¨s !";
  } catch (e) {
    Logger.log("ERREUR lors de la sauvegarde de la configuration: " + e.toString());
    throw new Error("Une erreur interne est survenue lors de la sauvegarde. " + e.message);
  }
}

// --- SECTION 3 : FONCTIONS POUR L'Ã‰DITION D'UN TEST EXISTANT ---

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

function showEditSidebar(rowIndex) {
  const template = HtmlService.createTemplateFromFile('ModifierTestUI');
  template.rowIndex = rowIndex;
  const html = template.evaluate().setTitle('Ã‰dition du Test (Ligne ' + rowIndex + ')').setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getTestDataForEdit(rowIndex) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG).getSheetByName("ParamÃ¨tres GÃ©nÃ©raux");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { headers: headers, values: values };
}

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

// --- SECTION 4 : FONCTIONS POUR LA DUPLICATION D'UN TEST ---

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

// --- SECTION 5 : FONCTIONS POUR LA GÃ‰NÃ‰RATION DE DOCUMENTS ---

function showPrintableSheetUI() {
  if (ID_MODELE_FICHE_TEST === "METTEZ_ICI_L_ID_DE_VOTRE_MODELE_GOOGLE_DOC") {
    SpreadsheetApp.getUi().alert("Configuration requise", "Veuillez d'abord renseigner l'ID de votre modÃ¨le Google Doc dans le script Menu.gs (variable ID_MODELE_FICHE_TEST).", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('GÃ©nÃ©rer une Fiche de Test', 'Veuillez entrer le numÃ©ro de la ligne Ã  utiliser pour gÃ©nÃ©rer le document :', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowIndex = parseInt(response.getResponseText());
    if (!isNaN(rowIndex) && rowIndex > 1) {
      try {
        const fileUrl = generatePrintableSheet(rowIndex);
        const htmlOutput = HtmlService.createHtmlOutput(`<p>La fiche de test a Ã©tÃ© gÃ©nÃ©rÃ©e avec succÃ¨s.</p><a href="${fileUrl}" target="_blank">Cliquez ici pour ouvrir le document</a>`).setWidth(300).setHeight(100);
        ui.showModalDialog(htmlOutput, 'Document CrÃ©Ã©');
      } catch (e) {
        ui.alert('Erreur', e.message, ui.ButtonSet.OK);
      }
    } else {
      ui.alert('NumÃ©ro de ligne invalide.');
    }
  }
}

function generatePrintableSheet(rowIndex) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG).getSheetByName("ParamÃ¨tres GÃ©nÃ©raux");
  if (rowIndex > sheet.getLastRow()) throw new Error("La ligne spÃ©cifiÃ©e n'existe pas.");
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

// --- SECTION 6 : FONCTIONS UTILITAIRES ---

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

