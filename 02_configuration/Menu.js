// =================================================================================
// == FICHIER : Menu.gs
// == VERSION : 4.6 - Sauvegarde des nouveaux champs de configuration.
// == RÔLE  : Logique côté serveur pour l'application web de configuration.
// =================================================================================

const ID_FEUILLE_CONFIG = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";
const ID_MODELE_FICHE_TEST = "1W_amKwp5kyyGWmg5LTaIQe5K8Gzxf_qvcjGskRy1Sq8";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏭 Usine')
    .addSubMenu(ui.createMenu('Configuration')
      .addItem('Configurer un nouveau test...', 'showConfigurationSidebar')
      .addItem('Modifier un test existant...', 'showEditSidebar_UI')
      .addItem('Dupliquer un test existant...', 'showDuplicateUI'))
    .addSubMenu(ui.createMenu('Validation')
      .addItem('Vérifier les en-têtes (CONFIG, BDD, TEMPLATE)', 'validateAllHeaders'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Documents')
      .addItem('Générer la fiche de test (imprimable)...', 'showPrintableSheetUI'))
    .addToUi();
}

function showConfigurationSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('FormulaireUI').setTitle('Configuration Usine à Tests').setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

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
      console.error("Impossible de charger les blocs méta : " + e.message);
    }
    optionsMap.availableMetaBlocks = availableMetaBlocks;
    return optionsMap;
  } catch (err) {
    Logger.log("ERREUR FATALE dans getInitialData: " + err.stack);
    throw new Error("Erreur côté serveur : " + err.message);
  }
}

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

function processNewTestConfiguration(formObject) {
  try {
    const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
    const paramsSheet = ss.getSheetByName("Paramètres Généraux");
    if (!paramsSheet) throw new Error("L'onglet 'Paramètres Généraux' est introuvable.");
    
    let headers = paramsSheet.getRange(1, 1, 1, paramsSheet.getLastColumn()).getValues()[0];
    
    // Assure la présence des nouvelles colonnes
    const requiredHeaders = [
        'Blocs_Meta_A_Inclure', 'ID_Gabarit_Email_Repondant', 'Email_Alias', 'Moteur_Calcul',
        'PAYMENT_PROVIDER', 'BYPASS_PAYMENT', 'REQUIRE_PASSWORD', 'FORM2_PASSWORD',
        'DELIVERABLE_TYPE', 'DELIVERABLE_TTL_MIN', 'CTX_ASK_ROLE', 'CTX_ASK_DEPARTMENT',
        'CTX_ASK_RGPD', 'CTX_COUNTRY_SOURCE', 'Mode_Acces_Test'
    ];
    let newHeadersAdded = false;
    requiredHeaders.forEach(headerName => {
        if (headers.indexOf(headerName) === -1) {
            paramsSheet.getRange(1, paramsSheet.getLastColumn() + 1).setValue(headerName);
            newHeadersAdded = true;
        }
    });
    if(newHeadersAdded) headers = paramsSheet.getRange(1, 1, 1, paramsSheet.getLastColumn()).getValues()[0];
    
    let emailDev = formObject.devEmail || "chanenam@gmail.com";
    const limiteLignes = getQuestionCountForTestType(formObject.type);
    const blocsMetaString = formObject.blocsMeta.join(',');

    let idGabaritRepondant = '';
    if (String(formObject.repondantContenu).includes('Niveau1')) idGabaritRepondant = 'RESULTATS_N1';
    else if (String(formObject.repondantContenu).includes('Niveau2')) idGabaritRepondant = 'RESULTATS_N2';
    else if (String(formObject.repondantContenu).includes('Niveau3')) idGabaritRepondant = 'RESULTATS_N3';

    // Fusionne les anciennes et nouvelles données
    const dataRow = { ...formObject }; // Commence avec toutes les nouvelles données
    
    // Ajoute ou écrase avec les données traitées
    Object.assign(dataRow, {
      'Id_Unique': '', 'Titre_Formulaire_Utilisateur': formObject.titre, 'Nom_Fichier_Complet': '',
      'Statut': 'En construction', 'Type_Test': formObject.type, 
      'Moteur_Calcul': formObject.moteur,
      'Blocs_Meta_A_Inclure': blocsMetaString, 'ID_Gabarit_Email_Repondant': idGabaritRepondant,
      'ID_Dossier_Cible': '', 'Limite_Lignes_A_Traiter': limiteLignes, 'nbQuestions': formObject.nbQuestions,
      'Repondant_Email_Actif': formObject.repondantActif ? "Oui" : "Non",
      'Patron_Email_Mode': formObject.patronActif ? "Oui" : "Non",
      'Formateur_Email_Actif': formObject.formateurActif ? "Oui" : "Non",
      'Developpeur_Email': emailDev, 'ID_Formulaire_Cible': '', 'ID_Sheet_Cible': ''
    });

    const nouvelleLigne = headers.map(header => dataRow[header] !== undefined ? dataRow[header] : '');
    paramsSheet.appendRow(nouvelleLigne);
    return "Configuration enregistrée avec succès !";
  } catch (e) {
    Logger.log("ERREUR processNewTestConfiguration: " + e.stack);
    throw new Error("Erreur interne lors de la sauvegarde : " + e.message);
  }
}


// --- SECTIONS 3, 4, 5, 6 (Édition, Duplication, Documents, Utilitaires) restent inchangées ---

function showEditSidebar_UI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Modifier un Test', 'Veuillez entrer le numéro de la ligne à modifier :', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowIndex = parseInt(response.getResponseText());
    if (!isNaN(rowIndex) && rowIndex > 1) {
      showEditSidebar(rowIndex);
    } else {
      ui.alert('Numéro de ligne invalide.');
    }
  }
}

function showEditSidebar(rowIndex) {
  const template = HtmlService.createTemplateFromFile('ModifierTestUI');
  template.rowIndex = rowIndex;
  const html = template.evaluate().setTitle('Édition du Test (Ligne ' + rowIndex + ')').setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getTestDataForEdit(rowIndex) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG).getSheetByName("Paramètres Généraux");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { headers: headers, values: values };
}

function updateTestData(rowIndex, updatedData) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG).getSheetByName("Paramètres Généraux");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndexMap = new Map(headers.map((h, i) => [h, i]));
  for (const header in updatedData) {
    if (colIndexMap.has(header)) {
      const colIndex = colIndexMap.get(header);
      sheet.getRange(rowIndex, colIndex + 1).setValue(updatedData[header]);
    }
  }
  return "Modifications enregistrées avec succès !";
}

function showDuplicateUI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Dupliquer une Configuration de Test', 'Veuillez entrer le numéro de la ligne à dupliquer :', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowIndex = parseInt(response.getResponseText());
    if (!isNaN(rowIndex) && rowIndex > 1) {
      try {
        const newRowIndex = duplicateTestConfiguration(rowIndex);
        ui.alert('Succès', `La ligne ${rowIndex} a été dupliquée à la fin de la feuille (nouvelle ligne : ${newRowIndex}).`, ui.ButtonSet.OK);
      } catch (e) {
        ui.alert('Erreur', e.message, ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Numéro de ligne invalide.');
    }
  }
}

function duplicateTestConfiguration(rowIndex) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG).getSheetByName("Paramètres Généraux");
  if (rowIndex > sheet.getLastRow()) {
    throw new Error("La ligne spécifiée n'existe pas.");
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const sourceValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fieldsToClear = ['Id_Unique', 'Nom_Fichier_Complet', 'Lien_Formulaire_Public', 'ID_Formulaire_Cible', 'ID_Sheet_Cible', 'Accès Direct Formulaire'];
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

function showPrintableSheetUI() {
  if (ID_MODELE_FICHE_TEST === "METTEZ_ICI_L_ID_DE_VOTRE_MODELE_GOOGLE_DOC") {
    SpreadsheetApp.getUi().alert("Configuration requise", "Veuillez d'abord renseigner l'ID de votre modèle Google Doc.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Générer une Fiche de Test', 'Veuillez entrer le numéro de la ligne :', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowIndex = parseInt(response.getResponseText());
    if (!isNaN(rowIndex) && rowIndex > 1) {
      try {
        const fileUrl = generatePrintableSheet(rowIndex);
        const htmlOutput = HtmlService.createHtmlOutput(`<p>La fiche de test a été générée.</p><a href="${fileUrl}" target="_blank">Ouvrir le document</a>`).setWidth(300).setHeight(100);
        ui.showModalDialog(htmlOutput, 'Document Créé');
      } catch (e) {
        ui.alert('Erreur', e.message, ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Numéro de ligne invalide.');
    }
  }
}

function generatePrintableSheet(rowIndex) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG).getSheetByName("Paramètres Généraux");
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
  Logger.log(`Document généré : ${newFile.getName()} (ID: ${newFile.getId()})`);
  return newFile.getUrl();
}

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