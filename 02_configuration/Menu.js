// =================================================================================
// == FICHIER : Menu.js
// == VERSION : 4.7 - Finalisation de la sauvegarde de tous les champs.
// == RÔLE  : Logique côté serveur pour l'application web de configuration.
// =================================================================================

const ID_FEUILLE_CONFIG = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";
const ID_MODELE_FICHE_TEST = "1W_amKwp5kyyGWmg5LTaIQe5K8Gzxf_qvcjGskRy1Sq8";

/**
 * Crée le menu de l'application à l'ouverture de la feuille de calcul.
 */
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

/**
 * Affiche la barre latérale de configuration (FormulaireUI.html).
 */
function showConfigurationSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('FormulaireUI').setTitle('Configuration Usine à Tests').setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Récupère les données initiales pour peupler les listes déroulantes de la barre latérale.
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
      console.error("Impossible de charger les blocs méta : " + e.message);
    }
    optionsMap.availableMetaBlocks = availableMetaBlocks;
    return optionsMap;
  } catch (err) {
    Logger.log("ERREUR FATALE dans getInitialData: " + err.stack);
    throw new Error("Erreur côté serveur : " + err.message);
  }
}

/**
 * Récupère le nombre de questions disponibles pour un type de test donné depuis la BDD.
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
 * Traite les données soumises depuis la barre latérale et crée une nouvelle ligne de configuration.
 * @param {Object} formObject L'objet contenant les données du formulaire.
 * @returns {string} Un message de succès.
 */
function processNewTestConfiguration(formObject) {
  try {
    const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
    const paramsSheet = ss.getSheetByName("Paramètres Généraux");
    if (!paramsSheet) throw new Error("L'onglet 'Paramètres Généraux' est introuvable.");

    const headers = paramsSheet.getRange(1, 1, 1, paramsSheet.getLastColumn()).getValues()[0];
    
    const dataRow = {};
    
    // Cartographie complète des données du formulaire vers les colonnes de la feuille
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

    // Moteur & Accès
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
    return "Configuration enregistrée avec succès !";
  } catch (e) {
    Logger.log("ERREUR processNewTestConfiguration: " + e.stack);
    throw new Error("Erreur interne lors de la sauvegarde : " + e.message);
  }
}

/**
 * Affiche une UI pour demander la ligne à modifier.
 */
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

/**
 * Affiche la barre latérale d'édition pour une ligne donnée.
 * @param {number} rowIndex Le numéro de la ligne à éditer.
 */
function showEditSidebar(rowIndex) {
  const template = HtmlService.createTemplateFromFile('ModifierTestUI');
  template.rowIndex = rowIndex;
  const html = template.evaluate().setTitle('Édition du Test (Ligne ' + rowIndex + ')').setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Récupère les données d'une ligne de configuration pour l'édition.
 * @param {number} rowIndex Le numéro de la ligne.
 * @returns {Object} Un objet avec les en-têtes et les valeurs de la ligne.
 */
function getTestDataForEdit(rowIndex) {
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG).getSheetByName("Paramètres Généraux");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { headers: headers, values: values };
}

/**
 * Met à jour les données d'une ligne de configuration.
 * @param {number} rowIndex Le numéro de la ligne à mettre à jour.
 * @param {Object} updatedData Un objet contenant les nouvelles données.
 * @returns {string} Un message de succès.
 */
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

/**
 * Affiche une UI pour demander la ligne à dupliquer.
 */
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

/**
 * Duplique une ligne de configuration.
 * @param {number} rowIndex Le numéro de la ligne à dupliquer.
 * @returns {number} Le numéro de la nouvelle ligne créée.
 */
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

/**
 * Affiche une UI pour générer une fiche de test imprimable.
 */
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

/**
 * Génère une fiche de test à partir d'un modèle Google Doc.
 * @param {number} rowIndex Le numéro de la ligne de configuration à utiliser.
 * @returns {string} L'URL du document généré.
 */
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

/**
 * Récupère les IDs des fichiers système (BDD, etc.) depuis l'onglet sys_ID_Fichiers.
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