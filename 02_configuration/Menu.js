// =================================================================================
// == FICHIER : menu.gs
// == VERSION : 3.6 (Valeur par défaut "Universel" pour Moteur_Calcul)
// == RÔLE  : Logique côté serveur pour l'application web de configuration.
// =================================================================================

const ID_FEUILLE_CONFIG = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";

// --- SECTION 1 : INTERFACE UTILISATEUR ---

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Actions Usine')
    .addItem('Configurer un nouveau test...', 'showConfigurationSidebar')
    .addToUi();
}

function showConfigurationSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('FormulaireUI')
      .setTitle('Configuration Usine à Tests')
      .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}


// --- SECTION 2 : FONCTIONS APPELÉES PAR L'INTERFACE HTML ---

function getInitialData() {
  const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
  const optionsSheet = ss.getSheetByName("sys_Options_Parametres");
  if (!optionsSheet) {
    throw new Error("L'onglet 'sys_Options_Parametres' est introuvable.");
  }

  const optionsData = optionsSheet.getDataRange().getValues();
  const headers = optionsData.shift();
  const optionsMap = {};
  headers.forEach((header, i) => {
    const options = optionsData.map(row => row[i]).filter(String);
    optionsMap[header] = options;
  });

  // Charger la liste des blocs méta disponibles depuis la BDD
  let availableMetaBlocks = [];
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const metaSheet = bdd.getSheetByName('Questions_META_FR');
    if (metaSheet) {
        const metaData = metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, 3).getValues(); // ID, Type, Titre
        availableMetaBlocks = metaData.map(row => ({ id: row[0], title: row[2] })).filter(block => block.id && block.title);
    }
  } catch(e) {
    console.error("Impossible de charger les blocs méta depuis la BDD : " + e.message);
  }

  return {
    typesDeTest: optionsMap['Type_Test'] || [],
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


// --- SECTION 3 : TRAITEMENT DE LA SOUMISSION ---

function processNewTestConfiguration(formObject) {
  try {
    const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
    const paramsSheet = ss.getSheetByName("Paramètres Généraux");
    if (!paramsSheet) { throw new Error("L'onglet 'Paramètres Généraux' est introuvable."); }
    
    let headers = paramsSheet.getRange(1, 1, 1, paramsSheet.getLastColumn()).getValues()[0];
    
    // ==================== DÉBUT MODIFICATION ====================
    // On s'assure que les colonnes requises, y compris Moteur_Calcul, existent.
    const requiredHeaders = ['Blocs_Meta_A_Inclure', 'ID_Gabarit_Email_Repondant', 'Email_Alias', 'Moteur_Calcul'];
    // ===================== FIN MODIFICATION =====================

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

    let idGabaritRepondant = ''; // Valeur par défaut
    if (formObject.repondantContenu && formObject.repondantContenu.includes('Niveau1')) {
        idGabaritRepondant = 'RESULTATS_N1';
    } else if (formObject.repondantContenu && formObject.repondantContenu.includes('Niveau2')) {
        idGabaritRepondant = 'RESULTATS_N2';
    } else if (formObject.repondantContenu && formObject.repondantContenu.includes('Niveau3')) {
        idGabaritRepondant = 'RESULTATS_N3';
    }

    const dataRow = {
      'Id_Unique': '',
      'Titre_Formulaire_Utilisateur': formObject.titre,
      'Nom_Fichier_Complet': '',
      'Statut': 'En construction',
      'Type_Test': formObject.type,
      // ==================== DÉBUT MODIFICATION ====================
      'Moteur_Calcul': 'Universel', // On force le moteur Universel par défaut
      // ===================== FIN MODIFICATION =====================
      'Blocs_Meta_A_Inclure': blocsMetaString,
      'ID_Gabarit_Email_Repondant': idGabaritRepondant,
      'ID_Dossier_Cible': '',
      'Limite_Lignes_A_Traiter': limiteLignes,
      'nbQuestions': formObject.nbQuestions,
      'Repondant_Email_Actif': formObject.repondantActif ? "Oui" : "Non",
      'Repondant_Quand': formObject.repondantQuand,
      'Repondant_Contenu': formObject.repondantContenu,
      'Patron_Email_Mode': formObject.patronActif ? "Oui" : "Non",
      'Patron_Quand': formObject.patronQuand,
      'Patron_Contenu': formObject.patronContenu,
      'Patron_Email': formObject.patronEmail,
      'Formateur_Email_Actif': formObject.formateurActif ? "Oui" : "Non",
      'Formateur_Quand': formObject.formateurQuand,
      'Formateur_Contenu': formObject.formateurContenu,
      'Formateur_Email': formObject.formateurEmail,
      'Developpeur_Email': emailDev,
      'ID_Formulaire_Cible': '',
      'ID_Sheet_Cible': '',
      'Email_Alias': formObject.emailAlias
    };
    
    const nouvelleLigne = headers.map(header => dataRow[header] !== undefined ? dataRow[header] : '');
    paramsSheet.appendRow(nouvelleLigne);
    return "Configuration enregistrée avec succès !";
  } catch (e) {
    Logger.log("ERREUR lors de la sauvegarde de la configuration: " + e.toString());
    throw new Error("Une erreur interne est survenue lors de la sauvegarde. " + e.message);
  }
}


// --- SECTION 4 : FONCTIONS UTILITAIRES ---

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