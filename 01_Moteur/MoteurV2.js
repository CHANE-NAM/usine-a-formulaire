// =================================================================================
// FICHIER : Moteur V2.js
// RÔLE : Fonctions principales de création et d'orchestration des tests.
// VERSION : 6.2 - Ne crée pas la sélection de langue si un seul onglet de question est trouvé.
// =================================================================================

/**
 * Gère le déploiement complet (création + mise à jour du statut + lien public).
 */
function lancerDeploiementComplet(rowIndex) {
  Logger.log(`Lancement du déploiement complet pour la ligne ${rowIndex}...`);
  
  try {
    const config = getConfigurationFromRow(rowIndex);

    if (config['Statut'].toLowerCase() !== 'en construction') {
      Logger.log(`La création pour la ligne ${rowIndex} a été ignorée (statut non valide).`);
      return null;
    }

    const nomFichierComplet = "[" + config['Type_Test'] + "] " + config['Titre_Formulaire_Utilisateur'];
    const systemIds = getSystemIds();
    if (!systemIds.ID_TEMPLATE_TRAITEMENT_V2) throw new Error("ID_TEMPLATE_TRAITEMENT_V2 introuvable.");

    let dossierCible;
    if (config['ID_Dossier_Cible']) {
      dossierCible = DriveApp.getFolderById(config['ID_Dossier_Cible']);
    } else {
      if (!systemIds.ID_DOSSIER_CIBLE_GEN) throw new Error("ID_DOSSIER_CIBLE_GEN introuvable.");
      dossierCible = DriveApp.getFolderById(systemIds.ID_DOSSIER_CIBLE_GEN);
    }

    const templateFile = DriveApp.getFileById(systemIds.ID_TEMPLATE_TRAITEMENT_V2);
    const sheetFile = templateFile.makeCopy(nomFichierComplet, dossierCible);
    const reponsesSheetId = sheetFile.getId();

    const form = FormApp.create(nomFichierComplet);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, reponsesSheetId);
    form.setProgressBar(true);
    
    const sousTitre = config['Sous-Titre_Formulaire']; 
    form.setDescription(sousTitre || ""); 

    const formFile = DriveApp.getFileById(form.getId());
    formFile.moveTo(dossierCible);

    const formUrl = form.getPublishedUrl();
    Logger.log("URL longue obtenue avec succès via getPublishedUrl() : " + formUrl);
    
    // --- Génération des questions ---
    if (!systemIds.ID_BDD) throw new Error("ID_BDD introuvable.");
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    
    const blocsMetaConfig = config['Blocs_Meta_A_Inclure'];
    if (blocsMetaConfig && blocsMetaConfig.trim() !== '') {
      const metaIds = blocsMetaConfig.split(',').map(id => id.trim());
      const metaSheet = bdd.getSheetByName('Questions_META_FR'); 
      if (metaSheet) {
        const metaData = metaSheet.getDataRange().getValues();
        const metaHeaders = metaData.shift();
        const idCol = metaHeaders.indexOf('ID');
        const metaQuestionsMap = metaData.reduce((acc, row) => { acc[row[idCol]] = row; return acc; }, {});
        
        metaIds.forEach(id => {
          if (metaQuestionsMap[id]) {
            const [q_id, q_type_old, q_titre, q_options, q_logique, q_description, q_params_json] = metaQuestionsMap[id];
            let final_meta_type = q_type_old;
            if (q_params_json) { try { const p = JSON.parse(q_params_json); if(p.mode) final_meta_type = p.mode; } catch(e){} }
            creerItemFormulaire(form, final_meta_type, q_titre, q_options, q_description, q_params_json);
          }
        });
      }
    }

    // ==================== DÉBUT DE LA MODIFICATION V6.2 ====================
    // Étape 1 : On identifie toutes les langues disponibles pour ce test
    const toutesLesFeuillesBDD = bdd.getSheets();
    const regexLangues = new RegExp('^Questions_' + config['Type_Test'] + '_([A-Z]{2})$', 'i');
    const languesAInclure = [];
    toutesLesFeuillesBDD.forEach(feuille => {
      const match = feuille.getName().match(regexLangues);
      if (match && match[1]) languesAInclure.push({ code: match[1].toUpperCase(), nomComplet: getLangueFullName(match[1]), feuille: feuille });
    });

    if (languesAInclure.length === 0) {
        throw new Error("Aucune feuille de questions trouvée pour le type '" + config['Type_Test'] + "'.");
    }

    // Étape 2 : On applique une logique différente s'il y a une ou plusieurs langues
    if (languesAInclure.length > 1) {
        // CAS 1 : MULTI-LANGUES (comportement original)
        // On crée la question de sélection de langue et les pages associées.
        Logger.log(`Mode multi-langues détecté (${languesAInclure.length} langues). Création de la sélection de langue.`);
        
        const itemLangue = form.addMultipleChoiceItem().setTitle("Langue / Language").setRequired(true);
        const choices = [];
        languesAInclure.forEach(langue => {
            const page = form.addPageBreakItem().setTitle("Questions (" + langue.nomComplet + ")");
            choices.push(itemLangue.createChoice(langue.nomComplet, page));
            
            const nbQuestionsDisponibles = langue.feuille.getLastRow() - 1;
            let nbQuestionsAUtiliser = (config['nbQuestions'] && config['nbQuestions'] > 0) ? Math.min(config['nbQuestions'], nbQuestionsDisponibles) : nbQuestionsDisponibles;
            if (nbQuestionsAUtiliser <= 0) return;

            const questionsData = langue.feuille.getRange(2, 1, nbQuestionsAUtiliser, 7).getValues();
            questionsData.forEach((q_data, index) => {
                const [id, type_old, titre, options, logique, description, params_json] = q_data;
                let final_type = type_old;
                if (params_json) { try { const p = JSON.parse(params_json); if(p.mode) final_type = p.mode; } catch(e){} }
                creerItemFormulaire(form, final_type, id + ': ' + titre, options, description, params_json);
                if (index === questionsData.length - 1) page.setGoToPage(FormApp.PageNavigationType.SUBMIT);
            });
        });
        itemLangue.setChoices(choices);

    } else {
        // CAS 2 : LANGUE UNIQUE (nouveau comportement)
        // On n'ajoute pas la question de langue, on insère directement les questions du test.
        Logger.log(`Mode langue unique détecté. Insertion directe des questions.`);
        
        const uniqueLangue = languesAInclure[0];
        const nbQuestionsDisponibles = uniqueLangue.feuille.getLastRow() - 1;
        let nbQuestionsAUtiliser = (config['nbQuestions'] && config['nbQuestions'] > 0) ? Math.min(config['nbQuestions'], nbQuestionsDisponibles) : nbQuestionsDisponibles;

        if (nbQuestionsAUtiliser > 0) {
            const questionsData = uniqueLangue.feuille.getRange(2, 1, nbQuestionsAUtiliser, 7).getValues();
            questionsData.forEach((q_data) => {
                const [id, type_old, titre, options, logique, description, params_json] = q_data;
                let final_type = type_old;
                if (params_json) { try { const p = JSON.parse(params_json); if(p.mode) final_type = p.mode; } catch(e){} }
                creerItemFormulaire(form, final_type, id + ': ' + titre, options, description, params_json);
            });
        }
    }
    // ===================== FIN DE LA MODIFICATION V6.2 =====================

    // --- MISE À JOUR DANS LA FEUILLE CONFIG ---
    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("Paramètres Généraux");
    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    headers.forEach((header, i) => { if (header) colIndex[header] = i; });

    const STATUT_COL = colIndex['Statut'];
    const ID_UNIQUE_COL = colIndex['Id_Unique'];
    const NOM_FICHIER_COL = colIndex['Nom_Fichier_Complet'];
    const ID_FORM_COL = colIndex['ID_Formulaire_Cible'];
    const ID_SHEET_COL = colIndex['ID_Sheet_Cible'];
    const LIEN_FORM_COL = colIndex['Lien_Formulaire_Public'];

    const idUnique = sheetFile.getId().slice(0, 8) + '-' + formFile.getId().slice(0, 8);
    
    configSheet.getRange(rowIndex, STATUT_COL + 1).setValue('Actif - Déclencheur à activer'); 
    configSheet.getRange(rowIndex, ID_UNIQUE_COL + 1).setValue(idUnique);
    configSheet.getRange(rowIndex, NOM_FICHIER_COL + 1).setValue(nomFichierComplet);
    if (ID_FORM_COL !== undefined) configSheet.getRange(rowIndex, ID_FORM_COL + 1).setValue(formFile.getId());
    if (ID_SHEET_COL !== undefined) configSheet.getRange(rowIndex, ID_SHEET_COL + 1).setValue(sheetFile.getId());
    if (LIEN_FORM_COL !== undefined) configSheet.getRange(rowIndex, LIEN_FORM_COL + 1).setValue(formUrl);
    
    SpreadsheetApp.flush();
    Logger.log(`Ligne ${rowIndex} mise à jour avec le statut 'Actif - Déclencheur à activer'.`);
    
    return { nomFichier: nomFichierComplet, urlSheet: sheetFile.getUrl(), urlForm: formUrl };

  } catch(e) {
    console.error("ERREUR (ligne " + rowIndex + ") : " + e.toString() + "\n" + e.stack);
    SpreadsheetApp.getUi().alert("Une erreur est survenue lors du déploiement pour la ligne " + rowIndex + ": " + e.message);
    return null;
  }
}
