// =================================================================================
// == PROJET [MOTEUR] - LOGIQUE MÉTIER
// == VERSION : 8.0 - Architecture multi-fichiers stable
// == RÔLE    : Contient la logique principale de création et de déploiement
// ==           des formulaires et des kits de traitement.
// =================================================================================

/**
 * Gère le déploiement complet (création + mise à jour du statut + liens).
 * Appelé par `orchestrateurDeploiementComplet_UI` depuis le fichier codeV3.gs.
 */
function lancerDeploiementComplet(rowIndex) {
  Logger.log(`Lancement du déploiement complet pour la ligne ${rowIndex}...`);
  
  try {
    const config = getConfigurationFromRow(rowIndex);

    if (config['Statut'].toLowerCase() !== 'en construction') {
      Logger.log(`La création pour la ligne ${rowIndex} a été ignorée (statut non valide).`);
      return null;
    }
    
    // --- Logique de nommage automatique ---
    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("Paramètres Généraux");
    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    headers.forEach((header, i) => { if (header) colIndex[header.trim()] = i; });

    const typeTest = config['Type_Test'] || 'TypeInconnu';
    const nbQuestions = config['nbQuestions'] || 'Nq';
    const statutDeploiement = config['Statut_Deploiement'] || 'alpha';
    const today = new Date();
    const dateStr = today.getFullYear() + 
                    ('0' + (today.getMonth() + 1)).slice(-2) + 
                    ('0' + today.getDate()).slice(-2);
    const nomFichierComplet = `${dateStr}_${typeTest}_${statutDeploiement}_${nbQuestions}q`;
    Logger.log(`Nom de fichier technique généré : ${nomFichierComplet}`);
    // --- Fin de la logique de nommage ---

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
    
    const form = FormApp.create(config['Titre_Formulaire_Utilisateur']);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, reponsesSheetId);
    form.setProgressBar(true);
    
    const sousTitre = config['Sous-Titre_Formulaire']; 
    form.setDescription(sousTitre || ""); 

    const formFile = DriveApp.getFileById(form.getId());
    formFile.moveTo(dossierCible);

    const formUrl = form.getPublishedUrl();
    const editUrl = form.getEditUrl();
    Logger.log("URL publique obtenue : " + formUrl);
    Logger.log("URL d'édition obtenue : " + editUrl);
    
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
    
    const languesAInclure = _identifierLangues(bdd, config['Type_Test']);
    _construireQuestionsFormulaire(form, languesAInclure, config['nbQuestions']);

    // --- Mise à jour de la feuille CONFIG ---
    const idUnique = sheetFile.getId().slice(0, 8) + '-' + formFile.getId().slice(0, 8);
    
    configSheet.getRange(rowIndex, colIndex['Statut'] + 1).setValue('Actif - Déclencheur à activer');
    configSheet.getRange(rowIndex, colIndex['Id_Unique'] + 1).setValue(idUnique);
    configSheet.getRange(rowIndex, colIndex['Nom_Fichier_Complet'] + 1).setValue(nomFichierComplet);
    if (colIndex['ID_Formulaire_Cible'] !== undefined) configSheet.getRange(rowIndex, colIndex['ID_Formulaire_Cible'] + 1).setValue(formFile.getId());
    if (colIndex['ID_Sheet_Cible'] !== undefined) configSheet.getRange(rowIndex, colIndex['ID_Sheet_Cible'] + 1).setValue(sheetFile.getId());
    if (colIndex['Lien_Formulaire_Public'] !== undefined) configSheet.getRange(rowIndex, colIndex['Lien_Formulaire_Public'] + 1).setValue(formUrl);
    
    const colNameEditUrl = Object.keys(colIndex).find(k => k.toLowerCase().includes('accès direct formulaire'));
    if (colNameEditUrl) {
      configSheet.getRange(rowIndex, colIndex[colNameEditUrl] + 1).setFormula(`=HYPERLINK("${editUrl}"; "Ouvrir le formulaire")`);
    }
    
    SpreadsheetApp.flush();
    Logger.log(`Ligne ${rowIndex} mise à jour avec le statut 'Actif - Déclencheur à activer'.`);
    return { nomFichier: nomFichierComplet, urlSheet: sheetFile.getUrl(), urlForm: formUrl };

  } catch(e) {
    console.error("ERREUR (ligne " + rowIndex + ") : " + e.toString() + "\n" + e.stack);
    SpreadsheetApp.getUi().alert("Une erreur est survenue lors du déploiement pour la ligne " + rowIndex + ": " + e.message);
    return null;
  }
}

