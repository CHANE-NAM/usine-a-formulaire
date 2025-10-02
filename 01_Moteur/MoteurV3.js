// =================================================================================
// == PROJET [MOTEUR] - LOGIQUE MÉTIER
// == VERSION : 8.1 - Architecture multi-fichiers stable
// =================================================================================


function lancerDeploiement_V4(rowIndex) {
  // On conserve la ligne de sécurité, par précaution.
  const FormApp = this.FormApp;

  // --- ESPION 1 : Démarrage de la fonction ---
  Logger.log(`--- DÉBUT DU TRAITEMENT pour la ligne ${rowIndex} ---`);

  try {
    const config = getConfigurationFromRow(rowIndex);

    // --- ESPION 2 : Vérification de la configuration lue ---
    Logger.log("CONFIG LUE : " + JSON.stringify(config));

    if (config['Statut'].trim().toLowerCase() !== 'en construction') {
      // --- ESPION 3 : Le script s'arrête à cause du statut ---
      Logger.log(`!!! SORTIE VOLONTAIRE : Le statut '${config['Statut']}' n'est pas 'en construction'.`);
      return null;
    }

    // --- ESPION 4 : Le statut est validé, on continue ---
    Logger.log(">>> Statut validé. Le processus de création continue...");


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

    const systemIds = getSystemIds();

    let dossierCible;
    if (config['ID_Dossier_Cible']) {
      dossierCible = DriveApp.getFolderById(config['ID_Dossier_Cible']);
    } else {
      if (!systemIds.ID_DOSSIER_CIBLE_GEN) throw new Error("ID_DOSSIER_CIBLE_GEN introuvable.");
      dossierCible = DriveApp.getFolderById(systemIds.ID_DOSSIER_CIBLE_GEN);
    }

    // --- ÉTAPE 1: Création du formulaire et partage via DriveApp ---
    const form = FormApp.create(config['Titre_Formulaire_Utilisateur']);
    
    DriveApp.getFileById(form.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    form.setProgressBar(true);
    const sousTitre = config['Sous-Titre_Formulaire'];
    form.setDescription(sousTitre || "");

    const formFile = DriveApp.getFileById(form.getId());
    formFile.moveTo(dossierCible);

    const formUrl = form.getPublishedUrl();
    const editUrl = form.getEditUrl();
    Logger.log("URL publique obtenue : " + formUrl);
    Logger.log("URL d'édition obtenue : " + editUrl);

    // --- Génération du lien chiffré ---
    const lienObfusque = generateEncryptedUrl(typeTest, rowIndex);

    // --- ÉTAPE 2: Génération des questions ---
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
        const metaQuestionsMap = metaData.reduce((acc, row) => {
          acc[row[idCol]] = row;
          return acc;
        }, {});
        metaIds.forEach(id => {
          if (metaQuestionsMap[id]) {
            const [q_id, q_type_old, q_titre, q_options, q_logique, q_description, q_params_json] = metaQuestionsMap[id];
            let final_meta_type = q_type_old;
            if (q_params_json) {
              try {
                const p = JSON.parse(q_params_json);
                if (p.mode) final_meta_type = p.mode;
              } catch (e) {}
            }
            creerItemFormulaire(form, final_meta_type, q_titre, q_options, q_description, q_params_json);
          }
        });
      }
    }
    const languesAInclure = _identifierLangues(bdd, config['Type_Test']);
    _construireQuestionsFormulaire(form, languesAInclure, config['nbQuestions']);

    // --- ÉTAPE 3: Mise à jour de la feuille CONFIG ---
    configSheet.getRange(rowIndex, colIndex['Statut'] + 1).setValue('Actif - Liaison manuelle requise');
    configSheet.getRange(rowIndex, colIndex['Nom_Fichier_Complet'] + 1).setValue(nomFichierComplet);
    if (colIndex['ID_Formulaire_Cible'] !== undefined) configSheet.getRange(rowIndex, colIndex['ID_Formulaire_Cible'] + 1).setValue(formFile.getId());
    if (colIndex['ID_Sheet_Cible'] !== undefined) configSheet.getRange(rowIndex, colIndex['ID_Sheet_Cible'] + 1).setValue('');
    if (colIndex['Lien_Formulaire_Public'] !== undefined) configSheet.getRange(rowIndex, colIndex['Lien_Formulaire_Public'] + 1).setValue(formUrl);
    if (colIndex['Lien_Formulaire_Obfusqué'] !== undefined) configSheet.getRange(rowIndex, colIndex['Lien_Formulaire_Obfusqué'] + 1).setValue(lienObfusque);
    const colNameEditUrl = Object.keys(colIndex).find(k => k.toLowerCase().includes('accès direct formulaire'));
    if (colNameEditUrl) {
      configSheet.getRange(rowIndex, colIndex[colNameEditUrl] + 1).setFormula(`=HYPERLINK("${editUrl}"; "Ouvrir le formulaire")`);
    }
    
    SpreadsheetApp.flush();
    Logger.log(`Ligne ${rowIndex} mise à jour avec le statut 'Actif - Liaison manuelle requise'.`);

    return {
      nomFichier: config['Titre_Formulaire_Utilisateur'],
      editUrl: editUrl
    };

  } catch (e) {
    // --- ESPION 5 : Une erreur inattendue a été capturée ---
    Logger.log(`!!! ERREUR CAPTURÉE : ${e.toString()} | Ligne: ${e.lineNumber} | Stack: ${e.stack}`);
    console.error("ERREUR V4 (ligne " + rowIndex + ") : " + e.toString() + "\n" + e.stack);
    SpreadsheetApp.getUi().alert("Une erreur est survenue lors du déploiement V4 pour la ligne " + rowIndex + ": " + e.message);
    return null;
  }
}


// --- Fonction d'encodage du lien sécurisé ---
function generateEncryptedUrl(typeTest, rowIndex) {
  var payload = {
    typeTest: typeTest,
    rowIndex: rowIndex,
    ts: Date.now()
  };
  var SECRET_KEY = "FELIX QUI POTUIT RERUM COGNOCERE CAUSA VIC";
  var ciphertext = CryptoJS.AES.encrypt(JSON.stringify(payload), SECRET_KEY).toString();
  var formId = "1dePVbG8jnVeNfgArdFHc1-Fz488NwvDPOjc2Wob2yFM";
  var baseFormUrl = "https://docs.google.com/forms/d/" + formId + "/viewform";
  var lienFinal = baseFormUrl + "?code=" + encodeURIComponent(ciphertext);
  return lienFinal;
}
