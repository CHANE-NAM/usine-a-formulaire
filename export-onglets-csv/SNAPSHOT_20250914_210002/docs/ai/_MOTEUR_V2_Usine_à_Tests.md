# _MOTEUR_V2_Usine_à_Tests

> Généré automatiquement depuis **scripts__MOTEUR_V2_Usine_à_Tests.txt** — snapshot: **SNAPSHOT_20250914_210002**.

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\01_Moteur\appsscript.json

```json

{
  "timeZone": "Indian/Mauritius",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Drive",
        "version": "v3",
        "serviceId": "drive"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/forms",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.container.ui"
  ]
}
```

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\01_Moteur\MigrationV1.js

```javascript

// =================================================================================
// FONCTION DE MIGRATION V1 -> V2 (JSON)
// RÃ”LE : Convertit les questions d'un ancien format (Options/Logique)
//         vers le nouveau format V2 (ParamÃ¨tres (JSON)).
// VERSION : 1.4 - Version finale et corrigÃ©e
// =================================================================================

/**
 * Fonction principale appelÃ©e depuis le menu de l'interface utilisateur.
 */
function lancerMigrationV1versV2() {
  try {
    const ID_BDD = '1m2MGBd0nyiAl3qw032B6Nfj7zQL27bRSBexiOPaRZd8';

    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Outil de Migration V1 -> V2',
      'Veuillez entrer le nom exact de l\'onglet dans la BDD Ã  migrer :',
      ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() == ui.Button.OK && response.getResponseText() != '') {
      const sheetName = response.getResponseText().trim();
      
      const bdd = SpreadsheetApp.openById(ID_BDD);
      if (!bdd) { throw new Error(`Impossible d'ouvrir la BDD avec l'ID fourni.`); }
      const sheet = bdd.getSheetByName(sheetName);

      if (!sheet) { throw new Error(`L'onglet "${sheetName}" est introuvable dans la BDD.`); }

      const resultat = convertirQuestionsEnJSON(sheet);
      
      ui.alert(
        'Migration TerminÃ©e',
        `Rapport pour l'onglet "${sheetName}":\n\n` +
        `- Lignes traitÃ©es : ${resultat.lignesTraitees}\n` +
        `- Questions converties : ${resultat.questionsConverties}\n` +
        `- Lignes ignorÃ©es : ${resultat.lignesIgnorees}\n` +
        `- Erreurs rencontrÃ©es : ${resultat.erreurs.length}` +
        (resultat.erreurs.length > 0 ? `\n\nConsultez les logs ("Affichage > Journaux") pour le dÃ©tail des erreurs.` : ''),
        ui.ButtonSet.OK);
    }
  } catch (e) {
    SpreadApp.getUi().alert(`Une erreur est survenue : ${e.message}`);
    console.error(`Erreur lors du lancement de la migration : ${e.stack}`);
  }
}

/**
 * CÅ“ur de la logique de conversion. Lit une feuille et met Ã  jour les lignes.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La feuille de calcul Ã  traiter.
 * @returns {object} Un objet contenant les statistiques de la migration.
 */
function convertirQuestionsEnJSON(sheet) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values.shift(); 

  const colIndex = {
    type: headers.indexOf('TypeQuestion'),
    options: headers.indexOf('Options'),
    logique: headers.indexOf('Logique'),
    description: headers.indexOf('Description'),
    json: headers.indexOf('ParamÃ¨tres (JSON)')
  };

  if (colIndex.type === -1 || colIndex.options === -1 || colIndex.logique === -1 || colIndex.json === -1) {
    throw new Error("Colonnes requises ('TypeQuestion', 'Options', 'Logique', 'ParamÃ¨tres (JSON)') manquantes.");
  }
  
  let questionsConverties = 0;
  let lignesIgnorees = 0;
  const erreurs = [];

  values.forEach((row, index) => {
    const jsonCell = row[colIndex.json];
    if (jsonCell) {
      lignesIgnorees++;
      return;
    }

    const typeQuestion = row[colIndex.type];
    const optionsStr = row[colIndex.options];
    const logiqueStr = row[colIndex.logique];
    const descriptionStr = colIndex.description !== -1 ? row[colIndex.description] : "";
    let jsonPayload = null;

    try {
      switch (typeQuestion) {
        case 'CHOIX_BINAIRE':
          if (optionsStr && logiqueStr) {
            const optionsArray = optionsStr.toString().split(';').map(s => s.trim());
            const logiqueArray = logiqueStr.toString().split(';').map(s => s.trim());
            if (optionsArray.length !== logiqueArray.length) {
              throw new Error(`CHOIX_BINAIRE: Le nombre d'options (${optionsArray.length}) et de logiques (${logiqueArray.length}) ne correspond pas.`);
            }
            jsonPayload = {
              mode: 'QRM_CAT',
              options: optionsArray.map((libelle, i) => ({ libelle: libelle, profil: logiqueArray[i], valeur: 1 }))
            };
          }
          break;

        case 'ECHELLE':
          if (optionsStr && logiqueStr) { // La description n'est pas bloquante
            const echelle = optionsStr.toString().split(';').map(s => parseInt(s.trim(), 10));
            const labels = descriptionStr ? descriptionStr.toString().split(';').map(s => s.trim()) : ["", ""];
            
            jsonPayload = {
              mode: 'ECHELLE_NOTE',
              profil: logiqueStr.toString().trim(),
              echelle_min: Math.min(...echelle),
              echelle_max: Math.max(...echelle),
              label_min: labels[0] || "",
              label_max: labels[1] || ""
            };
          }
          break;

        default:
          lignesIgnorees++;
          break;
      }

      if (jsonPayload) {
        sheet.getRange(index + 2, colIndex.json + 1).setValue(JSON.stringify(jsonPayload));
        questionsConverties++;
      } else {
        lignesIgnorees++;
      }

    } catch (e) {
      const errorMessage = `Erreur Ã  la ligne ${index + 2}: ${e.message}`;
      console.error(errorMessage);
      erreurs.push(errorMessage);
    }
  });

  return {
    lignesTraitees: values.length,
    questionsConverties: questionsConverties,
    lignesIgnorees: lignesIgnorees,
    erreurs: erreurs
  };
}
```

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\01_Moteur\Diagnostic.js

```javascript

/**
 * Ce script est un outil de diagnostic Ã  usage unique.
 * Il va crÃ©er un formulaire et inspecter l'objet retournÃ© pour
 * comprendre pourquoi la fonction .getShortUrl() n'est pas trouvÃ©e.
 */
function testCreationFormulaire() {
  try {
    Logger.log("--- DÃ©but du test de diagnostic de crÃ©ation de formulaire ---");
    
    // Ã‰tape 1 : On crÃ©e un formulaire de test.
    const form = FormApp.create("Test de Diagnostic Ultime");
    Logger.log("Objet 'form' crÃ©Ã©.");

    // Ã‰tape 2 : On vÃ©rifie si la fonction qui pose problÃ¨me existe VRAIMENT sur cet objet.
    if (form && typeof form.getShortUrl === 'function') {
      Logger.log("--> RÃ‰SULTAT POSITIF : La fonction .getShortUrl() a Ã©tÃ© trouvÃ©e !");
      Logger.log("    Lien court obtenu : " + form.getShortUrl());
    } else {
      Logger.log("--> RÃ‰SULTAT NÃ‰GATIF : La fonction .getShortUrl() est INTROUVABLE sur l'objet 'form'.");
    }
    
    // Ã‰tape 3 : On liste toutes les propriÃ©tÃ©s et mÃ©thodes que l'on trouve sur l'objet.
    // Cela nous dira ce qu'est rÃ©ellement l'objet 'form'.
    let properties = [];
    for (var name in form) {
      properties.push(name);
    }
    Logger.log("Liste de toutes les propriÃ©tÃ©s trouvÃ©es sur l'objet : " + properties.join(', '));

    // On supprime le formulaire de test pour ne pas polluer votre Drive.
    DriveApp.getFileById(form.getId()).setTrashed(true);
    Logger.log("Formulaire de test supprimÃ©.");

  } catch (e) {
    Logger.log("ERREUR CATASTROPHIQUE lors du test de diagnostic : " + e.toString());
    Logger.log(e.stack);
  }
  Logger.log("--- Fin du test de diagnostic ---");
}
```

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\01_Moteur\EmailCompositionUtils.js

```javascript

function normalizeAndDedupeCompositionEmails_(rows) {
  const seen = new Set();
  return rows
    .map(r => {
      const out = Object.assign({}, r);
      out.Element = (out.Element || '').toString().trim();
      return out;
    })
    .filter(r => {
      const key = [
        r.Type_Test || '',
        r.Code_Langue || '',
        r.Code_Niveau_Email || '',
        r.Code_Profil || '',
        r.Element || '',
        r.Ordre || ''
      ].join('|');
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}
```

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\01_Moteur\CodeV3.js

```javascript

// =================================================================================
// == PROJET [MOTEUR] - FICHIER PRINCIPAL (POINTS D'ENTRÃ‰E)
// == VERSION : 8.0 - Architecture multi-fichiers stable
// == RÃ”LE    : GÃ¨re l'interface utilisateur (menus) et orchestre les appels
// ==           vers les scripts de logique mÃ©tier.
// =================================================================================

/**
Â * CrÃ©e le menu personnalisÃ© dans l'interface utilisateur de Google Sheets Ã  l'ouverture.
Â * C'est le SEUL onOpen() du projet.
Â */
function onOpen() {
Â  SpreadsheetApp.getUi()
Â  Â  .createMenu('ðŸ­ Usine Ã  Tests')
Â  Â  .addItem("ðŸš€ DÃ©ployer un test de A Ã  Z...", "orchestrateurDeploiementComplet_UI")
Â  Â  .addToUi();
}

/**
Â * Orchestre le dÃ©ploiement complet d'un test depuis l'UI.
Â * Appelle la fonction de logique mÃ©tier `lancerDeploiementComplet` du fichier Moteur.gs.
Â */
function orchestrateurDeploiementComplet_UI() {
Â  const ui = SpreadsheetApp.getUi();
Â  
Â  const response = ui.prompt(
Â  Â  'ðŸš€ DÃ©ploiement de A Ã  Z',
Â  Â  'Entrez le numÃ©ro de la ligne Ã  dÃ©ployer entiÃ¨rement :',
Â  Â  ui.ButtonSet.OK_CANCEL
Â  );

Â  if (response.getSelectedButton() !== ui.Button.OK || response.getResponseText() === '') {
Â  Â  return; // Annulation par l'utilisateur
Â  }

Â  const rowIndex = parseInt(response.getResponseText(), 10);
Â  if (isNaN(rowIndex) || rowIndex <= 1) {
Â  Â  ui.alert('NumÃ©ro de ligne invalide. Veuillez entrer un nombre supÃ©rieur Ã  1.');
Â  Â  return;
Â  }
Â  
Â  ui.alert('Lancement du dÃ©ploiement complet... Cette opÃ©ration peut prendre un moment.');

Â  try {
    // Appel Ã  la fonction de logique mÃ©tier
Â  Â  const resultats = lancerDeploiementComplet(rowIndex);

Â  Â  if (resultats && resultats.urlSheet && resultats.urlForm) {
Â  Â  Â  const htmlOutput = HtmlService.createHtmlOutput(
Â  Â  Â  Â  `<h4>âœ… DÃ©ploiement RÃ©ussi !</h4>` +
Â  Â  Â  Â  `<p>Le kit "<b>${resultats.nomFichier}</b>" a Ã©tÃ© gÃ©nÃ©rÃ©.</p><hr>` +
Â  Â  Â  Â  `<p><b>1. Voici le lien public du formulaire Ã  partager :</b></p>` +
Â  Â  Â  Â  `<p style="margin-top:10px;"><a href="${resultats.urlForm}" target="_blank" style="background-color:#34A853; color:white; padding:8px 12px; text-decoration:none; border-radius:4px;">Copier ou ouvrir le lien du Formulaire</a></p><br>` +
Â  Â  Â  Â  `<p><b>2. ACTION FINALE REQUISE (pour que le test fonctionne) :</b></p>` +
Â  Â  Â  Â  `<p>Cliquez sur le lien ci-dessous, puis dans le menu :<br>` +
Â  Â  Â  Â  `<b>&nbsp;&nbsp;&nbsp;âš™ï¸ Actions du Kit -> Activer le traitement des rÃ©ponses</b>.</p>` +
Â  Â  Â  Â  `<p style="margin-top:10px;"><a href="${resultats.urlSheet}" target="_blank" style="background-color:#4285F4; color:white; padding:8px 12px; text-decoration:none; border-radius:4px;">Ouvrir le Kit pour l'activer</a></p>`
Â  Â  Â  )
Â  Â  Â  .setWidth(500)
Â  Â  Â  .setHeight(520);
Â  Â  Â  ui.showModalDialog(htmlOutput, "DÃ©ploiement TerminÃ©");

Â  Â  } else {
Â  Â  Â  ui.alert(`â„¹ï¸ Le dÃ©ploiement pour la ligne ${rowIndex} a Ã©tÃ© ignorÃ© (le statut n'Ã©tait probablement pas 'En construction').`);
Â  Â  }

Â  } catch (e) {
Â  Â  Logger.log(`ERREUR Critique lors du dÃ©ploiement complet (ligne ${rowIndex}) : ${e.toString()}`);
Â  Â  ui.alert(`âŒ ERREUR : Le dÃ©ploiement a Ã©chouÃ© pour la ligne ${rowIndex}. Consultez les logs pour les dÃ©tails. Message : ${e.message}`);
Â  }
}
```

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\01_Moteur\MoteurV3.js

```javascript

// =================================================================================
// == PROJET [MOTEUR] - LOGIQUE MÃ‰TIER
// == VERSION : 8.0 - Architecture multi-fichiers stable
// == RÃ”LE    : Contient la logique principale de crÃ©ation et de dÃ©ploiement
// ==           des formulaires et des kits de traitement.
// =================================================================================

/**
Â * GÃ¨re le dÃ©ploiement complet (crÃ©ation + mise Ã  jour du statut + liens).
 * AppelÃ© par `orchestrateurDeploiementComplet_UI` depuis le fichier codeV3.gs.
Â */
function lancerDeploiementComplet(rowIndex) {
Â  Logger.log(`Lancement du dÃ©ploiement complet pour la ligne ${rowIndex}...`);
Â  
Â  try {
Â  Â  const config = getConfigurationFromRow(rowIndex);

Â  Â  if (config['Statut'].toLowerCase() !== 'en construction') {
Â  Â  Â  Logger.log(`La crÃ©ation pour la ligne ${rowIndex} a Ã©tÃ© ignorÃ©e (statut non valide).`);
Â  Â  Â  return null;
Â  Â  }
    
    // --- Logique de nommage automatique ---
    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("ParamÃ¨tres GÃ©nÃ©raux");
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
    Logger.log(`Nom de fichier technique gÃ©nÃ©rÃ© : ${nomFichierComplet}`);
    // --- Fin de la logique de nommage ---

Â  Â  const systemIds = getSystemIds();
Â  Â  if (!systemIds.ID_TEMPLATE_TRAITEMENT_V2) throw new Error("ID_TEMPLATE_TRAITEMENT_V2 introuvable.");

Â  Â  let dossierCible;
Â  Â  if (config['ID_Dossier_Cible']) {
Â  Â  Â  dossierCible = DriveApp.getFolderById(config['ID_Dossier_Cible']);
Â  Â  } else {
Â  Â  Â  if (!systemIds.ID_DOSSIER_CIBLE_GEN) throw new Error("ID_DOSSIER_CIBLE_GEN introuvable.");
Â  Â  Â  dossierCible = DriveApp.getFolderById(systemIds.ID_DOSSIER_CIBLE_GEN);
Â  Â  }

Â  Â  const templateFile = DriveApp.getFileById(systemIds.ID_TEMPLATE_TRAITEMENT_V2);
Â  Â  const sheetFile = templateFile.makeCopy(nomFichierComplet, dossierCible);
Â  Â  const reponsesSheetId = sheetFile.getId();
Â  Â  
Â  Â  const form = FormApp.create(config['Titre_Formulaire_Utilisateur']);
Â  Â  form.setDestination(FormApp.DestinationType.SPREADSHEET, reponsesSheetId);
Â  Â  form.setProgressBar(true);
Â  Â  
Â  Â  const sousTitre = config['Sous-Titre_Formulaire']; 
Â  Â  form.setDescription(sousTitre || ""); 

Â  Â  const formFile = DriveApp.getFileById(form.getId());
Â  Â  formFile.moveTo(dossierCible);

Â  Â  const formUrl = form.getPublishedUrl();
Â  Â  const editUrl = form.getEditUrl();
Â  Â  Logger.log("URL publique obtenue : " + formUrl);
Â  Â  Logger.log("URL d'Ã©dition obtenue : " + editUrl);
Â  Â  
Â  Â  // --- GÃ©nÃ©ration des questions ---
Â  Â  if (!systemIds.ID_BDD) throw new Error("ID_BDD introuvable.");
Â  Â  const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
Â  Â  
Â  Â  const blocsMetaConfig = config['Blocs_Meta_A_Inclure'];
Â  Â  if (blocsMetaConfig && blocsMetaConfig.trim() !== '') {
Â  Â  Â  const metaIds = blocsMetaConfig.split(',').map(id => id.trim());
Â  Â  Â  const metaSheet = bdd.getSheetByName('Questions_META_FR'); 
Â  Â  Â  if (metaSheet) {
Â  Â  Â  Â  const metaData = metaSheet.getDataRange().getValues();
Â  Â  Â  Â  const metaHeaders = metaData.shift();
Â  Â  Â  Â  const idCol = metaHeaders.indexOf('ID');
Â  Â  Â  Â  const metaQuestionsMap = metaData.reduce((acc, row) => { acc[row[idCol]] = row; return acc; }, {});
Â  Â  Â  Â  
Â  Â  Â  Â  metaIds.forEach(id => {
Â  Â  Â  Â  Â  if (metaQuestionsMap[id]) {
Â  Â  Â  Â  Â  Â  const [q_id, q_type_old, q_titre, q_options, q_logique, q_description, q_params_json] = metaQuestionsMap[id];
Â  Â  Â  Â  Â  Â  let final_meta_type = q_type_old;
Â  Â  Â  Â  Â  Â  if (q_params_json) { try { const p = JSON.parse(q_params_json); if(p.mode) final_meta_type = p.mode; } catch(e){} }
Â  Â  Â  Â  Â  Â  creerItemFormulaire(form, final_meta_type, q_titre, q_options, q_description, q_params_json);
Â  Â  Â  Â  Â  }
Â  Â  Â  Â  });
Â  Â  Â  }
Â  Â  }
Â  Â  
    const languesAInclure = _identifierLangues(bdd, config['Type_Test']);
    _construireQuestionsFormulaire(form, languesAInclure, config['nbQuestions']);

Â  Â  // --- Mise Ã  jour de la feuille CONFIG ---
Â  Â  const idUnique = sheetFile.getId().slice(0, 8) + '-' + formFile.getId().slice(0, 8);
Â  Â  
Â  Â  configSheet.getRange(rowIndex, colIndex['Statut'] + 1).setValue('Actif - DÃ©clencheur Ã  activer');
Â  Â  configSheet.getRange(rowIndex, colIndex['Id_Unique'] + 1).setValue(idUnique);
Â  Â  configSheet.getRange(rowIndex, colIndex['Nom_Fichier_Complet'] + 1).setValue(nomFichierComplet);
Â  Â  if (colIndex['ID_Formulaire_Cible'] !== undefined) configSheet.getRange(rowIndex, colIndex['ID_Formulaire_Cible'] + 1).setValue(formFile.getId());
Â  Â  if (colIndex['ID_Sheet_Cible'] !== undefined) configSheet.getRange(rowIndex, colIndex['ID_Sheet_Cible'] + 1).setValue(sheetFile.getId());
Â  Â  if (colIndex['Lien_Formulaire_Public'] !== undefined) configSheet.getRange(rowIndex, colIndex['Lien_Formulaire_Public'] + 1).setValue(formUrl);
Â  Â  
Â  Â  const colNameEditUrl = Object.keys(colIndex).find(k => k.toLowerCase().includes('accÃ¨s direct formulaire'));
Â  Â  if (colNameEditUrl) {
Â  Â  Â  configSheet.getRange(rowIndex, colIndex[colNameEditUrl] + 1).setFormula(`=HYPERLINK("${editUrl}"; "Ouvrir le formulaire")`);
Â  Â  }
Â  Â  
Â  Â  SpreadsheetApp.flush();
Â  Â  Logger.log(`Ligne ${rowIndex} mise Ã  jour avec le statut 'Actif - DÃ©clencheur Ã  activer'.`);
Â  Â  return { nomFichier: nomFichierComplet, urlSheet: sheetFile.getUrl(), urlForm: formUrl };

Â  } catch(e) {
Â  Â  console.error("ERREUR (ligne " + rowIndex + ") : " + e.toString() + "\n" + e.stack);
Â  Â  SpreadsheetApp.getUi().alert("Une erreur est survenue lors du dÃ©ploiement pour la ligne " + rowIndex + ": " + e.message);
Â  Â  return null;
Â  }
}
```

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\01_Moteur\UtilsV3.js

```javascript

// =================================================================================
// == FICHIER : UtilsV3.gs
// == PROJET [MOTEUR] - FONCTIONS UTILITAIRES
// == VERSION : 8.1 - Correction du bug setGoToPage en mode multi-langues.
// == RÃ”LE    : Contient toutes les fonctions de support, appelÃ©es par les
// ==           autres scripts du projet.
// =================================-===============================================

// âš™ï¸ ID de la feuille de configuration centrale (CONFIG)
const ID_FEUILLE_CONFIGURATION = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";

// ------------------------------------
// IDs systÃ¨me (CONFIG â†’ onglet sys_ID_Fichiers)
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
  const sheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("ParamÃ¨tres GÃ©nÃ©raux");
  if (!sheet) throw new Error("L'onglet 'ParamÃ¨tres GÃ©nÃ©raux' est introuvable.");

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
// Fonctions liÃ©es aux questions du formulaire
// ------------------------------------

/**
 * Identifie les langues disponibles pour un type de test donnÃ©.
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
    const ordreLangues = ['FR', 'EN', 'ES', 'DE']; // Ordre de prioritÃ©
    languesAInclure.sort((a, b) => {
        const indexA = ordreLangues.indexOf(a.code);
        const indexB = ordreLangues.indexOf(b.code);
        if (indexA === -1) return 1; // Mettre les langues non listÃ©es Ã  la fin
        if (indexB === -1) return -1;
        return indexA - indexB;
    });
    Logger.log(`Langues triÃ©es pour la gÃ©nÃ©ration : ${languesAInclure.map(l => l.code).join(', ')}`);
    // --- FIN DE L'AJOUT ---

    if (languesAInclure.length === 0) {
        throw new Error("Aucune feuille de questions trouvÃ©e pour le type '" + typeTest + "'.");
    }
    return languesAInclure;
}


/**
 * Construit les questions dans le formulaire, en gÃ©rant le multi-langues.
 */
function _construireQuestionsFormulaire(form, languesAInclure, nbQuestionsConfig) {
    if (languesAInclure.length > 1) {
        Logger.log(`Mode multi-langues dÃ©tectÃ© (${languesAInclure.length} langues).`);
        const itemLangue = form.addMultipleChoiceItem().setTitle("Langue / Language").setRequired(true);
        const choices = [];
        languesAInclure.forEach(langue => {
            const page = form.addPageBreakItem().setTitle("Questions (" + langue.nomComplet + ")");
            choices.push(itemLangue.createChoice(langue.nomComplet, page));
            
            _ajouterQuestionsDepuisFeuille(form, langue.feuille, nbQuestionsConfig);
            
            // --- DÃ‰BUT DE LA CORRECTION v8.1 ---
        
            // On s'assure que la redirection vers la page de soumission est bien appliquÃ©e
            // Ã  l'objet 'page' que nous venons de crÃ©er, qui est un PageBreakItem.
            if (page && typeof page.setGoToPage === 'function') {
                page.setGoToPage(FormApp.PageNavigationType.SUBMIT);
            }
            
            // --- FIN DE LA CORRECTION v8.1 ---
        });
        itemLangue.setChoices(choices);
    } else {
        Logger.log(`Mode langue unique dÃ©tectÃ©. Insertion directe des questions.`);
        const uniqueLangue = languesAInclure[0];
        _ajouterQuestionsDepuisFeuille(form, uniqueLangue.feuille, nbQuestionsConfig);
    }
}

/**
 * Ajoute une sÃ©rie de questions Ã  un formulaire Ã  partir d'une feuille de calcul.
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
 * CrÃ©e un item (question) dans le formulaire en fonction de ses spÃ©cifications.
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
  const map = { FR: 'FranÃ§ais', EN: 'English', ES: 'EspaÃ±ol', DE: 'Deutsch' };
  return map[String(code || '').toUpperCase()] || code;
}
```

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\01_Moteur\forcerAutorisation.js

```javascript

function forcerAutorisation() {
  // Cette simple ligne est suffisante pour demander les autorisations Drive.
  DriveApp.getRootFolder(); 
  SpreadsheetApp.getUi().alert('Autorisation accordÃ©e ! Vous pouvez maintenant retourner Ã  votre feuille de calcul et relancer le dÃ©ploiement.');
}
```

