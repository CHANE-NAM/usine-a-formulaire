// =================================================================================
// FONCTION DE MIGRATION V1 -> V2 (JSON)
// RÔLE : Convertit les questions d'un ancien format (Options/Logique)
//         vers le nouveau format V2 (Paramètres (JSON)).
// VERSION : 1.4 - Version finale et corrigée
// =================================================================================

/**
 * Fonction principale appelée depuis le menu de l'interface utilisateur.
 */
function lancerMigrationV1versV2() {
  try {
    const ID_BDD = '1m2MGBd0nyiAl3qw032B6Nfj7zQL27bRSBexiOPaRZd8';

    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Outil de Migration V1 -> V2',
      'Veuillez entrer le nom exact de l\'onglet dans la BDD à migrer :',
      ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() == ui.Button.OK && response.getResponseText() != '') {
      const sheetName = response.getResponseText().trim();
      
      const bdd = SpreadsheetApp.openById(ID_BDD);
      if (!bdd) { throw new Error(`Impossible d'ouvrir la BDD avec l'ID fourni.`); }
      const sheet = bdd.getSheetByName(sheetName);

      if (!sheet) { throw new Error(`L'onglet "${sheetName}" est introuvable dans la BDD.`); }

      const resultat = convertirQuestionsEnJSON(sheet);
      
      ui.alert(
        'Migration Terminée',
        `Rapport pour l'onglet "${sheetName}":\n\n` +
        `- Lignes traitées : ${resultat.lignesTraitees}\n` +
        `- Questions converties : ${resultat.questionsConverties}\n` +
        `- Lignes ignorées : ${resultat.lignesIgnorees}\n` +
        `- Erreurs rencontrées : ${resultat.erreurs.length}` +
        (resultat.erreurs.length > 0 ? `\n\nConsultez les logs ("Affichage > Journaux") pour le détail des erreurs.` : ''),
        ui.ButtonSet.OK);
    }
  } catch (e) {
    SpreadApp.getUi().alert(`Une erreur est survenue : ${e.message}`);
    console.error(`Erreur lors du lancement de la migration : ${e.stack}`);
  }
}

/**
 * Cœur de la logique de conversion. Lit une feuille et met à jour les lignes.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La feuille de calcul à traiter.
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
    json: headers.indexOf('Paramètres (JSON)')
  };

  if (colIndex.type === -1 || colIndex.options === -1 || colIndex.logique === -1 || colIndex.json === -1) {
    throw new Error("Colonnes requises ('TypeQuestion', 'Options', 'Logique', 'Paramètres (JSON)') manquantes.");
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
      const errorMessage = `Erreur à la ligne ${index + 2}: ${e.message}`;
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