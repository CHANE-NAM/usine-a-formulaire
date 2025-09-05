/**
 * =================================================================================
 * == FICHIER : Graphiques.gs
 * == VERSION : 1.0 - Création initiale.
 * == RÔLE    : Génère des graphiques sous forme d'images pour les rapports.
 * =================================================================================
 */

/**
 * Génère une image de graphique radar à partir des scores des axes de résilience.
 * Utilise la méthode de création via une feuille Google Sheets temporaire.
 *
 * @param {object} axesData - Un objet contenant les scores des 5 axes. 
 * Ex: { Echec: { score: 4.6 }, Changement: { score: 4.4 }, ... }
 * @returns {Blob} L'image du graphique généré au format PNG.
 */
function creerGraphiqueRadar(axesData) {
  let tempSheet = null;
  try {
    // 1. Créer une feuille de calcul temporaire
    tempSheet = SpreadsheetApp.create("Graphique Radar Temporaire");
    const sheet = tempSheet.getSheets()[0];

    // 2. Préparer les données pour le graphique
    const data = [
      ['Axe', 'Score'],
      ['Réaction à l\'échec', parseFloat(axesData.Echec.score)],
      ['Adaptation au changement', parseFloat(axesData.Changement.score)],
      ['Utilisation des ressources', parseFloat(axesData.Ressources.score)],
      ['Gestion de crise', parseFloat(axesData.Crise.score)],
      ['Projection et objectifs', parseFloat(axesData.Objectifs.score)]
    ];
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    // 3. Construire le graphique radar
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.RADAR)
      .addRange(sheet.getRange("A1:B6"))
      .setOption('title', 'Profil de Résilience')
      .setOption('legend', { position: 'none' })
      .setOption('vAxis', { minValue: 0, maxValue: 10 }) // Fixe l'échelle de 0 à 10
      .setPosition(5, 5, 0, 0)
      .build();

    sheet.insertChart(chart);

    // 4. Récupérer le graphique en tant qu'image
    const chartImage = chart.getAs('image/png');
    
    return chartImage;

  } catch (e) {
    Logger.log("Erreur lors de la génération du graphique radar : " + e.message);
    return null; // Retourne null en cas d'erreur
  } finally {
    // 5. S'assurer que le fichier temporaire est bien supprimé
    if (tempSheet) {
      DriveApp.getFileById(tempSheet.getId()).setTrashed(true);
    }
  }
}