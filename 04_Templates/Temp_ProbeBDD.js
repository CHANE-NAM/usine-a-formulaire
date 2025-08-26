// ============================================================================
// FICHIER TEMPORAIRE : Temp_ProbeBDD.gs  (à supprimer après debug)
// Rôle : vérifier que Questions_* contient des JSON valides et des 'mode'
// ============================================================================

function Temp_probeQuestions(typeTest, langue) {
  try {
    const ids = getSystemIds();
    const bdd = SpreadsheetApp.openById(ids.ID_BDD);
    const name = `Questions_${typeTest}_${langue}`;
    const sh = bdd.getSheetByName(name);
    if (!sh) { Logger.log('❌ Feuille introuvable: ' + name); return; }

    const values = sh.getDataRange().getValues();
    const headers = values.shift();
    const colID     = headers.indexOf('ID');
    const colParams = headers.indexOf('Paramètres (JSON)');
    const colTypeQ  = headers.indexOf('TypeQuestion');

    let total=0, ok=0, badJson=0, noMode=0;
    const sample = [];

    values.forEach(r => {
      const id = r[colID];
      const js = r[colParams];
      const tq = r[colTypeQ];
      if (!id) return;
      total++;
      try {
        const p = JSON.parse(js || '{}');
        if (p && p.mode) ok++; else noMode++;
        if (sample.length < 5) sample.push({id, type:tq, params:p});
      } catch(e) {
        badJson++;
        if (sample.length < 5) sample.push({id, type:tq, params:'<JSON invalide>'});
      }
    });

    Logger.log('Feuille %s : total=%s, ok=%s, badJson=%s, sansMode=%s', name, total, ok, badJson, noMode);
    Logger.log('Échantillon 1..5 → ' + JSON.stringify(sample));
  } catch (e) {
    Logger.log('ERREUR Temp_probeQuestions: ' + e.message);
  }
}

// Wrappers pratiques (sélectionne & Exécute)
function run_probe_ADA_FR(){ Temp_probeQuestions('r&K_Adaptabilite','FR'); }
// Tu pourras ajouter ensuite si besoin :
// function run_probe_RESI_FR(){ Temp_probeQuestions('r&K_Resilience','FR'); }
// function run_probe_CREA_FR(){ Temp_probeQuestions('r&K_Creativite','FR'); }
