/**
 * Moteur de calcul — r&K_Environnement (échelle 1..10)
 * Regroupe 60 items ENV001..ENV060 en 15 thèmes (paquets de 4 : 2 "K", 2 "r").
 * Retourne K et r globaux, les scores par thème, et des champs à plat pour les gabarits.
 */

/**
 * Charge les recommandations depuis la BDD pour le test r&K Environnement.
 * VERSION AMÉLIORÉE : Nettoie (normalise) les clés pour une correspondance robuste.
 * @param {string} langueCible - Le code de la langue (ex: 'FR').
 * @returns {Map<string, string>} Une map associant un nom de thème normalisé à sa recommandation.
 */
function _chargerRecommandationsTheme_rK_Env(langueCible) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const sheetName = `Profils_r&K_Environnement_${langueCible}`;
    const sheet = bdd.getSheetByName(sheetName);
    if (!sheet) return new Map();

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const axeCol = headers.indexOf('Axe');
    const recoCol = headers.indexOf('Recommandation');

    if (axeCol === -1 || recoCol === -1) return new Map();

    const recommandationsMap = new Map();
    for (const row of data) {
      const axe = row[axeCol];
      const reco = row[recoCol];
      if (axe && reco) {
        // AJOUT : On normalise la clé avant de la stocker
        const cleNormalisee = String(axe).trim().toLowerCase();
        recommandationsMap.set(cleNormalisee, reco);
      }
    }
    return recommandationsMap;
  } catch (e) {
    Logger.log(`Erreur lors du chargement des recommandations pour r&K Environnement : ${e.message}`);
    return new Map();
  }
}

function calculerResultats_rK_Environnement(reponse, langueCible, config) {
  // AJOUT : Charger les recommandations au début
  const recommandationsMap = _chargerRecommandationsTheme_rK_Env(langueCible);

  // 1) Récupère toutes les valeurs numériques des items ENVxxx (clés = n° 1..60)
  const envVals = {};
  for (const k in reponse) {
    const m = String(k).match(/^ENV(\d{3})/);
    if (m) {
      const n = parseInt(m[1], 10);
      const v = Number(String(reponse[k]).replace(',', '.'));
      if (!isNaN(v)) envVals[n] = v;
    }
  }

  // 2) Thèmes (15 x 4 items)
  const THEMES = [
    "Concurrence & Pression du marché",
    "Clients & Demande",
    "Technologies & Innovation",
    "Réglementation & Cadre juridique",
    "Ressources humaines & Compétences",
    "Financement & Accès aux capitaux",
    "Fournisseurs & Logistique",
    "Ressources & Infrastructures matérielles",
    "Image & Réputation sectorielle",
    "Partenariats & Réseaux",
    "Territoire & Environnement géographique",
    "Tendances sociétales & culturelles",
    "Contexte économique global",
    "Risques & Sécurité",
    "Opportunités de croissance & Marchés",
  ];

  // helpers
  const avg = (a,b) => (a+b)/2;
  const clamp = (x,min,max)=>Math.max(min,Math.min(max,x));
  const interpK = (x) => x>=7 ? "Environnement plutôt stable et prévisible" : x<=3 ? "Environnement plutôt instable / changeant" : "Stabilité modérée avec quelques variations";
  const interpr = (x) => x>=7 ? "Changements rapides / forte dynamique" : x<=3 ? "Changements lents / faible dynamique" : "Vitesse de changement modérée";

  // 3) Boucle par thème
  const themes = [];
  let sumK = 0, sumR = 0, filledK = 0, filledR = 0;
  for (let t=0; t<15; t++) {
    const base = t*4;
    const K1 = envVals[base+1], K2 = envVals[base+2];
    const R1 = envVals[base+3], R2 = envVals[base+4];

    const hasK = (K1!=null && K2!=null);
    const hasR = (R1!=null && R2!=null);
    const k = hasK ? avg(K1, K2) : null;
    const r = hasR ? avg(R1, R2) : null;
    if (k!=null) { sumK += k; filledK++; }
    if (r!=null) { sumR += r; filledR++; }

    // MODIFICATION : Chercher la recommandation dans la map de manière robuste
    const nomThemeActuel = THEMES[t];
    const cleNormalisee = String(nomThemeActuel).trim().toLowerCase(); // On normalise aussi ici
    const recommandationTrouvee = recommandationsMap.get(cleNormalisee) || "Aucune recommandation spécifique pour ce thème.";

    themes.push({
      name: nomThemeActuel,
      stabilite: k!=null ? +k.toFixed(2) : "",
      vitesse:   r!=null ? +r.toFixed(2) : "",
      interpretStab: k!=null ? interpK(k) : "",
      interpretVit:  r!=null ? interpr(r) : "",
      reco: recommandationTrouvee // Utiliser la recommandation trouvée
    });
  }
  
  const scoreK = filledK ? +(sumK/filledK).toFixed(2) : 0;
  const scoreR = filledR ? +(sumR/filledR).toFixed(2) : 0;

  const hi = 6.5, lo = 3.5;
  let titreProfil = "";
  if (scoreK >= hi && scoreR <= lo) titreProfil = "Stable & Lent";
  else if (scoreK >= hi && scoreR >= hi) titreProfil = "Stable & Rapide";
  else if (scoreK <= lo && scoreR >= hi) titreProfil = "Instable & Rapide";
  else if (scoreK <= lo && scoreR <= lo) titreProfil = "Instable & Lent";
  else if (scoreK >= scoreR) titreProfil = "Plutôt Stable";
  else titreProfil = "Plutôt Rapide";

  const flat = {
    Score_Stabilite: scoreK,
    Interpretation_Stabilite: interpK(scoreK),
    Score_Vitesse: scoreR,
    Interpretation_Vitesse: interpr(scoreR),
    Titre_Profil: titreProfil,
    profilFinal: titreProfil
  };

  themes.forEach((th, i) => {
    const n = i+1;
    flat[`Nom_Theme_${n}`] = th.name;
    flat[`Score_Stabilite_Theme_${n}`] = th.stabilite;
    flat[`Interpretation_Stabilite_Theme_${n}`] = th.interpretStab;
    flat[`Score_Vitesse_Theme_${n}`] = th.vitesse;
    flat[`Interpretation_Vitesse_Theme_${n}`] = th.interpretVit;
    flat[`Recommandations_Theme_${n}`] = th.reco;
  });

  return {
    scoresData: { K: scoreK, r: scoreR },
    sousTotauxParMode: { K: scoreK, r: scoreR },
    mapCodeToName: { K: "Stabilité (K)", r: "Vitesse (r)" },
    themes,
    ...flat
  };
}