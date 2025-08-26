function testFusionRapportFull() {
  const templateId = '1F-vPh9xhtWlF2eAHEfzwgwo3cmGbIyJXrMgmCePaDKQ';

  const themes = [
    'Concurrence & Marché','Clients & Demande','Technologies & Innovation',
    'Réglementation & Cadre juridique','Ressources humaines & Compétences',
    'Financement & Accès aux capitaux','Fournisseurs & Logistique',
    'Ressources & Infrastructures matérielles','Image & Réputation sectorielle',
    'Partenariats & Réseaux','Territoire & Environnement géographique',
    'Tendances sociétales & culturelles','Contexte économique global',
    'Risques & Sécurité','Opportunités de croissance & Marchés'
  ];

  const vars = {
    Nom_Entreprise: 'ACME SA',
    Score_Stabilite: 7,
    Interpretation_Stabilite: 'Plutôt K (stable)',
    Score_Vitesse: 4,
    Interpretation_Vitesse: 'Changements lents'
  };

  // Remplir les 15 thèmes
  for (let i = 0; i < themes.length; i++) {
    const j = i + 1;
    vars['Nom_Theme_' + j] = themes[i];
    // exemples de scores ; mets tes vraies valeurs si dispo
    vars['Score_Stabilite_Theme_' + j] = 5 + (i % 3);          // 5..7
    vars['Interpretation_Stabilite_Theme_' + j] = ['Instable','Modéré','Assez stable'][(i)%3];
    vars['Score_Vitesse_Theme_' + j] = 3 + ((i+1) % 5);        // 3..7
    vars['Interpretation_Vitesse_Theme_' + j] = ['Lent','Modéré','Rapide','Très rapide','Modéré'][(i+1)%5];
    vars['Recommandations_Theme_' + j] = 'Recommandations ciblées pour le thème ' + j + '.';
  }

  const pdf = genererPdfDepuisModele(templateId, vars, 'Test_Rapport_Expert_FULL');
  DriveApp.createFile(pdf).setName('Test_Rapport_Expert_FULL.pdf');
}
