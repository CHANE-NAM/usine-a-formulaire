function test_lancerLaGeneration() {
  Logger.log("--- ESPION A : Démarrage de la fonction de test ---");

  const numeroDeLigne = 19; 
  Logger.log(`--- ESPION B : Ligne cible = ${numeroDeLigne} ---`);

  lancerDeploiement_V4(numeroDeLigne);

  Logger.log("--- ESPION C : Fin de la fonction de test ---");
}