# Snapshot: SNAPSHOT_20250906_210001

- Genere le: 2025-09-07T07:42:48Z
- CSV: 52 fichiers
- Projets GAS concatenes: 4

## Google Sheets
- BDD_V2_Tests_Profils_1m2MGB (onglets: 43, lignes: 1920) - id_hint: BDD_V2_Tests_Profils_1m2MGB
  - Copie_de_Questions_MBTI_EN: 75 lignes, 6 colonnes - BDD_V2_Tests_Profils_1m2MGB\Copie_de_Questions_MBTI_EN.csv
    - colonnes: ID, TypeQuestion, TitreQuestion, Options, Logique, description/libÃ©lÃ©s
  - ex_sys_PiecesJointes: 76 lignes, 6 colonnes - BDD_V2_Tests_Profils_1m2MGB\ex_sys_PiecesJointes.csv
    - colonnes: Type_Test, Profil_Code, Email_Niveau, Langue, Nom_Pour_Info, ID_Fichier_Drive
  - Feuille_36: 9 lignes, 1 colonnes - BDD_V2_Tests_Profils_1m2MGB\Feuille_36.csv
  - Gabarits_Emails: 6 lignes, 6 colonnes - BDD_V2_Tests_Profils_1m2MGB\Gabarits_Emails.csv
    - colonnes: ID_Gabarit, Langue, Sujet, Niveau_Details_Resultats, Niveau_Pieces_Jointes, Corps_HTML
  - Liste_Fichiers_Drive: 51 lignes, 2 colonnes - BDD_V2_Tests_Profils_1m2MGB\Liste_Fichiers_Drive.csv
    - colonnes: Nom du Fichier, ID du Fichier
  - Nomenclature: 39 lignes, 8 colonnes - BDD_V2_Tests_Profils_1m2MGB\Nomenclature.csv
    - colonnes: Questions_[Couleurs]_[FR], , , , , , , a
  - Nomenclature2: 14 lignes, 3 colonnes - BDD_V2_Tests_Profils_1m2MGB\Nomenclature2.csv
    - colonnes: Nom du Placeholder {{...}}, Origine de la DonnÃ©e, Mode de Calcul / Logique d'Obtention
  - Profils_ANCRES_EN: 8 lignes, 11 colonnes - BDD_V2_Tests_Profils_1m2MGB\Profils_ANCRES_EN.csv
    - colonnes: Code_Profil, Nom_Complet, Titre_Profil, Description_Profil, Detail_1, Detail_2, Detail_3, Detail_4, Detail_5, Detail_6, Detail_7
- CONFIG_V2_Usine_Tests_1kLBqI (onglets: 8, lignes: 118) - id_hint: CONFIG_V2_Usine_Tests_1kLBqI
  - Formulaires: 0 lignes, 3 colonnes - CONFIG_V2_Usine_Tests_1kLBqI\Formulaires.csv
    - colonnes: Titre du Test, Lien vers le Formulaire, Date de CrÃ©ation
  - nomenclature: 60 lignes, 1 colonnes - CONFIG_V2_Usine_Tests_1kLBqI\nomenclature.csv
    - colonnes: ParamÃ¨tres GÃ©nÃ©raux
  - Param_tres_G_n_raux: 13 lignes, 38 colonnes - CONFIG_V2_Usine_Tests_1kLBqI\Param_tres_G_n_raux.csv
    - colonnes: ID_Gabarit_Email_Repondant, Id_Unique, Titre_Formulaire_Utilisateur, Sous-Titre_Formulaire, Nom_Fichier_Complet, Statut, Lien_Formulaire_Public, AccÃ¨s Direct Formulaire, lien_form_entier, Racc Public, Type_Test, Blocs_Meta_A_Inclure
  - R_ponses_au_formulaire_1: 6 lignes, 42 colonnes - CONFIG_V2_Usine_Tests_1kLBqI\R_ponses_au_formulaire_1.csv
    - colonnes: Horodateur, Titre du Test (pour l'utilisateur), Type de Test, Nombre de questions minimum, Nombre de questions Ã  utiliser, Envoyer un email au rÃ©pondant ?, Quand envoyer l'email ?, Quel niveau de contenu ?, Envoyer un email au patron ?, Quand envoyer l'email ?, Quel niveau de contenu ?, Envoyer un email au formateur ?
  - ref_Modes_Traitement: 12 lignes, 5 colonnes - CONFIG_V2_Usine_Tests_1kLBqI\ref_Modes_Traitement.csv
    - colonnes: Code_Unique, Mode_Questionnement, Mode_Traitement, Description Longue, ParamÃ¨tres NÃ©cessaires (Exemple de structure JSON pour une question)
  - sys_Contenu_Emails: 6 lignes, 4 colonnes - CONFIG_V2_Usine_Tests_1kLBqI\sys_Contenu_Emails.csv
    - colonnes: Type_Test, Niveau_Contenu, Sujet_Email, Corps_Email
  - sys_ID_Fichiers: 11 lignes, 5 colonnes - CONFIG_V2_Usine_Tests_1kLBqI\sys_ID_Fichiers.csv
    - colonnes: ClÃ©, ID, ID Script, nom_fich_google_Sheet, ID google Sheet
  - sys_Options_Parametres: 10 lignes, 11 colonnes - CONFIG_V2_Usine_Tests_1kLBqI\sys_Options_Parametres.csv
    - colonnes: Statut, Type_Test, Repondant_Email_Actif, Repondant_Quand, Repondant_Contenu, Patron_Email_Mode, Patron_Quand, Patron_Contenu, Formateur_Email_Actif, Formateur_Quand, Formateur_Contenu
- TEMPLATE_V2_Kit_de_Traitement_1XwyTt (onglets: , lignes: 3) - id_hint: TEMPLATE_V2_Kit_de_Traitement_1XwyTt
  - Feuille_1: 3 lignes, 1 colonnes - TEMPLATE_V2_Kit_de_Traitement_1XwyTt\Feuille_1.csv
    - colonnes: q

## Projets GAS (fonctions detectees)
- scripts__MOTEUR_V2_Usine_à_Tests.txt - 586 lignes
  - fonctions: lancerMigrationV1versV2, convertirQuestionsEnJSON, testCreationFormulaire, normalizeAndDedupeCompositionEmails_, onOpen, orchestrateurDeploiementComplet_UI, lancerDeploiementComplet, getSystemIds, getConfigurationFromRow, _identifierLangues, _construireQuestionsFormulaire, _ajouterQuestionsDepuisFeuille, nbQuestionsAUtiliser, creerItemFormulaire, resolvedType, choices, getLangueFullName, forcerAutorisation
- scripts__CONFIG_V2_Usine_à_Tests.txt - 530 lignes
  - fonctions: onOpen, showConfigurationSidebar, getInitialData, getQuestionCountForTestType, processNewTestConfiguration, showEditSidebar_UI, showEditSidebar, getTestDataForEdit, updateTestData, showDuplicateUI, duplicateTestConfiguration, showPrintableSheetUI, generatePrintableSheet, getSystemIds, convertirLiensExistantsEnCourts, addValidationMenu_, normalizeHeader_, getHeaderRow_, findSheetByVariants_, assertHeaders_
- scripts__BDD_V2_Tests_Profils.txt - 113 lignes
  - fonctions: onOpen, listFilesFromFolder, shouldRecurse, getFilesRecursive
- scripts__TEMPLATE_V2_Kit_de_Traitement.txt - 2890 lignes
  - fonctions: calculerResultats, _executerCalcul, _aiguillerCalcul, _traiterQCU_CAT, valeur, _traiterECHELLE_NOTE, _determinerProfilFinal, _chargerProfils, _creerMapCodeVersNom, _chargerQuestions, headers, _calculerScoresMaxPossibles, mode, _normStr, _normLang, onOpen, onInstall, retraiterReponse_UI, ouvrirSidebarPourLigne, activerTraitementAutomatique

