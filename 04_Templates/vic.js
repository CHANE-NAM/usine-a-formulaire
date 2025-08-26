function dupliquerCompoEmailsDepuisEnvironnement() {
  const SOURCE_TYPE = 'r&K_Environnement';
  const CIBLE_TYPES = ['r&K_Adaptabilite','r&K_Resilience','r&K_Creativite']; // modifiable
  const systemIds = getSystemIds();
  const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
  const sh = bdd.getSheetByName('sys_Composition_Emails');
  if (!sh) throw new Error("Onglet sys_Composition_Emails introuvable");

  const data = sh.getDataRange().getValues();
  const headers = data.shift();
  const idx = {
    typeTest: headers.indexOf('Type_Test'),
    langue: headers.indexOf('Code_Langue'),
    niveau: headers.indexOf('Code_Niveau_Email'),
    profil: headers.indexOf('Code_Profil'),
    element: headers.indexOf('Element'),
    ordre: headers.indexOf('Ordre'),
    contenu: headers.indexOf('Contenu / ID_Document')
  };

  const srcRows = data.filter(r => (r[idx.typeTest]||'').toString().trim() === SOURCE_TYPE);
  if (srcRows.length === 0) throw new Error("Aucune ligne source pour "+SOURCE_TYPE);

  let appended = 0;
  CIBLE_TYPES.forEach(tgt => {
    const has = data.some(r => (r[idx.typeTest]||'').toString().trim() === tgt);
    if (has) {
      Logger.log('Déjà présent: '+tgt+' (aucune copie)');
      return;
    }
    const rowsToAppend = srcRows.map(r => { const clone = r.slice(); clone[idx.typeTest] = tgt; return clone; });
    const oldLast = sh.getLastRow();
    sh.insertRowsAfter(oldLast, rowsToAppend.length);
    sh.getRange(oldLast+1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
    appended += rowsToAppend.length;
    Logger.log('Ajouté pour '+tgt+' : '+rowsToAppend.length+' lignes');
  });

  Logger.log('Terminé. Lignes ajoutées: '+appended);
}

function diagnostic_Compo_rK(options) {
  options = options || {};
  const langue = (options.langue || 'FR').trim();      // ex: 'FR' ou 'EN'
  const niveau = (options.niveau || 'N1').trim();      // ex: 'N1' / 'N3'…
  const systemIds = getSystemIds();
  const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
  const sh = bdd.getSheetByName('sys_Composition_Emails');
  if (!sh) throw new Error("sys_Composition_Emails introuvable");

  const data = sh.getDataRange().getValues();
  const headers = data.shift();
  const idx = {
    typeTest: headers.indexOf('Type_Test'),
    langue:   headers.indexOf('Code_Langue'),
    niveau:   headers.indexOf('Code_Niveau_Email'),
    profil:   headers.indexOf('Code_Profil'),
    element:  headers.indexOf('Element'),
    ordre:    headers.indexOf('Ordre')
  };

  const rowsNorm = normalizeAndDedupeCompositionEmailsRows_(data, idx);
  const TYPES = ['r&K_Environnement','r&K_Adaptabilite','r&K_Resilience','r&K_Creativite'];

  Logger.log(`► Vérif composition — langue=${langue} | niveau=${niveau}`);
  TYPES.forEach(T => {
    const matches = rowsNorm.filter(r => {
      const okType = (String(r[idx.typeTest]||'').trim() === T);
      const okLang = (String(r[idx.langue]||'').trim() === langue);
      const lvl    = String(r[idx.niveau]||'');
      const lvList = lvl.split(',').map(s=>s.trim()).filter(Boolean);
      const okLvl  = lvList.length ? lvList.includes(niveau) : lvl.includes(niveau);
      return okType && okLang && okLvl;
    });

    const byElement = {};
    matches.forEach(r => {
      const el = String(r[idx.element]||'').trim();
      byElement[el] = (byElement[el]||0) + 1;
    });

    Logger.log(`• ${T} → ${matches.length} ligne(s). Répartition: ` + JSON.stringify(byElement));
    if (matches.length === 0) Logger.log(`⚠ Aucune brique trouvée pour ${T} (${langue}/${niveau})`);
  });
}


