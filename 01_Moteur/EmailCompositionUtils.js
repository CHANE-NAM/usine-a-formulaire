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
