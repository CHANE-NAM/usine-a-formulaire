// Patch CryptoJS.WordArray.random pour Google Apps Script
// Remplace la source d'aléa native (inexistante dans GAS) par Utilities.getUuid() + MD5.
// Charge ce fichier APRES AAA_CryptoJS_AES.gs (d'où le préfixe ZZZ).
(function () {
  if (typeof CryptoJS === 'undefined' || !CryptoJS.lib || !CryptoJS.lib.WordArray || !CryptoJS.MD5) {
    throw new Error('CryptoJS non chargé avant le patch WordArray.random (vérifie AAA_CryptoJS_AES.gs)');
  }
  var WA = CryptoJS.lib.WordArray;
  // Génère nBytes d’entropie pseudo-aléatoire (suffisant pour notre usage de salage OpenSSL).
  WA.random = function (nBytes) {
    var words = [];
    var produced = 0;
    // Chaque MD5(uuid) = 16 octets = 4 words
    while (produced < nBytes) {
      var uuid = Utilities.getUuid();           // 36 chars pseudo-aléatoires fournis par GAS
      var digest = CryptoJS.MD5(uuid).words;    // 4 words (16 bytes)
      for (var i = 0; i < digest.length && produced < nBytes; i++) {
        words.push(digest[i]);
        produced += 4;
      }
    }
    // Ajuste la longueur exacte (sigBytes) à nBytes
    return CryptoJS.lib.WordArray.create(words, nBytes);
  };
})();