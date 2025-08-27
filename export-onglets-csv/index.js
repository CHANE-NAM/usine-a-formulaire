// export-onglets-csv/index.js
// Usages :
//   node index.js --out "<snapshotDir>" --id "ID1" --id "ID2"
//   node index.js --out "<snapshotDir>" --url "https://docs.google.com/spreadsheets/d/ID/edit" [...]
//   node index.js --out "<snapshotDir>" --gsheet "C:\\chemin\\classeur.gsheet" [...]

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets.readonly',
  'https://www.googleapis.com/auth/drive.readonly',
];

const CREDENTIALS = path.join(__dirname, 'credentials.json');
const TOKEN_PATH  = path.join(__dirname, 'token.json');

/* ---------------------- Utils ---------------------- */
function arrToCsv(rows) {
  return (rows || []).map(r =>
    (r || []).map(cell => {
      if (cell == null) return '';
      const s = String(cell).replace(/"/g, '""');
      return /[",\n]/.test(s) ? `"${s}"` : s;
    }).join(',')
  ).join('\n');
}

function safeName(s) {
  return String(s || '').replace(/[^\w\-]+/g, '_').replace(/^_+|_+$/g, '');
}

function ensureFileExists(p, hint) {
  if (!fs.existsSync(p)) {
    throw new Error(`${hint || 'Fichier manquant'} : ${p}`);
  }
}

/* -------------------- Arguments -------------------- */
function parseArgs() {
  const args = process.argv.slice(2);
  const opts = { out: null, gsheet: [], url: [], id: [] };
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--out')    { opts.out = args[++i]; continue; }
    if (a === '--gsheet') { opts.gsheet.push(args[++i]); continue; }
    if (a === '--url')    { opts.url.push(args[++i]); continue; }
    if (a === '--id')     { opts.id.push(args[++i]); continue; }
  }
  if (!opts.out) throw new Error('Argument requis : --out "<dossierSnapshot>"');
  if (!(opts.gsheet.length || opts.url.length || opts.id.length)) {
    throw new Error('Fournir au moins un classeur via --id "<ID>" OU --url "<lien>" OU --gsheet "<fichier.gsheet>".');
  }
  return opts;
}

/* ---------- ID depuis .gsheet, URL ou ID direct ----- */
function spreadsheetIdFromAny(input) {
  // 1) ID direct
  if (/^[a-zA-Z0-9_-]{20,}$/.test(input) && !input.startsWith('http')) return input;

  // 2) URL
  if (input.startsWith('http')) {
    const m = input.match(/\/d\/([a-zA-Z0-9_-]+)\//);
    if (m && m[1]) return m[1];
    throw new Error(`URL non reconnue : ${input}`);
  }

  // 3) .gsheet local (attention : peut être "virtuel" sur Google Drive)
  const stat = fs.lstatSync(input);
  if (stat.isDirectory()) {
    throw new Error(`Chemin détecté comme dossier, pas .gsheet : ${input}. Utilise plutôt --url ou --id.`);
  }
  const raw = fs.readFileSync(input, 'utf8');
  const json = JSON.parse(raw);
  const url = json.url || json.doc_id || '';
  const m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)\//);
  if (m && m[1]) return m[1];
  if (json.doc_id) return json.doc_id;
  throw new Error(`Impossible d'extraire l'ID depuis : ${input}`);
}

/* ------------------- Auth Google ------------------- */
async function authorize() {
  ensureFileExists(CREDENTIALS, 'credentials.json introuvable');
  const credentials = JSON.parse(fs.readFileSync(CREDENTIALS, 'utf8'));
  const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web || {};
  if (!client_id || !client_secret) {
    throw new Error('credentials.json invalide (client_id/client_secret manquants).');
  }
  const redirect = (redirect_uris && redirect_uris[0]) || 'http://localhost';
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect);

  // Jeton déjà présent ?
  if (fs.existsSync(TOKEN_PATH)) {
    oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8')));
    return oAuth2Client;
  }

  // Génère l'URL d'autorisation
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent',
  });

  console.log('\nAutorisez l’accès en suivant ce lien :\n', authUrl, '\n');

  // Essaie d'ouvrir le navigateur automatiquement (Windows)
  try {
    const { exec } = require('child_process');
    if (process.platform === 'win32') exec(`start "" "${authUrl}"`);
  } catch (_) {}

  // Demande le code d'autorisation
  const rl = require('readline').createInterface({ input: process.stdin, output: process.stdout });
  const code = await new Promise(res => rl.question('Code d’autorisation : ', ans => { rl.close(); res(ans.trim()); }));
  if (!code) throw new Error('Aucun code saisi. Relancez et collez le code fourni par Google.');

  // Échange code ↔ token
  const { tokens } = await oAuth2Client.getToken(code);
  oAuth2Client.setCredentials(tokens);
  fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));
  console.log('Token enregistré -> token.json');
  return oAuth2Client;
}

/* ---------------- Export d’un classeur -------------- */
async function exportSpreadsheet(auth, spreadsheetId, outDir) {
  const sheets = google.sheets({ version: 'v4', auth });

  // Métadonnées du classeur
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const bookTitle = meta.data.properties?.title || spreadsheetId;
  const container = path.join(outDir, `${safeName(bookTitle)}_${spreadsheetId.slice(0, 6)}`);
  fs.mkdirSync(container, { recursive: true });

  // Pour chaque onglet : values.get(range = 'NomOnglet')
  for (const sh of (meta.data.sheets || [])) {
    const tabName = sh.properties?.title || `Sheet_${sh.properties?.sheetId}`;
    const range = `'${tabName.replace(/'/g, "''")}'`;
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
    const csv = arrToCsv(res.data.values || []);
    const f = path.join(container, safeName(tabName) + '.csv');
    fs.writeFileSync(f, csv, 'utf8');
    console.log(`OK ${bookTitle} -> ${tabName}.csv`);
  }
}

/* ------------------- Programme main ---------------- */
(async () => {
  try {
    const opts = parseArgs();
    const outDir = path.resolve(opts.out);
    fs.mkdirSync(outDir, { recursive: true });

    const ids = [
      ...opts.gsheet.map(spreadsheetIdFromAny),
      ...opts.url.map(spreadsheetIdFromAny),
      ...opts.id.map(spreadsheetIdFromAny),
    ];

    const auth = await authorize();
    for (const id of ids) {
      await exportSpreadsheet(auth, id, outDir);
    }
    console.log(`\nCSV déposés dans : ${outDir}`);
  } catch (e) {
    console.error('ERREUR :', e.message);
    process.exit(1);
  }
})();
