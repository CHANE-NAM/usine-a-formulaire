// Tools/backup_menu.js
"use strict";

/*
 But : point d’entrée unique pour lancer, au choix,
   [1] backup_gas.ps1  (sauvegarde GAS → GitHub)
   [2] snapshot_rk.ps1 (snapshot complet : GAS + CSV + ZIP)
   [3] les deux (1 puis 2)

 Finalité : éviter d’ouvrir PowerShell à la main, et rendre le menu
 robuste face aux lignes “parasites” injectées par certains outils
 (ex. activation auto d’un venv Python).

 Utilisation :
   node .\Tools\backup_menu.js           # menu interactif
   node .\Tools\backup_menu.js 1|2|3     # choix direct, sans menu
*/

const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");
const readline = require("readline");

// === Racine du dépôt (un cran au-dessus du dossier contenant ce fichier) ===
const repoRoot = path.resolve(__dirname, "..");

// === Choisit le premier chemin existant dans une liste ===
function pickExisting(paths) {
  for (const p of paths) {
    try {
      if (fs.existsSync(p)) return p;
    } catch {}
  }
  return null;
}

// === Candidats de chemins pour retrouver les scripts PS ===
// (fonctionne que le menu soit dans Tools/ ou Tools/menu/)
const backupCandidates = [
  path.resolve(__dirname, "backup_gas.ps1"),
  path.resolve(__dirname, "backup", "backup_gas.ps1"),
  path.resolve(repoRoot, "Tools", "backup_gas.ps1"),
  path.resolve(repoRoot, "Tools", "backup", "backup_gas.ps1"),
];

const snapshotCandidates = [
  path.resolve(__dirname, "snapshot_rk.ps1"),
  path.resolve(__dirname, "snapshot", "snapshot_rk.ps1"),
  path.resolve(repoRoot, "Tools", "snapshot_rk.ps1"),
  path.resolve(repoRoot, "Tools", "snapshot", "snapshot_rk.ps1"),
];

const psBackup = pickExisting(backupCandidates);
const psSnapshot = pickExisting(snapshotCandidates);

function assertFound(label, file, tried = []) {
  if (file) return;
  console.error(`\n✗ Introuvable: ${label}`);
  if (tried.length) {
    console.error("Chemins testés :");
    for (const t of tried) console.error(" - " + t);
  }
  process.exit(2);
}

// Exige que les deux scripts existent (si besoin, relâcher cette contrainte)
assertFound("backup_gas.ps1", psBackup, backupCandidates);
assertFound("snapshot_rk.ps1", psSnapshot, snapshotCandidates);

// === Lance un script PowerShell en héritant de l’I/O du terminal ===
function runPS(psPath) {
  return new Promise((resolve, reject) => {
    const ps = spawn(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", psPath],
      { cwd: repoRoot, windowsHide: false }
    );
    ps.stdout.on("data", (d) => process.stdout.write(d));
    ps.stderr.on("data", (d) => process.stderr.write(d));
    ps.on("close", (code) =>
      code === 0 ? resolve() : reject(new Error(`Exit ${code}`))
    );
  });
}

// === Exécute selon un choix (1/2/3) ===
async function executeChoice(choice) {
  if (choice === "1") {
    await runPS(psBackup);
  } else if (choice === "2") {
    await runPS(psSnapshot);
  } else if (choice === "3") {
    await runPS(psBackup);
    await runPS(psSnapshot);
  } else {
    console.log("Annulé.");
    process.exit(0);
  }
  console.log("\n✓ Terminé.");
  process.exit(0);
}

// --- Menu interactif ROBUSTE (ignore les lignes parasites : Activate.ps1 & cie) ---
function askChoice() {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    const showMenu = () => {
      console.log("\n=== Menu sauvegardes r&K ===");
      console.log("[1] Backup GAS → GitHub");
      console.log("[2] Snapshot complet (GAS + CSV + ZIP)");
      console.log("[3] Les deux (1 puis 2)");
    };
    const prompt = () =>
      rl.question("\nTon choix ? ", (raw) => {
        const ans = String(raw).trim();
        // N'accepte QUE 1/2/3 ; toute autre ligne (ex. & ...Activate.ps1) est ignorée
        if (!/^[123]$/.test(ans)) {
          console.log("Choix invalide. Tape 1, 2 ou 3.");
          return prompt();
        }
        rl.close();
        resolve(ans);
      });

    showMenu();
    prompt();
  });
}

// === Point d'entrée ===
(async () => {
  const argChoice = (process.argv[2] || "").trim();
  const choice = /^[123]$/.test(argChoice) ? argChoice : await askChoice();
  try {
    await executeChoice(choice);
  } catch (e) {
    console.error("\n✗ Erreur:", e.message);
    process.exit(1);
  }
})();
