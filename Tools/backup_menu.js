// Tools/backup_menu.js
"use strict";

const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");
const readline = require("readline");

// === Repo root (un cran au-dessus du dossier contenant ce fichier) ===
const repoRoot = path.resolve(__dirname, "..");

// === Résolution robuste des chemins (fonctionne que le menu soit dans Tools/ ou Tools/menu/) ===
function pickExisting(paths) {
  for (const p of paths) if (fs.existsSync(p)) return p;
  return null;
}

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

assertFound("backup_gas.ps1", psBackup, backupCandidates);
assertFound("snapshot_rk.ps1", psSnapshot, snapshotCandidates);

// === Lance un script PowerShell avec héritage d'E/S ===
function runPS(psPath) {
  return new Promise((resolve, reject) => {
    const ps = spawn(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", psPath],
      { cwd: repoRoot, windowsHide: false }
    );
    ps.stdout.on("data", (d) => process.stdout.write(d));
    ps.stderr.on("data", (d) => process.stderr.write(d));
    ps.on("close", (code) => (code === 0 ? resolve() : reject(new Error(`Exit ${code}`))));
  });
}

// === Exécution selon un choix (1/2/3) ===
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

// === Entrée : arg facultatif pour bypasser le menu (node backup_menu.js 1|2|3) ===
const argChoice = (process.argv[2] || "").trim();
if (["1", "2", "3"].includes(argChoice)) {
  executeChoice(argChoice).catch((e) => {
    console.error("\n✗ Erreur:", e.message);
    process.exit(1);
  });
} else {
  // === Menu interactif ===
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  console.log("\n=== Menu sauvegardes r&K ===");
  console.log("[1] Backup GAS → GitHub");
  console.log("[2] Snapshot complet (GAS + CSV + ZIP)");
  console.log("[3] Les deux (1 puis 2)");
  console.log("[Autre] Quitter");
  rl.question("\nTon choix ? ", (ans) => {
    rl.close();
    executeChoice(ans.trim()).catch((e) => {
      console.error("\n✗ Erreur:", e.message);
      process.exit(1);
    });
  });
}
