const fs = require('fs');
const path = require('path');

// --- CONFIGURATION ---
// Extensions de fichiers à inclure
const scriptExtensions = [
    '.js', '.ts', '.jsx', '.tsx', '.py', '.html', '.css', '.scss', '.less',
    '.xml', '.php', '.rb', '.java', '.c', '.cpp', '.cs', '.go',
    '.sh', '.ps1', '.bat', '.cmd', '.sql', '.vue', '.svelte', '.astro'
];

// Le dossier à scanner
const folderToScan = './'; // À ajuster si besoin
// --- FIN CONFIGURATION ---

// Récupérer le nom du répertoire scanné
const resolvedFolderPath = path.resolve(folderToScan);
const dirName = path.basename(resolvedFolderPath);

// Créer une date au format aammjj
const now = new Date();
const year = String(now.getFullYear()).slice(2); // 2 derniers chiffres de l'année
const month = String(now.getMonth() + 1).padStart(2, '0');
const day = String(now.getDate()).padStart(2, '0');
const dateStamp = `${year}${month}${day}`;

// Nom du fichier de sortie
const outputFileName = `${dirName}_Script_${dateStamp}.txt`;

// Créer une chaîne avec date et heure en français
const dateTimeString = `// Fichier généré le ${now.toLocaleDateString('fr-FR')} à ${now.toLocaleTimeString('fr-FR')}\n\n`;

// Initialiser la variable qui accumulera le contenu
let allContent = dateTimeString;

// Fonction récursive
function readFilesRecursively(directory) {
    fs.readdirSync(directory).forEach(file => {
        const absolutePath = path.join(directory, file);
        if (fs.statSync(absolutePath).isDirectory()) {
            readFilesRecursively(absolutePath);
        } else {
            const fileExtension = path.extname(file).toLowerCase();
            if (scriptExtensions.includes(fileExtension)) {
                allContent += `// --- Début du fichier: ${absolutePath} ---\n`;
                allContent += fs.readFileSync(absolutePath, 'utf8');
                allContent += `\n// --- Fin du fichier: ${absolutePath} ---\n\n`;
            }
        }
    });
}

// Supprimer le fichier de sortie existant
if (fs.existsSync(outputFileName)) {
    fs.unlinkSync(outputFileName);
    console.log(`Ancien fichier '${outputFileName}' supprimé.`);
}

// Lancer le traitement
try {
    readFilesRecursively(folderToScan);
    fs.writeFileSync(outputFileName, allContent, 'utf8');
    console.log(`Succès : Tous les scripts ont été exportés dans '${outputFileName}'`);
} catch (error) {
    console.error(`Erreur lors de la lecture ou de l'écriture des fichiers : ${error.message}`);
}
