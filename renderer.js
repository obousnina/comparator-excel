const { ipcRenderer } = require('electron');

const file1Btn = document.getElementById('selectFile1');
const file2Btn = document.getElementById('selectFile2');
const compareBtn = document.getElementById('compareBtn');
const exportExcelBtn = document.getElementById('exportExcelBtn');
const file1Display = document.getElementById('file1');
const file2Display = document.getElementById('file2');
const differencesDisplay = document.getElementById('differences');
const loader = document.getElementById('loader');

let file1Path = '';
let file2Path = '';
let differences = [];

file1Btn.addEventListener('click', async () => {
    file1Path = await ipcRenderer.invoke('select-file');
    file1Display.textContent = file1Path || 'Aucun fichier sélectionné';
});

file2Btn.addEventListener('click', async () => {
    file2Path = await ipcRenderer.invoke('select-file');
    file2Display.textContent = file2Path || 'Aucun fichier sélectionné';
});

compareBtn.addEventListener('click', async () => {
    if (!file1Path || !file2Path) {
        alert('Veuillez sélectionner les deux fichiers');
        return;
    }

    // Afficher le loader
    loader.style.display = 'block';
    differencesDisplay.innerHTML = '';

    differences = await ipcRenderer.invoke('compare-files', file1Path, file2Path);

    // Masquer le loader après la comparaison
    loader.style.display = 'none';

    if (differences.length === 0) {
        differencesDisplay.innerHTML = '<li>Aucune différence trouvée.</li>';
    } else {
        differences.forEach(diff => {
            const listItem = document.createElement('li');
            listItem.innerHTML = `
                <strong>Ligne ${diff.row}</strong><br>
                <span style="color: red;">Fichier 1:</span> ${JSON.stringify(diff.file1)}<br>
                <span style="color: green;">Fichier 2:</span> ${JSON.stringify(diff.file2)}
            `;
            differencesDisplay.appendChild(listItem);
        });
        exportExcelBtn.style.display = 'inline';  // Affiche le bouton Exporter vers Excel
    }
});

exportExcelBtn.addEventListener('click', async () => {
    if (differences.length === 0) {
        alert("Aucune différence à exporter.");
        return;
    }

    await ipcRenderer.invoke('export-to-excel', differences);
    alert('Fichier Excel généré avec succès.');
});
