const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const xlsx = require('xlsx');
const fs = require('fs');

function createWindow() {
    const win = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            preload: path.join(__dirname, 'renderer.js'),
            nodeIntegration: true,
            contextIsolation: false
        }
    });

    win.loadFile('index.html');
}

app.whenReady().then(createWindow);

ipcMain.handle('compare-files', async (event, file1, file2) => {
    const workbook1 = xlsx.readFile(file1);
    const workbook2 = xlsx.readFile(file2);

    const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
    const sheet2 = workbook2.Sheets[workbook2.SheetNames[0]];

    const data1 = xlsx.utils.sheet_to_json(sheet1, { header: 1 });
    const data2 = xlsx.utils.sheet_to_json(sheet2, { header: 1 });

    const differences = [];

    const maxLength = Math.max(data1.length, data2.length);

    for (let i = 0; i < maxLength; i++) {
        const row1 = data1[i] || [];
        const row2 = data2[i] || [];

        if (JSON.stringify(row1) !== JSON.stringify(row2)) {
            differences.push({
                row: i + 1,
                file1: row1,
                file2: row2
            });
        }
    }

    return differences;
});

ipcMain.handle('select-file', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }]
    });

    if (canceled) {
        return null;
    } else {
        return filePaths[0];
    }
});

ipcMain.handle('export-to-excel', async (event, differences) => {
    const workbook = xlsx.utils.book_new();
    const worksheetData = [["Row", "File 1 Content", "File 2 Content"]];

    differences.forEach(diff => {
        worksheetData.push([diff.row, JSON.stringify(diff.file1), JSON.stringify(diff.file2)]);
    });

    const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Differences');

    const { filePath } = await dialog.showSaveDialog({
        title: 'Enregistrer le fichier Excel',
        defaultPath: 'differences.xlsx',
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
    });

    if (filePath) {
        xlsx.writeFile(workbook, filePath);
    }
});
