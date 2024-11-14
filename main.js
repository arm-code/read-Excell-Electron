const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const XLSX = require('xlsx');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true,
        }
    });

    mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

// IPC para abrir el cuadro de diálogo de selección de archivo
ipcMain.handle('selectExcelFile', async () => {
    const result = await dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [{ name: 'Archivos Excel', extensions: ['xlsx', 'xls'] }]
    });
    return result.filePaths[0];  // Retorna la ruta del archivo seleccionado
});

// IPC para leer el archivo Excel
ipcMain.handle('readExcelFile', (event, filePath) => {
    console.log('Leyendo archivo Excel:', filePath);

    try {
        const workbook = XLSX.readFile(filePath);
        const data = {};
        const hojas = workbook.SheetNames;

         // Iterar sobre todas las hojas presentes en el archivo
         hojas.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];

            if (sheet) {
                // Leer el rango B2:AE2 de la hoja
                const dataRow = XLSX.utils.sheet_to_json(sheet, {
                    header: 1,
                    range: 'B2:AE2',  // Leer el rango B2:AE2
                });

                // Guardar los datos de la primera fila de la hoja
                data[sheetName] = dataRow[0]; // Solo tomamos la primera fila
            }
        });

        return data;

    } catch (error) {
        console.error('Error al leer el archivo Excel:', error);
        return { error: 'Error al leer el archivo Excel' };
    }
});
