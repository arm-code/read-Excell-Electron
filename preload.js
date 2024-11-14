const { contextBridge, ipcRenderer } = require('electron');

// Exponer la función de seleccionar archivo Excel
contextBridge.exposeInMainWorld('electron', {
    selectExcelFile: () => ipcRenderer.invoke('selectExcelFile'),
    readExcelFile: (filePath) => ipcRenderer.invoke('readExcelFile', filePath)
});
