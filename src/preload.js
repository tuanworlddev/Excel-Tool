const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
  openFileDialog: (name, extensions) =>
    ipcRenderer.invoke("openFileDialog", name, extensions),
  readExcelFile: (filePath) => ipcRenderer.invoke("readExcelFile", filePath),
  exportData: (data) => ipcRenderer.invoke("exportData", data),
});
