const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  // Ventana
  minimize: ()  => ipcRenderer.send('window-minimize'),
  maximize: ()  => ipcRenderer.send('window-maximize'),
  close:    ()  => ipcRenderer.send('window-close'),

  // Configuración
  getConfig:  ()    => ipcRenderer.invoke('get-config'),
  saveConfig: (cfg) => ipcRenderer.invoke('save-config', cfg),

  // Microsoft Graph
  getLicenses:     (cfg)     => ipcRenderer.invoke('get-licenses', cfg),
  listUsers:       (cfg)     => ipcRenderer.invoke('list-users', cfg),
  createUser:      (payload) => ipcRenderer.invoke('create-user', payload),
  toggleUser:      (payload) => ipcRenderer.invoke('toggle-user', payload),
  updatePhone:     (payload) => ipcRenderer.invoke('update-phone', payload),
  assignLicense:   (payload) => ipcRenderer.invoke('assign-license', payload),
  previewUsername: (payload) => ipcRenderer.invoke('preview-username', payload),

  // Registro
  getLog:      ()  => ipcRenderer.invoke('get-log'),
  downloadLog: ()  => ipcRenderer.invoke('download-log'),
  clearLog:    ()  => ipcRenderer.invoke('clear-log'),

  // Importación CSV
  openCsvDialog:       ()          => ipcRenderer.invoke('open-csv-dialog'),
  downloadCsvTemplate: ()          => ipcRenderer.invoke('download-csv-template'),
  processCsv:          (payload)   => ipcRenderer.invoke('process-csv', payload),
});
