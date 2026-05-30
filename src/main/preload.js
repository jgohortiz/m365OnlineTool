const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  // Bloqueo / desbloqueo
  lockCheckExists:  ()         => ipcRenderer.invoke('lock-check-exists'),
  lockTryUnlock:    (pass)     => ipcRenderer.invoke('lock-try-unlock', pass),
  lockInitConf:     (payload)  => ipcRenderer.invoke('lock-init-conf', payload),
  lockMigrateConf:  (payload)  => ipcRenderer.invoke('lock-migrate-conf', payload),
  lockExit:             ()         => ipcRenderer.invoke('lock-exit'),
  changeMasterPassword: (payload)  => ipcRenderer.invoke('change-master-password', payload),

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

  // Eliminar
  deleteUser:          (payload)   => ipcRenderer.invoke('delete-user', payload),

  // Notificaciones
  listTemplates:            ()         => ipcRenderer.invoke('list-templates'),
  readTemplate:             (fn)       => ipcRenderer.invoke('read-template', fn),
  openTemplatesFolder:      ()         => ipcRenderer.invoke('open-templates-folder'),
  openNotifCsv:             ()         => ipcRenderer.invoke('open-notif-csv'),
  downloadNotifTemplate:    ()         => ipcRenderer.invoke('download-notif-template'),
  downloadNotifCsvTemplate: ()         => ipcRenderer.invoke('download-notif-csv-template'),
  sendNotifications:        (payload)  => ipcRenderer.invoke('send-notifications', payload),
  onNotifProgress:          (cb)       => ipcRenderer.on('notif-progress', (_, data) => cb(data)),
});
