// Suprimir advertencia de deprecación de 'punycode' generada por dependencias transitivas
process.removeAllListeners('warning');
process.on('warning', w => { if (w.name === 'DeprecationWarning' && w.message.includes('punycode')) return; process.stderr.write(w.stack + '\n'); });

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs   = require('fs');

// ─── nodemailer (requerido para notificaciones) ───────────────────────────────
let nodemailer;
try {
  nodemailer = require('nodemailer');
} catch (_) {
  nodemailer = null; // Se instala con: npm install
}
// ─── Archivo de configuración cifrado m365OnlineTool.conf ───────────────────
const crypto = require('crypto');

const CONF_PATH = path.join(
  app.isPackaged
    ? path.dirname(process.execPath)
    : path.resolve(__dirname, '../../'),
  'm365OnlineTool.conf'
);

// Valor de control constante para verificar que la clave es correcta
const CONTROL_VALUE = 'm365OnlineTool:OK';
// Clave maestra en memoria (se establece al desbloquear)
let masterKey = null;

function deriveKey(password, salt) {
  return crypto.pbkdf2Sync(password, salt, 100000, 32, 'sha256');
}

function encrypt(text, key) {
  const iv   = crypto.randomBytes(16);
  const cipher = crypto.createCipheriv('aes-256-cbc', key, iv);
  const enc  = Buffer.concat([cipher.update(text, 'utf8'), cipher.final()]);
  return iv.toString('hex') + ':' + enc.toString('hex');
}

function decrypt(data, key) {
  const [ivHex, encHex] = data.split(':');
  const iv      = Buffer.from(ivHex, 'hex');
  const enc     = Buffer.from(encHex, 'hex');
  const decipher = crypto.createDecipheriv('aes-256-cbc', key, iv);
  return Buffer.concat([decipher.update(enc), decipher.final()]).toString('utf8');
}

function confExists() { return fs.existsSync(CONF_PATH); }

function readConf() {
  if (!masterKey) return {};
  try {
    const raw  = JSON.parse(fs.readFileSync(CONF_PATH, 'utf8'));
    const plain = decrypt(raw.data, masterKey);
    return JSON.parse(plain);
  } catch { return {}; }
}

function writeConf(data) {
  if (!masterKey) throw new Error('No hay clave maestra establecida.');
  const salt    = confExists()
    ? JSON.parse(fs.readFileSync(CONF_PATH, 'utf8')).salt
    : crypto.randomBytes(16).toString('hex');
  const plain   = JSON.stringify({ ...data, __control: CONTROL_VALUE }, null, 2);
  const encoded = encrypt(plain, masterKey);
  fs.writeFileSync(CONF_PATH, JSON.stringify({ salt, data: encoded }), 'utf8');
}

// Intentar desbloquear con la clave dada; retorna 'ok', 'wrong', 'new'
function tryUnlock(password) {
  if (!confExists()) return 'new'; // archivo no existe
  try {
    const raw = JSON.parse(fs.readFileSync(CONF_PATH, 'utf8'));
    // Detectar archivo sin cifrar (formato antiguo: no tiene campo 'salt')
    if (!raw.salt || !raw.data) return 'migrate';
    const key    = deriveKey(password, raw.salt);
    const plain  = decrypt(raw.data, key);
    const parsed = JSON.parse(plain);
    if (parsed.__control !== CONTROL_VALUE) return 'wrong';
    masterKey = key;
    return 'ok';
  } catch { return 'wrong'; }
}

// Migrar conf sin cifrar: leer datos, cifrar con nueva clave
function migrateConf(password) {
  try {
    const old = JSON.parse(fs.readFileSync(CONF_PATH, 'utf8'));
    // Guardar backup
    fs.writeFileSync(CONF_PATH + '.bak', JSON.stringify(old, null, 2), 'utf8');
    const salt = crypto.randomBytes(16).toString('hex');
    masterKey  = deriveKey(password, salt);
    const plain  = JSON.stringify({ ...old, __control: CONTROL_VALUE });
    const encoded = encrypt(plain, masterKey);
    fs.writeFileSync(CONF_PATH, JSON.stringify({ salt, data: encoded }));
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Crear nueva conf con nueva clave
function initConf(password) {
  const salt = crypto.randomBytes(16).toString('hex');
  masterKey  = deriveKey(password, salt);
  const plain  = JSON.stringify({ __control: CONTROL_VALUE });
  const encoded = encrypt(plain, masterKey);
  fs.writeFileSync(CONF_PATH, JSON.stringify({ salt, data: encoded }));
}

let mainWindow;
let lockWindow;

// ─── Ventana principal ────────────────────────────────────────────────────────
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280, height: 820, minWidth: 1100, minHeight: 700,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    frame: false,
    titleBarStyle: 'hidden',
    backgroundColor: '#ffffff',
    icon: path.join(__dirname, '../../resources/icon.ico'),
    show: false,
  });
  mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'));
  mainWindow.once('ready-to-show', () => {
    mainWindow.maximize();
    mainWindow.show();
  });
  mainWindow.on('closed', () => { mainWindow = null; });
}

// ─── Ventana de bloqueo / desbloqueo ─────────────────────────────────────────
function createLockWindow() {
  lockWindow = new BrowserWindow({
    width: 420, height: 460, resizable: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    frame: false,
    backgroundColor: '#0078d4',
    icon: path.join(__dirname, '../../resources/icon.ico'),
    show: false,
    center: true,
  });
  const mode = confExists()
    ? ((() => { try { const r=JSON.parse(fs.readFileSync(CONF_PATH,'utf8')); return (r.salt&&r.data)?'unlock':'migrate'; } catch { return 'migrate'; } })())
    : 'new';
  lockWindow.loadFile(path.join(__dirname, '../renderer/lock.html'), { hash: mode });
  lockWindow.once('ready-to-show', () => lockWindow.show());
  lockWindow.on('closed', () => { lockWindow = null; });
}

app.whenReady().then(() => {
  createLockWindow();
});
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
app.on('activate', () => { if (!mainWindow && !lockWindow) createLockWindow(); });

// ─── Controles de ventana ─────────────────────────────────────────────────────
ipcMain.on('window-minimize', () => mainWindow.minimize());
ipcMain.on('window-maximize', () => mainWindow.isMaximized() ? mainWindow.unmaximize() : mainWindow.maximize());
ipcMain.on('window-close',    () => mainWindow.close());

// ─── Handlers de bloqueo ─────────────────────────────────────────────────────
ipcMain.handle('lock-check-exists', () => confExists());

ipcMain.handle('lock-try-unlock', (_, password) => {
  const result = tryUnlock(password);
  if (result === 'ok') {
    appendLog('ACCESO | Sesión iniciada correctamente');
    // Abrir ventana principal y cerrar lock
    createWindow();
    if (lockWindow) { lockWindow.close(); lockWindow = null; }
  } else if (result === 'wrong') {
    appendLog('ACCESO FALLIDO | Clave incorrecta');
  }
  return result;
});

ipcMain.handle('lock-init-conf', (_, { password }) => {
  try {
    initConf(password);
    appendLog('CONFIGURACIÓN INICIALIZADA | Archivo cifrado creado');
    createWindow();
    if (lockWindow) { lockWindow.close(); lockWindow = null; }
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

ipcMain.handle('lock-exit', () => { app.quit(); });

ipcMain.handle('lock-migrate-conf', (_, { password }) => {
  const res = migrateConf(password);
  if (res.ok) {
    appendLog('MIGRACIÓN | Configuración existente cifrada con nueva clave maestra');
    createWindow();
    if (lockWindow) { lockWindow.close(); lockWindow = null; }
  }
  return res;
});

// ─── Configuración ────────────────────────────────────────────────────────────
ipcMain.handle('get-config',  ()      => { try { return readConf(); } catch { return {}; } });
ipcMain.handle('save-config', (_, cfg) => {
  try {
    writeConf(cfg);
    appendLog(`CONFIGURACIÓN GUARDADA | Dominio: ${cfg.defaultDomain||'-'} | SMTP: ${cfg.smtpHost||'-'} | Usuario SMTP: ${cfg.smtpUser||'-'}`);
    return true;
  } catch (e) {
    appendLog(`ERROR GUARDAR CONFIGURACIÓN | ${e.message}`);
    return false;
  }
});

// ─── Auxiliar Graph API ───────────────────────────────────────────────────────
async function getAccessToken(cfg) {
  const { ConfidentialClientApplication } = require('@azure/msal-node');
  const cca = new ConfidentialClientApplication({
    auth: {
      clientId:     cfg.clientId,
      clientSecret: cfg.clientSecret,
      authority:    `https://login.microsoftonline.com/${cfg.tenantId}`,
    },
  });
  const result = await cca.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'],
  });
  return result.accessToken;
}

async function graphRequest(token, method, endpoint, body = null) {
  const fetch = require('node-fetch');
  const opts = {
    method,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
  };
  if (body) opts.body = JSON.stringify(body);
  const res  = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, opts);
  if (res.status === 204) return { success: true };
  const data = await res.json();
  if (!res.ok) throw new Error(data.error?.message || `HTTP ${res.status}`);
  return data;
}

// ─── Generador de ID de lote ─────────────────────────────────────────────────
// Equivalente a GetProcessLoteUUID -Time
// Formato: LT-AAMMDD:HHmmss-xxxxxxxxx
function generateLoteUUID() {
  const now    = new Date();
  const pad    = n => String(n).padStart(2, '0');
  const yy     = String(now.getFullYear()).slice(-2);
  const mm     = pad(now.getMonth() + 1);
  const dd     = pad(now.getDate());
  const HH     = pad(now.getHours());
  const mi     = pad(now.getMinutes());
  const ss     = pad(now.getSeconds());
  const stamp  = `${yy}${mm}${dd}:${HH}${mi}${ss}`;
  // 9 caracteres aleatorios (hex) para el sufijo
  const suffix = [...Array(9)].map(() => Math.floor(Math.random() * 16).toString(16)).join('');
  return `LT-${stamp}-${suffix}`;
}

// ── Formateador de teléfono móvil — equivalente a SetPhoneMobileNumber ─────────
// Recorta a 10 dígitos (solo numéricos) y antepone el prefijo configurado.
function formatMobilePhone(rawPhone, phonePrefix) {
  const prefix = (phonePrefix || '+57').trim();
  // Dejar solo dígitos
  let digits = String(rawPhone || '').replace(/\D/g, '');
  // Recortar a 10 dígitos (comportamiento del script original)
  if (digits.length > 10) digits = digits.slice(0, 10);
  if (!digits) return null;
  return `${prefix} ${digits}`;
}

// ── JobTitle según tipo de cuenta ─────────────────────────────────────────────
function getJobTitle(tipo) {
  const t = String(tipo || '').trim().toLowerCase();
  if (t === 'maestro'    || t === 'm') return 'DOCENTE';
  if (t === 'estudiante' || t === 'e') return 'ESTUDIANTE';
  return 'OTRO';
}


// ── Generador de contraseña aleatoria para CSV (usa parámetros de configuración) ──
function generateCsvPassword(cfg) {
  const len     = parseInt(cfg.pwLength  || 12);
  const upper   = cfg.pwUpper   !== false;
  const lower   = cfg.pwLower   !== false;
  const numbers = cfg.pwNumbers !== false;
  const symbols = cfg.pwSymbols !== false;
  const UPPER   = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const LOWER   = 'abcdefghijklmnopqrstuvwxyz';
  const NUMS    = '0123456789';
  const SYMS    = '!@#$%&*+-=?';
  let pool = ''; let required = [];
  if (upper)   { pool += UPPER;   required.push(UPPER[Math.floor(Math.random()*UPPER.length)]); }
  if (lower)   { pool += LOWER;   required.push(LOWER[Math.floor(Math.random()*LOWER.length)]); }
  if (numbers) { pool += NUMS;    required.push(NUMS[Math.floor(Math.random()*NUMS.length)]); }
  if (symbols) { pool += SYMS;    required.push(SYMS[Math.floor(Math.random()*SYMS.length)]); }
  if (!pool) pool = LOWER + NUMS;
  let pw = [...required];
  while (pw.length < len) pw.push(pool[Math.floor(Math.random() * pool.length)]);
  for (let i = pw.length-1; i > 0; i--) { const j = Math.floor(Math.random()*(i+1)); [pw[i],pw[j]]=[pw[j],pw[i]]; }
  return pw.join('');
}

// ─── Generador de UPN ─────────────────────────────────────────────────────────

// Normaliza texto: quita tildes, caracteres especiales, espacios extra
// Equivalente a GetTextWithFormat -Replace -Clean -Trim -ToLower
function normalizeText(s) {
  return String(s || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')   // quitar diacríticos/tildes
    .replace(/[^a-z0-9\s]/gi, '')       // quitar caracteres especiales
    .replace(/\s+/g, ' ')               // colapsar espacios múltiples
    .trim()
    .toLowerCase();
}

// Prefijo según tipo de cuenta
function getTipoPrefix(tipo) {
  const t = String(tipo).trim().toLowerCase();
  if (t === 'maestro'    || t === 'm') return 'm';
  if (t === 'estudiante' || t === 'e') return 'e';
  return '';  // tipo "otro" — sin prefijo
}

// Construye un UPN candidato a partir de sus partes
function buildUPN(prefix, parts, domain) {
  const upn = (prefix ? prefix + '.' : '') + parts.join('') + '@' + domain;
  return upn.toLowerCase();
}

/**
 * Genera el UPN con lógica en cascada (3 métodos), verificando disponibilidad
 * en Microsoft Graph antes de pasar al siguiente método.
 *
 * Método 1: prefijo.L1Apellido1            → m.jdoe
 * Método 2: prefijo.L1L2Apellido1          → m.jmdoe       (si hay 2+ nombres)
 * Método 3: prefijo.L1Apellido1L2Apellido2 → m.jdoes       (si hay 2+ apellidos)
 * Fallback:  prefijo.L1Apellido1+timestamp  → m.jdoe1704...
 *
 * Si el UPN ya existe Y el displayName coincide con el usuario actual,
 * se asume que es el mismo usuario y se devuelve ese UPN sin avanzar.
 *
 * @param {string} token        - Access token de Graph API
 * @param {string} tipo         - maestro | estudiante | otro
 * @param {string} nombres      - Nombres completos del usuario
 * @param {string} apellidos    - Apellidos completos del usuario
 * @param {string} displayName  - Nombre para mostrar (para detectar duplicado propio)
 * @param {string} domain       - Dominio sin @ (ej: escuela.org)
 * @returns {Promise<{upn: string, metodo: number, log: string[]}>}
 */
async function generateUPN(token, tipo, nombres, apellidos, displayName, domain) {
  const log = [];
  const prefix     = getTipoPrefix(tipo);
  const nombresN   = normalizeText(nombres);
  const apellidosN = normalizeText(apellidos);
  const displayN   = String(displayName || '').trim().toUpperCase();

  const partesNombres   = nombresN.split(' ').filter(Boolean);
  const partesApellidos = apellidosN.split(' ').filter(Boolean);

  // ── Método 1: L(primer nombre) + primer apellido ──────────────────────────
  const l1 = partesNombres[0]?.[0] || '';
  const a1  = partesApellidos[0]   || '';
  const upn1 = buildUPN(prefix, [l1, a1], domain);
  log.push(`Método 1: ${upn1}`);

  const exist1 = await checkUPNExists(token, upn1);
  if (!exist1) {
    log.push(`Disponible: ${upn1}`);
    return { upn: upn1, metodo: 1, log };
  }
  // Verificar si el existente es el mismo usuario (mismo displayName)
  const owner1 = await getUPNDisplayName(token, upn1);
  if (owner1.toUpperCase() === displayN) {
    log.push(`UPN ${upn1} pertenece al mismo usuario — reutilizando`);
    return { upn: upn1, metodo: 1, log };
  }
  log.push(`En uso por otro usuario: ${upn1}`);

  // ── Método 2: L(primer nombre) + L(segundo nombre) + primer apellido ──────
  if (partesNombres.length >= 2) {
    const l2   = partesNombres[1][0];
    const upn2 = buildUPN(prefix, [l1, l2, a1], domain);
    log.push(`Método 2: ${upn2}`);

    const exist2 = await checkUPNExists(token, upn2);
    if (!exist2) {
      log.push(`Disponible: ${upn2}`);
      return { upn: upn2, metodo: 2, log };
    }
    const owner2 = await getUPNDisplayName(token, upn2);
    if (owner2.toUpperCase() === displayN) {
      log.push(`UPN ${upn2} pertenece al mismo usuario — reutilizando`);
      return { upn: upn2, metodo: 2, log };
    }
    log.push(`En uso por otro usuario: ${upn2}`);
  } else {
    log.push('Método 2 omitido — solo hay un nombre');
  }

  // ── Método 3: L(primer nombre) + primer apellido + L(segundo apellido) ────
  if (partesApellidos.length >= 2) {
    const la2  = partesApellidos[1][0];
    const upn3 = buildUPN(prefix, [l1, a1, la2], domain);
    log.push(`Método 3: ${upn3}`);

    const exist3 = await checkUPNExists(token, upn3);
    if (!exist3) {
      log.push(`Disponible: ${upn3}`);
      return { upn: upn3, metodo: 3, log };
    }
    const owner3 = await getUPNDisplayName(token, upn3);
    if (owner3.toUpperCase() === displayN) {
      log.push(`UPN ${upn3} pertenece al mismo usuario — reutilizando`);
      return { upn: upn3, metodo: 3, log };
    }
    log.push(`En uso por otro usuario: ${upn3}`);
  } else {
    log.push('Método 3 omitido — solo hay un apellido');
  }

  // ── Fallback: método 1 + sufijo numérico incremental ─────────────────────
  log.push('Aplicando fallback con sufijo numérico');
  for (let i = 2; i <= 99; i++) {
    const upnF = buildUPN(prefix, [l1, a1, String(i)], domain);
    const existF = await checkUPNExists(token, upnF);
    if (!existF) {
      log.push(`Fallback disponible: ${upnF}`);
      return { upn: upnF, metodo: 0, log };
    }
  }

  // Último recurso: timestamp
  const upnTs = buildUPN(prefix, [l1, a1, Date.now()], domain);
  log.push(`Último recurso (timestamp): ${upnTs}`);
  return { upn: upnTs, metodo: -1, log };
}

// Verifica si un UPN ya existe en el tenant
/**
 * checkMailboxExists — Equivalente a ChekAccountExchangeOnlineExist
 * Verifica si el UPN tiene un buzón activo en Exchange Online consultando
 * el recurso /users/{upn}/mailboxSettings a través de Graph API.
 * Retorna { exists: bool, displayName: string, accountEnabled: bool }
 */
async function checkMailboxExists(token, upn) {
  try {
    const enc  = encodeURIComponent(upn);
    const data = await graphRequest(token, 'GET',
      `/users/${enc}?$select=id,displayName,accountEnabled,assignedLicenses,mail,userPrincipalName`
    );
    // Si el usuario existe en Entra ID y tiene licencia con Exchange, el buzón existe
    const hasExchangeLicense = (data.assignedLicenses || []).length > 0;
    return {
      exists:         true,
      hasMailbox:     hasExchangeLicense,
      displayName:    data.displayName || '',
      accountEnabled: data.accountEnabled ?? true,
      mail:           data.mail || upn,
      id:             data.id || '',
    };
  } catch (e) {
    // 404 = usuario no existe
    return { exists: false, hasMailbox: false, displayName: '', accountEnabled: false, mail: '', id: '' };
  }
}

/**
 * checkDisplayNameAvailability — Equivalente a ChekDisplayNameAvailability
 * Verifica si el displayName o la combinación givenName+surname ya están en uso
 * por un usuario diferente al que se está procesando.
 *
 * Retorna { available: bool, conflictUser: string|null }
 *   available    = true  → el nombre está libre o pertenece al mismo usuario
 *   conflictUser = UPN del usuario que ya usa ese nombre (si hay conflicto)
 */
async function checkDisplayNameAvailability(token, givenName, surname, displayName) {
  try {
    const dnNorm = String(displayName || '').trim().replace(/\s+/g, ' ').toUpperCase();
    const gnNorm = String(givenName   || '').trim().toUpperCase();
    const snNorm = String(surname     || '').trim().toUpperCase();

    // Buscar por displayName exacto
    const encDN = encodeURIComponent(displayName.trim());
    const byDN  = await graphRequest(token, 'GET',
      `/users?$filter=displayName eq '${encDN}'&$select=id,displayName,givenName,surname,userPrincipalName&$top=5`
    ).catch(() => ({ value: [] }));

    for (const u of (byDN.value || [])) {
      const existDN = String(u.displayName || '').trim().replace(/\s+/g, ' ').toUpperCase();
      if (existDN === dnNorm) {
        // Mismo displayName → conflicto con otro usuario
        return { available: false, conflictUser: u.userPrincipalName, conflictName: u.displayName };
      }
    }

    // Buscar por startsWith de givenName + surname (equivale al Like de PS)
    const encGN = encodeURIComponent(givenName.trim());
    const encSN = encodeURIComponent(surname.trim());
    const byName = await graphRequest(token, 'GET',
      `/users?$filter=startsWith(givenName,'${encGN}') and startsWith(surname,'${encSN}')&$select=id,displayName,givenName,surname,userPrincipalName&$top=10`
    ).catch(() => ({ value: [] }));

    for (const u of (byName.value || [])) {
      const existDN = String(u.displayName || '').trim().replace(/\s+/g, ' ').toUpperCase();
      if (existDN !== dnNorm) {
        // Mismo nombre/apellido pero distinto displayName → posible duplicado
        return { available: false, conflictUser: u.userPrincipalName, conflictName: u.displayName };
      }
    }

    return { available: true, conflictUser: null, conflictName: null };
  } catch {
    // Si falla la consulta, asumimos disponible para no bloquear la creación
    return { available: true, conflictUser: null, conflictName: null };
  }
}

// checkUPNExists — versión simplificada usada internamente por generateUPN
async function checkUPNExists(token, upn) {
  const result = await checkMailboxExists(token, upn);
  return result.exists;
}

// getUPNDisplayName — obtiene el displayName de un UPN existente
async function getUPNDisplayName(token, upn) {
  const result = await checkMailboxExists(token, upn);
  return result.displayName;
}

// Vista previa sin verificación de disponibilidad (para el formulario)
function generateUsernameFast(tipo, nombres, apellidos, domain) {
  const prefix  = getTipoPrefix(tipo);
  const nN      = normalizeText(nombres).split(' ').filter(Boolean);
  const aN      = normalizeText(apellidos).split(' ').filter(Boolean);
  const l1      = nN[0]?.[0] || '';
  const a1      = aN[0]      || '';
  return buildUPN(prefix, [l1, a1], domain);
}

ipcMain.handle('preview-username', (_, { tipo, firstName, lastName, domain }) =>
  generateUsernameFast(tipo, firstName, lastName, domain)
);

// ─── Registro de actividad ────────────────────────────────────────────────────
const LOG_PATH = path.join(app.getPath('userData'), 'activity.log');

function appendLog(entry) {
  fs.appendFileSync(LOG_PATH, `[${new Date().toISOString()}] ${entry}\n`);
}

ipcMain.handle('get-log', () => fs.existsSync(LOG_PATH) ? fs.readFileSync(LOG_PATH, 'utf8') : '');

ipcMain.handle('download-log', async () => {
  const { filePath } = await dialog.showSaveDialog(mainWindow, {
    title: 'Guardar registro de actividad',
    defaultPath: `registro-m365OnlineTool-${Date.now()}.txt`,
    filters: [{ name: 'Archivo de texto', extensions: ['txt', 'log'] }],
  });
  if (filePath) {
    fs.copyFileSync(LOG_PATH, filePath);
    appendLog(`REGISTRO DESCARGADO | ${filePath}`);
    return { saved: true, path: filePath };
  }
  return { saved: false };
});

ipcMain.handle('clear-log', () => { fs.writeFileSync(LOG_PATH, ''); return true; });

// ─── Licencias ────────────────────────────────────────────────────────────────
ipcMain.handle('get-licenses', async (_, cfg) => {
  try {
    const token = await getAccessToken(cfg);
    const data  = await graphRequest(token, 'GET', '/subscribedSkus');
    appendLog(`LICENCIAS CARGADAS | ${data.value.length} SKU(s) encontrados`);
    return { ok: true, data: data.value };
  } catch (e) {
    appendLog(`ERROR LICENCIAS | ${e.message}`);
    return { ok: false, error: e.message };
  }
});

// ─── Listar usuarios ──────────────────────────────────────────────────────────
ipcMain.handle('list-users', async (_, cfg) => {
  try {
    const token    = await getAccessToken(cfg);
    const allUsers = [];

    // Intentar primero con signInActivity (requiere AuditLog.Read.All + Azure AD Premium P1/P2)
    // Si falla (403/400) se reintenta sin ese campo
    const selectBase    = 'id,displayName,userPrincipalName,accountEnabled,mobilePhone,assignedLicenses,createdDateTime,onPremisesExtensionAttributes';
    const selectFull    = selectBase; // sin signInActivity (requiere Premium P1/P2)
    async function fetchAllPages(selectFields) {
      const pages = [];
      let url = `/users?$select=${selectFields}&$top=100`;
      while (url) {
        const data = await graphRequest(token, 'GET', url);
        if (data.value) pages.push(...data.value);
        const next = data['@odata.nextLink'];
        url = next ? next.replace('https://graph.microsoft.com/v1.0', '') : null;
      }
      return pages;
    }

    const users = await fetchAllPages(selectFull);
    allUsers.push(...users);

    appendLog(`USUARIOS CARGADOS | ${allUsers.length} usuario(s)`);
    return { ok: true, data: allUsers };
  } catch (e) {
    appendLog(`ERROR CARGAR USUARIOS | ${e.message}`);
    return { ok: false, error: e.message };
  }
});

// ─── Buscar usuario por UPN ───────────────────────────────────────────────────
async function findUserByUPN(token, upn) {
  try {
    const enc  = encodeURIComponent(upn);
    const data = await graphRequest(token, 'GET',
      `/users?$filter=userPrincipalName eq '${enc}'&$select=id,displayName,givenName,surname,userPrincipalName,accountEnabled,mobilePhone,assignedLicenses`
    );
    return data.value && data.value.length > 0 ? data.value[0] : null;
  } catch { return null; }
}

// ─── Crear usuario ────────────────────────────────────────────────────────────
ipcMain.handle('create-user', async (_, { cfg, userData }) => {
  try {
    const token = await getAccessToken(cfg);
    const displayName = `${userData.firstName} ${userData.lastName}`.trim().toUpperCase();

    // ── Validación 1: disponibilidad del displayName / nombre completo ────────
    const dnCheck = await checkDisplayNameAvailability(
      token, userData.firstName, userData.lastName, displayName
    );
    if (!dnCheck.available) {
      const msg = `El nombre completo "${displayName}" ya está en uso por ${dnCheck.conflictUser} ("${dnCheck.conflictName}"). No se puede crear la cuenta.`;
      appendLog(`BLOQUEADO NOMBRE | ${msg}`);
      return { ok: false, error: msg };
    }

    // ── Generación de UPN con verificación de buzón en cascada ───────────────
    const { upn, metodo, log: upnLog } = await generateUPN(
      token, userData.tipo, userData.firstName, userData.lastName, displayName, userData.domain
    );
    upnLog.forEach(msg => appendLog(`UPN | ${msg}`));
    appendLog(`UPN seleccionado método ${metodo}: ${upn}`);

    // ── Validación 2: verificar que el buzón/cuenta no exista ya ─────────────
    const mailCheck = await checkMailboxExists(token, upn);
    if (mailCheck.exists) {
      appendLog(`CUENTA EXISTENTE | ${upn} | ${mailCheck.displayName} | Activa: ${mailCheck.accountEnabled}`);
      return {
        ok: false,
        error: `La cuenta ${upn} ya existe (${mailCheck.displayName}). ${!mailCheck.accountEnabled ? 'Está inactiva — puede reactivarla desde Usuarios.' : ''}`.trim(),
      };
    }

    // ── Generar ID de lote para esta operación ───────────────────────────────
    const loteId = generateLoteUUID();
    appendLog(`LOTE | ${loteId} | ${upn}`);

    const formattedPhone = formatMobilePhone(userData.phone, cfg.phonePrefix);

    // Nombres y apellidos siempre en mayúsculas
    const fnUpper = String(userData.firstName || '').trim().toUpperCase();
    const lnUpper = String(userData.lastName  || '').trim().toUpperCase();

    const created = await graphRequest(token, 'POST', '/users', {
      accountEnabled:    true,
      displayName:       `${fnUpper} ${lnUpper}`,
      givenName:         fnUpper,
      surname:           lnUpper,
      userPrincipalName: upn,
      mailNickname:      upn.split('@')[0],
      jobTitle:          getJobTitle(userData.tipo),
      passwordProfile:   { forceChangePasswordNextSignIn: true, password: userData.password },
      // mobilePhone (Teléfono móvil): vacío — solo se usa en autenticación MFA
      mobilePhone:       null,
      // businessPhones (Teléfono de oficina): valor de phoneNumber en configuración
      businessPhones:    cfg.phoneNumber ? [cfg.phoneNumber] : [],
      // Campos organizacionales desde configuración
      city:              cfg.city          || null,
      country:           cfg.country       || null,
      department:        cfg.department    || null,
      officeLocation:    cfg.office        || null,
      postalCode:        cfg.postalCode    || null,
      state:             cfg.state         || null,
      streetAddress:     cfg.streetAddress || null,
      usageLocation:     cfg.usageLocation || null,
      // Atributos personalizados (CustomAttribute en Exchange)
      onPremisesExtensionAttributes: {
        extensionAttribute1: loteId,
        extensionAttribute2: userData.cedula       || null,
        extensionAttribute3: userData.mailPersonal || null,
      },
    });

    if (userData.skuId) {
      await graphRequest(token, 'POST', `/users/${created.id}/assignLicense`, {
        addLicenses: [{ skuId: userData.skuId }], removeLicenses: [],
      });
    }
    // StrongAuthenticationMethods: SMS (OneWaySMS) como método predeterminado
    // Equivalente a: $SMS.MethodType = "OneWaySMS"; $SMS.IsDefault = $true
    if (formattedPhone) {
      try {
        await graphRequest(token, 'POST', `/users/${created.id}/authentication/phoneMethods`,
          { phoneNumber: formattedPhone, phoneType: 'mobile' }
        );
        // Registrar también para SMS (OneWaySMS = método predeterminado MFA)
        await graphRequest(token, 'POST', `/users/${created.id}/authentication/phoneMethods`,
          { phoneNumber: formattedPhone, phoneType: 'alternateMobile' }
        );
      } catch (_) { /* no crítico — el buzón puede tardar en provisionarse */ }
    }

    appendLog(`CREAR | ${upn} | Lote: ${loteId} | Tipo: ${userData.tipo} | Cédula: ${userData.cedula || '-'} | Tel: ${userData.phone || '-'} | Licencia: ${userData.skuId || 'ninguna'}`);
    return { ok: true, user: created, upn };
  } catch (e) {
    appendLog(`ERROR AL CREAR | ${userData.firstName} ${userData.lastName} | ${e.message}`);
    return { ok: false, error: e.message };
  }
});

// ─── Activar / Inactivar cuenta ───────────────────────────────────────────────
ipcMain.handle('toggle-user', async (_, { cfg, userId, enabled, upn }) => {
  try {
    const token = await getAccessToken(cfg);
    await graphRequest(token, 'PATCH', `/users/${userId}`, { accountEnabled: enabled });
    appendLog(`${enabled ? 'ACTIVAR' : 'INACTIVAR'} | ${upn}`);
    return { ok: true };
  } catch (e) {
    appendLog(`ERROR AL CAMBIAR ESTADO | ${upn} | ${e.message}`);
    return { ok: false, error: e.message };
  }
});

// ─── Actualizar teléfono ──────────────────────────────────────────────────────
// mobilePhone → null (Teléfono móvil vacío)
// businessPhones → phoneNumber de configuración (Teléfono de oficina)
// authentication/phoneMethods → celular formateado (solo MFA)
ipcMain.handle('update-phone', async (_, { cfg, userId, phone, upn }) => {
  try {
    const token = await getAccessToken(cfg);
    const formattedPhone = formatMobilePhone(phone, cfg.phonePrefix);

    // Información de contacto: móvil vacío, oficina = config
    await graphRequest(token, 'PATCH', `/users/${userId}`, {
      mobilePhone:   null,
      businessPhones: cfg.phoneNumber ? [cfg.phoneNumber] : [],
    });

    // Métodos de autenticación MFA: celular del usuario
    if (formattedPhone) {
      try {
        // Intentar actualizar si ya existe, o crear si no
        const existing = await graphRequest(token, 'GET',
          `/users/${userId}/authentication/phoneMethods`
        ).catch(() => ({ value: [] }));
        const mobileMethod = (existing.value || []).find(m => m.phoneType === 'mobile');
        if (mobileMethod) {
          await graphRequest(token, 'PUT',
            `/users/${userId}/authentication/phoneMethods/${mobileMethod.id}`,
            { phoneNumber: formattedPhone, phoneType: 'mobile' }
          );
        } else {
          await graphRequest(token, 'POST',
            `/users/${userId}/authentication/phoneMethods`,
            { phoneNumber: formattedPhone, phoneType: 'mobile' }
          );
        }
      } catch (_) { /* no crítico */ }
    }

    appendLog(`ACTUALIZAR TELÉFONO | ${upn} | MFA: ${formattedPhone || phone} | Oficina: ${cfg.phoneNumber || '-'}`);
    return { ok: true };
  } catch (e) {
    appendLog(`ERROR ACTUALIZAR TELÉFONO | ${upn} | ${e.message}`);
    return { ok: false, error: e.message };
  }
});

// ─── Asignar licencia ─────────────────────────────────────────────────────────
ipcMain.handle('assign-license', async (_, { cfg, userId, skuId, upn }) => {
  try {
    const token = await getAccessToken(cfg);
    await graphRequest(token, 'POST', `/users/${userId}/assignLicense`, {
      addLicenses: [{ skuId }], removeLicenses: [],
    });
    appendLog(`ASIGNAR LICENCIA | ${upn} | SKU: ${skuId}`);
    return { ok: true };
  } catch (e) { return { ok: false, error: e.message }; }
});


// ─── Eliminar usuario ─────────────────────────────────────────────────────────
ipcMain.handle('delete-user', async (_, { cfg, userId, upn }) => {
  try {
    const token = await getAccessToken(cfg);
    await graphRequest(token, 'DELETE', `/users/${userId}`);
    appendLog(`ELIMINAR | ${upn}`);
    return { ok: true };
  } catch (e) {
    appendLog(`ERROR AL ELIMINAR | ${upn} | ${e.message}`);
    return { ok: false, error: e.message };
  }
});

// ─── Seleccionar archivo CSV ──────────────────────────────────────────────────
ipcMain.handle('open-csv-dialog', async () => {
  const { filePaths } = await dialog.showOpenDialog(mainWindow, {
    title: 'Seleccionar archivo CSV',
    filters: [{ name: 'Archivo CSV', extensions: ['csv'] }],
    properties: ['openFile'],
  });
  if (!filePaths || filePaths.length === 0) return { ok: false };
  try {
    const content = fs.readFileSync(filePaths[0], 'utf8');
    return { ok: true, content, filePath: filePaths[0] };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

// ─── Descargar plantilla CSV ──────────────────────────────────────────────────
ipcMain.handle('download-csv-template', async () => {
  const { filePath } = await dialog.showSaveDialog(mainWindow, {
    title: 'Guardar plantilla CSV',
    defaultPath: 'plantilla-usuarios-m365.csv',
    filters: [{ name: 'Archivo CSV', extensions: ['csv'] }],
  });
  if (!filePath) return { saved: false };

  const header  = 'accion,tipo,nombres,apellidos,telefono,cuenta_institucional,cedula,mailpersonal,observacion\n';
  const ejemplo = [
    'crear,maestro,John,Doe,+57 3001234567,,10000001,john.doe@correo.com,Ejemplo maestro',
    'crear,estudiante,Jane,Doe,+57 3109876543,,10000002,jane.doe@correo.com,Ejemplo estudiante',
    'crear,otro,Admin,Doe,+57 3201111111,,10000003,admin.doe@correo.com,Ejemplo otro',
    'editar,,,,+57 3001111111,m.jdoe@escuela.org,10000001,,Actualizar teléfono',
    'desactivar,,,,, e.jdoe@escuela.org,,,',
    'activar,,,,,adoe@escuela.org,,,',
    'eliminar,,,,,m.jdoe@escuela.org,,,',
  ].join('\n');

  fs.writeFileSync(filePath, '\uFEFF' + header + ejemplo, 'utf8'); // BOM para Excel
  appendLog(`PLANTILLA CSV DESCARGADA | ${filePath}`);
  return { saved: true, path: filePath };
});

// ─── Procesar importación CSV ─────────────────────────────────────────────────
ipcMain.handle('process-csv', async (_, { cfg, rows, skuId: defaultSku }) => {
  const results = [];
  appendLog(`CSV INICIO | ${rows.length} fila(s) a procesar`);
  let token;
  try { token = await getAccessToken(cfg); }
  catch (e) {
    appendLog(`ERROR AUTENTICACIÓN CSV | ${e.message}`);
    return { ok: false, error: `Error de autenticación: ${e.message}` };
  }

  for (const row of rows) {
    const accion = (row.accion || '').trim().toLowerCase();
    const result = {
      fila: row._fila, accion,
      lote: '', cedula: '', nombres: '', apellidos: '',
      mailpersonal: '', alias: '', celular: '',
      estado: '', upn: '', displayName: '',
      tipo: '', id: '', creado: '', accountEnabled: null,
      observacion: (row.observacion || '').trim(),
    };

    try {
      // ── CREAR ──────────────────────────────────────────────────────────────
      if (accion === 'crear') {
        const fn     = (row.nombres   || '').trim();
        const ln     = (row.apellidos || '').trim();
        const cedula      = (row.cedula       || '').trim();
        const mailPersonal= (row.mailpersonal || '').trim();
        const observacion = (row.observacion  || '').trim();
        const domain = (cfg.defaultDomain || '').trim().replace(/^@/, '');
        const tipo   = (row.tipo          || 'otro').trim().toLowerCase();
        const phone  = (row.telefono      || '').trim();
        // Contraseña siempre aleatoria — generada en el proceso
        const pass   = generateCsvPassword(cfg);
        // SKU: fila > SKU predeterminado del tipo (config) > SKU predeterminado global
        const skuFromTipo = tipo === 'maestro'    ? (cfg.skuMaestro    || '')
                          : tipo === 'estudiante' ? (cfg.skuEstudiante || '')
                          : (cfg.skuOtro || '');
        const sku    = (skuFromTipo || defaultSku || '').trim();

        if (!fn || !ln) throw new Error('Faltan: nombres o apellidos');
        if (!domain)   throw new Error('Configure el Dominio predeterminado en Configuración antes de importar.');

        const displayNameCsv = `${fn} ${ln}`.trim().toUpperCase();

        // ── Validación 1: disponibilidad del nombre completo ─────────────────
        const dnCheckCsv = await checkDisplayNameAvailability(token, fn, ln, displayNameCsv);
        if (!dnCheckCsv.available) {
          const msgDN = `El nombre completo "${displayNameCsv}" ya está en uso por ${dnCheckCsv.conflictUser} ("${dnCheckCsv.conflictName}")`;
          appendLog(`CSV BLOQUEADO NOMBRE | ${msgDN}`);
          throw new Error(msgDN);
        }

        // ── Generación de UPN con verificación de buzón en cascada ───────────
        const { upn, metodo: metodoUpn, log: upnLogCsv } = await generateUPN(
          token, tipo, fn, ln, displayNameCsv, domain
        );
        upnLogCsv.forEach(msg => appendLog(`UPN CSV | ${msg}`));
        appendLog(`UPN seleccionado método ${metodoUpn}: ${upn}`);

        // ── Validación 2: cuenta/buzón ya existe ─────────────────────────────
        const mailCheckCsv = await checkMailboxExists(token, upn);
        if (mailCheckCsv.exists) {
          appendLog(`CSV CUENTA EXISTENTE | ${upn} | ${mailCheckCsv.displayName}`);
          throw new Error(`La cuenta ${upn} ya existe (${mailCheckCsv.displayName})${!mailCheckCsv.accountEnabled ? ' — inactiva' : ''}`);
        }

        result.upn = upn;

        // ── Generar ID de lote para esta fila ────────────────────────────────
        const loteIdCsv = generateLoteUUID();
        appendLog(`CSV LOTE | ${loteIdCsv} | ${upn}`);

        const formattedPhoneCsv = formatMobilePhone(phone, cfg.phonePrefix);

        // Nombres y apellidos siempre en mayúsculas
        const fnUpper = String(fn || '').trim().toUpperCase();
        const lnUpper = String(ln || '').trim().toUpperCase();

        const created = await graphRequest(token, 'POST', '/users', {
          accountEnabled:    true,
          displayName:       `${fnUpper} ${lnUpper}`,
          givenName:         fnUpper,
          surname:           lnUpper,
          userPrincipalName: upn,
          mailNickname:      upn.split('@')[0],
          jobTitle:          getJobTitle(tipo),
          passwordProfile:   { forceChangePasswordNextSignIn: true, password: pass },
          // mobilePhone (Teléfono móvil): vacío — solo se usa en autenticación MFA
          mobilePhone:       null,
          // businessPhones (Teléfono de oficina): valor de phoneNumber en configuración
          businessPhones:    cfg.phoneNumber ? [cfg.phoneNumber] : [],
          // Campos organizacionales desde configuración
          city:              cfg.city          || null,
          country:           cfg.country       || null,
          department:        cfg.department    || null,
          officeLocation:    cfg.office        || null,
          postalCode:        cfg.postalCode    || null,
          state:             cfg.state         || null,
          streetAddress:     cfg.streetAddress || null,
          usageLocation:     cfg.usageLocation || null,
          // Atributos personalizados (CustomAttribute en Exchange)
          onPremisesExtensionAttributes: {
            extensionAttribute1: loteIdCsv,
            extensionAttribute2: cedula       || null,
            extensionAttribute3: mailPersonal || null,
          },
        });

        if (sku) {
          await graphRequest(token, 'POST', `/users/${created.id}/assignLicense`, {
            addLicenses: [{ skuId: sku }], removeLicenses: [],
          });
        }
        // StrongAuthenticationMethods: SMS (OneWaySMS) predeterminado
        if (formattedPhoneCsv) {
          try {
            await graphRequest(token, 'POST', `/users/${created.id}/authentication/phoneMethods`,
              { phoneNumber: formattedPhoneCsv, phoneType: 'mobile' }
            );
            await graphRequest(token, 'POST', `/users/${created.id}/authentication/phoneMethods`,
              { phoneNumber: formattedPhoneCsv, phoneType: 'alternateMobile' }
            );
          } catch (_) {}
        }

        appendLog(`CSV CREAR | ${upn} | Lote: ${loteIdCsv} | Tipo: ${tipo} | Cédula: ${cedula || '-'} | Tel: ${phone || '-'} | MailPersonal: ${mailPersonal || '-'} | Licencia: ${sku || 'ninguna'}${observacion ? ' | Obs: ' + observacion : ''}`);
        result.estado       = 'éxito';
        result.lote         = loteIdCsv;
        result.cedula       = cedula;
        result.nombres      = fnUpper;
        result.apellidos    = lnUpper;
        result.mailpersonal = mailPersonal;
        result.alias        = upn.split('@')[0];
        result.celular      = formattedPhoneCsv || phone;
        result.upn          = upn;
        result.displayName  = `${fnUpper} ${lnUpper}`;
        result.tipo         = getJobTitle(tipo);
        result.id           = created.id;
        result.creado       = new Date().toISOString();
        result.accountEnabled = true;
      }

      // ── EDITAR ─────────────────────────────────────────────────────────────
      else if (accion === 'editar') {
        const upnExistente = (row.cuenta_institucional || '').trim();
        if (!upnExistente) throw new Error('cuenta_institucional es obligatorio para editar');

        const user = await findUserByUPN(token, upnExistente);
        if (!user) throw new Error(`Usuario no encontrado: ${upnExistente}`);

        result.upn = upnExistente;
        const phone = (row.telefono || '').trim();
        const sku   = (row.skuId    || '').trim();

        // mobilePhone → null; businessPhones → config; MFA → celular formateado
        if (phone) {
          const formattedPhoneEdit = formatMobilePhone(phone, cfg.phonePrefix);
          await graphRequest(token, 'PATCH', `/users/${user.id}`, {
            mobilePhone:    null,
            businessPhones: cfg.phoneNumber ? [cfg.phoneNumber] : [],
          });
          if (formattedPhoneEdit) {
            try {
              const existing = await graphRequest(token, 'GET',
                `/users/${user.id}/authentication/phoneMethods`
              ).catch(() => ({ value: [] }));
              const mobileMethod = (existing.value || []).find(m => m.phoneType === 'mobile');
              if (mobileMethod) {
                await graphRequest(token, 'PUT',
                  `/users/${user.id}/authentication/phoneMethods/${mobileMethod.id}`,
                  { phoneNumber: formattedPhoneEdit, phoneType: 'mobile' }
                );
              } else {
                await graphRequest(token, 'POST',
                  `/users/${user.id}/authentication/phoneMethods`,
                  { phoneNumber: formattedPhoneEdit, phoneType: 'mobile' }
                );
              }
            } catch (_) { /* no crítico */ }
          }
          appendLog(`CSV EDITAR TELÉFONO | ${upnExistente} | MFA: ${formattedPhoneEdit || phone} | Oficina: ${cfg.phoneNumber || '-'}`);
        }
        if (sku) {
          await graphRequest(token, 'POST', `/users/${user.id}/assignLicense`, {
            addLicenses: [{ skuId: sku }], removeLicenses: [],
          });
          appendLog(`CSV ASIGNAR LICENCIA | ${upnExistente} | SKU: ${sku}`);
        }

        result.estado       = 'éxito';
        result.upn          = upnExistente;
        result.alias        = upnExistente.split('@')[0];
        result.celular      = (typeof formattedPhoneEdit !== 'undefined' ? formattedPhoneEdit : formatMobilePhone(phone, cfg.phonePrefix)) || phone || '';
        result.displayName  = user.displayName || '';
        result.nombres      = user.givenName  || (row.nombres   || '').trim();
        result.apellidos    = user.surname    || (row.apellidos || '').trim();
        result.cedula       = (row.cedula       || '').trim();
        result.mailpersonal = (row.mailpersonal || '').trim();
        result.id           = user.id;
        result.accountEnabled = user.accountEnabled ?? true;
        result.detalle      = `Actualizado: ${[phone ? 'teléfono' : '', sku ? 'licencia' : ''].filter(Boolean).join(', ') || 'sin cambios'}`;
      }

      // ── DESACTIVAR ─────────────────────────────────────────────────────────
      else if (accion === 'desactivar') {
        const upnExistente = (row.cuenta_institucional || '').trim();
        if (!upnExistente) throw new Error('cuenta_institucional es obligatorio para desactivar');

        const user = await findUserByUPN(token, upnExistente);
        if (!user) throw new Error(`Usuario no encontrado: ${upnExistente}`);

        await graphRequest(token, 'PATCH', `/users/${user.id}`, { accountEnabled: false });
        appendLog(`CSV INACTIVAR | ${upnExistente}`);
        result.upn          = upnExistente;
        result.alias        = upnExistente.split('@')[0];
        result.displayName  = user.displayName || '';
        result.nombres      = user.givenName  || '';
        result.apellidos    = user.surname    || '';
        result.id           = user.id;
        result.accountEnabled = false;
        result.estado       = 'éxito';
        result.detalle      = 'Cuenta desactivada';
      }

      // ── ACTIVAR ────────────────────────────────────────────────────────────
      else if (accion === 'activar') {
        const upnExistente = (row.cuenta_institucional || '').trim();
        if (!upnExistente) throw new Error('cuenta_institucional es obligatorio para activar');

        const user = await findUserByUPN(token, upnExistente);
        if (!user) throw new Error(`Usuario no encontrado: ${upnExistente}`);

        await graphRequest(token, 'PATCH', `/users/${user.id}`, { accountEnabled: true });
        appendLog(`CSV ACTIVAR | ${upnExistente}`);
        result.upn          = upnExistente;
        result.alias        = upnExistente.split('@')[0];
        result.displayName  = user.displayName || '';
        result.nombres      = user.givenName  || '';
        result.apellidos    = user.surname    || '';
        result.id           = user.id;
        result.accountEnabled = true;
        result.estado       = 'éxito';
        result.detalle      = 'Cuenta activada';
      }

      // ── ELIMINAR ──────────────────────────────────────────────────────────────
      else if (accion === 'eliminar') {
        const upnExistente = (row.cuenta_institucional || '').trim();
        if (!upnExistente) throw new Error('cuenta_institucional es obligatorio para eliminar');

        const user = await findUserByUPN(token, upnExistente);
        if (!user) throw new Error(`Usuario no encontrado: ${upnExistente}`);

        await graphRequest(token, 'DELETE', `/users/${user.id}`);
        appendLog(`CSV ELIMINAR | ${upnExistente}`);
        result.upn          = upnExistente;
        result.alias        = upnExistente.split('@')[0];
        result.displayName  = user.displayName || '';
        result.nombres      = user.givenName  || '';
        result.apellidos    = user.surname    || '';
        result.id           = user.id;
        result.accountEnabled = null;
        result.estado       = 'éxito';
        result.detalle      = 'Cuenta eliminada permanentemente';
      }

      else {
        throw new Error(`Acción desconocida: "${accion}". Use: crear, editar, desactivar, activar, eliminar`);
      }

    } catch (e) {
      result.estado  = 'error';
      result.detalle = e.message;
      appendLog(`CSV ERROR | Fila ${row._fila} | ${accion} | ${e.message}`);
    }

    results.push(result);
  }

  const cOk  = results.filter(r => r.estado === 'éxito').length;
  const cErr = results.filter(r => r.estado === 'error').length;
  appendLog(`CSV COMPLETADO | Éxito: ${cOk} | Error: ${cErr} | Total: ${results.length}`);
  return { ok: true, results };
});

// ═══════════════════════════════════════════════════════════════════════════════
// MÓDULO DE NOTIFICACIONES
// ═══════════════════════════════════════════════════════════════════════════════

// Carpeta de plantillas: al mismo nivel que el ejecutable o raíz del proyecto
function getTemplatesPath() {
  const base = app.isPackaged
    ? path.dirname(process.execPath)
    : path.resolve(__dirname, '../../');
  const dir = path.join(base, 'plantillas');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  return dir;
}

// ─── Listar plantillas .html ───────────────────────────────────────────────────
ipcMain.handle('list-templates', () => {
  try {
    const dir   = getTemplatesPath();
    const files = fs.readdirSync(dir).filter(f => f.toLowerCase().endsWith('.html'));
    appendLog(`PLANTILLAS CARGADAS | ${files.length} plantilla(s) en: ${dir}`);
    return { ok: true, files, dir };
  } catch (e) {
    appendLog(`ERROR PLANTILLAS | ${e.message}`);
    return { ok: false, error: e.message };
  }
});

// ─── Leer contenido de una plantilla ──────────────────────────────────────────
ipcMain.handle('read-template', (_, filename) => {
  try {
    const filePath = path.join(getTemplatesPath(), filename);
    const content  = fs.readFileSync(filePath, 'utf8');
    appendLog(`PLANTILLA SELECCIONADA | ${filename}`);
    return { ok: true, content };
  } catch (e) {
    appendLog(`ERROR LEER PLANTILLA | ${filename} | ${e.message}`);
    return { ok: false, error: e.message };
  }
});

// ─── Abrir carpeta de plantillas en el explorador ─────────────────────────────
ipcMain.handle('open-templates-folder', () => {
  const dir = getTemplatesPath();
  require('electron').shell.openPath(dir);
  return { ok: true, dir };
});

// ─── Seleccionar CSV de notificaciones ────────────────────────────────────────
ipcMain.handle('open-notif-csv', async () => {
  const { filePaths } = await dialog.showOpenDialog(mainWindow, {
    title: 'Seleccionar CSV de notificaciones',
    filters: [{ name: 'Archivo CSV', extensions: ['csv'] }],
    properties: ['openFile'],
  });
  if (!filePaths || !filePaths.length) return { ok: false };
  try {
    const content = fs.readFileSync(filePaths[0], 'utf8');
    appendLog(`CSV NOTIFICACIONES CARGADO | ${filePaths[0]}`);
    return { ok: true, content, filePath: filePaths[0] };
  } catch (e) {
    appendLog(`ERROR CSV NOTIFICACIONES | ${e.message}`);
    return { ok: false, error: e.message };
  }
});

// ─── Enviar notificaciones ────────────────────────────────────────────────────
ipcMain.handle('send-notifications', async (_, { smtpCfg, rows, templateContent, subject, destField }) => {
  if (!nodemailer) return { ok: false, error: 'nodemailer no está instalado. Ejecute: npm install' };

  const transporter = nodemailer.createTransport({
    host:   smtpCfg.host,
    port:   parseInt(smtpCfg.port) || 587,
    secure: smtpCfg.port === '465',
    auth:   { user: smtpCfg.user, pass: smtpCfg.pass },
    tls:    { rejectUnauthorized: false },
  });

  // Verificar conexión antes de enviar
  try {
    await transporter.verify();
    appendLog(`SMTP CONEXIÓN OK | ${smtpCfg.host}:${smtpCfg.port} | Usuario: ${smtpCfg.user}`);
  } catch (e) {
    appendLog(`ERROR SMTP CONEXIÓN | ${smtpCfg.host}:${smtpCfg.port} | ${e.message}`);
    return { ok: false, error: `Error de conexión SMTP: ${e.message}` };
  }

  const results = [];
  if (!rows || rows.length === 0) return { ok: true, results };

  const total = rows.length;
  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx];
    // Notificar progreso al renderer
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('notif-progress', { current: idx + 1, total });
    }
    // destField: columna del CSV que se usa como dirección de envío
    const destRaw = destField ? row[destField] : (row['email_personal'] || row['email_institucional'] || '');
    const dest = String(destRaw || '').trim();
    if (!dest) {
      results.push({ ...row, envio_estado: 'Omitido', envio_detalle: 'Sin correo destino' });
      continue;
    }

    // Rellenar plantilla con los valores de la fila
    const FIELDS = [
      'lote','cedula','nombres','apellidos','email_personal','alias','celular',
      'estado','email_institucional','nombre_completo','tipo','id','creado',
      'estado_cuenta','observacion',
    ];

    let html = templateContent;
    FIELDS.forEach(field => {
      const val = String(row[field] || '');
      // Soportar {{Campo}}, {Campo} y [Campo]
      const safe = field.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      html = html
        .replace(new RegExp(`\\{\\{${safe}\\}\\}`, 'gi'), val)
        .replace(new RegExp(`\\{${safe}\\}`,   'gi'), val)
        .replace(new RegExp(`\\[${safe}\\]`,   'gi'), val);
    });

    try {
      await transporter.sendMail({
        from:    smtpCfg.user,
        to:      dest,
        subject: subject,
        html:    html,
      });
      appendLog(`NOTIF ENVIADO | ${dest} | ${subject}`);
      results.push({ ...row, envio_estado: 'Enviado', envio_detalle: '' });
    } catch (e) {
      appendLog(`NOTIF ERROR | ${dest} | ${e.message}`);
      results.push({ ...row, envio_estado: 'Error', envio_detalle: e.message });
    }
  }

  return { ok: true, results };
});

// ─── Descargar plantilla de notificación (en blanco) ─────────────────────────
ipcMain.handle('download-notif-template', async () => {
  const { filePath } = await dialog.showSaveDialog(mainWindow, {
    title: 'Guardar plantilla de notificación',
    defaultPath: 'plantilla-notificacion.html',
    filters: [{ name: 'Plantilla HTML', extensions: ['html'] }],
  });
  if (!filePath) return { saved: false };

  const template = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"/>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; background:#f3f2f1; margin:0; padding:20px; }
  .container { max-width:600px; margin:0 auto; background:white; border-radius:6px; overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,.1); }
  .header { background:#0078d4; color:white; padding:28px 32px; }
  .header h1 { margin:0; font-size:22px; font-weight:600; }
  .body { padding:28px 32px; color:#201f1e; line-height:1.8; }
  .footer { background:#faf9f8; padding:16px 32px; font-size:11px; color:#605e5c; border-top:1px solid #e1dfdd; }
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>Título del mensaje</h1>
  </div>
  <div class="body">
    <p>Estimado/a {{nombre_completo}},</p>
    <p>Escriba aquí el cuerpo del mensaje.</p>
  </div>
  <div class="footer">Este mensaje fue generado automáticamente.</div>
</div>
</body>
</html>`;

  fs.writeFileSync(filePath, template, 'utf8');
  appendLog(`PLANTILLA NOTIFICACIÓN DESCARGADA | ${filePath}`);
  return { saved: true, path: filePath };
});

// ─── Descargar plantilla CSV para notificaciones ──────────────────────────────
ipcMain.handle('download-notif-csv-template', async () => {
  const { filePath } = await dialog.showSaveDialog(mainWindow, {
    title: 'Guardar plantilla CSV de notificaciones',
    defaultPath: 'plantilla-notificaciones.csv',
    filters: [{ name: 'Archivo CSV', extensions: ['csv'] }],
  });
  if (!filePath) return { saved: false };

  const header = 'lote,cedula,nombres,apellidos,email_personal,alias,celular,estado,email_institucional,nombre_completo,tipo,id,creado,estado_cuenta,observacion\n';
  fs.writeFileSync(filePath, '\uFEFF' + header, 'utf8'); // BOM para Excel, sin filas de ejemplo
  appendLog(`PLANTILLA CSV NOTIFICACIONES DESCARGADA | ${filePath}`);
  return { saved: true, path: filePath };
});

// ─── Cambiar contraseña maestra ───────────────────────────────────────────────
ipcMain.handle('change-master-password', async (_, { currentPass, newPass }) => {
  try {
    // Verificar contraseña actual
    const result = tryUnlock(currentPass);
    if (result !== 'ok') {
      return { ok: false, error: 'La contraseña actual es incorrecta.' };
    }
    // Leer configuración actual (ya descifrada con masterKey)
    const currentCfg = readConf();
    // Reencifrar con la nueva contraseña
    const salt    = crypto.randomBytes(16).toString('hex');
    masterKey     = deriveKey(newPass, salt);
    const plain   = JSON.stringify({ ...currentCfg, __control: CONTROL_VALUE });
    const encoded = encrypt(plain, masterKey);
    fs.writeFileSync(CONF_PATH, JSON.stringify({ salt, data: encoded }));
    appendLog('CAMBIO DE CONTRASEÑA MAESTRA | Completado exitosamente');
    return { ok: true };
  } catch (e) {
    appendLog(`ERROR CAMBIO CONTRASEÑA | ${e.message}`);
    return { ok: false, error: e.message };
  }
});
