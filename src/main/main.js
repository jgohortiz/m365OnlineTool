const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs   = require('fs');
// ─── Archivo de configuración m365OnlineTool.conf ────────────────────────────
// Se guarda al mismo nivel del ejecutable (en producción) o del proyecto (dev)
const CONF_PATH = path.join(
  app.isPackaged
    ? path.dirname(process.execPath)          // junto al .exe instalado
    : path.resolve(__dirname, '../../'),      // raíz del proyecto en desarrollo
  'm365OnlineTool.conf'
);

function readConf() {
  try {
    if (!fs.existsSync(CONF_PATH)) return {};
    return JSON.parse(fs.readFileSync(CONF_PATH, 'utf8'));
  } catch { return {}; }
}

function writeConf(data) {
  fs.writeFileSync(CONF_PATH, JSON.stringify(data, null, 2), 'utf8');
}

let mainWindow;

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

app.whenReady().then(createWindow);
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
app.on('activate', () => { if (!mainWindow) createWindow(); });

// ─── Controles de ventana ─────────────────────────────────────────────────────
ipcMain.on('window-minimize', () => mainWindow.minimize());
ipcMain.on('window-maximize', () => mainWindow.isMaximized() ? mainWindow.unmaximize() : mainWindow.maximize());
ipcMain.on('window-close',    () => mainWindow.close());

// ─── Configuración ────────────────────────────────────────────────────────────
ipcMain.handle('get-config',  ()      => readConf());
ipcMain.handle('save-config', (_, c)  => { writeConf(c); return true; });

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
  if (filePath) { fs.copyFileSync(LOG_PATH, filePath); return { saved: true, path: filePath }; }
  return { saved: false };
});

ipcMain.handle('clear-log', () => { fs.writeFileSync(LOG_PATH, ''); return true; });

// ─── Licencias ────────────────────────────────────────────────────────────────
ipcMain.handle('get-licenses', async (_, cfg) => {
  try {
    const token = await getAccessToken(cfg);
    const data  = await graphRequest(token, 'GET', '/subscribedSkus');
    return { ok: true, data: data.value };
  } catch (e) { return { ok: false, error: e.message }; }
});

// ─── Listar usuarios ──────────────────────────────────────────────────────────
ipcMain.handle('list-users', async (_, cfg) => {
  try {
    const token    = await getAccessToken(cfg);
    const allUsers = [];
    // Graph API devuelve máximo 100 usuarios por página; se pagina con @odata.nextLink
    let url = '/users?$select=id,displayName,userPrincipalName,accountEnabled,mobilePhone,assignedLicenses&$top=100';
    while (url) {
      const data = await graphRequest(token, 'GET', url);
      if (data.value) allUsers.push(...data.value);
      const next = data['@odata.nextLink'];
      url = next ? next.replace('https://graph.microsoft.com/v1.0', '') : null;
    }
    return { ok: true, data: allUsers };
  } catch (e) { return { ok: false, error: e.message }; }
});

// ─── Buscar usuario por UPN ───────────────────────────────────────────────────
async function findUserByUPN(token, upn) {
  try {
    const enc  = encodeURIComponent(upn);
    const data = await graphRequest(token, 'GET',
      `/users?$filter=userPrincipalName eq '${enc}'&$select=id,displayName,userPrincipalName,accountEnabled,mobilePhone,assignedLicenses`
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

    const created = await graphRequest(token, 'POST', '/users', {
      accountEnabled:    true,
      displayName:       `${userData.firstName} ${userData.lastName}`,
      givenName:         userData.firstName,
      surname:           userData.lastName,
      userPrincipalName: upn,
      mailNickname:      upn.split('@')[0],
      jobTitle:          getJobTitle(userData.tipo),
      passwordProfile:   { forceChangePasswordNextSignIn: true, password: userData.password },
      mobilePhone:       formattedPhone,
      // AlternateMobilePhones: teléfono del usuario con prefijo configurado
      businessPhones:    formattedPhone ? [formattedPhone] : (cfg.phoneNumber ? [cfg.phoneNumber] : []),
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
ipcMain.handle('update-phone', async (_, { cfg, userId, phone, upn }) => {
  try {
    const token = await getAccessToken(cfg);
    await graphRequest(token, 'PATCH', `/users/${userId}`, { mobilePhone: phone });
    appendLog(`ACTUALIZAR TELÉFONO | ${upn} | ${phone}`);
    return { ok: true };
  } catch (e) { return { ok: false, error: e.message }; }
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

  const header  = 'accion,tipo,nombres,apellidos,dominio,telefono,contrasena,skuId,upn_existente,cedula,mailpersonal,observacion\n';
  const ejemplo = [
    'crear,maestro,John,Doe,escuela.org,+57 3001234567,Cambiar2024!,,,10000001,john.doe@correo.com,Ejemplo maestro',
    'crear,estudiante,Jane,Doe,escuela.org,+57 3109876543,Cambiar2024!,,,10000002,jane.doe@correo.com,Ejemplo estudiante',
    'crear,otro,Admin,Doe,escuela.org,+57 3201111111,Cambiar2024!,SKU-ID-AQUI,,10000003,admin.doe@correo.com,Ejemplo otro',
    'editar,,,,,+57 3001111111,,,,10000001,,Actualizar teléfono,m.jdoe@escuela.org',
    'desactivar,,,,,,,,e.jdoe@escuela.org,,,,',
    'activar,,,,,,,,adoe@escuela.org,,,,',
  ].join('\n');

  fs.writeFileSync(filePath, '\uFEFF' + header + ejemplo, 'utf8'); // BOM para Excel
  return { saved: true, path: filePath };
});

// ─── Procesar importación CSV ─────────────────────────────────────────────────
ipcMain.handle('process-csv', async (_, { cfg, rows, skuId: defaultSku }) => {
  const results = [];
  let token;
  try { token = await getAccessToken(cfg); }
  catch (e) { return { ok: false, error: `Error de autenticación: ${e.message}` }; }

  for (const row of rows) {
    const accion = (row.accion || '').trim().toLowerCase();
    const result = { fila: row._fila, accion, upn: '', estado: '', detalle: '' };

    try {
      // ── CREAR ──────────────────────────────────────────────────────────────
      if (accion === 'crear') {
        const fn     = (row.nombres   || '').trim();
        const ln     = (row.apellidos || '').trim();
        const cedula      = (row.cedula       || '').trim();
        const mailPersonal= (row.mailpersonal || '').trim();
        const observacion = (row.observacion  || '').trim();
        const domain = (row.dominio || cfg.defaultDomain || '').trim().replace(/^@/, '');
        const tipo   = (row.tipo          || 'otro').trim().toLowerCase();
        const phone  = (row.telefono      || '').trim();
        const pass   = (row.contrasena    || 'Cambiar2024!').trim();
        // SKU: fila > SKU predeterminado del tipo (config) > SKU predeterminado global
        const skuFromTipo = tipo === 'maestro'    ? (cfg.skuMaestro    || '')
                          : tipo === 'estudiante' ? (cfg.skuEstudiante || '')
                          : (cfg.skuOtro || '');
        const sku    = (row.skuId || skuFromTipo || defaultSku || '').trim();

        if (!fn || !ln || !domain) throw new Error('Faltan: nombres, apellidos o dominio');

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

        const created = await graphRequest(token, 'POST', '/users', {
          accountEnabled:    true,
          displayName:       `${fn} ${ln}`,
          givenName:         fn,
          surname:           ln,
          userPrincipalName: upn,
          mailNickname:      upn.split('@')[0],
          jobTitle:          getJobTitle(tipo),
          passwordProfile:   { forceChangePasswordNextSignIn: true, password: pass },
          mobilePhone:       formattedPhoneCsv,
          // AlternateMobilePhones con prefijo configurado
          businessPhones:    formattedPhoneCsv ? [formattedPhoneCsv] : (cfg.phoneNumber ? [cfg.phoneNumber] : []),
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
        result.estado  = 'éxito';
        result.detalle = `Cuenta creada: ${upn}`;
      }

      // ── EDITAR ─────────────────────────────────────────────────────────────
      else if (accion === 'editar') {
        const upnExistente = (row.upn_existente || '').trim();
        if (!upnExistente) throw new Error('upn_existente es obligatorio para editar');

        const user = await findUserByUPN(token, upnExistente);
        if (!user) throw new Error(`Usuario no encontrado: ${upnExistente}`);

        result.upn = upnExistente;
        const phone = (row.telefono || '').trim();
        const sku   = (row.skuId    || '').trim();

        if (phone) {
          await graphRequest(token, 'PATCH', `/users/${user.id}`, { mobilePhone: phone });
          appendLog(`CSV EDITAR TELÉFONO | ${upnExistente} | ${phone}`);
        }
        if (sku) {
          await graphRequest(token, 'POST', `/users/${user.id}/assignLicense`, {
            addLicenses: [{ skuId: sku }], removeLicenses: [],
          });
          appendLog(`CSV ASIGNAR LICENCIA | ${upnExistente} | SKU: ${sku}`);
        }

        result.estado  = 'éxito';
        result.detalle = `Actualizado: ${[phone ? 'teléfono' : '', sku ? 'licencia' : ''].filter(Boolean).join(', ') || 'sin cambios'}`;
      }

      // ── DESACTIVAR ─────────────────────────────────────────────────────────
      else if (accion === 'desactivar') {
        const upnExistente = (row.upn_existente || '').trim();
        if (!upnExistente) throw new Error('upn_existente es obligatorio para desactivar');

        const user = await findUserByUPN(token, upnExistente);
        if (!user) throw new Error(`Usuario no encontrado: ${upnExistente}`);

        await graphRequest(token, 'PATCH', `/users/${user.id}`, { accountEnabled: false });
        appendLog(`CSV INACTIVAR | ${upnExistente}`);
        result.upn     = upnExistente;
        result.estado  = 'éxito';
        result.detalle = 'Cuenta desactivada';
      }

      // ── ACTIVAR ────────────────────────────────────────────────────────────
      else if (accion === 'activar') {
        const upnExistente = (row.upn_existente || '').trim();
        if (!upnExistente) throw new Error('upn_existente es obligatorio para activar');

        const user = await findUserByUPN(token, upnExistente);
        if (!user) throw new Error(`Usuario no encontrado: ${upnExistente}`);

        await graphRequest(token, 'PATCH', `/users/${user.id}`, { accountEnabled: true });
        appendLog(`CSV ACTIVAR | ${upnExistente}`);
        result.upn     = upnExistente;
        result.estado  = 'éxito';
        result.detalle = 'Cuenta activada';
      }

      else {
        throw new Error(`Acción desconocida: "${accion}". Use: crear, editar, desactivar, activar`);
      }

    } catch (e) {
      result.estado  = 'error';
      result.detalle = e.message;
      appendLog(`CSV ERROR | Fila ${row._fila} | ${accion} | ${e.message}`);
    }

    results.push(result);
  }

  return { ok: true, results };
});
