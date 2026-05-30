# m365OnlineTool — Gestor de Cuentas Microsoft 365 Plan A1

<p align="center">
  <img src="https://img.shields.io/badge/Electron-28.x-47848F?style=for-the-badge&logo=electron&logoColor=white"/>
  <img src="https://img.shields.io/badge/Node.js-18+-339933?style=for-the-badge&logo=node.js&logoColor=white"/>
  <img src="https://img.shields.io/badge/Microsoft%20365-Plan%20A1-0078D4?style=for-the-badge&logo=microsoft&logoColor=white"/>
  <img src="https://img.shields.io/badge/Licencia-MIT-green?style=for-the-badge"/>
  <img src="https://img.shields.io/badge/Plataforma-Windows-0078D4?style=for-the-badge&logo=windows&logoColor=white"/>
  <img src="https://img.shields.io/badge/Estado-Activo-brightgreen?style=for-the-badge"/>
</p>

<p align="center">
  Aplicación de escritorio para Windows que permite a instituciones educativas gestionar cuentas de usuario en Microsoft 365 Plan A1 a través de Microsoft Graph API, con interfaz gráfica en estilo Fluent UI.
</p>

---

## Tabla de contenidos

- [Descripción general](#descripción-general)
- [Características](#características)
- [Requisitos previos](#requisitos-previos)
- [Instalación y configuración del entorno](#instalación-y-configuración-del-entorno)
- [Configuración en Azure Active Directory](#configuración-en-azure-active-directory)
- [Uso de la aplicación](#uso-de-la-aplicación)
- [Formato del CSV — Acciones masivas](#formato-del-csv--acciones-masivas)
- [Formato del CSV — Notificaciones](#formato-del-csv--notificaciones)
- [Plantillas HTML de notificación](#plantillas-html-de-notificación)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Compilar el instalador](#compilar-el-instalador)
- [Contribuir](#contribuir)
- [Licencia](#licencia)
- [Aviso legal](#aviso-legal)

---

## Descripción general

**m365OnlineTool** es una herramienta de código abierto orientada a administradores de TI en instituciones educativas que utilizan **Microsoft 365 Plan A1**. Permite gestionar el ciclo completo de vida de cuentas de usuario directamente desde el escritorio de Windows, sin necesidad de abrir el portal web de Microsoft 365 ni escribir scripts de PowerShell.

La aplicación se comunica con **Microsoft Graph API v1.0** mediante un registro de aplicación en Azure Active Directory con autenticación de tipo *client credentials*. La configuración se almacena localmente cifrada con **AES-256-CBC**.

---

## Características

### Seguridad
- Configuración protegida con **contraseña maestra** y cifrado **AES-256-CBC**
- Clave derivada con **PBKDF2** (100.000 iteraciones, SHA-256) + salt aleatorio
- Pantalla de desbloqueo al iniciar con detección de archivo nuevo, existente o sin cifrar (migración automática)
- Opción para cambiar la contraseña maestra desde la interfaz

### Panel
- Estadísticas del tenant: total de usuarios, activos, inactivos, con licencia
- Tabla de licencias A1 disponibles con unidades consumidas y disponibles
- Carga automática al abrir la aplicación (solo la primera vez)

### Nueva Cuenta
- Generación automática de UPN con **3 métodos en cascada** que verifican disponibilidad en Exchange Online:
  - Método 1: `prefijo.L1Apellido1` → `m.jdoe@escuela.org`
  - Método 2: `prefijo.L1L2Apellido1` → `m.jmdoe@escuela.org` (si hay ≥ 2 nombres)
  - Método 3: `prefijo.L1Apellido1L2Apellido2` → `m.jdoes@escuela.org` (si hay ≥ 2 apellidos)
  - Fallback con sufijo numérico incremental
- Prefijos automáticos por tipo: `m.` maestros, `e.` estudiantes, sin prefijo para otros
- Validación de nombre completo y correo institucional antes de crear (bloquea duplicados)
- Contraseña temporal aleatoria con parámetros configurables
- Teléfono celular como método MFA (autenticación SMS)
- Asignación automática de licencia según el tipo configurado
- `JobTitle` automático: `DOCENTE`, `ESTUDIANTE` u `OTRO`
- Nombres y apellidos guardados en **mayúsculas**
- Campos organizacionales desde la configuración (ciudad, país, dirección, etc.)
- `CustomAttribute1` = ID de lote (`LT-AAMMDD:HHmmss-xxxxxxxxx`)
- `CustomAttribute2` = Cédula
- `CustomAttribute3` = Correo personal

### Usuarios
- Lista completa con paginación automática (sin límite de registros)
- Filtros por columna en tiempo real: nombre, UPN, tipo, teléfono, licencia, estado, fecha, atributos
- Ordenamiento ascendente/descendente por cualquier columna
- Columnas: nombre, UPN, tipo, teléfono, licencia, estado, fecha de creación, Lote (CA1), Cédula (CA2), Correo personal (CA3)
- Activar o inactivar cuentas
- Editar teléfono y asignar licencia
- Eliminar cuentas con **token de confirmación aleatorio** de 8 caracteres
- Exportar lista filtrada a `.csv`

### Acciones masivas (CSV)
- Importación de usuarios desde archivo `.csv` con cinco acciones: `crear`, `editar`, `desactivar`, `activar`, `eliminar`
- Eliminaciones masivas requieren token de confirmación
- Vista previa de todas las filas antes de procesar
- Tabla de resultados con 15 columnas por fila
- Exportación de resultados a `.csv`

### Notificaciones
- Envío de correos desde plantillas HTML a destinatarios de un `.csv`
- Variables en plantillas con formato `{{variable}}`, `{variable}` o `[variable]`
- Indicador de progreso durante el envío (`Enviando correo N de M`)
- Tabla de destinatarios con búsqueda en tiempo real
- Resultados del envío exportables a `.csv`

### Registro de Actividad
- Log persistente de todas las operaciones con fecha/hora ISO 8601
- Coloreado por tipo de operación en el visor integrado
- Descarga como archivo `.txt`

### Configuración
- **Conexión API**: Tenant ID, Client ID, Client Secret, Dominio predeterminado + probar conexión
- **Organización**: Datos que se asignan a cada cuenta nueva (ciudad, país, dirección, prefijo telefónico, etc.)
- **Licencias**: SKU predeterminado por tipo de cuenta + parámetros de contraseña aleatoria
- **Correo saliente**: Servidor SMTP, puerto, usuario + probar conexión SMTP
- **Contraseña maestra**: Cambio de clave con verificación y medidor de fortaleza
- **Acerca de**: Versión, API, ruta de configuración

---

## Requisitos previos

### En la máquina de desarrollo

| Herramienta | Versión mínima | Enlace |
|---|---|---|
| Node.js | 18.x LTS | https://nodejs.org |
| npm | 9.x (incluido con Node.js) | — |
| Visual Studio Code | Cualquier versión reciente | https://code.visualstudio.com |
| Git | 2.x | https://git-scm.com |
| Windows | 10 / 11 (64 bits) | — |

### En Microsoft Azure

- Cuenta de **administrador global** en el tenant de Microsoft 365 Plan A1
- Acceso a **portal.azure.com**

---

## Instalación y configuración del entorno

### 1. Clonar el repositorio

```bash
git clone https://github.com/tu-usuario/m365OnlineTool.git
cd m365OnlineTool
```

### 2. Instalar dependencias

```bash
npm install
```

### 3. Ejecutar en modo desarrollo

```bash
npm start
```

Al iniciar por primera vez se muestra la pantalla de configuración inicial para crear la contraseña maestra. Los datos se cifran con AES-256-CBC antes de guardarse en `m365OnlineTool.conf`.

---

## Configuración en Azure Active Directory

### Paso 1 — Registrar la aplicación

1. Inicia sesión en https://portal.azure.com
2. Ve a **Azure Active Directory → Registros de aplicaciones → Nueva registración**
3. Nombre: `m365OnlineTool` | Tipo: solo este directorio | Sin URI de redireccionamiento
4. Haz clic en **Registrar**

### Paso 2 — Copiar las credenciales

- **Id. de directorio (inquilino)** → Tenant ID
- **Id. de aplicación (cliente)** → Client ID

### Paso 3 — Crear el secreto

Ve a **Certificados y secretos → Nuevo secreto de cliente**. Copia el **Valor** inmediatamente (solo se muestra una vez).

### Paso 4 — Asignar permisos

Ve a **Permisos de API → Agregar permiso → Microsoft Graph → Permisos de aplicación** y agrega:

| Permiso | Motivo |
|---|---|
| `User.ReadWrite.All` | Crear, leer y modificar usuarios |
| `Directory.ReadWrite.All` | Asignar licencias |
| `UserAuthenticationMethod.ReadWrite.All` | Registrar teléfono para MFA |

Haz clic en **Conceder consentimiento de administrador**.

### Paso 5 — Ingresar credenciales en la app

Ve a **Configuración → Conexión API**, ingresa los datos y haz clic en **Probar conexión Graph API**.

---

## Uso de la aplicación

### Crear una cuenta

1. Ve a **Nueva Cuenta**
2. Selecciona el tipo (Maestro / Estudiante / Otro)
3. Ingresa nombres, apellidos, teléfono, cédula y correo personal
4. El correo institucional se genera automáticamente
5. Haz clic en **Crear cuenta**

### Acciones masivas

1. Ve a **Acciones masivas**
2. Descarga la plantilla CSV con **Descargar plantilla**
3. Completa el CSV con las acciones requeridas
4. Selecciona el archivo y haz clic en **Procesar importación**

Para acciones de eliminación se solicita un token de confirmación aleatorio de 8 caracteres.

### Notificaciones

1. Ve a **Notificaciones**
2. Configura las credenciales SMTP en **Configuración → Correo saliente**
3. Selecciona una plantilla HTML de la carpeta `plantillas/`
4. Carga el CSV de destinatarios
5. Ingresa la contraseña del remitente y haz clic en **Enviar notificaciones**

---

## Formato del CSV — Acciones masivas

| Columna | Acciones que la usan | Descripción |
|---|---|---|
| `accion` | todas | `crear` / `editar` / `desactivar` / `activar` / `eliminar` |
| `tipo` | `crear` | `maestro` / `estudiante` / `otro` |
| `nombres` | `crear` | Nombres completos |
| `apellidos` | `crear` | Apellidos completos |
| `telefono` | `crear`, `editar` | Celular con prefijo (ej: `+57 3001234567`) |
| `cuenta_institucional` | `editar`, `desactivar`, `activar`, `eliminar` | UPN de la cuenta |
| `cedula` | `crear`, `editar` | Número de cédula → `CustomAttribute2` |
| `mailpersonal` | `crear`, `editar` | Correo personal → `CustomAttribute3` |
| `observacion` | todas | Nota libre visible en resultados y log |

El dominio, la contraseña y el SKU de licencia se toman automáticamente de la Configuración.

---

## Formato del CSV — Notificaciones

```
lote,cedula,nombres,apellidos,email_personal,alias,celular,estado,
email_institucional,nombre_completo,tipo,id,creado,estado_cuenta,observacion
```

Este formato corresponde exactamente al CSV exportado desde la sección **Resultados de la importación** en Acciones masivas.

---

## Plantillas HTML de notificación

Las plantillas se guardan en la carpeta `plantillas/` al mismo nivel que el ejecutable. Las variables se insertan con cualquiera de estos formatos:

```
{{variable}}   {variable}   [variable]
```

Variables disponibles:

| Variable | Descripción |
|---|---|
| `{{nombre_completo}}` | Nombre completo del destinatario |
| `{{email_institucional}}` | Correo institucional |
| `{{email_personal}}` | Correo personal |
| `{{alias}}` | Alias (parte del UPN antes del @) |
| `{{cedula}}` | Número de cédula |
| `{{celular}}` | Teléfono celular |
| `{{tipo}}` | DOCENTE / ESTUDIANTE / OTRO |
| `{{lote}}` | ID de lote de creación |
| `{{creado}}` | Fecha de creación |
| `{{estado_cuenta}}` | Activa / Inactiva |
| `{{observacion}}` | Observación del CSV |

---

## Estructura del proyecto

```
m365OnlineTool/
├── src/
│   ├── main/
│   │   ├── main.js          ← Proceso principal (Electron + Graph API + cifrado)
│   │   └── preload.js       ← Puente seguro IPC (contextIsolation)
│   └── renderer/
│       ├── index.html       ← Interfaz principal (Fluent UI)
│       └── lock.html        ← Pantalla de desbloqueo / primera configuración
├── plantillas/
│   └── Bienvenida cuenta institucional.html   ← Plantilla de ejemplo
├── resources/
│   └── icon.ico             ← Ícono de la aplicación (256x256)
├── package.json
├── README.md
└── LICENSE.md
```

### Archivo de configuración

```
m365OnlineTool.conf          ← Al mismo nivel del ejecutable o raíz del proyecto
```

Formato cifrado:
```json
{ "salt": "hex_aleatorio", "data": "iv_hex:datos_aes256_hex" }
```

### Dependencias principales

| Paquete | Versión | Uso |
|---|---|---|
| `electron` | 28.x | Shell de escritorio para Windows |
| `@azure/msal-node` | 2.x | Autenticación OAuth2 con Azure AD |
| `@microsoft/microsoft-graph-client` | 3.x | Cliente oficial de Microsoft Graph |
| `node-fetch` | 2.x | Peticiones HTTP desde el proceso principal |
| `nodemailer` | 6.x | Envío de correos SMTP |
| `electron-builder` | 24.x | Compilación del instalador `.exe` |

---

## Compilar el instalador

```bash
npm run build
```

Requiere el archivo `resources/icon.ico` (256x256 px). Al finalizar:

```
dist/
├── m365OnlineTool Setup 1.0.0.exe    ← Instalador NSIS
└── win-unpacked/                      ← Versión portable
```

---

## Contribuir

1. Haz un **fork** del repositorio
2. Crea una rama: `git checkout -b funcionalidad/nombre-descriptivo`
3. Realiza tus cambios y haz commit: `git commit -m "feat: descripción"`
4. Sube tu rama: `git push origin funcionalidad/nombre-descriptivo`
5. Abre un **Pull Request**

**Convenciones de commits:** `feat:` nueva función · `fix:` corrección · `docs:` documentación · `refactor:` reestructuración

**Ideas de mejoras futuras**
- Importación masiva desde Excel
- Restablecimiento de contraseña desde la interfaz
- Gestión de grupos de Microsoft 365
- Soporte para múltiples tenants
- Notificaciones de licencias próximas a vencer

---

## Licencia

Distribuido bajo la licencia **MIT**. Consulta el archivo [LICENSE.md](LICENSE.md) para ver el texto completo.

---

## Aviso legal

Este proyecto es una herramienta independiente de código abierto y **no está afiliado, patrocinado ni respaldado por Microsoft Corporation**. "Microsoft 365", "Azure", "Fluent UI" y "Microsoft Graph" son marcas registradas de Microsoft Corporation.

---

<p align="center">
  Hecho con dedicación para la comunidad educativa de habla hispana.<br/>
  Si este proyecto te resulta útil, considera darle una estrella en GitHub.
</p>
