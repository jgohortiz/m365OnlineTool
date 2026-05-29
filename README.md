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
- [Capturas de pantalla](#capturas-de-pantalla)
- [Requisitos previos](#requisitos-previos)
- [Instalación y configuración del entorno](#instalación-y-configuración-del-entorno)
- [Configuración en Azure Active Directory](#configuración-en-azure-active-directory)
- [Uso de la aplicación](#uso-de-la-aplicación)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Compilar el instalador](#compilar-el-instalador)
- [Contribuir](#contribuir)
- [Código de conducta](#código-de-conducta)
- [Licencia](#licencia)
- [Aviso legal](#aviso-legal)

---

## Descripción general

**m365OnlineTool** es una herramienta de código abierto orientada a administradores de TI en instituciones educativas que utilizan **Microsoft 365 Plan A1**. Permite gestionar el ciclo completo de vida de cuentas de usuario directamente desde el escritorio de Windows, sin necesidad de abrir el portal web de Microsoft 365 ni escribir scripts de PowerShell.

La aplicación se comunica con la API de **Microsoft Graph v1.0** mediante un registro de aplicación en Azure Active Directory con autenticación de tipo *client credentials* (sin intervención del usuario final).

---

## Características

### Gestión de cuentas
- **Creación de cuentas** con nomenclatura automática según el tipo de usuario:
  - Maestros → prefijo `m.` → `m.pgonzalez@escuela.org`
  - Estudiantes → prefijo `e.` → `e.mgomez@escuela.org`
  - Otros → inicial del nombre + apellido → `jrodriguez@escuela.org`
- Normalización automática de caracteres especiales y acentos en el nombre de usuario
- Vista previa del UPN generado en tiempo real antes de crear la cuenta
- Contraseña temporal con obligación de cambio en el primer inicio de sesión

### Seguridad y recuperación
- Registro del **número de teléfono celular** como método de autenticación MFA y recuperación de cuenta
- El número queda registrado en Microsoft Entra ID como `phoneMethods` para recuperación de acceso

### Licencias
- Listado de todas las SKUs disponibles en el tenant con unidades consumidas y disponibles
- Asignación de licencia A1 al momento de crear la cuenta o en cualquier momento posterior desde la vista de usuarios

### Administración de usuarios
- Lista completa de usuarios del tenant con búsqueda en tiempo real
- Filtro por tipo (maestros, estudiantes, otros) y por estado (activo / inactivo)
- **Activar o inactivar** cuentas con un solo clic (bloqueo/desbloqueo de inicio de sesión)
- Edición rápida de teléfono y licencia desde un diálogo contextual


### Importación masiva desde CSV
- Soporta cuatro acciones en un mismo archivo: **crear**, **editar**, **desactivar** y **activar**
- Generación automática de UPN según el tipo al usar la acción `crear`
- Vista previa de todas las filas antes de procesar
- Tabla de resultados por fila con estado (éxito / error) y detalle del error
- Soporte de arrastrar y soltar el archivo directamente en la interfaz
- Descarga de plantilla `.csv` de ejemplo desde la misma pantalla
- Licencia A1 predeterminada aplicable a todas las cuentas nuevas del lote
- El archivo acepta BOM UTF-8 (compatible con Excel en español)

### Registro de actividad
- Log persistente de todas las operaciones realizadas (creación, activación, desactivación, cambio de teléfono, asignación de licencia)
- Cada entrada incluye fecha y hora ISO 8601, tipo de operación, UPN afectado y detalles
- Descarga del log como archivo `.txt` desde la interfaz
- Coloreado por tipo de operación en el visor integrado

### Interfaz gráfica
- Diseño fiel a **Fluent UI** (sistema de diseño de Microsoft 365): colores, tipografía, controles y espaciado
- Barra de título personalizada con controles de ventana nativos
- Navegación lateral tipo NavigationView con indicador de sección activa
- CommandBar contextual por cada vista
- MessageBar para notificaciones de éxito y error
- Sin emojis ni íconos de terceros — solo SVG geométricos puros

---

## Capturas de pantalla

> Las capturas se generan al compilar y ejecutar la aplicación en su entorno local.

| Vista | Descripción |
|---|---|
| Panel | Estadísticas del tenant y estado de licencias A1 |
| Nueva Cuenta | Formulario con selección de tipo y vista previa del UPN |
| Usuarios | Tabla filtrable con acciones por fila |
| Registro | Consola de actividad con descarga |
| Configuración | Credenciales de Azure y prueba de conexión |

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

O si descargaste el ZIP, descomprímelo y abre la carpeta en la terminal.

### 2. Instalar dependencias

```bash
npm install
```

Este comando descarga Electron, las librerías de Microsoft Graph y todas las dependencias declaradas en `package.json`. La primera ejecución puede tardar entre 3 y 8 minutos según la velocidad de la conexión.

### 3. Ejecutar en modo desarrollo

```bash
npm start
```

La ventana de la aplicación se abrirá. Para inspeccionar errores del proceso de interfaz, presiona **Ctrl + Shift + I** dentro de la app para abrir las DevTools de Chromium.

Los errores del proceso principal (Node.js / Electron) aparecen directamente en la terminal de VS Code.

### 4. Extensiones recomendadas para VS Code

Instálalas desde el panel de extensiones (**Ctrl + Shift + X**):

- `dbaeumer.vscode-eslint` — análisis estático de JavaScript
- `esbenp.prettier-vscode` — formateo automático de código
- `formulahendry.auto-rename-tag` — útil al editar el HTML del renderer

---

## Configuración en Azure Active Directory

Este paso es necesario para que la aplicación pueda comunicarse con Microsoft Graph API.

### Paso 1 — Registrar la aplicación

1. Inicia sesión en https://portal.azure.com con una cuenta de administrador global
2. Busca y abre **Azure Active Directory** (o **Microsoft Entra ID**)
3. En el menú lateral selecciona **Registros de aplicaciones**
4. Haz clic en **Nueva registración**
5. Completa el formulario:
   - **Nombre:** `m365OnlineTool`
   - **Tipos de cuenta admitidos:** Solo las cuentas de este directorio organizativo
   - **URI de redireccionamiento:** dejar vacío
6. Haz clic en **Registrar**

### Paso 2 — Copiar las credenciales

En la página de la aplicación recién creada encontrarás:

- **Id. de directorio (inquilino)** → es el **Tenant ID**
- **Id. de aplicación (cliente)** → es el **Client ID**

Copia ambos valores y guárdalos temporalmente.

### Paso 3 — Crear el secreto de cliente

1. En el menú lateral de la app ve a **Certificados y secretos**
2. Haz clic en **Nuevo secreto de cliente**
3. Descripción: `m365OnlineTool-Secret`, expiración: **24 meses**
4. Haz clic en **Agregar**
5. **Copia el campo Valor inmediatamente** — solo se muestra una vez. Si lo pierdes deberás crear uno nuevo.

### Paso 4 — Asignar permisos de Microsoft Graph

1. En el menú lateral ve a **Permisos de API**
2. Haz clic en **Agregar un permiso → Microsoft Graph → Permisos de aplicación**
3. Busca y selecciona los siguientes permisos:

| Permiso | Motivo |
|---|---|
| `User.ReadWrite.All` | Crear, leer y modificar usuarios |
| `Directory.ReadWrite.All` | Leer y modificar directorio (necesario para asignar licencias) |
| `UserAuthenticationMethod.ReadWrite.All` | Registrar el teléfono como método de autenticación MFA |

4. Haz clic en **Agregar permisos**
5. Haz clic en el botón **Conceder consentimiento de administrador para [nombre del tenant]** y confirma

Los tres permisos deben mostrar una palomita verde en la columna de estado.

### Paso 5 — Ingresar credenciales en la aplicación

1. Abre la app con `npm start`
2. Ve a la sección **Configuración** en el menú lateral
3. Ingresa:
   - **ID de directorio (Tenant ID)**
   - **ID de aplicación (Client ID)**
   - **Secreto de cliente**
   - **Dominio predeterminado** (ej: `miescuela.org`)
4. Haz clic en **Guardar configuración**
5. Haz clic en **Probar conexión** — el indicador en la barra lateral debe volverse verde

---

## Uso de la aplicación

### Crear una cuenta nueva

1. Ve a **Nueva Cuenta** en el menú lateral
2. Selecciona el tipo: **Maestro**, **Estudiante** u **Otro**
3. Ingresa el primer nombre y el apellido — el UPN se genera automáticamente en tiempo real
4. Ingresa el dominio si no está precargado desde la configuración
5. (Opcional) Ingresa el teléfono celular para recuperación de cuenta
6. Selecciona una licencia A1 de la lista desplegable
7. Haz clic en **Crear cuenta**

El sistema normalizará automáticamente los acentos y caracteres especiales del nombre al generar el UPN.

### Activar o inactivar una cuenta

1. Ve a **Usuarios**
2. Localiza al usuario usando la búsqueda o los filtros
3. Haz clic en **Activar** o **Desactivar** en la columna de acciones

### Editar teléfono o asignar licencia a cuenta existente

1. Ve a **Usuarios**
2. Haz clic en **Editar** en la fila del usuario
3. Modifica el teléfono y/o selecciona una nueva licencia
4. Haz clic en **Guardar cambios**

### Descargar el registro de actividad

1. Ve a **Registro de Actividad**
2. Haz clic en **Descargar registro**
3. Selecciona la ubicación y nombre del archivo `.txt`

El archivo de log también se encuentra en:
```
C:\Users\[usuario]\AppData\Roaming\m365OnlineTool\activity.log
```

---

## Estructura del proyecto

```
m365OnlineTool/
│
├── src/
│   ├── main/
│   │   ├── main.js          # Proceso principal de Electron
│   │   │                    # Maneja ventana, IPC, Graph API, log
│   │   └── preload.js       # Puente seguro entre renderer y Node.js
│   │                        # (contextIsolation habilitado)
│   │
│   └── renderer/
│       └── index.html       # Interfaz gráfica completa (HTML + CSS + JS)
│                            # Fluent UI, sin frameworks externos
│
├── resources/
│   └── icon.ico             # Ícono de la aplicación (256x256)
│
├── package.json             # Dependencias, scripts y config de electron-builder
├── README.md                # Este archivo
├── LICENSE.md               # Licencia MIT
└── setup.bat                # Script de instalación para Windows
```

### Dependencias principales

| Paquete | Versión | Uso |
|---|---|---|
| `electron` | 28.x | Shell de escritorio para Windows |
| `@azure/msal-node` | 2.x | Autenticación OAuth2 con Azure AD |
| `@microsoft/microsoft-graph-client` | 3.x | Cliente oficial de Microsoft Graph |
| `electron-store` | 8.x | Almacenamiento local cifrado de configuración |
| `node-fetch` | 2.x | Peticiones HTTP desde el proceso principal |
| `electron-builder` | 24.x | Compilación del instalador `.exe` |

---

## Compilar el instalador

Para generar el instalador `.exe` listo para distribuir:

```bash
npm run build
```

Antes de compilar, asegúrate de tener el archivo `resources/icon.ico` (256x256 px). Puedes convertir cualquier imagen PNG en https://convertio.co/es/png-ico/.

El proceso tardará entre 5 y 10 minutos. Al finalizar encontrarás:

```
dist/
├── m365OnlineTool Setup 1.0.0.exe    # Instalador NSIS (recomendado para distribuir)
└── win-unpacked/                 # Versión portable, sin necesidad de instalar
```

El instalador generado no requiere que Node.js esté instalado en el equipo de destino. Electron incluye su propio runtime de Chromium y Node.js empaquetado.

---

## Contribuir

Las contribuciones son bienvenidas. Este proyecto es de código abierto y está hecho para la comunidad educativa.

### Cómo contribuir

1. Haz un **fork** del repositorio
2. Crea una rama para tu cambio:
   ```bash
   git checkout -b funcionalidad/nombre-descriptivo
   ```
3. Realiza tus cambios y haz commit con mensajes descriptivos:
   ```bash
   git commit -m "feat: agregar soporte para grupos de Microsoft 365"
   ```
4. Sube tu rama:
   ```bash
   git push origin funcionalidad/nombre-descriptivo
   ```
5. Abre un **Pull Request** describiendo el cambio propuesto

### Convenciones de commits

Se recomienda el formato [Conventional Commits](https://www.conventionalcommits.org/es/):

| Prefijo | Uso |
|---|---|
| `feat:` | Nueva funcionalidad |
| `fix:` | Corrección de error |
| `docs:` | Cambios en documentación |
| `style:` | Cambios de formato sin impacto funcional |
| `refactor:` | Reestructuración de código sin cambio de comportamiento |
| `chore:` | Cambios en configuración, dependencias o scripts |

### Reportar errores

Abre un **Issue** en GitHub con:
- Descripción del problema
- Pasos para reproducirlo
- Versión de Windows y de Node.js
- Mensaje de error completo (si aplica)

### Ideas de mejoras futuras

- Importación masiva de usuarios desde archivo CSV o Excel
- Restablecimiento de contraseña desde la interfaz
- Gestión de grupos de Microsoft 365
- Asignación de roles administrativos
- Soporte para múltiples tenants
- Exportación de reporte de usuarios en PDF o Excel
- Notificaciones de licencias próximas a vencer

---

## Código de conducta

Este proyecto adopta el [Contributor Covenant](https://www.contributor-covenant.org/es/version/2/1/code_of_conduct/) como código de conducta. Se espera que todos los participantes mantengan un entorno respetuoso e inclusivo.

---

## Licencia

Distribuido bajo la licencia **MIT**. Consulta el archivo [LICENSE.md](LICENSE.md) para ver el texto completo.

En resumen: puedes usar, copiar, modificar y distribuir este software libremente, incluso con fines comerciales, siempre que incluyas el aviso de copyright original.

---

## Aviso legal

Este proyecto es una herramienta independiente de código abierto y **no está afiliado, patrocinado ni respaldado por Microsoft Corporation**. "Microsoft 365", "Azure", "Fluent UI" y "Microsoft Graph" son marcas registradas de Microsoft Corporation.

El uso de este software implica la aceptación de los términos de servicio de Microsoft 365 y Microsoft Azure aplicables a tu organización.

---

<p align="center">
  Hecho con dedicación para la comunidad educativa de habla hispana.<br/>
  Si este proyecto te resulta útil, considera darle una estrella en GitHub.
</p>
