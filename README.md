# La Experiencia SharePoint — Código de ejemplos

Repositorio oficial de código del libro **"La Experiencia SharePoint: Desarrollo con Microsoft Graph"**.

> Fuente única de verdad del código de ejemplos. El libro embebe el código inline en los capítulos como extractos didácticos; este repo contiene los archivos **completos, compilables y ejecutables**.

## 📚 Lenguajes

- **C#** (.NET 8.0+)
- **PowerShell** (7.4+)
- **JavaScript/Node.js** (18+)
- **Python** (3.10+)
- **Java** (17+)

## 📁 Estructura

```
├── common/                        # Módulo de autenticación compartido (DRY + DI)
│   ├── SharePointGraphAuth/        #   C# (GraphAuthOptions + SharePointGraphClientFactory)
│   ├── SharePointGraphAuth.Tests/  #   C# xUnit tests
│   ├── laexperiencia_sharepoint/   #   Python (get_graph_client / get_access_token) + tests
│   ├── graphAuth.js                #   JavaScript (getGraphClient / getAccessToken) + Jest tests
│   ├── SharePointGraph.psm1        #   PowerShell (Connect-SharePointGraph) + Pester tests
│   └── sharepoint-graph-auth/      #   Java (GraphServiceClientFactory) + JUnit tests
├── chapter-02-auth/               # Autenticación (interactiva/delegada + certificado + cache)
├── chapter-03-sites/              # Sitios y bibliotecas
├── chapter-04-documents/          # Documentos
├── chapter-05-permissions/        # Permisos y sharing
├── chapter-06-metadata/           # Metadata y content types (Term Store)
├── chapter-07-automation/         # Webhooks y delta queries
├── chapter-14-spfx/               # 12 proyectos SharePoint Framework (SPFx 1.20, Node 18 LTS)
├── chapter-17-teams/              # Apps de Teams: tab app standalone + bot que lee SharePoint via Graph
├── la-experiencia-sharepoint.slnx  # Solución C# (dotnet)
├── pom.xml                         # Reactor Java (Maven multi-module)
├── package.json                    # Node.js (deps + Jest)
├── requirements.txt                # Python (raíz)
├── .env.example                    # Plantilla de variables de entorno
└── smoke-tests.sh                  # Smoke tests en vivo contra el tenant
```

## 🔐 Variables de entorno

Copia `.env.example` a `.env` y rellena:

```
TENANT_ID=...
CLIENT_ID=...
CLIENT_SECRET=...
SHAREPOINT_SITE_URL=https://TU-TENANT.sharepoint.com/sites/book-test
SHAREPOINT_SITE_ID=...
```

> ⚠️ Nunca commitees `.env` con credenciales reales (ya está en `.gitignore`).

## 🚀 Instalación y build

**C# (.NET):**
```bash
dotnet build la-experiencia-sharepoint.slnx     # 0 errores
dotnet test common/SharePointGraphAuth.Tests    # 5 tests unitarios
```

**Java (Maven):**
```bash
mvn clean compile                                 # reactor completo, 0 errores
mvn -pl common/sharepoint-graph-auth test         # 2 tests JUnit
```

**JavaScript (Node.js):**
```bash
npm install
npx jest                                          # 3 tests Jest
```

**Python:**
```bash
pip install -r requirements.txt
pip install -e common                              # módulo de auth compartido
python -m pytest common/tests/                     # 4 tests pytest
```

**PowerShell:**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
pwsh -NoProfile -Command "Invoke-Pester -Path ./common/SharePointGraph.Tests.ps1 -Output Detailed"  # 2 tests Pester
```

**SPFx (cap14) — requiere Node 18 LTS:**
```bash
nvm use 18
cd chapter-14-spfx/01-webpart-hello-graph   # cualquier proyecto
npm install
gulp build                                   # 0 warnings, 0 TS errors (los 11 proyectos)
gulp bundle --ship && gulp package-solution --ship   # -> sharepoint/solution/*.sppkg
```
Los 11 proyectos SPFx de `chapter-14-spfx/` construyen verdes (0 warnings). Ver
`chapter-14-spfx/README.md` para el índice de los 12 ejemplos y los permisos Graph
que cada uno declara.

**Teams (cap17):**
```bash
cd chapter-17-teams/02-bot-sharepoint
npm install
npm test          # 4 tests Jest (graphAuth, sin red)
npm run build     # tsc, 0 errores
```
`01-tab-app-sharepoint` es un package de manifest + iconos (sin build; `bash package.sh` genera el .zip). `02-bot-sharepoint` es un bot Node (botbuilder + restify) que responde `sites` / `docs <site>` leyendo Graph. Ver los README de cada uno para el ciclo de vida completo (Entra ID, Azure Bot, ngrok, sideload).

## ▶️ Ejecutar los ejemplos

Cada capítulo tiene un `main`/`Main` de **solo lectura** que opera contra el sitio de pruebas (resuelto por path, por defecto `olddogsoft1.sharepoint.com/sites/book-test` — sobreescribible con `SHAREPOINT_HOSTNAME`/`SHAREPOINT_SITE_PATH`).

```bash
# C#
dotnet run --project chapter-03-sites/csharp/Chapter03Sites.csproj
# Python
python chapter-03-sites/python/site_operations.py
# JavaScript
node chapter-03-sites/javascript/SiteOperations.js
# PowerShell
pwsh -NoProfile -File chapter-03-sites/powershell/SiteOperations.ps1
# Java
mvn -q install -DskipTests
mvn -q -pl chapter-03-sites/java exec:java -Dexec.mainClass=com.sharepointexperience.chapter03.SiteOperations
```

## 🧪 Tests automatizados

```bash
bash smoke-tests.sh        # 24 smoke tests en vivo contra el tenant (C#/Py/JS/PS/Java)
```

Suite total: **40 pruebas** (16 unit sin red + 24 smoke en vivo); **40/40 validadas en verde** con mínimo privilegio (`Sites.Selected`). Ver `PRUEBAS-AUTOMATIZADAS.md`.

## 📖 Notas

- La autenticación **app-only con client secret** (Microsoft Graph) es el patrón soportado para automatización no interactiva. Los ejemplos del cap. 2 usan auth **interactiva/delegada** (navegador) para enseñar los flujos.
- El módulo `common/` centraliza la auth (DRY); los ejemplos reciben el cliente/token por **inyección de dependencias**.
- `complete-workflow/` (proyecto integrador) es un stub pendiente de portar.

## 📄 Licencia

MIT License (código). El contenido del libro es CC BY-NC-SA 4.0.

## 💝 Apoya el proyecto

Este código es gratuito y open source. Si te ayuda, puedes apoyar de varias maneras:

- **GitHub Sponsors** — pulsa el botón **Sponsor** arriba (requiere que el autor tenga el perfil de Sponsors habilitado en `olddogsoftllc`).
- **Comprar el libro** — *La Experiencia SharePoint* (4 volúmenes) en Leanpub (enlace próximamente; mientras tanto en [olddogsoft.com/blog](https://www.olddogsoft.com/blog)).
- **PayPal / Ko-fi** — vías one-click configuradas vía `.github/FUNDING.yml` (ver ese archivo para activarlas).
- **Compartir el repo** — una estrella ⭐ y compartir con tu equipo también ayudan.

> El código siempre será gratis y MIT. El libro (contenido editorial) es de pago para sostener el proyecto. Tu apoyo financia más ejemplos, validación en vivo y mantenimiento frente a los cambios de Microsoft Graph.