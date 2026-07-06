# La Experiencia SharePoint — Example Code

Official code repository for the book **"La Experiencia SharePoint: Desarrollo con Microsoft Graph"**.

> Single source of truth for the example code. The book inlines code excerpts within chapters for didactic purposes; this repo contains the **complete, compilable, and runnable** files.

## 📚 Languages

- **C#** (.NET 8.0+)
- **PowerShell** (7.4+)
- **JavaScript/Node.js** (18+)
- **Python** (3.10+)
- **Java** (17+)

## 📁 Structure

```
├── common/                        # Shared authentication module (DRY + DI)
│   ├── SharePointGraphAuth/        #   C# (GraphAuthOptions + SharePointGraphClientFactory)
│   ├── SharePointGraphAuth.Tests/  #   C# xUnit tests
│   ├── laexperiencia_sharepoint/   #   Python (get_graph_client / get_access_token) + tests
│   ├── graphAuth.js                #   JavaScript (getGraphClient / getAccessToken) + Jest tests
│   ├── SharePointGraph.psm1        #   PowerShell (Connect-SharePointGraph) + Pester tests
│   └── sharepoint-graph-auth/      #   Java (GraphServiceClientFactory) + JUnit tests
├── chapter-02-auth/               # Authentication (interactive/delegated + certificate + cache)
├── chapter-03-sites/              # Sites and libraries
├── chapter-04-documents/          # Documents
├── chapter-05-permissions/        # Permissions and sharing
├── chapter-06-metadata/           # Metadata and content types (Term Store)
├── chapter-07-automation/         # Webhooks and delta queries
├── chapter-14-spfx/               # 12 SharePoint Framework projects (SPFx 1.20, Node 18 LTS)
├── chapter-17-teams/              # Teams apps: standalone tab app + bot that reads SharePoint via Graph
├── la-experiencia-sharepoint.slnx  # C# solution (dotnet)
├── pom.xml                         # Java reactor (Maven multi-module)
├── package.json                    # Node.js (deps + Jest)
├── requirements.txt                # Python (root)
├── .env.example                    # Environment variables template
└── smoke-tests.sh                  # Live smoke tests against the tenant
```

## 🔐 Environment variables

Copy `.env.example` to `.env` and fill it in:

```
TENANT_ID=...
CLIENT_ID=...
CLIENT_SECRET=...
SHAREPOINT_SITE_URL=https://YOUR-TENANT.sharepoint.com/sites/book-test
SHAREPOINT_SITE_ID=...
```

> ⚠️ Never commit `.env` with real credentials (it is already in `.gitignore`).

## 🚀 Installation and build

**C# (.NET):**
```bash
dotnet build la-experiencia-sharepoint.slnx     # 0 errors
dotnet test common/SharePointGraphAuth.Tests    # 5 unit tests
```

**Java (Maven):**
```bash
mvn clean compile                                 # full reactor, 0 errors
mvn -pl common/sharepoint-graph-auth test         # 2 JUnit tests
```

**JavaScript (Node.js):**
```bash
npm install
npx jest                                          # 3 Jest tests
```

**Python:**
```bash
pip install -r requirements.txt
pip install -e common                              # shared auth module
python -m pytest common/tests/                     # 4 pytest tests
```

**PowerShell:**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
pwsh -NoProfile -Command "Invoke-Pester -Path ./common/SharePointGraph.Tests.ps1 -Output Detailed"  # 2 Pester tests
```

**SPFx (cap14) — requires Node 18 LTS:**
```bash
nvm use 18
cd chapter-14-spfx/01-webpart-hello-graph   # any project
npm install
gulp build                                   # 0 warnings, 0 TS errors (all 11 projects)
gulp bundle --ship && gulp package-solution --ship   # -> sharepoint/solution/*.sppkg
```
The 11 SPFx projects in `chapter-14-spfx/` build green (0 warnings). See
`chapter-14-spfx/README.md` for the index of the 12 examples and the Graph
permissions each one declares.

**Teams (cap17):**
```bash
cd chapter-17-teams/02-bot-sharepoint
npm install
npm test          # 4 Jest tests (graphAuth, no network)
npm run build     # tsc, 0 errors
```
`01-tab-app-sharepoint` is a manifest + icons package (no build; `bash package.sh` generates the .zip). `02-bot-sharepoint` is a Node bot (botbuilder + restify) that responds to `sites` / `docs <site>` by reading Graph. See each project's README for the full lifecycle (Entra ID, Azure Bot, ngrok, sideload).

## ▶️ Run the examples

Each chapter has a read-only `main`/`Main` that operates against the test site (resolved by path, default `olddogsoft1.sharepoint.com/sites/book-test` — overridable via `SHAREPOINT_HOSTNAME`/`SHAREPOINT_SITE_PATH`).

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

## 🧪 Automated tests

```bash
bash smoke-tests.sh        # 24 live smoke tests against the tenant (C#/Py/JS/PS/Java)
```

Total suite: **57 tests** (33 network-less unit + 24 live smoke); **57/57 validated green** with least privilege (`Sites.Selected`). See `PRUEBAS-AUTOMATIZADAS.md`.

## 📖 Notes

- **App-only with client secret** authentication (Microsoft Graph) is the supported pattern for non-interactive automation. The chapter 2 examples use **interactive/delegated** auth (browser) to teach the flows.
- The `common/` module centralizes auth (DRY); examples receive the client/token via **dependency injection**.
- `complete-workflow/` (integrating project) is a stub pending porting.

## 📄 License

MIT License (code). The book content is CC BY-NC-SA 4.0.

## 💝 Support the project

This code is free and open source. If it helps you, you can support it in several ways:

- **GitHub Sponsors** — click the **Sponsor** button above (requires the author to have the Sponsors profile enabled under `olddogsoftllc`).
- **Buy the book** — *La Experiencia SharePoint* (4 volumes) on Leanpub (link coming soon; in the meantime at [olddogsoft.com/blog](https://www.olddogsoft.com/blog)).
- **PayPal / Ko-fi** — one-click avenues configured via `.github/FUNDING.yml` (see that file to enable them).
- **Share the repo** — a star ⭐ and sharing with your team also help.

> The code will always be free and MIT. The book (editorial content) is paid to sustain the project. Your support funds more examples, live validation, and maintenance against Microsoft Graph changes.