# Graph Report - .  (2026-08-17)

## Corpus Check
- 387 files · ~81,205 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 1457 nodes · 1787 edges · 133 communities (96 shown, 37 thin omitted)
- Extraction: 99% EXTRACTED · 1% INFERRED · 0% AMBIGUOUS · INFERRED: 25 edges (avg confidence: 0.75)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- other — siteoperations (53)
- other — eslintrc (47)
- other — eslintrc (47)
- other — eslintrc (47)
- other — eslintrc (47)
- other — eslintrc (47)
- other — eslintrc (47)
- other — eslintrc (47)
- other — eslintrc (47)
- other — eslintrc (47)
- other — eslintrc (47)
- other — eslintrc (47)
- common (shared auth factories) (37)
- other — graphauthexample (36)
- other — ipermissionspanelwebpartprops (35)
- other — site (31)
- other — documentsexplorerwebpart (29)
- other — any (28)
- other — cardview (27)
- other — hellographwebpart (26)
- other — iownapiwebpartprops (25)
- other — operations (20)
- other — operations (20)
- other — flow (19)
- other — certificate (19)
- other — example (18)
- other — operations (18)
- other — iteamsenabledwebpartprops (18)
- other — authwithtokencache (17)
- other — queries (16)
- common (shared auth factories) (16)
- other — documentoperations (15)
- other — authwithcertificate (14)
- other — noframeworkutils (14)
- other — authwithtokencache (13)
- other — permissionoperations (13)
- other — permissionoperations (13)
- other — graphauthexample (12)
- other — siteoperations (12)
- other — deltaqueries (12)
- other — footerapplicationcustomizerapp (12)
- other — index (12)
- other — authwithcertificate (11)
- other — priorityfieldcustomizerfieldcu (11)
- other — actionscommandsetcommandset (11)
- other — authwithcertificate (10)
- other — webhookmanager (10)
- other — webhookmanager (10)
- other — authwithtokencache (9)
- other — deltaquerymanager (9)
- other — webhookmanager (9)
- other — documentoperations (8)
- other — permissionoperations (8)
- other — tests (8)
- other — siteoperations (7)
- other — deltachanges (7)
- other — manager (7)
- other — contenttypewithmetadata (6)
- other — termstoreexplorer (6)
- other — exploretermstore (6)
- other — authwithcertificate (5)
- other — authwithtokencache (5)
- other — graphauthexample (4)
- common (shared auth factories) (4)
- other — flow (3)
- other — d (3)
- other — d (3)
- other — d (3)
- other — d (3)
- other — d (3)
- other — d (3)
- other — d (3)
- other — d (3)
- other — d (3)
- other — d (3)
- other — d (3)
- other — gulpfile (2)
- other — scss (2)
- other — gulpfile (2)
- other — scss (2)
- other — gulpfile (2)
- other — gulpfile (2)
- other — gulpfile (2)
- other — gulpfile (2)
- other — scss (2)
- other — gulpfile (2)
- other — scss (2)
- other — gulpfile (2)
- other — scss (2)
- other — gulpfile (2)
- other — gulpfile (2)
- other — scss (2)
- other — gulpfile (2)
- other — scss (2)
- other — package (2)

## God Nodes (most connected - your core abstractions)
1. `get_graph_client()` - 14 edges
2. `PermissionOperations` - 13 edges
3. `SiteOperations` - 11 edges
4. `DocumentOperations` - 11 edges
5. `DocumentOperations` - 11 edges
6. `PermissionOperations` - 11 edges
7. `AuthWithTokenCache` - 10 edges
8. `SiteOperations` - 10 edges
9. `DeltaQueryManager` - 10 edges
10. `OwnApiWebPartWebPart` - 10 edges

## Surprising Connections (you probably didn't know these)
- `getGraphClient()` --references--> `{ ClientSecretCredential }`  [EXTRACTED]
  chapter-17-teams/02-bot-sharepoint/src/graphAuth.ts → chapter-02-auth/javascript/auth-flow.js
- `main()` --calls--> `get_graph_client()`  [INFERRED]
  chapter-06-metadata/python/01-termstore/explore_term_store.py → common/laexperiencia_sharepoint/graph_auth.py
- `test_get_access_token_raises_when_env_missing()` --calls--> `get_access_token()`  [INFERRED]
  common/tests/test_graph_auth.py → common/laexperiencia_sharepoint/graph_auth.py
- `test_get_graph_client_cert_raises_when_path_missing()` --calls--> `get_graph_client()`  [INFERRED]
  common/tests/test_graph_auth.py → common/laexperiencia_sharepoint/graph_auth.py
- `test_get_graph_client_forces_secret_when_cert_env_present()` --calls--> `get_graph_client()`  [INFERRED]
  common/tests/test_graph_auth.py → common/laexperiencia_sharepoint/graph_auth.py

## Import Cycles
- None detected.

## Communities (133 total, 37 thin omitted)

### Community 0 - "other — siteoperations (53)"
Cohesion: 0.06
Nodes (22): GraphServiceClient, Task, SiteOperations, DriveItem, GraphServiceClient, Task, DocumentOperations, GraphServiceClient (+14 more)

### Community 1 - "other — eslintrc (47)"
Cohesion: 0.04
Nodes (46): RATIONALE: The "module" keyword is deprecated except when describing legacy…, RATIONALE: This rule warns if setters are defined without getters, which is…, RATIONALE: In TypeScript, if you write x["y"] instead of x.y, it disables type…, RATIONALE: Catches code that is likely to be incorrect, RATIONALE: Catches a common coding mistake., RATIONALE: If you have more than 2,000 lines in a single source file, it's…, RATIONALE: Deprecated language feature., RATIONALE: Catches code that is likely to be incorrect (+38 more)

### Community 2 - "other — eslintrc (47)"
Cohesion: 0.04
Nodes (46): RATIONALE: The "module" keyword is deprecated except when describing legacy…, RATIONALE: This rule warns if setters are defined without getters, which is…, RATIONALE: In TypeScript, if you write x["y"] instead of x.y, it disables type…, RATIONALE: Catches code that is likely to be incorrect, RATIONALE: Catches a common coding mistake., RATIONALE: If you have more than 2,000 lines in a single source file, it's…, RATIONALE: Deprecated language feature., RATIONALE: Catches code that is likely to be incorrect (+38 more)

### Community 3 - "other — eslintrc (47)"
Cohesion: 0.04
Nodes (46): RATIONALE: The "module" keyword is deprecated except when describing legacy…, RATIONALE: This rule warns if setters are defined without getters, which is…, RATIONALE: In TypeScript, if you write x["y"] instead of x.y, it disables type…, RATIONALE: Catches code that is likely to be incorrect, RATIONALE: Catches a common coding mistake., RATIONALE: If you have more than 2,000 lines in a single source file, it's…, RATIONALE: Deprecated language feature., RATIONALE: Catches code that is likely to be incorrect (+38 more)

### Community 4 - "other — eslintrc (47)"
Cohesion: 0.04
Nodes (46): RATIONALE: The "module" keyword is deprecated except when describing legacy…, RATIONALE: This rule warns if setters are defined without getters, which is…, RATIONALE: In TypeScript, if you write x["y"] instead of x.y, it disables type…, RATIONALE: Catches code that is likely to be incorrect, RATIONALE: Catches a common coding mistake., RATIONALE: If you have more than 2,000 lines in a single source file, it's…, RATIONALE: Deprecated language feature., RATIONALE: Catches code that is likely to be incorrect (+38 more)

### Community 5 - "other — eslintrc (47)"
Cohesion: 0.04
Nodes (46): RATIONALE: The "module" keyword is deprecated except when describing legacy…, RATIONALE: This rule warns if setters are defined without getters, which is…, RATIONALE: In TypeScript, if you write x["y"] instead of x.y, it disables type…, RATIONALE: Catches code that is likely to be incorrect, RATIONALE: Catches a common coding mistake., RATIONALE: If you have more than 2,000 lines in a single source file, it's…, RATIONALE: Deprecated language feature., RATIONALE: Catches code that is likely to be incorrect (+38 more)

### Community 6 - "other — eslintrc (47)"
Cohesion: 0.04
Nodes (46): RATIONALE: The "module" keyword is deprecated except when describing legacy…, RATIONALE: This rule warns if setters are defined without getters, which is…, RATIONALE: In TypeScript, if you write x["y"] instead of x.y, it disables type…, RATIONALE: Catches code that is likely to be incorrect, RATIONALE: Catches a common coding mistake., RATIONALE: If you have more than 2,000 lines in a single source file, it's…, RATIONALE: Deprecated language feature., RATIONALE: Catches code that is likely to be incorrect (+38 more)

### Community 7 - "other — eslintrc (47)"
Cohesion: 0.04
Nodes (46): RATIONALE: The "module" keyword is deprecated except when describing legacy…, RATIONALE: This rule warns if setters are defined without getters, which is…, RATIONALE: In TypeScript, if you write x["y"] instead of x.y, it disables type…, RATIONALE: Catches code that is likely to be incorrect, RATIONALE: Catches a common coding mistake., RATIONALE: If you have more than 2,000 lines in a single source file, it's…, RATIONALE: Deprecated language feature., RATIONALE: Catches code that is likely to be incorrect (+38 more)

### Community 8 - "other — eslintrc (47)"
Cohesion: 0.04
Nodes (46): RATIONALE: The "module" keyword is deprecated except when describing legacy…, RATIONALE: This rule warns if setters are defined without getters, which is…, RATIONALE: In TypeScript, if you write x["y"] instead of x.y, it disables type…, RATIONALE: Catches code that is likely to be incorrect, RATIONALE: Catches a common coding mistake., RATIONALE: If you have more than 2,000 lines in a single source file, it's…, RATIONALE: Deprecated language feature., RATIONALE: Catches code that is likely to be incorrect (+38 more)

### Community 9 - "other — eslintrc (47)"
Cohesion: 0.04
Nodes (46): RATIONALE: The "module" keyword is deprecated except when describing legacy…, RATIONALE: This rule warns if setters are defined without getters, which is…, RATIONALE: In TypeScript, if you write x["y"] instead of x.y, it disables type…, RATIONALE: Catches code that is likely to be incorrect, RATIONALE: Catches a common coding mistake., RATIONALE: If you have more than 2,000 lines in a single source file, it's…, RATIONALE: Deprecated language feature., RATIONALE: Catches code that is likely to be incorrect (+38 more)

### Community 10 - "other — eslintrc (47)"
Cohesion: 0.04
Nodes (46): RATIONALE: The "module" keyword is deprecated except when describing legacy…, RATIONALE: This rule warns if setters are defined without getters, which is…, RATIONALE: In TypeScript, if you write x["y"] instead of x.y, it disables type…, RATIONALE: Catches code that is likely to be incorrect, RATIONALE: Catches a common coding mistake., RATIONALE: If you have more than 2,000 lines in a single source file, it's…, RATIONALE: Deprecated language feature., RATIONALE: Catches code that is likely to be incorrect (+38 more)

### Community 11 - "other — eslintrc (47)"
Cohesion: 0.04
Nodes (46): RATIONALE: The "module" keyword is deprecated except when describing legacy…, RATIONALE: This rule warns if setters are defined without getters, which is…, RATIONALE: In TypeScript, if you write x["y"] instead of x.y, it disables type…, RATIONALE: Catches code that is likely to be incorrect, RATIONALE: Catches a common coding mistake., RATIONALE: If you have more than 2,000 lines in a single source file, it's…, RATIONALE: Deprecated language feature., RATIONALE: Catches code that is likely to be incorrect (+38 more)

### Community 12 - "common (shared auth factories) (37)"
Cohesion: 0.09
Nodes (31): main(), GraphServiceClient, TermStoreExplorer, _build_credential(), _env_or_none(), get_access_token(), get_graph_client(), _load_cert_material() (+23 more)

### Community 13 - "other — graphauthexample (36)"
Cohesion: 0.10
Nodes (11): GraphAuthExample, HttpClient, Gson, HttpClient, SiteListResponse, SiteOperations, DocumentOperations, Gson (+3 more)

### Community 14 - "other — ipermissionspanelwebpartprops (35)"
Cohesion: 0.10
Nodes (14): IPermissionsPanelWebPartProps, PermissionsPanel(), IGraphQueryState, useGraphQuery(), ISitePermission, IPermissionsPanelWebPartWebPartProps, PermissionsPanelWebPartWebPart, GraphService (+6 more)

### Community 15 - "other — site (31)"
Cohesion: 0.12
Nodes (12): Site, GraphServiceClient, Permission, PermissionOperations, DriveItem, GraphServiceClient, Task, DeltaQueryManager (+4 more)

### Community 16 - "other — documentsexplorerwebpart (29)"
Cohesion: 0.12
Nodes (12): DocumentsExplorer(), IDocumentsExplorerWebPartProps, DocumentsExplorerWebPartWebPart, IDocumentsExplorerWebPartWebPartProps, IGraphQueryState, useGraphQuery(), IDrive, IDriveItem (+4 more)

### Community 17 - "other — any (28)"
Cohesion: 0.11
Nodes (16): Any, AuthWithTokenCache, main(), auth_with_token_cache.py Chapter 02: Authentication Token Cache Authentication…, Gets authenticated headers with token caching, Clears the token cache, Main execution function, Simple in-memory token cache implementation (+8 more)

### Community 18 - "other — cardview (27)"
Cohesion: 0.14
Nodes (11): CardView, IQuickViewData, QuickView, ITicket, ITicketsAceAdaptiveCardExtensionProps, ITicketsAceAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID, TicketsAceAdaptiveCardExtension (+3 more)

### Community 19 - "other — hellographwebpart (26)"
Cohesion: 0.13
Nodes (10): HelloGraph(), IHelloGraphWebPartProps, HelloGraphWebPartWebPart, IHelloGraphWebPartWebPartProps, IGraphQueryState, useGraphQuery(), ISite, GraphService (+2 more)

### Community 20 - "other — iownapiwebpartprops (25)"
Cohesion: 0.13
Nodes (8): IOwnApiWebPartProps, OwnApiWebPart(), IAsyncState, useAsync(), IOwnApiWebPartWebPartProps, OwnApiWebPartWebPart, OwnApiService, FakeResponse

### Community 21 - "other — operations (20)"
Cohesion: 0.15
Nodes (11): DocumentOperations, main(), document_operations.py Chapter 04: Documents SharePoint Document Operations…, Downloads a file from SharePoint Args: site_id: The site ID drive_id: The drive…, Searches for files across SharePoint Args: query: Search query Returns: Search…, Lists files in a specific folder Args: site_id: The site ID drive_id: The drive…, SharePoint Document Operations handler., Gets file metadata Args: site_id: The site ID drive_id: The drive ID item_id:… (+3 more)

### Community 22 - "other — operations (20)"
Cohesion: 0.15
Nodes (11): main(), PermissionOperations, permission_operations.py Chapter 05: Permissions SharePoint Permission…, Creates an organization sharing link Args: site_id: The site ID drive_id: The…, Grants access to a specific user Args: site_id: The site ID drive_id: The drive…, Lists all permissions for an item Args: site_id: The site ID drive_id: The…, SharePoint Permission Operations handler., Deletes a permission Args: site_id: The site ID drive_id: The drive ID item_id:… (+3 more)

### Community 23 - "other — flow (19)"
Cohesion: 0.14
Nodes (13): { Client }, { ClientSecretCredential }, { TokenCredentialAuthenticationProvider }, getGraphClient(), IGraphDrive, IGraphDriveItem, IGraphSite, IGraphSiteCollection (+5 more)

### Community 24 - "other — certificate (19)"
Cohesion: 0.13
Nodes (12): AuthWithCertificate, CertificateLoader, main(), auth_with_certificate.py Chapter 02: Authentication Certificate-Based…, Certificate-based authentication handler, Validates configuration, Acquires access token using certificate authentication Returns: str: Access…, Gets authenticated headers (+4 more)

### Community 25 - "other — example (18)"
Cohesion: 0.15
Nodes (11): GraphAuthConfig, GraphAuthExample, main(), graph_auth_example.py Chapter 02: Authentication Microsoft Graph Client…, Main execution function, Configuration for Microsoft Graph authentication, Validates that all required environment variables are present, Microsoft Graph Authentication handler using client credentials (+3 more)

### Community 26 - "other — operations (18)"
Cohesion: 0.16
Nodes (10): main(), site_operations.py Chapter 03: Sites SharePoint Site Operations Example…, Searches for sites by keyword., Builds the token with the common module and runs a read-only demo against book-…, SharePoint Site Operations handler., Lists all sites in the organization., Gets a specific site by hostname and site path (e.g. contoso.sharepoint.com /…, Gets site by its unique identifier. (+2 more)

### Community 27 - "other — iteamsenabledwebpartprops (18)"
Cohesion: 0.20
Nodes (7): ITeamsEnabledWebPartProps, TeamsEnabledWebPart(), ITeamsEnabledWebPartWebPartProps, TeamsEnabledWebPartWebPart, hostRunningSuffix(), resolveThemeName(), ThemeName

### Community 28 - "other — authwithtokencache (17)"
Cohesion: 0.15
Nodes (5): AuthWithTokenCache, axios, main(), qs, TokenCache

### Community 29 - "other — queries (16)"
Cohesion: 0.23
Nodes (8): DeltaQueryManager, main(), Consulta delta para items de una lista de SharePoint., Extrae el token delta de la URL., Muestra los tokens delta almacenados., Gestiona consultas delta para sincronizacion incremental., Realiza consulta delta inicial y almacena token., Sincroniza cambios usando token delta almacenado.

### Community 30 - "common (shared auth factories) (16)"
Cohesion: 0.19
Nodes (13): buildCredential(), { Client }, { ClientSecretCredential, ClientCertificateCredential }, envOrNone(), getAccessToken(), getGraphClient(), requireEnv(), ENV_KEYS (+5 more)

### Community 31 - "other — documentoperations (15)"
Cohesion: 0.21
Nodes (6): axios, DocumentOperations, fs, { getAccessToken }, main(), path

### Community 32 - "other — authwithcertificate (14)"
Cohesion: 0.20
Nodes (6): AuthWithCertificate, CertificateLoader, { ConfidentialClientApplication }, crypto, fs, main()

### Community 33 - "other — noframeworkutils (14)"
Cohesion: 0.19
Nodes (4): buildMarkup(), IMarkupOptions, INoFrameworkWebPartWebPartProps, NoFrameworkWebPartWebPart

### Community 34 - "other — authwithtokencache (13)"
Cohesion: 0.26
Nodes (3): AuthWithTokenCache, HttpClient, TokenCacheEntry

### Community 35 - "other — permissionoperations (13)"
Cohesion: 0.27
Nodes (5): GraphServiceClient, Permission, Task, PermissionOperations, LaExperiencia.SharePoint.Chapter05.Permissions

### Community 36 - "other — permissionoperations (13)"
Cohesion: 0.26
Nodes (4): axios, { getAccessToken }, main(), PermissionOperations

### Community 37 - "other — graphauthexample (12)"
Cohesion: 0.23
Nodes (5): axios, GraphAuthConfig, GraphAuthExample, main(), qs

### Community 38 - "other — siteoperations (12)"
Cohesion: 0.27
Nodes (4): axios, { getAccessToken }, main(), SiteOperations

### Community 39 - "other — deltaqueries (12)"
Cohesion: 0.33
Nodes (4): axios, DeltaQueryManager, { getAccessToken }, main()

### Community 40 - "other — footerapplicationcustomizerapp (12)"
Cohesion: 0.30
Nodes (6): FooterApplicationCustomizerApplicationCustomizer, IFooterApplicationCustomizerApplicationCustomizerProperties, DEFAULT_FOOTER_TEXT, footerBarStyle(), IFooterBarStyle, resolveFooterText()

### Community 41 - "other — index (12)"
Cohesion: 0.29
Nodes (5): SharedLibraryLibrary, formatRelativeDate(), truncate(), classifyHttpStatus(), HttpStatusCategory

### Community 43 - "other — priorityfieldcustomizerfieldcu (11)"
Cohesion: 0.29
Nodes (5): IPriorityFieldCustomizerFieldCustomizerProperties, PriorityFieldCustomizerFieldCustomizer, buildBadgeText(), getPriorityClass(), PriorityClass

### Community 44 - "other — actionscommandsetcommandset (11)"
Cohesion: 0.35
Nodes (5): ActionsCommandSetCommandSet, IActionsCommandSetCommandSetProperties, buildExportMessage(), buildNotifyMessage(), shouldShowExport()

### Community 45 - "other — authwithcertificate (10)"
Cohesion: 0.29
Nodes (5): GraphServiceClient, string, AuthWithCertificate, SharePointExperience.Chapter02, X509Certificate2

### Community 46 - "other — webhookmanager (10)"
Cohesion: 0.24
Nodes (6): WebhookManager, GraphServiceClient, string, Subscription, Task, AutomationExamples

### Community 47 - "other — webhookmanager (10)"
Cohesion: 0.31
Nodes (3): GraphServiceClient, Subscription, WebhookManager

### Community 48 - "other — authwithtokencache (9)"
Cohesion: 0.22
Nodes (5): AuthWithTokenCache, Task, GraphAuthExample, Task, Chapter02Auth

### Community 50 - "other — webhookmanager (9)"
Cohesion: 0.25
Nodes (3): { getGraphClient }, main(), WebhookManager

### Community 51 - "other — documentoperations (8)"
Cohesion: 0.46
Nodes (7): Connect-GraphApi(), Find-SharePointFiles(), Get-DriveFiles(), Get-DriveLibraries(), Get-FileFromSharePoint(), Get-FileMetadata(), Send-FileToSharePoint()

### Community 52 - "other — permissionoperations (8)"
Cohesion: 0.46
Nodes (7): Connect-GraphApi(), Get-ItemPermissions(), Get-SitePermissions(), Grant-UserAccess(), New-AnonymousSharingLink(), New-OrganizationSharingLink(), Remove-Permission()

### Community 53 - "other — tests (8)"
Cohesion: 0.29
Nodes (7): CLIENT_ID, CLIENT_SECRET, run_smoke(), smoke-tests.sh script, SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH, TENANT_ID

### Community 54 - "other — siteoperations (7)"
Cohesion: 0.52
Nodes (6): Connect-GraphApi(), Get-AllSites(), Get-RootSite(), Get-SiteById(), Get-SiteByPath(), Search-Sites()

### Community 55 - "other — deltachanges (7)"
Cohesion: 0.62
Nodes (5): Extract-DeltaToken(), Get-DeltaChanges(), Get-GraphHeaders(), Get-InitialDelta(), Get-ListDelta()

### Community 57 - "other — contenttypewithmetadata (6)"
Cohesion: 0.33
Nodes (4): GraphServiceClient, Task, ContentTypeWithMetadata, ManagedMetadataExamples

### Community 59 - "other — exploretermstore (6)"
Cohesion: 0.40
Nodes (3): { getGraphClient }, main(), TermStoreExplorer

### Community 60 - "other — authwithcertificate (5)"
Cohesion: 0.70
Nodes (4): Get-CertificateFromFile(), Get-CertificateFromStore(), Get-GraphTokenWithCertificate(), Test-Certificate()

### Community 62 - "other — graphauthexample (4)"
Cohesion: 0.83
Nodes (3): Connect-GraphWithToken(), Get-GraphAccessToken(), Test-GraphConnection()

## Knowledge Gaps
- **116 isolated node(s):** `SharePointExperience.Chapter02`, `{ Client }`, `{ TokenCredentialAuthenticationProvider }`, `fs`, `crypto` (+111 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **37 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `LaExperiencia.SharePoint.Common` connect `other — siteoperations (53)` to `other — permissionoperations (13)`, `other — site (31)`?**
  _High betweenness centrality (0.003) - this node is a cross-community bridge._
- **Are the 7 inferred relationships involving `get_graph_client()` (e.g. with `main()` and `test_get_graph_client_auto_detects_cert()`) actually correct?**
  _`get_graph_client()` has 7 INFERRED edges - model-reasoned connections that need verification._
- **What connects `SharePointExperience.Chapter02`, `{ Client }`, `{ TokenCredentialAuthenticationProvider }` to the rest of the system?**
  _116 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `other — siteoperations (53)` be split into smaller, more focused modules?**
  _Cohesion score 0.062409288824383166 - nodes in this community are weakly interconnected._
- **Should `other — eslintrc (47)` be split into smaller, more focused modules?**
  _Cohesion score 0.0425531914893617 - nodes in this community are weakly interconnected._
- **Should `other — eslintrc (47)` be split into smaller, more focused modules?**
  _Cohesion score 0.0425531914893617 - nodes in this community are weakly interconnected._
- **Should `other — eslintrc (47)` be split into smaller, more focused modules?**
  _Cohesion score 0.0425531914893617 - nodes in this community are weakly interconnected._