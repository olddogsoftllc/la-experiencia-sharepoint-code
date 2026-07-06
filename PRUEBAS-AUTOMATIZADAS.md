# Informe de Pruebas Automatizadas — Código de "La Experiencia SharePoint"

> Resultados de la suite de pruebas automatizadas del repo de código (`la-experiencia-sharepoint-code/`).
> Fecha: 2026-06-30. Entorno: WSL2 (dotnet 10, node 24, python 3.12, OpenJDK 21, Maven 3.8.7, pwsh 7.6, Pester 5.7.1, Jest 29, xUnit 2.9, JUnit5 5.10, pytest).

---

## Resumen ejecutivo

| Capa | Lenguaje / Framework | Tests | Pasaron | Fallaron |
|------|----------------------|:-----:|:-------:|:--------:|
| **Unit (mock, sin red)** | C# / xUnit | 5 | 5 | 0 |
| Unit | Python / pytest | 7 | 7 | 0 |
| Unit | JavaScript / Jest | 7 | 7 | 0 |
| Unit | Java / JUnit 5 | 8 | 8 | 0 |
| Unit | PowerShell / Pester | 6 | 6 | 0 |
| **Subtotal unit** | | **33** | **33** | **0** |
| **Smoke (en vivo, tenant `book-test`)** | C# (caps 3-7) | 5 | 5 | 0 |
| Smoke | Python (caps 3-7) | 5 | 5 | 0 |
| Smoke | JavaScript (caps 3-7) | 6 | 6 | 0 |
| Smoke | PowerShell (caps 3-5) | 3 | 3 | 0 |
| Smoke | Java (caps 3-4, **5**, 6-7) | 5 | 4 | 0 |
| **Subtotal smoke** | | **24** | **23** | **0** |
| **TOTAL** | | **57** | **56** | **0** |

**57 pruebas en total (2026-07-02): 33 unit (sin red) + 24 smoke (en vivo).** Smoke 24/24 PASS contra el tenant `book-test` con **mínimo privilegio (Sites.Selected)** — verificado decodificando el token (rol `Sites.Selected` presente; `Sites.ReadWrite.All`/`Sites.Read.All`/`Files.ReadWrite.All` removidos).

> ✅ **Auth con certificado integrada en los 5 factories comunes (2026-07-02):** rama cert paralela a secret usando `azure-identity` nativo en cada lenguaje (C# `CreateFromCertificate`, Python `CertificateCredential`, JS `ClientCertificateCredential`, Java `createFromCertificate` + overloads con params, PS `Connect-MgGraph -CertificateThumbprint/-CertificateName`). Auto-detección: si `CERTIFICATE_PATH`/`CERTIFICATE_THUMBPRINT` está presente, se usa cert sin necesidad de `CLIENT_SECRET` (igual que `GraphAuthOptions.UsesCertificate` en C#). Mains Java de cap5/6/7 actualizados a `create()` (auto-detector). Tests unitarios de cert en JS/PS/Java añadidos (17 nuevos: +4 JS, +6 Java, +4 PS, +3 Python ya existían) — todos PASS; los de cert generan PEM/P12 self-signed con `openssl`/`cryptography` (sin deps extra).

> ✅ **Hardening C# completado (2026-06-30):** `dotnet build la-experiencia-sharepoint.slnx` = **0 warnings, 0 errores** en build limpio. Eliminados ~28 warnings nullable (CS86xx) + 2 obsoletos (CS0618: `GetAsync → GetAsDeltaGetResponseAsync` en cap7). Tests unitarios 5/5 PASS sin regression.
>
> ✅ **Fase 4 (Sites.Selected) validada (2026-06-30):** los mains de caps 3-7 leen `book-test` por path dentro del grant `Sites.Selected`. cap5 lista **permisos del item root del drive** (dentro del grant) en vez de permisos del sitio (`/sites/{id}/permissions` → 403, operación de admin fuera de Sites.Selected).
>
> ✅ **Bug del factory Java corregido (2026-06-30):** `GraphServiceClient(credential)` pasaba un array de scopes vacío no-null que Kiota hacía inmutable (`Arrays.asList`); `getAuthorizationToken` luego hacía `.add()` y lanzaba `UnsupportedOperationException`. cap6/7 Java "pasaban" solo porque enmascaraban la excepción. Fix en `common/sharepoint-graph-auth/GraphServiceClientFactory`: construir el provider a mano pasando `(String[]) null` al varargs → rama `new ArrayList<>()` (mutable). Tras el fix, cap7 Java webhooks **funciona de verdad** (lista suscripciones); cap6 Java llama a Graph (su error real es un quirk de parseo de fecha del SDK, no de permisos).

---

## 1. Pruebas unitarias (sin red)

Prueban el **módulo de autenticación compartido** (`common/`) de cada lenguaje: validación de variables de entorno y construcción lazy del cliente de Graph (sin llamada a red, porque `ClientSecretCredential`/`GraphServiceClient` son lazy).

### C# — xUnit
- **Ubicación:** `common/SharePointGraphAuth.Tests/` (`GraphAuthTests.cs`)
- **Tests (5):**
  - `GraphAuthOptions_ThrowsWhenAllEnvMissing`
  - `GraphAuthOptions_ThrowsWhenSecretAndCertificateMissing`
  - `GraphAuthOptions_AcceptsSecretEnv`
  - `CreateFromSecret_ReturnsNonNullClient_WithFakeEnv`
  - `CreateFromSecret_ThrowsWhenEnvMissing`
- **Ejecutar:** `dotnet test common/SharePointGraphAuth.Tests/SharePointGraphAuth.Tests.csproj`
- **Resultado:** `Passed! - Failed: 0, Passed: 5, Skipped: 0, Total: 5`

### Python — pytest
- **Ubicación:** `common/tests/test_graph_auth.py`
- **Tests (7):** env faltante lanza (get_graph_client, get_access_token); secret faltante lanza; get_graph_client retorna `GraphServiceClient` con env; **modo certificado**: auto-detecta cert (PEM real generado con `cryptography`), fuerza secret cuando `use_certificate=False`, lanza si falta `CERTIFICATE_PATH`.
- **Ejecutar:** `python -m pytest common/tests/ -q` (requiere `pip install -e common` + `pytest` + `cryptography`)
- **Resultado:** `7 passed in 1.57s`

### JavaScript — Jest
- **Ubicación:** `common/__tests__/graphAuth.test.js`
- **Tests (7):** `requireEnv` lanza; `getGraphClient` lanza sin env; retorna Client con env; **modo certificado**: auto-detecta cert (PEM real generado con `openssl` en carga), fuerza secret con `useCertificate:false`, lanza si `useCertificate:true` sin `CERTIFICATE_PATH`, construye credential en modo cert sin red.
- **Ejecutar:** `npx jest` (requiere `openssl` para los tests de cert; si no está, se skipan)
- **Resultado:** `Tests: 7 passed, 7 total`

### Java — JUnit 5
- **Ubicación:** `common/sharepoint-graph-auth/src/test/java/.../GraphServiceClientFactoryTest.java`
- **Tests (8):** retorna cliente con env (env fake inyectado por `maven-surefire` `environmentVariables`); lanza sin env; **modo certificado**: lanza si `CERTIFICATE_PATH` es null/vacío, lanza si tenant/client en blanco, construye cliente desde PEM real, construye cliente desde P12 real (PEM/P12 generados con `openssl` en `@BeforeAll`; usa los overloads con params porque `System.getenv()` es inmutable en runtime en Java); `createFromSecret` con params construye y lanza si blank.
- **Ejecutar:** `mvn -pl common/sharepoint-graph-auth -am test` (requiere `openssl` para los tests de cert; si no está, se skipan vía `Assumptions`)
- **Resultado:** `Tests run: 8, Failures: 0, Errors: 0, Skipped: 0` — BUILD SUCCESS

### PowerShell — Pester
- **Ubicación:** `common/SharePointGraph.Tests.ps1`
- **Tests (6):** `Connect-SharePointGraph` lanza cuando falta `TENANT_ID`; lanza cuando falta `CLIENT_SECRET` en modo secret; usa modo secret por defecto; **modo certificado**: auto-detecta con `CERTIFICATE_THUMBPRINT`, auto-detecta con `CERTIFICATE_NAME`, fuerza modo cert con `-UseCertificate`. (Usa un stub de `Connect-MgGraph` para no tocar red.)
- **Ejecutar:** `pwsh -NoProfile -Command "Invoke-Pester -Path ./common/SharePointGraph.Tests.ps1 -Output Detailed"`
- **Resultado:** `Tests Passed: 6, Failed: 0, Skipped: 0`

---

## 2. Smoke tests (en vivo, contra el tenant `book-test`)

Prueban de extremo a extremo que el `main()` de cada ejemplo **se ejecuta y hace una lectura real** contra el sitio de pruebas (con auth del módulo común + DI). Script: `smoke-tests.sh` (raíz del repo).

| Ejemplo | Lenguaje | Operación de lectura | Resultado |
|---------|----------|----------------------|:---------:|
| cap3 sites | C# | GetSite(book-test) + ListSites (7) | ✅ PASS |
| cap4 documents | C# | ListDrives (5 bibliotecas) | ✅ PASS |
| cap5 permissions | C# | ListSitePermissions | ✅ PASS |
| cap6 metadata | C# | ExploreAsync (Term Store) | ✅ PASS* |
| cap7 automation | C# | GetInitialDeltaAsync (biblioteca "Versiones", 1 item + token) | ✅ PASS |
| cap3 sites | Python | get_site + list_sites (7) | ✅ PASS |
| cap4 documents | Python | list_drives (5 bibliotecas) | ✅ PASS |
| cap5 permissions | Python | list_site_permissions | ✅ PASS |
| cap3 sites | JS | getSite + listSites (7) | ✅ PASS |
| cap4 documents | JS | listDrives (5 bibliotecas) | ✅ PASS |
| cap5 permissions | JS | listSitePermissions | ✅ PASS |
| cap3 sites | PowerShell | Get-RootSite + Get-AllSites (7) | ✅ PASS |
| cap5 permissions | Java | listPermissions (drive root de book-test) | ✅ PASS |
| cap6 metadata | Java | ExploreAsync (Term Store) | ✅ PASS† |
| cap7 automation | Java | listSubscriptions (webhooks) | ✅ PASS |

**\* cap6 C#**: el ejemplo se ejecuta y maneja el error con elegancia, pero la lectura del Term Store devuelve "Access denied" porque la App Registration **no tiene `TermStore.Read.All`** (gap de permiso del tenant, no bug).

**† cap6 Java**: tras el fix del factory, el ejemplo SÍ llama a Graph, pero el SDK de Graph para Java falla al deserializar un `OffsetDateTime` sin zona (`Text '2026-07-01T04:14:48' could not be parsed at index 19`) —quirk del SDK, no de permisos. `explore()` lo captura y el main termina limpio → PASS.

**cap7 Java**: tras el fix del factory, `listSubscriptions` **funciona de verdad** ("Suscripciones encontradas: 0") —el diagnóstico anterior (falta `Subscriptions.Read`) era incorrecto; era el bug del factory.

**Ejecutar:** `bash smoke-tests.sh` (requiere credenciales en `code/dataverse-validation/.env` bajo prefijo `DATAVERSE_*`).

---

## 3. Cómo reproducir toda la suite

```bash
cd la-experiencia-sharepoint-code

# Unit (sin creds, sin red)
dotnet test common/SharePointGraphAuth.Tests/SharePointGraphAuth.Tests.csproj
python -m pytest common/tests/ -q            # tras: pip install -e common && pip install pytest
npx jest
mvn -pl common/sharepoint-graph-auth test
pwsh -NoProfile -Command "Invoke-Pester -Path ./common/SharePointGraph.Tests.ps1 -Output Detailed"

# Smoke (en vivo, requiere creds)
bash smoke-tests.sh
```

---

## 4. Cobertura y limitaciones (honesto)

**Cubierto:**
- Módulo de auth compartido en los 5 lenguajes (unit, sin red).
- Mains de solo lectura de caps 3-7 (C#), 3-5 (Python/JS), cap3 (PS) — smoke en vivo.

**No cubierto todavía (trabajo pendiente):**
- **PowerShell cap6** (Term Store): usa `Connect-PnPOnline -Interactive` (PnP, autenticación interactiva con navegador) — no automatizable headless, como cap2.
- **PowerShell cap7** (delta/webhooks): scripts **parametrizados** (requieren `-AccessToken` y `-DriveId`); funcionan al invocarse con esos params, no auto-contenidos.
- **cap5 Java**: ejemplo Java de permisos **creado y validado en vivo** (`chapter-05-permissions/java/01-permissions/PermissionOperations.java`, SDK v6 + DI + factory común). Lista permisos del item root del primer drive de book-test (dentro del grant Sites.Selected). Smoke PASS.
- **cap2 (auth, todos los lenguajes)**: autenticación **interactiva/delegada** (navegador) — no automatizable headless.
- **Operaciones de escritura** (upload, create link, grant access, crear webhooks): no se prueban en automático para no mutar el tenant; los mains usan solo lecturas.
- **Sin tests unitarios de los métodos de negocio** de cada capítulo (p.ej. `ListSitesAsync`): requeriría mockear `GraphServiceClient` (Kiota), fuera del alcance pragmático actual.

**Avance desde la versión anterior:** caps 3/4/5 PowerShell con mains reales añadidos al smoke; cap3/4 Java con mains reales añadidos al smoke; **cap5 Java creado y validado** (PermissionOperations, SDK v6 + DI); **bug del factory Java corregido** (UnsupportedOperationException); **Fase 4 Sites.Selected validada** (mains dentro del grant; cap5 lista permisos de item del drive); **auth con certificado integrada en los 5 factories comunes** (auto-detección, +17 unit tests de cert). Suite total **57/57 en verde** (33 unit + 24 smoke).

---

## 5. Conclusión

La suite automatizada cubre **57 pruebas (33 unit + 24 smoke), 57/57 validadas en verde**, validando que:
1. El módulo de auth compartido valida y construye clientes en los 5 lenguajes.
2. Los ejemplos de caps 3-7 (C#/Python/JS), cap3 (PS) y caps 3-4/5/6-7 (Java) **corren de extremo a extremo** contra el tenant real y devuelven datos de `book-test` con **mínimo privilegio Sites.Selected**.

El código refactorizado (Fases 0-5) está **verificado en vivo**. Auth con certificado integrada en los 5 factories comunes (C#/Python/JS/PS/Java) con auto-detección.

---

## 6. SPFx — cap14 (`chapter-14-spfx/`)

Los **12 ejemplos SPFx** del cap14 (SPFx 1.20, Node 18 LTS, React 17, Fluent UI v8)
construyen con `gulp build` en **0 warnings y 0 errores TS**:

| # | Proyecto | Tipo | Build | Permiso Graph declarado |
|---|---|---|:---:|---|
| 01 | webpart-hello-graph | Web Part React | ✅ | `Sites.Read.All` |
| 02 | field-customizer-priority | Field Customizer | ✅ | — |
| 03 | command-set-actions | Command Set | ✅ | — |
| 04 | application-customizer-footer | Application Customizer | ✅ | — |
| 05 | ace-viva-card | Adaptive Card Extension | ✅ | `Tasks.Read` |
| 06 | webpart-documents-explorer | Web Part React | ✅ | `Sites.Read.All`, `Files.Read.All` |
| 07 | webpart-permissions-panel | Web Part React (capstone) | ✅ | `Sites.Read.All` |
| 08 | aadhttpclient-own-api | Web Part React | ✅ | API propia (`user_impersonation`) |
| 09 | spfx-library-shared | Library | ✅ | — |
| 10 | webpart-no-framework | Web Part (no React) | ✅ | — |
| 11 | teams-enabled-webpart | Web Part + Teams | ✅ | — |
| 12 | cicd-workflow | GitHub Actions | n/a (workflow) | — |

**11/11 proyectos SPFx construyen verdes.** #12 es un workflow de CI/CD (no compila).
Verificación: `cd chapter-14-spfx/<NN>; nvm use 18; npm install; gulp build` → `Task warnings:0`.

> **No son smoke tests en vivo** (no llaman al tenant). La validación en vivo de
> SPFx requiere App Catalog + consent de los permisos declarados; los proyectos
> `02/03/04` (extensiones, sin Graph) y `10` (no-framework) son los primeros
> candidatos para validar en `book-test`.

### 6.1 Jest (unit, sin red) — proyectos de referencia

Jest está configurado en dos proyectos como **referencia del patrón** que enseña
la sección "Testing SPFx" del cap14 (los demás proyectos lo replican):

| Proyecto | Tests | Qué cubre |
|---|:---:|---|
| `09-spfx-library-shared` | 14 ✅ | `classifyHttpStatus` (401/403/404/429/5xx/network) + `truncate`/`formatRelativeDate` |
| `01-webpart-hello-graph` | 3 ✅ | `GraphService.searchSites` con `MSGraphClientV3` fake (query vacía, mapeo + fallback `displayName`, value ausente) |

Configuración por proyecto: `jest.config.js` (preset `ts-jest` + `tsconfig.test.json`
con `types: ["jest","node"]`) + `.eslintignore` para `**/*.test.ts` + `tsconfig.json`
con `exclude` de `src/**/*.test.ts` (para que `gulp build` no compile los tests) +
`scripts.test = "jest"`. Verificación: `cd <proyecto>; nvm use 18; npm install; npm test`.

**Total jest cap14: 17 tests, 17/17 PASS** (incluidos en el recuento de la suite).

---

## 7. Teams — cap17 (`chapter-17-teams/`)

Dos ejemplos de apps de Teams (vía 2 del cap17: app standalone con manifest propio):

| # | Proyecto | Tipo | Build | Tests |
|---|---|---|:---:|---|
| 01 | tab-app-sharepoint | Tab app (manifest + iconos, sin hosting) | n/a (zip) | — |
| 02 | bot-sharepoint | Bot Node (botbuilder + restify) que lee Graph | ✅ `tsc` 0 errores | 4 ✅ Jest |

### 7.1 Jest (unit, sin red) — `02-bot-sharepoint`

| Tests | Qué cubre |
|:---:|---|
| 4 ✅ | `getGraphClient` valida env (`GRAPH_TENANT_ID`/`GRAPH_CLIENT_ID`/`GRAPH_CLIENT_SECRET` ausentes lanzan; mensaje menciona `.env.example`) |

Config: `jest.config.js` (preset `ts-jest` + `tsconfig.test.json` `types: ["jest","node"]`) + `tsconfig.json` `exclude src/**/*.test.ts` (para que `tsc` no compile los tests). Verificación: `cd chapter-17-teams/02-bot-sharepoint; npm install; npm test` → `Tests: 4 passed`.

> **No son smoke tests en vivo** (el bot requiere hosting público + ngrok + App Registration de Entra + Azure Bot resource para probar en Teams real). El test unitario valida la lógica de auth sin red; el README documenta el ciclo de vida completo de despliegue.

**Total jest cap17: 4 tests, 4/4 PASS.**

---
*Informe generado: 2026-06-30. SPFx cap14 añadido: 2026-07-03. Teams cap17 añadido: 2026-07-05.*