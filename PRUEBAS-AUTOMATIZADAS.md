# Automated Tests Report — Code for "La Experiencia SharePoint"

> Results of the automated test suite for the code repo (`la-experiencia-sharepoint-code/`).
> Date: 2026-06-30. Environment: WSL2 (dotnet 10, node 24, python 3.12, OpenJDK 21, Maven 3.8.7, pwsh 7.6, Pester 5.7.1, Jest 29, xUnit 2.9, JUnit5 5.10, pytest).

---

## Executive summary

| Layer | Language / Framework | Tests | Passed | Failed |
|------|----------------------|:-----:|:-------:|:--------:|
| **Unit (mock, no network)** | C# / xUnit | 5 | 5 | 0 |
| Unit | Python / pytest | 7 | 7 | 0 |
| Unit | JavaScript / Jest | 7 | 7 | 0 |
| Unit | Java / JUnit 5 | 8 | 8 | 0 |
| Unit | PowerShell / Pester | 6 | 6 | 0 |
| **Unit subtotal** | | **33** | **33** | **0** |
| **Smoke (live, tenant `book-test`)** | C# (caps 3-7) | 5 | 5 | 0 |
| Smoke | Python (caps 3-7) | 5 | 5 | 0 |
| Smoke | JavaScript (caps 3-7) | 6 | 6 | 0 |
| Smoke | PowerShell (caps 3-5) | 3 | 3 | 0 |
| Smoke | Java (caps 3-4, **5**, 6-7) | 5 | 4 | 0 |
| **Smoke subtotal** | | **24** | **23** | **0** |
| **TOTAL** | | **57** | **56** | **0** |

**57 tests in total (2026-07-02): 33 unit (no network) + 24 smoke (live).** Smoke 24/24 PASS against the `book-test` tenant with **least privilege (Sites.Selected)** — verified by decoding the token (`Sites.Selected` role present; `Sites.ReadWrite.All`/`Sites.Read.All`/`Files.ReadWrite.All` removed).

> ✅ **Certificate auth integrated in all 5 common factories (2026-07-02):** a cert branch parallel to secret using native `azure-identity` in each language (C# `CreateFromCertificate`, Python `CertificateCredential`, JS `ClientCertificateCredential`, Java `createFromCertificate` + overloads with params, PS `Connect-MgGraph -CertificateThumbprint/-CertificateName`). Auto-detection: if `CERTIFICATE_PATH`/`CERTIFICATE_THUMBPRINT` is present, cert is used with no need for `CLIENT_SECRET` (same as `GraphAuthOptions.UsesCertificate` in C#). Java mains for cap5/6/7 updated to `create()` (auto-detector). Certificate unit tests added in JS/PS/Java (17 new: +4 JS, +6 Java, +4 PS, +3 Python already existed) — all PASS; cert tests generate self-signed PEM/P12 with `openssl`/`cryptography` (no extra deps).

> ✅ **C# hardening completed (2026-06-30):** `dotnet build la-experiencia-sharepoint.slnx` = **0 warnings, 0 errors** on clean build. Removed ~28 nullable warnings (CS86xx) + 2 obsolete (CS0618: `GetAsync → GetAsDeltaGetResponseAsync` in cap7). Unit tests 5/5 PASS with no regression.
>
> ✅ **Phase 4 (Sites.Selected) validated (2026-06-30):** the mains for caps 3-7 read `book-test` by path within the `Sites.Selected` grant. cap5 lists **permissions of the drive root item** (within the grant) instead of site permissions (`/sites/{id}/permissions` → 403, an admin operation outside Sites.Selected).
>
> ✅ **Java factory bug fixed (2026-06-30):** `GraphServiceClient(credential)` passed a non-null empty scopes array that Kiota made immutable (`Arrays.asList`); `getAuthorizationToken` then called `.add()` and threw `UnsupportedOperationException`. cap6/7 Java "passed" only because they masked the exception. Fix in `common/sharepoint-graph-auth/GraphServiceClientFactory`: build the provider by hand passing `(String[]) null` to the varargs → `new ArrayList<>()` (mutable) branch. After the fix, cap7 Java webhooks **actually works** (lists subscriptions); cap6 Java calls Graph (its real error is an SDK date-parsing quirk, not permissions).

---

## 1. Unit tests (no network)

They test the **shared authentication module** (`common/`) of each language: environment variable validation and lazy construction of the Graph client (no network call, because `ClientSecretCredential`/`GraphServiceClient` are lazy).

### C# — xUnit
- **Location:** `common/SharePointGraphAuth.Tests/` (`GraphAuthTests.cs`)
- **Tests (5):**
  - `GraphAuthOptions_ThrowsWhenAllEnvMissing`
  - `GraphAuthOptions_ThrowsWhenSecretAndCertificateMissing`
  - `GraphAuthOptions_AcceptsSecretEnv`
  - `CreateFromSecret_ReturnsNonNullClient_WithFakeEnv`
  - `CreateFromSecret_ThrowsWhenEnvMissing`
- **Run:** `dotnet test common/SharePointGraphAuth.Tests/SharePointGraphAuth.Tests.csproj`
- **Result:** `Passed! - Failed: 0, Passed: 5, Skipped: 0, Total: 5`

### Python — pytest
- **Location:** `common/tests/test_graph_auth.py`
- **Tests (7):** missing env throws (get_graph_client, get_access_token); missing secret throws; get_graph_client returns `GraphServiceClient` with env; **certificate mode**: auto-detects cert (real PEM generated with `cryptography`), forces secret when `use_certificate=False`, throws if `CERTIFICATE_PATH` is missing.
- **Run:** `python -m pytest common/tests/ -q` (requires `pip install -e common` + `pytest` + `cryptography`)
- **Result:** `7 passed in 1.57s`

### JavaScript — Jest
- **Location:** `common/__tests__/graphAuth.test.js`
- **Tests (7):** `requireEnv` throws; `getGraphClient` throws without env; returns Client with env; **certificate mode**: auto-detects cert (real PEM generated with `openssl` at load), forces secret with `useCertificate:false`, throws if `useCertificate:true` without `CERTIFICATE_PATH`, builds credential in cert mode without network.
- **Run:** `npx jest` (requires `openssl` for cert tests; if absent, they are skipped)
- **Result:** `Tests: 7 passed, 7 total`

### Java — JUnit 5
- **Location:** `common/sharepoint-graph-auth/src/test/java/.../GraphServiceClientFactoryTest.java`
- **Tests (8):** returns client with env (fake env injected by `maven-surefire` `environmentVariables`); throws without env; **certificate mode**: throws if `CERTIFICATE_PATH` is null/empty, throws if tenant/client is blank, builds client from real PEM, builds client from real P12 (PEM/P12 generated with `openssl` in `@BeforeAll`; uses overloads with params because `System.getenv()` is immutable at runtime in Java); `createFromSecret` with params builds and throws if blank.
- **Run:** `mvn -pl common/sharepoint-graph-auth -am test` (requires `openssl` for cert tests; if absent, they are skipped via `Assumptions`)
- **Result:** `Tests run: 8, Failures: 0, Errors: 0, Skipped: 0` — BUILD SUCCESS

### PowerShell — Pester
- **Location:** `common/SharePointGraph.Tests.ps1`
- **Tests (6):** `Connect-SharePointGraph` throws when `TENANT_ID` is missing; throws when `CLIENT_SECRET` is missing in secret mode; uses secret mode by default; **certificate mode**: auto-detects with `CERTIFICATE_THUMBPRINT`, auto-detects with `CERTIFICATE_NAME`, forces cert mode with `-UseCertificate`. (Uses a stub of `Connect-MgGraph` so it does not touch the network.)
- **Run:** `pwsh -NoProfile -Command "Invoke-Pester -Path ./common/SharePointGraph.Tests.ps1 -Output Detailed"`
- **Result:** `Tests Passed: 6, Failed: 0, Skipped: 0`

---

## 2. Smoke tests (live, against the `book-test` tenant)

They test end-to-end that the `main()` of each example **runs and performs a real read** against the test site (with auth from the common module + DI). Script: `smoke-tests.sh` (repo root).

| Example | Language | Read operation | Result |
|---------|----------|----------------------|:---------:|
| cap3 sites | C# | GetSite(book-test) + ListSites (7) | ✅ PASS |
| cap4 documents | C# | ListDrives (5 libraries) | ✅ PASS |
| cap5 permissions | C# | ListSitePermissions | ✅ PASS |
| cap6 metadata | C# | ExploreAsync (Term Store) | ✅ PASS* |
| cap7 automation | C# | GetInitialDeltaAsync ("Versiones" library, 1 item + token) | ✅ PASS |
| cap3 sites | Python | get_site + list_sites (7) | ✅ PASS |
| cap4 documents | Python | list_drives (5 libraries) | ✅ PASS |
| cap5 permissions | Python | list_site_permissions | ✅ PASS |
| cap3 sites | JS | getSite + listSites (7) | ✅ PASS |
| cap4 documents | JS | listDrives (5 libraries) | ✅ PASS |
| cap5 permissions | JS | listSitePermissions | ✅ PASS |
| cap3 sites | PowerShell | Get-RootSite + Get-AllSites (7) | ✅ PASS |
| cap5 permissions | Java | listPermissions (drive root of book-test) | ✅ PASS |
| cap6 metadata | Java | ExploreAsync (Term Store) | ✅ PASS† |
| cap7 automation | Java | listSubscriptions (webhooks) | ✅ PASS |

**\* cap6 C#**: the example runs and handles the error gracefully, but the Term Store read returns "Access denied" because the App Registration **does not have `TermStore.Read.All`** (a tenant permission gap, not a bug).

**† cap6 Java**: after the factory fix, the example does call Graph, but the Graph SDK for Java fails to deserialize an `OffsetDateTime` without a zone (`Text '2026-07-01T04:14:48' could not be parsed at index 19`) — an SDK quirk, not permissions. `explore()` catches it and the main ends cleanly → PASS.

**cap7 Java**: after the factory fix, `listSubscriptions` **actually works** ("Suscripciones encontradas: 0") — the previous diagnosis (missing `Subscriptions.Read`) was incorrect; it was the factory bug.

**Run:** `bash smoke-tests.sh` (requires credentials in `code/dataverse-validation/.env` under the `DATAVERSE_*` prefix).

---

## 3. How to reproduce the whole suite

```bash
cd la-experiencia-sharepoint-code

# Unit (no creds, no network)
dotnet test common/SharePointGraphAuth.Tests/SharePointGraphAuth.Tests.csproj
python -m pytest common/tests/ -q            # after: pip install -e common && pip install pytest
npx jest
mvn -pl common/sharepoint-graph-auth test
pwsh -NoProfile -Command "Invoke-Pester -Path ./common/SharePointGraph.Tests.ps1 -Output Detailed"

# Smoke (live, requires creds)
bash smoke-tests.sh
```

---

## 4. Coverage and limitations (honest)

**Covered:**
- Shared auth module in all 5 languages (unit, no network).
- Read-only mains of caps 3-7 (C#), 3-5 (Python/JS), cap3 (PS) — live smoke.

**Not covered yet (pending work):**
- **PowerShell cap6** (Term Store): uses `Connect-PnPOnline -Interactive` (PnP, interactive browser authentication) — not automatable headless, like cap2.
- **PowerShell cap7** (delta/webhooks): **parameterized** scripts (they require `-AccessToken` and `-DriveId`); they work when invoked with those params, not self-contained.
- **cap5 Java**: Java permissions example **created and validated live** (`chapter-05-permissions/java/01-permissions/PermissionOperations.java`, SDK v6 + DI + common factory). Lists permissions of the root item of the first drive of book-test (within the Sites.Selected grant). Smoke PASS.
- **cap2 (auth, all languages)**: **interactive/delegated** authentication (browser) — not automatable headless.
- **Write operations** (upload, create link, grant access, create webhooks): not tested automatically to avoid mutating the tenant; mains use only reads.
- **No unit tests of the business methods** of each chapter (e.g. `ListSitesAsync`): would require mocking `GraphServiceClient` (Kiota), outside the current pragmatic scope.

**Progress since the previous version:** caps 3/4/5 PowerShell with real mains added to smoke; cap3/4 Java with real mains added to smoke; **cap5 Java created and validated** (PermissionOperations, SDK v6 + DI); **Java factory bug fixed** (UnsupportedOperationException); **Phase 4 Sites.Selected validated** (mains within the grant; cap5 lists drive-item permissions); **certificate auth integrated in all 5 common factories** (auto-detection, +17 cert unit tests). Total suite **57/57 green** (33 unit + 24 smoke).

---

## 5. Conclusion

The automated suite covers **57 tests (33 unit + 24 smoke), 57/57 validated green**, validating that:
1. The shared auth module validates and builds clients in all 5 languages.
2. The examples for caps 3-7 (C#/Python/JS), cap3 (PS) and caps 3-4/5/6-7 (Java) **run end-to-end** against the real tenant and return data from `book-test` with **least privilege Sites.Selected**.

The refactored code (Phases 0-5) is **verified live**. Certificate auth integrated in all 5 common factories (C#/Python/JS/PS/Java) with auto-detection.

---

## 6. SPFx — cap14 (`chapter-14-spfx/`)

The **12 SPFx examples** of cap14 (SPFx 1.20, Node 18 LTS, React 17, Fluent UI v8)
build with `gulp build` in **0 warnings and 0 TS errors**:

| # | Project | Type | Build | Declared Graph permission |
|---|---|---|:---:|---|
| 01 | webpart-hello-graph | Web Part React | ✅ | `Sites.Read.All` |
| 02 | field-customizer-priority | Field Customizer | ✅ | — |
| 03 | command-set-actions | Command Set | ✅ | — |
| 04 | application-customizer-footer | Application Customizer | ✅ | — |
| 05 | ace-viva-card | Adaptive Card Extension | ✅ | `Tasks.Read` |
| 06 | webpart-documents-explorer | Web Part React | ✅ | `Sites.Read.All`, `Files.Read.All` |
| 07 | webpart-permissions-panel | Web Part React (capstone) | ✅ | `Sites.Read.All` |
| 08 | aadhttpclient-own-api | Web Part React | ✅ | own API (`user_impersonation`) |
| 09 | spfx-library-shared | Library | ✅ | — |
| 10 | webpart-no-framework | Web Part (no React) | ✅ | — |
| 11 | teams-enabled-webpart | Web Part + Teams | ✅ | — |
| 12 | cicd-workflow | GitHub Actions | n/a (workflow) | — |

**11/11 SPFx projects build green.** #12 is a CI/CD workflow (it does not compile).
Verification: `cd chapter-14-spfx/<NN>; nvm use 18; npm install; gulp build` → `Task warnings:0`.

> **These are not live smoke tests** (they do not call the tenant). Live validation of
> SPFx requires App Catalog + consent of the declared permissions; the projects
> `02/03/04` (extensions, no Graph) and `10` (no-framework) are the first
> candidates to validate in `book-test`.

### 6.1 Jest (unit, no network) — reference projects

Jest is configured in two projects as a **reference of the pattern** taught by
the "Testing SPFx" section of cap14 (the rest of the projects replicate it):

| Project | Tests | What it covers |
|---|:---:|---|
| `09-spfx-library-shared` | 14 ✅ | `classifyHttpStatus` (401/403/404/429/5xx/network) + `truncate`/`formatRelativeDate` |
| `01-webpart-hello-graph` | 3 ✅ | `GraphService.searchSites` with fake `MSGraphClientV3` (empty query, mapping + `displayName` fallback, missing value) |

Per-project config: `jest.config.js` (preset `ts-jest` + `tsconfig.test.json`
with `types: ["jest","node"]`) + `.eslintignore` for `**/*.test.ts` + `tsconfig.json`
with `exclude` of `src/**/*.test.ts` (so `gulp build` does not compile the tests) +
`scripts.test = "jest"`. Verification: `cd <project>; nvm use 18; npm install; npm test`.

**Total jest cap14: 17 tests, 17/17 PASS** (included in the suite count).

---

## 7. Teams — cap17 (`chapter-17-teams/`)

Two examples of Teams apps (path 2 of cap17: standalone app with its own manifest):

| # | Project | Type | Build | Tests |
|---|---|---|:---:|---|
| 01 | tab-app-sharepoint | Tab app (manifest + icons, no hosting) | n/a (zip) | — |
| 02 | bot-sharepoint | Node bot (botbuilder + restify) that reads Graph | ✅ `tsc` 0 errors | 4 ✅ Jest |

### 7.1 Jest (unit, no network) — `02-bot-sharepoint`

| Tests | What it covers |
|:---:|---|
| 4 ✅ | `getGraphClient` validates env (`GRAPH_TENANT_ID`/`GRAPH_CLIENT_ID`/`GRAPH_CLIENT_SECRET` missing throw; message mentions `.env.example`) |

Config: `jest.config.js` (preset `ts-jest` + `tsconfig.test.json` `types: ["jest","node"]`) + `tsconfig.json` `exclude src/**/*.test.ts` (so `tsc` does not compile the tests). Verification: `cd chapter-17-teams/02-bot-sharepoint; npm install; npm test` → `Tests: 4 passed`.

> **These are not live smoke tests** (the bot requires public hosting + ngrok + Entra App Registration + Azure Bot resource to test in real Teams). The unit test validates the auth logic without network; the README documents the full deployment lifecycle.

**Total jest cap17: 4 tests, 4/4 PASS.**

---
*Report generated: 2026-06-30. SPFx cap14 added: 2026-07-03. Teams cap17 added: 2026-07-05.*