# Chapter 14 — SPFx Examples

Twelve self-contained SharePoint Framework (SPFx 1.20) projects that cover the
full cap14 syllabus of *La Experiencia SharePoint*. Each project is an
independent `yo @microsoft/sharepoint` solution you can build, package, and
deploy on its own.

> **Stack**: SPFx 1.20.x · Node 18 LTS · TypeScript 4.7 · React 17 · Fluent UI
> React v8 · gulp 4. Use `nvm use 18` before building any project.

## Projects

| # | Folder | Type | Cap14 section covered | Build |
|---|---|---|---|---|
| 01 | `01-webpart-hello-graph` | Web Part (React) | MSGraphClientV3 + hooks + Property Pane | ✅ green (0 warnings) |
| 02 | `02-field-customizer-priority` | Field Customizer | Tipos de Extensiones | ✅ green (0 warnings) |
| 03 | `03-command-set-actions` | Command Set | Tipos de Extensiones | ✅ green (0 warnings) |
| 04 | `04-application-customizer-footer` | Application Customizer | Tipos de Extensiones | ✅ green (0 warnings) |
| 05 | `05-ace-viva-card` | Adaptive Card Extension | ACE en detalle | ✅ green (0 warnings) |
| 06 | `06-webpart-documents-explorer` | Web Part (React) | Patrón de Acceso a Graph | ✅ green (0 warnings) |
| 07 | `07-webpart-permissions-panel` | Web Part (React) | Graph + permisos (capstone) | ✅ green (0 warnings) |
| 08 | `08-aadhttpclient-own-api` | Web Part (React) | AadHttpClient | ✅ green (0 warnings) |
| 09 | `09-spfx-library-shared` | Library | SPFx Library (shared code) | ✅ green (0 warnings) |
| 10 | `10-webpart-no-framework` | Web Part (no React) | Web Part sin React | ✅ green (0 warnings) |
| 11 | `11-teams-enabled-webpart` | Web Part + Teams | Integración con Teams | ✅ green (0 warnings) |
| 12 | `12-cicd-workflow` | GitHub Actions | CI/CD para SPFx | ✅ created (no build) |

Legend: ✅ all 11 SPFx projects build clean (0 warnings, 0 TS errors); #12 is a workflow (no build).

## Build any project

```bash
cd chapter-14-spfx/<NN-folder>
nvm use 18
npm install
gulp build              # DEBUG build (type-check + lint + bundle)
gulp bundle --ship      # production bundle
gulp package-solution --ship   # -> sharepoint/solution/*.sppkg
```

## Graph permissions

Projects that call Microsoft Graph (01, 06, 07) declare the scopes they need in
`config/package-solution.json` under `webApiPermissionRequests`. After uploading
the `.sppkg`, a tenant admin must approve each scope in **SharePoint Admin Center
→ API Access**. Scope policy follows the book's "mínimo privilegio" rule
(`Sites.Read.All`, `Files.Read.All`) — never `Sites.ReadWrite.All`.

## Deploy

Upload the `.sppkg` to the tenant App Catalog, enable it, and (for Web Parts)
add it to a page. Extensions are registered per-list/per-site with
`Add-PnPCustomAction` or column JSON — see each project's README.

## Tests

Jest is set up on two reference projects that demonstrate the pattern taught in
the cap14 "Testing SPFx" section (the others replicate it):

- `09-spfx-library-shared` — `npm test` → 14 tests (utils)
- `01-webpart-hello-graph` — `npm test` → 3 tests (`GraphService` with a fake
  `MSGraphClientV3`)

Each uses `jest.config.js` + `tsconfig.test.json` (`types: ["jest","node"]`) +
`.eslintignore` for `**/*.test.ts`, and excludes `src/**/*.test.ts` from the SPFx
`tsconfig.json` so `gulp build` stays clean.