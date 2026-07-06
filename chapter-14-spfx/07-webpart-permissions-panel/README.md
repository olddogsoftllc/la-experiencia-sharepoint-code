# 07 — Web Part: Permissions Panel (capstone)

The capstone Web Part: reads a site's permission entries via Microsoft Graph,
classifies errors, and emits structured telemetry — wiring together Graph,
DI, error handling, and the telemetry stub from cap14.

Connects to cap5 (Permissions) and demonstrates the cap14 "Manejo de errores y
telemetría" section end-to-end.

## Structure

```
src/webparts/permissionsPanelWebPart/
  PermissionsPanelWebPartWebPart.ts     # onInit (graph client + telemetry), render, Property Pane (siteId)
  components/PermissionsPanelWebPart.tsx     # useGraphQuery + roles/grantee rendering + 403 MessageBar
  services/GraphService.ts             # getSitePermissions(siteId) with try/catch + telemetry
  services/TelemetryClient.ts           # in-memory event recorder (wire to AppInsights in prod)
  hooks/useGraphQuery.ts
  models/ISitePermission.ts
```

## Build

```bash
nvm use 18
npm install
gulp build
gulp bundle --ship
gulp package-solution --ship     # sharepoint/solution/permissions-panel.sppkg
```

## Deploy

1. Upload the `.sppkg` to the App Catalog and enable it.
2. Approve **Microsoft Graph → Sites.Read.All** in API Access.
3. Add the Web Part to a page. Set **Graph site id** (defaults to the current site).

## Why this is the capstone

`/sites/{id}/permissions` is an **admin operation**. Under the book's
Sites.Selected / "mínimo privilegio" model it returns **403** — and the Web Part
surfaces that explicitly instead of silently failing. It also shows:
- `GraphService` receiving both the Graph client and a `TelemetryClient` (DI).
- `trackSuccess` / `trackError` called on every call.
- The 401 / 403 / 404 / 429 classification from the shared Library (`#09`).