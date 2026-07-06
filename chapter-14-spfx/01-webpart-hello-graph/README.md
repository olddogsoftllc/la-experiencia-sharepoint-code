# 01 — Web Part: Hello Graph

A React Web Part that searches SharePoint sites through Microsoft Graph using
`MSGraphClientV3`, a `useGraphQuery` hook, and the SPFx Property Pane.

Demonstrates (covers the book's chapter 14):

- `MSGraphClientV3` obtained in `onInit` and wrapped in a `GraphService`
  (dependency-injected into the component so the component is unit-testable).
- `useGraphQuery<T>(fetcher, deps)` hook for async state (loading / data / error).
- Property Pane with `PropertyPaneTextField` (query) + `PropertyPaneSlider` (max results).
- Fluent UI v8 `Spinner`, `MessageBar`, `DocumentCard` rendering of results.
- Themed via `semanticColors` / `onThemeChanged`; Teams-aware via `hasTeamsContext`.
- Graph permission `Sites.Read.All` declared in `config/package-solution.json`.

## Structure

```
src/webparts/helloGraphWebPart/
  HelloGraphWebPartWebPart.ts        # Web Part: onInit (graph client), render, Property Pane
  components/
    HelloGraphWebPart.tsx            # Function component using useGraphQuery + Fluent UI
    IHelloGraphWebPartProps.ts
    HelloGraphWebPart.module.scss
  hooks/
    useGraphQuery.ts                 # Reusable async-data hook
  services/
    GraphService.ts                  # Thin MSGraphClientV3 wrapper (DI-friendly)
  models/
    ISite.ts
  loc/                               # Resource strings for the Property Pane
```

## Build & test

Requires Node 18 LTS (SPFx 1.20). Use nvm:

```bash
nvm use 18
npm install
gulp build              # type-check + lint + bundle (DEBUG)
gulp bundle --ship      # production bundle (minified, no source maps)
gulp package-solution --ship   # produces sharepoint/solution/*.sppkg
```

## Deploy

1. `gulp package-solution --ship` → `sharepoint/solution/hello-graph.sppkg`.
2. Upload the `.sppkg` to the tenant App Catalog and enable it.
3. Approve the **Microsoft Graph → Sites.Read.All** permission in
   **SharePoint Admin Center → API Access**.
4. Add the Web Part to a page; configure the **query** and **max results**
   in the Property Pane.

> In development, `gulp serve` opens the local workbench. The first Graph call
> triggers a tenant-admin consent popup for `Sites.Read.All`.