# 04 — Application Customizer: Footer

An **Application Customizer** that injects a global footer bar into the
`Bottom` placeholder of every modern page where the extension is active.

Covers the book's cap14 "Tipos de Extensiones" (Application Customizer) and
shows `BaseApplicationCustomizer`, `placeholderProvider`, `PlaceholderName.Bottom`,
and `tryCreateContent` — including reacting to `changedEvent` when placeholders
become available after `onInit`.

## Build

```bash
nvm use 18
npm install
gulp build
gulp bundle --ship
gulp package-solution --ship     # sharepoint/solution/app-customizer-footer.sppkg
```

## Deploy & register

1. Upload the `.sppkg` to the App Catalog and enable it.
2. Activate tenant-wide (the `skipFeatureDeployment` flag in
   `package-solution.json` enables this) **or** register on a single site:

```powershell
Add-PnPCustomAction -Name "Footer" -Title "Footer" `
  -Location "ClientSideExtension.ApplicationCustomizer" `
  -ClientSideComponentId "<id-from-manifest>" `
  -ClientSideComponentProperties '{ "footerText": "Powered by La Experiencia SharePoint" }'
```

## Properties (ClientSideComponentProperties)

| Property | Default | Purpose |
|---|---|---|
| `footerText` | `Powered by SPFx` | Text shown in the footer bar |

> The footer text is rendered with `textContent` (not `innerHTML`), so it is
> safe even if the property comes from an admin-controlled JSON string.