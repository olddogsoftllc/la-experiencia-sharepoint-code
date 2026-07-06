# 11 — Web Part: Teams-Enabled

A React Web Part that runs in **both SharePoint and Microsoft Teams** from a
single SPFx bundle, detecting the host and theme at runtime.

Covers the book's cap14 "Integración SPFx con Teams".

## What it shows

- `this.context.sdks.microsoftTeams` detects whether the Web Part is running
  inside Teams / Office / Outlook vs SharePoint.
- `teamsJs.app.getContext()` provides the host name and theme.
- `onThemeChanged` adapts to SharePoint themes; the `.dark` class handles
  inverted themes in both hosts.
- The `teams/` folder ships the icons used in the generated Teams manifest.

## Build

```bash
nvm use 18
npm install
gulp build
gulp bundle --ship
gulp package-solution --ship     # sharepoint/solution/teams-enabled-webpart.sppkg
```

## Deploy

1. Upload the `.sppkg` to the App Catalog and enable it.
2. When prompted, choose **"Add to Teams"** to make it available as a personal
   or channel tab (the manifest's `supportedHosts` already includes Teams).
3. Add the Web Part to a SharePoint page or a Teams tab; set **Title** in the
   Property Pane.

## Properties

| Property | Default | Purpose |
|---|---|---|
| `title` | `Teams-Enabled Web Part` | Heading shown in the card |