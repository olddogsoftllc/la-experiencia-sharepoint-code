# 06 — Web Part: Documents Explorer

A React Web Part that lists a site's document libraries (Graph drives) and the
root items of the first library, rendered with Fluent UI `DocumentCard`.

Connects to the book's cap4 (Documents) and demonstrates the cap14
"Patrón de Acceso a Graph en SPFx": `MSGraphClientV3` in `onInit`, a
dependency-injected `GraphService`, the `useGraphQuery` hook, and Fluent UI.

## Structure

```
src/webparts/documentsExplorerWebPart/
  DocumentsExplorerWebPartWebPart.ts   # onInit (graph client), render, Property Pane (maxResults)
  components/DocumentsExplorerWebPart.tsx   # two useGraphQuery calls: drives + root items
  services/GraphService.ts             # getSiteDrives(siteId) + getDriveRootItems(driveId)
  hooks/useGraphQuery.ts
  models/IDrive.ts                      # IDrive, IDriveItem
```

## Build

```bash
nvm use 18
npm install
gulp build
gulp bundle --ship
gulp package-solution --ship     # sharepoint/solution/documents-explorer.sppkg
```

## Deploy

1. Upload the `.sppkg` to the App Catalog and enable it.
2. Approve **Microsoft Graph → Sites.Read.All** and **Files.Read.All** in
   SharePoint Admin Center → API Access.
3. Add the Web Part to a page; set **Max items per library** in the Property Pane.

## Graph calls

- `GET /sites/{siteId}/drives` — libraries of the current site (`pageContext.site.id`).
- `GET /drives/{driveId}/root/children?$top=N` — root items of the first library.