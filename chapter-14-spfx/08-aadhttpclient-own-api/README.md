# 08 — Web Part: AadHttpClient (own API)

A React Web Part that calls a **custom Azure AD-protected API** via
`AadHttpClient` — the counterpart to `MSGraphClientV3` for your own services.

Covers the book's chapter 14 "MSGraphClient vs AadHttpClient".

## Structure

```
src/webparts/ownApiWebPart/
  OwnApiWebPartWebPart.ts        # onInit (aadHttpClientFactory.getClient), render, Property Pane (apiId + apiUrl)
  components/OwnApiWebPart.tsx    # useAsync hook, renders JSON response
  services/OwnApiService.ts      # AadHttpClient.get + status check
  hooks/useAsync.ts
```

## Build

```bash
nvm use 18
npm install
gulp build
gulp bundle --ship
gulp package-solution --ship     # sharepoint/solution/aadhttpclient-own-api.sppkg
```

## Deploy

1. Register your API in Azure AD, expose an **App ID URI** (e.g.
   `api://contoso-api`) and a scope (e.g. `user_impersonation`).
2. Edit `config/package-solution.json` → `webApiPermissionRequests` to reference
   **your** API's App ID URI and scope (the scaffold ships a placeholder).
3. Upload the `.sppkg` to the App Catalog and enable it.
4. Approve the API permission in SharePoint Admin Center → API Access.
5. Add the Web Part to a page; set **API App ID URI** and **API endpoint URL**
   in the Property Pane.

## Properties

| Property | Default | Purpose |
|---|---|---|
| `apiId` | `api://your-api-app-id-uri` | App ID URI of the target API |
| `apiUrl` | `https://your-api.azurewebsites.net/api/data` | Endpoint to call |

> `AadHttpClient` obtains the token automatically using the App ID URI; you
> never handle tokens. Changing either property re-creates the service and
> re-renders (see `onPropertyPaneFieldChanged`).