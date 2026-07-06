# 02 — Bot de Teams que lee SharePoint via Graph

Bot de Teams (vía 2 del cap17) que responde comandos en el chat leyendo SharePoint con Microsoft Graph. Compilable y testeable offline; requiere hosting público + ngrok en dev para probar en Teams real.

## Comandos
- `sites` — lista hasta 10 sites del tenant
- `docs <nombre>` — busca un site por nombre y lista documentos de su primera biblioteca
- `help` — ayuda

## Archivos
```
02-bot-sharepoint/
├── src/
│   ├── index.ts          # bot (botbuilder + restify) + messaging endpoint
│   ├── graphAuth.ts      # factory Graph app-only (reutiliza patrón cap2)
│   └── graphAuth.test.ts # test unitario (sin red)
├── manifest.json         # manifest de Teams (edita botId + id)
├── package.json
├── tsconfig.json / tsconfig.test.json / jest.config.js
├── .env.example          # copia a .env y rellena
└── README.md
```

## Setup completo (paso a paso)

### 1. App Registration en Entra ID
```
Entra ID → App registrations → New registration
├── Name: "SharePoint Teams Bot"
├── Supported account types: Accounts in any org (multitenant)
├── Redirect URI (Web): https://token.botframework.com/.auth/web/redirect
└── Generate client secret → MICROSOFT_APP_PASSWORD

API permissions (Graph, Application):
├── Sites.Read.All  (o Sites.Selected con consent admin por site)
└── Grant admin consent
```

### 2. Azure Bot resource
```
Azure Portal → Create resource → "Azure Bot"
├── Type: Multi Tenant
├── Creation kind: existing App Registration (el de arriba)
└── Channels → Add "Microsoft Teams" channel
```

### 3. .env
```bash
cp .env.example .env
# rellena MICROSOFT_APP_ID, MICROSOFT_APP_PASSWORD, GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET
```

### 4. Build + test (offline, sin creds)
```bash
npm install
npm test      # 4 tests unitarios de graphAuth (sin red)
npm run build # tsc → dist/
```

### 5. Desarrollo local con ngrok
```bash
npm start                              # localhost:3978
ngrok http 3978                        # → https://abc.ngrok.io
# Azure Bot → Configuration → Messaging endpoint = https://abc.ngrok.io/api/messages
```

### 6. Manifest + sideload
- Edita `manifest.json`: `id` (UUID nuevo), `bots[0].botId` = `MICROSOFT_APP_ID` (clientId de Entra).
- Crea `color.png` (192×192) + `outline.png` (32×32 monocromo).
- `zip sp-bot.zip manifest.json color.png outline.png`
- Teams → Apps → Manage your apps → Upload a custom app.

### 7. Producción
- Despliega a Azure App Service (Node Linux): mismo código, `npm start`.
- Azure Bot → Messaging endpoint = `https://<prod>/api/messages`.
- Quita ngrok.

## Gobernanza
- **Auth**: app-only con `Sites.Selected` en producción (no `Sites.ReadWrite.All`). El bot es una app, no un usuario.
- **Secrets**: `.env` en `.gitignore`, jamás al repo. En App Service usa config settings (o Key Vault reference).
- **Auditoría**: loguea `context.activity.from.aadObjectId` + texto + resultado a Application Insights.
- **Sin escritura**: este bot solo lee. Añadir `create`/`delete` requiere confirmación por Adaptive Card (ver cap17).
- **Scopes**: `personal` + `team` — en `team`, todos los miembros invocan el bot; revisa que encaje con tu modelo de permisos.

## Notas honestas
- `@microsoft/microsoft-graph-client` v3 requiere `isomorphic-fetch` como polyfill de `fetch` en Node (ya en deps). Sin él, `client.api().get()` falla con "fetch is not defined".
- `Sites.Selected` requiere que el admin conceda acceso del app a sites concretos (`Add-PnPAzureADAppSitePermission`). Con `Sites.Read.All` funciona sin ese paso, pero es más privilegio del necesario — úsalo solo en dev.
- El bot usa **app-only** (no On-Behalf-Of del usuario que escribe). Eso significa que ve todos los sites que la app tenga permitido, **no** los del usuario. Para respetar permisos por usuario, implementa OBO flow (más complejo, ver cap17 §SSO).