# 02 — Teams Bot that reads SharePoint via Graph

Teams Bot (path 2 of cap17) that responds to commands in chat by reading SharePoint with Microsoft Graph. Compilable and testable offline; requires public hosting + ngrok in dev to test on a real Teams client.

## Commands
- `sites` — lists up to 10 sites in the tenant
- `docs <name>` — searches a site by name and lists documents from its first library
- `help` — help

## Files
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

## Full setup (step by step)

### 1. App Registration in Entra ID
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

### 5. Local development with ngrok
```bash
npm start                              # localhost:3978
ngrok http 3978                        # → https://abc.ngrok.io
# Azure Bot → Configuration → Messaging endpoint = https://abc.ngrok.io/api/messages
```

### 6. Manifest + sideload
- Edit `manifest.json`: `id` (new UUID), `bots[0].botId` = `MICROSOFT_APP_ID` (Entra clientId).
- Create `color.png` (192×192) + `outline.png` (32×32 monochrome).
- `zip sp-bot.zip manifest.json color.png outline.png`
- Teams → Apps → Manage your apps → Upload a custom app.

### 7. Production
- Deploy to Azure App Service (Node Linux): same code, `npm start`.
- Azure Bot → Messaging endpoint = `https://<prod>/api/messages`.
- Remove ngrok.

## Governance
- **Auth**: app-only with `Sites.Selected` in production (not `Sites.ReadWrite.All`). The bot is an app, not a user.
- **Secrets**: `.env` in `.gitignore`, never committed to the repo. In App Service use config settings (or a Key Vault reference).
- **Auditing**: log `context.activity.from.aadObjectId` + text + outcome to Application Insights.
- **No writes**: this bot only reads. Adding `create`/`delete` requires confirmation via an Adaptive Card (see cap17).
- **Scopes**: `personal` + `team` — in `team`, all members invoke the bot; verify that this fits your permission model.

## Honest notes
- `@microsoft/microsoft-graph-client` v3 requires `isomorphic-fetch` as a `fetch` polyfill in Node (already in deps). Without it, `client.api().get()` fails with "fetch is not defined".
- `Sites.Selected` requires the admin to grant the app access to specific sites (`Add-PnPAzureADAppSitePermission`). With `Sites.Read.All` it works without that step, but it is more privilege than necessary — use it only in dev.
- The bot uses **app-only** (not On-Behalf-Of the user who writes). This means it sees all the sites the app is allowed to access, **not** the user's sites. To respect per-user permissions, implement the OBO flow (more complex, see cap17 §SSO).