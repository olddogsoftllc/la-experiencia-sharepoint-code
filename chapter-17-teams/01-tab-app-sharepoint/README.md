# 01 — Teams Tab App that surfaces SharePoint

Standalone Teams App (path 2 of cap17) that displays a SharePoint site as a personal tab. No hosting of its own: the tab points to a SharePoint URL via `teamslogon.aspx`, so you can test the full lifecycle (manifest → packaging → sideload → publish) without deploying anything beyond the .zip.

## Files

```
01-tab-app-sharepoint/
├── manifest.json     # edita: id (UUID), TENANT, contentUrl
├── color.png         # 192×192 — ver ICONS.md
├── outline.png       # 32×32 monocromo — ver ICONS.md
├── package.sh        # genera el .zip válido
└── README.md
```

## Full lifecycle

### 1. Edit `manifest.json`
- `id`: generate a new UUID (`uuidgen` or `python -c "import uuid;print(uuid.uuid4())"`). This is **not** the Entra clientId — it is a separate UUID.
- `staticTabs[0].contentUrl` and `websiteUrl`: replace `TENANT` with your tenant and `sites/projects` with your site.
- `validDomains`: must include your `.sharepoint.com` domain or Teams will block the iframe.

### 2. Create the icons
Read `ICONS.md` and generate `color.png` (192×192) + `outline.png` (32×32 monochrome).

### 3. Package
```bash
bash package.sh
# → sp-projects-tab.zip (3 archivos en la raíz)
```

### 4. Sideload (development)
Teams → **Apps** → **Manage your apps** → **Upload an app** → **Upload a custom app** → select the .zip.
> Requires that the Teams Admin has enabled "Allow sideloading of custom apps". If the option does not appear, request it.

### 5. Publish to org store (production)
Teams → Apps → Manage your apps → **Submit an app to your org** → upload the .zip.
→ Goes to the Teams Admin Center (Pending approval) → the admin approves → available to the whole tenant.

## Governance
- `validDomains` scoped to your tenant (no wildcards).
- The tab inherits the user's permissions in SharePoint — `teamslogon.aspx` respects RBAC. If the user does not have access to the site, they see an access-denied page (no content is leaked).
- For production, replace the placeholder icons with real brand icons.