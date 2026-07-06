# 01 — Tab App de Teams que surfacea SharePoint

App de Teams **standalone** (vía 2 del cap17) que muestra un site de SharePoint como tab personal. Sin hosting propio: el tab apunta a una URL de SharePoint vía `teamslogon.aspx`, así que puedes probar el ciclo completo (manifest → empaquetado → sideload → publish) sin desplegar nada más que el .zip.

## Archivos

```
01-tab-app-sharepoint/
├── manifest.json     # edita: id (UUID), TENANT, contentUrl
├── color.png         # 192×192 — ver ICONS.md
├── outline.png       # 32×32 monocromo — ver ICONS.md
├── package.sh        # genera el .zip válido
└── README.md
```

## Ciclo de vida completo

### 1. Editar `manifest.json`
- `id`: genera un UUID nuevo (`uuidgen` o `python -c "import uuid;print(uuid.uuid4())"`). **No** es el clientId de Entra — es un UUID aparte.
- `staticTabs[0].contentUrl` y `websiteUrl`: cambia `TENANT` por tu tenant y `sites/projects` por tu site.
- `validDomains`: debe incluir tu dominio `.sharepoint.com` o Teams bloquea el iframe.

### 2. Crear los iconos
Lee `ICONS.md` y genera `color.png` (192×192) + `outline.png` (32×32 monocromo).

### 3. Empaquetar
```bash
bash package.sh
# → sp-projects-tab.zip (3 archivos en la raíz)
```

### 4. Sideload (desarrollo)
Teams → **Apps** → **Manage your apps** → **Upload an app** → **Upload a custom app** → selecciona el .zip.
> Requiere que el Teams Admin haya habilitado "Allow sideloading of custom apps". Si la opción no aparece, pídelo.

### 5. Publish a org store (producción)
Teams → Apps → Manage your apps → **Submit an app to your org** → sube el .zip.
→ Va al Teams Admin Center (Pending approval) → el admin aprueba → disponible para todo el tenant.

## Gobernanza
- `validDomains` acotado a tu tenant (no wildcards).
- El tab hereda los permisos del usuario en SharePoint — `teamslogon.aspx` respeta RBAC. Si el usuario no tiene acceso al site, ve una página de acceso denegado (no se filtra contenido).
- Para producción, sustituye los iconos placeholder por iconos reales de marca.