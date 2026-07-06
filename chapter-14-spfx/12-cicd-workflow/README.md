# 12 — CI/CD Workflow for SPFx

A reusable GitHub Actions workflow that builds, packages, and deploys an SPFx
solution to the tenant App Catalog. Copy `.github/workflows/deploy-spfx.yml`
into any of the `chapter-14-spfx` projects (or any SPFx repo) and configure the
secrets below.

Covers the book's chapter 14 "CI/CD for SPFx" and "Bundling, CDN and the `--ship` flag".

## What the workflow does

1. Checks out the repo and sets up **Node 18 LTS** (the version SPFx 1.20 targets).
2. `npm ci` — reproducible install from the lockfile.
3. `gulp bundle --ship` — minified, tree-shaken production bundle (no source maps).
4. `gulp package-solution --ship` — produces `sharepoint/solution/*.sppkg`.
5. Decodes a base64-encoded certificate (ephemeral, only on the runner).
6. Uploads the `.sppkg` to the App Catalog with `pnp/action-deploy-spfx` using a
   **service principal + certificate** — no username/password, matching the
   cert-auth approach used throughout the book.

## Required GitHub secrets

| Secret | Value |
|---|---|
| `SP_TENANT_ID` | Azure AD tenant id |
| `SP_CLIENT_ID` | App registration (client) id with App Catalog rights |
| `SP_CERTIFICATE_THUMBPRINT` | Thumbprint of the cert registered on the app |
| `SP_CERTIFICATE_BASE64` | The `.pfx` exported as base64 (`base64 cert.pfx > cert.b64`) |
| `APP_CATALOG_URL` | e.g. `https://contoso.sharepoint.com/sites/appcatalog` |

## Optional repository variables

| Variable | Default | Purpose |
|---|---|---|
| `SPFX_SOLUTION_FOLDER` | `.` | Path to the SPFx project when the repo holds more than one |

## Notes

- The certificate is **CI-only and ephemeral**: it lives on the runner for the
  duration of the step and is discarded. Rotate/revoke it as needed; the book's
  policy is that these test artifacts are temporary, not long-lived secrets.
- `overwrite: true` re-deploys on every push to `main`. Use `skip-feature-deployment`
  so the solution is tenant-wide immediately (no per-site app add).
- For multi-tenant or large bundles, add an Azure CDN step (`cdnBasePath` in
  `config/config.json`) before `gulp bundle --ship` — see chapter 14 "Bundling, CDN".