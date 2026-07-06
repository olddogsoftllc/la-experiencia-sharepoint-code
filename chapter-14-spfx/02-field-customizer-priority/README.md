# 02 — Field Customizer: Priority

A **Field Customizer** that renders a list column's value as a colored badge
based on the priority text: `high` → red, `low` → green, anything else → grey.

Covers the book's cap14 "Tipos de Extensiones" (Field Customizer) and shows the
`BaseFieldCustomizer.onRenderCell` lifecycle with CSS Modules.

## How it works

`onRenderCell` reads `event.fieldValue`, picks a CSS class by keyword, builds a
`<span>` badge with `textContent` (no `innerHTML` → no XSS from cell values),
and appends it to `event.domElement`.

## Build

```bash
nvm use 18
npm install
gulp build
gulp bundle --ship
gulp package-solution --ship     # sharepoint/solution/field-priority.sppkg
```

## Deploy & register

1. Upload the `.sppkg` to the App Catalog and enable it.
2. On a list, register the customizer on a text column by its component id
   (column JSON):

```json
{
  "ClientSideComponentId": "<id-from-manifest>"
}
```

   Or with PnP PowerShell:

```powershell
Set-PnPField -List "Tasks" -Identity "Priority" -Values @{
  CustomFormatter = '{"ClientSideComponentId":"<id-from-manifest>"}'
}
```

3. The column now shows colored priority badges.

## Properties (ClientSideComponentProperties)

| Property | Default | Purpose |
|---|---|---|
| `prefix` | _(none)_ | Text shown before the value, e.g. the field name |