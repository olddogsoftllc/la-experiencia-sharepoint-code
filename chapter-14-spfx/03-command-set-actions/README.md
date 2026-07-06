# 03 — Command Set: Actions

A **ListView Command Set** that adds two context buttons to a list/library
command bar:

- **Export selection** (`COMMAND_1`) — visible only when ≥ 1 row is selected.
- **Notify count** (`COMMAND_2`) — always visible; reports how many rows are
  selected.

Covers the book's cap14 "Tipos de Extensiones" (Command Set) and shows
`BaseListViewCommandSet`, `tryGetCommand`, `listViewStateChangedEvent`, and
`raiseOnChange` to toggle command visibility from the selection state.

## Build

```bash
nvm use 18
npm install
gulp build
gulp bundle --ship
gulp package-solution --ship     # sharepoint/solution/command-set-actions.sppkg
```

## Deploy & register

1. Upload the `.sppkg` to the App Catalog and enable it.
2. Register the command set on a list via PnP PowerShell:

```powershell
Add-PnPCustomAction -Name "Actions" -Title "Actions" `
  -Location "ClientSideExtension.ListViewCommandSet" `
  -ClientSideComponentId "<id-from-manifest>" -List "Tasks"
```

## Properties (ClientSideComponentProperties)

| Property | Default | Purpose |
|---|---|---|
| `notifyPrefix` | `Selected` | Prefix shown by the "Notify count" command |