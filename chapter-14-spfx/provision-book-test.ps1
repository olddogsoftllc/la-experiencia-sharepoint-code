# provision-book-test.ps1  (v2 — robusto)
# Provisioning del sitio book-test para validar los ejemplos SPFx del cap14.
# Idempotente. Usa try/catch y mensajes claros para cada paso.
#
# Uso:
#   Connect-PnPOnline -Url "https://olddogsoft1.sharepoint.com/sites/book-test" -Interactive -ClientId "<tu-app-id>"
#   .\provision-book-test.ps1                  # crea lista + columna + items
#   .\provision-book-test.ps1 -RegisterExtensions   # además registra 02 y 03
#   .\provision-book-test.ps1 -Reset           # borra y recrea Tareas (limpio)

[CmdletBinding()]
param(
    [switch]$RegisterExtensions,
    [switch]$Reset
)

$ListTitle = 'Tasks'
$FieldDisplayName = 'Priority'
$FieldInternalName = 'Priority'

# IDs de los manifests de los .sppkg ya en el App Catalog
$FieldCustomizerId = '53f8a795-a835-45fd-bfa7-01249906d75b'   # 02-field-customizer-priority
$CommandSetId       = 'ae7678f8-f77c-4a54-a9c7-f22dfd3cc70c'   # 03-command-set-actions

function Step($msg) { Write-Host "› $msg" -ForegroundColor Cyan }
function Ok($msg)   { Write-Host "  $msg" -ForegroundColor Green }
function Warn($msg) { Write-Host "  $msg" -ForegroundColor Yellow }

# --- 0) Conexión ---
$ctx = Get-PnPContext
if (-not $ctx) { throw "No hay conexión PnP. Ejecuta Connect-PnPOnline primero." }
Step "Conectado a $((Get-PnPProperty -ClientObject $ctx.Web -Property Url))"

# --- -Reset: borrar Tareas si existe ---
if ($Reset) {
    Step "Reset: borrando lista '$ListTitle' si existe"
    try {
        $existing = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
        if ($existing) { Remove-PnPList -Identity $ListTitle -Force; Ok "lista borrada." }
        else { Warn "no existía." }
    } catch { Warn "no se pudo borrar: $_" }
}

# --- 1) Crear la lista si no existe ---
Step "Lista '$ListTitle'"
$list = $null
try { $list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue } catch { }
if ($list) {
    Ok "ya existe (id=$($list.Id))."
} else {
    try {
        $list = New-PnPList -Title $ListTitle -Template GenericList -ErrorAction Stop
        Ok "lista creada."
    } catch { throw "No se pudo crear la lista: $_" }
}

# --- 2) Crear la columna "Prioridad" si no existe ---
Step "Columna '$FieldDisplayName'"
$field = $null
try { $field = Get-PnPField -List $ListTitle -Identity $FieldInternalName -ErrorAction SilentlyContinue } catch { }
if ($field) {
    Ok "ya existe (internal name: $($field.InternalName))."
} else {
    try {
        # Add-PnPFieldText crea un campo de texto; -AddToDefaultView lo muestra en la vista
        Add-PnPFieldText -List $ListTitle -DisplayName $FieldDisplayName -InternalName $FieldInternalName -AddToDefaultView -ErrorAction Stop | Out-Null
        Ok "columna creada."
        # refrescar referencia
        try { $field = Get-PnPField -List $ListTitle -Identity $FieldInternalName -ErrorAction SilentlyContinue } catch { }
    } catch {
        # fallback: crear desde esquema XML (garantiza el internal name)
        try {
            $xml = "<Field Type='Text' DisplayName='$FieldDisplayName' InternalName='$FieldInternalName' Name='$FieldInternalName' Required='FALSE' />"
            Add-PnPFieldFromXml -List $ListTitle -FieldXml $xml -ErrorAction Stop | Out-Null
            Ok "columna creada (vía XML)."
        } catch { throw "No se pudo crear la columna: $_" }
    }
}

# --- 3) Items de muestra ---
Step "Items de muestra"
$items = Get-PnPListItem -List $ListTitle -Fields "ID,Priority" -PageSize 100 -ErrorAction SilentlyContinue
$needSamples = (-not $items) -or ($items.Count -lt 4)
if ($needSamples) {
    $samples = @(
        @{ Title = 'Review PR #142';          Priority = 'high' },
        @{ Title = 'Fix login redirect';      Priority = 'normal' },
        @{ Title = 'Update dependencies';     Priority = 'low' },
        @{ Title = 'Document deployment';     Priority = 'high' }
    )
    foreach ($s in $samples) {
        try {
            Add-PnPListItem -List $ListTitle -Values $s -ErrorAction Stop | Out-Null
        } catch { Warn "No se añadió item '$($s.Title)': $_" }
    }
    Ok "items añadidos."
} else {
    Ok "ya hay $($items.Count) items; no se añaden más."
}

# --- 4) Registrar extensiones SPFx (opcional) ---
if ($RegisterExtensions) {
    Step "Registro de extensiones SPFx"

    # 02 Field Customizer — binding correcto: ClientSideComponentId + Properties
    # directamente sobre el campo (NO dentro de CustomFormatter, que es para
    # column-formatting declarativo y SharePoint ignoraría).
    try {
        Set-PnPField -List $ListTitle -Identity $FieldInternalName -Values @{
            ClientSideComponentId          = [guid]$FieldCustomizerId
            ClientSideComponentProperties   = '{"prefix":"Prio"}'
        } -ErrorAction Stop
        Ok "02 Field Customizer registrado en '$ListTitle > $FieldDisplayName'."
    } catch { Warn "02 no se registró: $_" }

    # 03 Command Set — CustomAction a nivel web (tu PnP no tiene -List en Add-PnPCustomAction)
    try {
        $prev = Get-PnPCustomAction -ErrorAction SilentlyContinue | Where-Object { $_.ClientSideComponentId -eq $CommandSetId }
        if ($prev) { $prev | ForEach-Object { Remove-PnPCustomAction -Identity $_.Id -Force -ErrorAction SilentlyContinue } }
        Add-PnPCustomAction -Name 'Actions' -Title 'Actions' `
            -Location 'ClientSideExtension.ListViewCommandSet' `
            -ClientSideComponentId $CommandSetId -ErrorAction Stop | Out-Null
        Ok "03 Command Set registrado a nivel web (aparece en todas las listas de book-test)."
    } catch { Warn "03 no se registró: $_" }
}

Step "Ver"
$site = (Get-PnPProperty -ClientObject (Get-PnPContext).Web -Property Url)
Write-Host "  Abre: $site/Lists/$ListTitle/AllItems.aspx" -ForegroundColor Yellow
Write-Host "`nListo." -ForegroundColor Green