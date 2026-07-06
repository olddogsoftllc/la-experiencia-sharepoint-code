# provision-book-test.ps1  (v2 — robust)
# Provisions the book-test site to validate the SPFx examples from chapter 14.
# Idempotent. Uses try/catch and clear messages for each step.
#
# Usage:
#   Connect-PnPOnline -Url "https://olddogsoft1.sharepoint.com/sites/book-test" -Interactive -ClientId "<your-app-id>"
#   .\provision-book-test.ps1                  # creates list + column + items
#   .\provision-book-test.ps1 -RegisterExtensions   # also registers 02 and 03
#   .\provision-book-test.ps1 -Reset           # deletes and recreates Tasks (clean)

[CmdletBinding()]
param(
    [switch]$RegisterExtensions,
    [switch]$Reset
)

$ListTitle = 'Tasks'
$FieldDisplayName = 'Priority'
$FieldInternalName = 'Priority'

# Manifest IDs of the .sppkg already in the App Catalog
$FieldCustomizerId = '53f8a795-a835-45fd-bfa7-01249906d75b'   # 02-field-customizer-priority
$CommandSetId       = 'ae7678f8-f77c-4a54-a9c7-f22dfd3cc70c'   # 03-command-set-actions

function Step($msg) { Write-Host "› $msg" -ForegroundColor Cyan }
function Ok($msg)   { Write-Host "  $msg" -ForegroundColor Green }
function Warn($msg) { Write-Host "  $msg" -ForegroundColor Yellow }

# --- 0) Connection ---
$ctx = Get-PnPContext
if (-not $ctx) { throw "No PnP connection. Run Connect-PnPOnline first." }
Step "Connected to $((Get-PnPProperty -ClientObject $ctx.Web -Property Url))"

# --- -Reset: delete Tasks if it exists ---
if ($Reset) {
    Step "Reset: deleting list '$ListTitle' if it exists"
    try {
        $existing = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
        if ($existing) { Remove-PnPList -Identity $ListTitle -Force; Ok "list deleted." }
        else { Warn "did not exist." }
    } catch { Warn "could not delete: $_" }
}

# --- 1) Create the list if it does not exist ---
Step "List '$ListTitle'"
$list = $null
try { $list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue } catch { }
if ($list) {
    Ok "already exists (id=$($list.Id))."
} else {
    try {
        $list = New-PnPList -Title $ListTitle -Template GenericList -ErrorAction Stop
        Ok "list created."
    } catch { throw "Could not create the list: $_" }
}

# --- 2) Create the "Priority" column if it does not exist ---
Step "Column '$FieldDisplayName'"
$field = $null
try { $field = Get-PnPField -List $ListTitle -Identity $FieldInternalName -ErrorAction SilentlyContinue } catch { }
if ($field) {
    Ok "already exists (internal name: $($field.InternalName))."
} else {
    try {
        # Add-PnPFieldText creates a text field; -AddToDefaultView shows it in the view
        Add-PnPFieldText -List $ListTitle -DisplayName $FieldDisplayName -InternalName $FieldInternalName -AddToDefaultView -ErrorAction Stop | Out-Null
        Ok "column created."
        # refresh reference
        try { $field = Get-PnPField -List $ListTitle -Identity $FieldInternalName -ErrorAction SilentlyContinue } catch { }
    } catch {
        # fallback: create from XML schema (guarantees the internal name)
        try {
            $xml = "<Field Type='Text' DisplayName='$FieldDisplayName' InternalName='$FieldInternalName' Name='$FieldInternalName' Required='FALSE' />"
            Add-PnPFieldFromXml -List $ListTitle -FieldXml $xml -ErrorAction Stop | Out-Null
            Ok "column created (via XML)."
        } catch { throw "Could not create the column: $_" }
    }
}

# --- 3) Sample items ---
Step "Sample items"
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
        } catch { Warn "Item '$($s.Title)' was not added: $_" }
    }
    Ok "items added."
} else {
    Ok "already $($items.Count) items; none added."
}

# --- 4) Register SPFx extensions (optional) ---
if ($RegisterExtensions) {
    Step "Registering SPFx extensions"

    # 02 Field Customizer — correct binding: ClientSideComponentId + Properties
    # directly on the field (NOT inside CustomFormatter, which is for
    # declarative column formatting and SharePoint would ignore).
    try {
        Set-PnPField -List $ListTitle -Identity $FieldInternalName -Values @{
            ClientSideComponentId          = [guid]$FieldCustomizerId
            ClientSideComponentProperties   = '{"prefix":"Prio"}'
        } -ErrorAction Stop
        Ok "02 Field Customizer registered on '$ListTitle > $FieldDisplayName'."
    } catch { Warn "02 was not registered: $_" }

    # 03 Command Set — web-level CustomAction (your PnP has no -List on Add-PnPCustomAction)
    try {
        $prev = Get-PnPCustomAction -ErrorAction SilentlyContinue | Where-Object { $_.ClientSideComponentId -eq $CommandSetId }
        if ($prev) { $prev | ForEach-Object { Remove-PnPCustomAction -Identity $_.Id -Force -ErrorAction SilentlyContinue } }
        Add-PnPCustomAction -Name 'Actions' -Title 'Actions' `
            -Location 'ClientSideExtension.ListViewCommandSet' `
            -ClientSideComponentId $CommandSetId -ErrorAction Stop | Out-Null
        Ok "03 Command Set registered at web level (appears on every list in book-test)."
    } catch { Warn "03 was not registered: $_" }
}

Step "Verify"
$site = (Get-PnPProperty -ClientObject (Get-PnPContext).Web -Property Url)
Write-Host "  Open: $site/Lists/$ListTitle/AllItems.aspx" -ForegroundColor Yellow
Write-Host "`nDone." -ForegroundColor Green