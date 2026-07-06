#Requires -Version 7.0
<#
.SYNOPSIS
    Ejemplo de Delta Queries con Microsoft Graph para sincronizacion incremental.

.DESCRIPTION
    Realiza consultas delta para detectar cambios en documentos de SharePoint
    sin necesidad de descargar todo el contenido.

.PARAMETER DriveId
    ID del drive de SharePoint.

.PARAMETER FolderId
    ID de la carpeta (por defecto: root).

.PARAMETER AccessToken
    Token de acceso a Microsoft Graph.

.EXAMPLE
    .\Get-DeltaChanges.ps1 -DriveId "b!abc123" -AccessToken "eyJ0..."

.NOTES
    Referencia: Capitulo 7 - Automatizacion y Flujos
    Autor: Efren Ignacio Garza Castillo - Old Dog Soft LLC
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DriveId,

    [Parameter()]
    [string]$FolderId = "root",

    [Parameter(Mandatory = $true)]
    [string]$AccessToken
)

$script:DeltaTokens = @{}
$BaseUrl = "https://graph.microsoft.com/v1.0"

function Get-GraphHeaders {
    return @{
        "Authorization" = "Bearer $AccessToken"
        "Accept"        = "application/json"
        "Content-Type"  = "application/json"
    }
}

function Get-InitialDelta {
    param([string]$DriveId, [string]$FolderId)

    Write-Host "🔄 Ejecutando consulta delta inicial en drive: $DriveId" -ForegroundColor Cyan

    $url = "$BaseUrl/drives/$DriveId/items/$FolderId/delta"
    $response = Invoke-RestMethod -Uri $url -Headers (Get-GraphHeaders) -Method GET

    $items = $response.value
    Write-Host "   Items obtenidos: $($items.Count)" -ForegroundColor Green

    # Guardar token delta
    if ($response.'@odata.deltaLink') {
        $token = Extract-DeltaToken -DeltaLink $response.'@odata.deltaLink'
        if ($token) {
            $script:DeltaTokens["$DriveId`_$FolderId"] = $token
            Write-Host "   Token delta guardado: $($token.Substring(0, [Math]::Min(50, $token.Length)))..." -ForegroundColor DarkGray
        }
    }

    return $items
}

function Get-DeltaChanges {
    param([string]$DriveId, [string]$FolderId)

    $key = "$DriveId`_$FolderId"

    if (-not $script:DeltaTokens.ContainsKey($key)) {
        Write-Host "⚠️  No existe token delta. Ejecutando consulta inicial..." -ForegroundColor Yellow
        return Get-InitialDelta -DriveId $DriveId -FolderId $FolderId
    }

    $token = $script:DeltaTokens[$key]
    Write-Host "🔄 Sincronizando cambios con token delta..." -ForegroundColor Cyan

    $url = "$BaseUrl/drives/$DriveId/items/$FolderId/delta?token=$token"
    $response = Invoke-RestMethod -Uri $url -Headers (Get-GraphHeaders) -Method GET

    $items = $response.value
    Write-Host "   Cambios detectados: $($items.Count)" -ForegroundColor Green

    foreach ($item in $items) {
        $name = $item.name ?? "unknown"
        if ($item.deleted) {
            Write-Host "   🗑️  Eliminado: $name" -ForegroundColor Red
        } else {
            Write-Host "   ✏️  Modificado: $name" -ForegroundColor Blue
        }
    }

    # Actualizar token
    if ($response.'@odata.deltaLink') {
        $newToken = Extract-DeltaToken -DeltaLink $response.'@odata.deltaLink'
        if ($newToken) {
            $script:DeltaTokens[$key] = $newToken
            Write-Host "   Token delta actualizado" -ForegroundColor DarkGray
        }
    }

    # Verificar paginacion
    if ($response.'@odata.nextLink') {
        Write-Host "   📄 Hay mas paginas: $($response.'@odata.nextLink')" -ForegroundColor Magenta
    }

    return $items
}

function Get-ListDelta {
    param([string]$SiteId, [string]$ListId)

    Write-Host "🔄 Consulta delta para lista: $ListId" -ForegroundColor Cyan

    $url = "$BaseUrl/sites/$SiteId/lists/$ListId/items/delta"
    $response = Invoke-RestMethod -Uri $url -Headers (Get-GraphHeaders) -Method GET

    $items = $response.value

    if ($response.'@odata.deltaLink') {
        $token = Extract-DeltaToken -DeltaLink $response.'@odata.deltaLink'
        if ($token) {
            $script:DeltaTokens["list_$SiteId`_$ListId"] = $token
        }
    }

    return $items
}

function Extract-DeltaToken {
    param([string]$DeltaLink)

    if (-not $DeltaLink) { return $null }

    $uri = [System.Uri]::new($DeltaLink)
    $query = [System.Web.HttpUtility]::ParseQueryString($uri.Query)

    return $query["token"] ?? $query["deltaToken"]
}

function Show-StoredTokens {
    Write-Host "`n📋 Tokens Delta Almacenados:" -ForegroundColor Yellow
    foreach ($entry in $script:DeltaTokens.GetEnumerator()) {
        $truncated = $entry.Value.Substring(0, [Math]::Min(50, $entry.Value.Length))
        Write-Host "   $($entry.Key): $truncated..." -ForegroundColor DarkGray
    }
}

# ============== Ejecucion Principal ==============

try {
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║     EJEMPLO: DELTA QUERIES CON MICROSOFT GRAPH              ║" -ForegroundColor Cyan
    Write-Host "║     Sincronizacion Incremental de Documentos                 ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    # Consulta inicial
    Write-Host "--- Consulta Delta Inicial ---" -ForegroundColor White
    $null = Get-InitialDelta -DriveId $DriveId -FolderId $FolderId

    # Sincronizacion de cambios
    Write-Host "`n--- Sincronizacion de Cambios ---" -ForegroundColor White
    $null = Get-DeltaChanges -DriveId $DriveId -FolderId $FolderId

    Show-StoredTokens

    Write-Host "`n✅ Example completed" -ForegroundColor Green

} catch {
    Write-Host "❌ Error: $_" -ForegroundColor Red
    Write-Host "   $($_.Exception.Message)" -ForegroundColor DarkRed
    exit 1
}
