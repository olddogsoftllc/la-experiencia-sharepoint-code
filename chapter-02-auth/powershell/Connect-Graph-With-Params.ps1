#Requires -Modules Microsoft.Graph.Authentication
<#
.SYNOPSIS
    Capítulo 2: Autenticación con Microsoft Graph
.DESCRIPTION
    Conecta a Microsoft Graph usando Client Credentials.
.EXAMPLE
    .\Connect-Graph.ps1
#>

$TenantId = $env:TENANT_ID
$ClientId = $env:CLIENT_ID
$ClientSecret = $env:CLIENT_SECRET

if (-not $TenantId -or -not $ClientId -or -not $ClientSecret) {
    Write-Error "Faltan variables de entorno requeridas"
    exit 1
}

Write-Host "=== Capítulo 2: Autenticación ===" -ForegroundColor Cyan

$SecureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($ClientId, $SecureSecret)

try {
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential
    $Context = Get-MgContext
    Write-Host "✓ Conexión exitosa" -ForegroundColor Green
    Write-Host "  Tenant: $($Context.TenantId)" -ForegroundColor Gray
    Write-Host "  App: $($Context.ClientId)" -ForegroundColor Gray
}
catch {
    Write-Error "✗ Error de conexión: $_"
}
