#Requires -Modules Microsoft.Graph.Authentication
<#
.SYNOPSIS
    Chapter 2: Authentication with Microsoft Graph
.DESCRIPTION
    Connects to Microsoft Graph using Client Credentials.
.EXAMPLE
    .\Connect-Graph.ps1
#>

$TenantId = $env:TENANT_ID
$ClientId = $env:CLIENT_ID
$ClientSecret = $env:CLIENT_SECRET

if (-not $TenantId -or -not $ClientId -or -not $ClientSecret) {
    Write-Error "Missing required environment variables"
    exit 1
}

Write-Host "=== Chapter 2: Authentication ===" -ForegroundColor Cyan

$SecureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($ClientId, $SecureSecret)

try {
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential
    $Context = Get-MgContext
    Write-Host "✓ Connection successful" -ForegroundColor Green
    Write-Host "  Tenant: $($Context.TenantId)" -ForegroundColor Gray
    Write-Host "  App: $($Context.ClientId)" -ForegroundColor Gray
}
catch {
    Write-Error "✗ Connection error: $_"
}
