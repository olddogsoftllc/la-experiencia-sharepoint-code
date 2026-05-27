#requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Site Operations Example

.DESCRIPTION
    Chapter 03: Sites
    Demonstrates listing, creating, and retrieving SharePoint sites

.NOTES
    Requires environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET
#>

[CmdletBinding()]
param()

#Requires -Modules MSAL.PS

# Global variables for authentication
$script:GraphHeaders = $null

function Connect-GraphApi {
    <cmdletbinding()
    param()

    try {
        $tenantId = $env:TENANT_ID
        $clientId = $env:CLIENT_ID
        $clientSecret = $env:CLIENT_SECRET

        if (-not ($tenantId -and $clientId -and $clientSecret)) {
            throw "Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET"
        }

        Write-Verbose "Authenticating to Microsoft Graph"

        $tokenResponse = Get-MsalToken `
            -ClientId $clientId `
            -ClientSecret (ConvertTo-SecureString $clientSecret -AsPlainText -Force) `
            -TenantId $tenantId `
            -Scopes "https://graph.microsoft.com/.default"

        $script:GraphHeaders = @{
            'Authorization' = "Bearer $($tokenResponse.AccessToken)"
            'Content-Type' = 'application/json'
        }

        Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect: $_"
        throw
    }
}

function Get-AllSites {
    <cmdletbinding()
    param()

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Fetching all sites..." -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites"
        $response = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Get

        Write-Host "`nFound $($response.value.Count) sites:" -ForegroundColor Green
        Write-Host "-" * 80

        foreach ($site in $response.value) {
            Write-Host "Name: $($site.name)" -ForegroundColor Yellow
            Write-Host "  ID: $($site.id)"
            Write-Host "  Web URL: $($site.webUrl)"
            Write-Host "  Display Name: $($site.displayName)"
            Write-Host "-" * 80
        }

        return $response.value
    }
    catch {
        Write-Error "Error listing sites: $_"
        throw
    }
}

function Get-SiteByPath {
    <cmdletbinding()
    param(
        [Parameter(Mandatory = $true)]
        [string]$Hostname,

        [Parameter(Mandatory = $true)]
        [string]$SitePath
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Fetching site: $Hostname/sites/$SitePath" -ForegroundColor Cyan

        $encodedPath = [System.Web.HttpUtility]::UrlEncode("sites/$SitePath")
        $uri = "https://graph.microsoft.com/v1.0/sites/$Hostname`:$encodedPath"
        $site = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Get

        Write-Host "`nSite found:" -ForegroundColor Green
        Write-Host "  Name: $($site.name)"
        Write-Host "  ID: $($site.id)"
        Write-Host "  Web URL: $($site.webUrl)"
        Write-Host "  Description: $($site.description)"

        return $site
    }
    catch {
        Write-Error "Error getting site: $_"
        throw
    }
}

function Get-SiteById {
    <cmdletbinding()
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Fetching site by ID: $SiteId" -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId"
        $site = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Get

        Write-Host "`nSite found:" -ForegroundColor Green
        Write-Host "  Name: $($site.name)"
        Write-Host "  Web URL: $($site.webUrl)"

        return $site
    }
    catch {
        Write-Error "Error getting site by ID: $_"
        throw
    }
}

function Get-RootSite {
    <cmdletbinding()
    param()

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Fetching root site..." -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/root"
        $site = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Get

        Write-Host "`nRoot site:" -ForegroundColor Green
        Write-Host "  Name: $($site.name)"
        Write-Host "  ID: $($site.id)"
        Write-Host "  Web URL: $($site.webUrl)"

        return $site
    }
    catch {
        Write-Error "Error getting root site: $_"
        throw
    }
}

function Search-Sites {
    <cmdletbinding()
    param(
        [Parameter(Mandatory = $true)]
        [string]$Keyword
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Searching for sites with keyword: '$Keyword'" -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites?search='$Keyword'"
        $response = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Get

        Write-Host "`nFound $($response.value.Count) matching sites:" -ForegroundColor Green
        Write-Host "-" * 80

        foreach ($site in $response.value) {
            Write-Host "Name: $($site.name)" -ForegroundColor Yellow
            Write-Host "  Web URL: $($site.webUrl)"
            Write-Host "  Description: $($site.description)"
            Write-Host "-" * 80
        }

        return $response.value
    }
    catch {
        Write-Error "Error searching sites: $_"
        throw
    }
}

# Main execution
try {
    Write-Host "=== SharePoint Site Operations Example ===" -ForegroundColor Yellow

    # Get root site
    Get-RootSite

    # List all sites
    Get-AllSites

    Write-Host "`nSite operations completed successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Script execution failed: $_"
    exit 1
}
