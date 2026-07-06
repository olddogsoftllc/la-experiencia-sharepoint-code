#requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Permission Operations Example

.DESCRIPTION
    Chapter 05: Permissions
    Demonstrates managing sharing links and permissions

.NOTES
    Requires environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET
#>

[CmdletBinding()]
param()

#Requires -Modules MSAL.PS

$script:GraphHeaders = $null

function Connect-GraphApi {
    [CmdletBinding()]
    param()

    try {
        $tenantId = $env:TENANT_ID
        $clientId = $env:CLIENT_ID
        $clientSecret = $env:CLIENT_SECRET

        if (-not ($tenantId -and $clientId -and $clientSecret)) {
            throw "Missing required environment variables"
        }

        $tokenResponse = Get-MsalToken `
            -ClientId $clientId `
            -ClientSecret (ConvertTo-SecureString $clientSecret -AsPlainText -Force) `
            -TenantId $tenantId `
            -Scopes "https://graph.microsoft.com/.default"

        $script:GraphHeaders = @{
            'Authorization' = "Bearer $($tokenResponse.AccessToken)"
            'Content-Type' = 'application/json'
        }

        Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect: $_"
        throw
    }
}

function New-AnonymousSharingLink {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [Parameter(Mandatory = $true)]
        [string]$ItemId,

        [Parameter()]
        [ValidateSet("view", "edit")]
        [string]$LinkType = "view"
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Creating anonymous sharing link for item: $ItemId" -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId/createLink"

        $body = @{
            type = $LinkType
            scope = "anonymous"
        } | ConvertTo-Json

        $response = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Post -Body $body

        Write-Host "Sharing link created successfully:" -ForegroundColor Green
        Write-Host "  Link: $($response.link.webUrl)"
        Write-Host "  Type: $($response.link.type)"
        Write-Host "  Scope: $($response.link.scope)"

        return $response
    }
    catch {
        Write-Error "Error creating sharing link: $_"
        throw
    }
}

function New-OrganizationSharingLink {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [Parameter(Mandatory = $true)]
        [string]$ItemId,

        [Parameter()]
        [ValidateSet("view", "edit")]
        [string]$LinkType = "view"
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Creating organization sharing link for item: $ItemId" -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId/createLink"

        $body = @{
            type = $LinkType
            scope = "organization"
        } | ConvertTo-Json

        $response = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Post -Body $body

        Write-Host "Organization sharing link created successfully:" -ForegroundColor Green
        Write-Host "  Link: $($response.link.webUrl)"
        Write-Host "  Type: $($response.link.type)"
        Write-Host "  Scope: $($response.link.scope)"

        return $response
    }
    catch {
        Write-Error "Error creating organization link: $_"
        throw
    }
}

function Grant-UserAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [Parameter(Mandatory = $true)]
        [string]$ItemId,

        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter()]
        [ValidateSet("write", "read")]
        [string]$Role = "write"
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Granting $Role access to $UserEmail for item: $ItemId" -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId/invite"

        $body = @{
            recipients = @(
                @{
                    email = $UserEmail
                }
            )
            roles = @($Role)
            sendNotification = $true
            message = "You have been granted access to this document."
        } | ConvertTo-Json -Depth 3

        $response = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Post -Body $body

        Write-Host "Access granted successfully:" -ForegroundColor Green
        Write-Host "  Permission ID: $($response.value[0].id)"
        Write-Host "  Roles: $($response.value[0].roles -join ', ')"

        return $response
    }
    catch {
        Write-Error "Error granting access: $_"
        throw
    }
}

function Get-ItemPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [Parameter(Mandatory = $true)]
        [string]$ItemId
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Listing permissions for item: $ItemId" -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId/permissions"
        $response = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Get

        Write-Host "Found $($response.value.Count) permissions:" -ForegroundColor Green
        Write-Host "-" * 80

        foreach ($permission in $response.value) {
            Write-Host "Permission ID: $($permission.id)" -ForegroundColor Yellow
            Write-Host "  Roles: $($permission.roles -join ', ')"

            if ($permission.link) {
                Write-Host "  Link Type: $($permission.link.type)"
                Write-Host "  Link Scope: $($permission.link.scope)"
                Write-Host "  Web URL: $($permission.link.webUrl)"
            }

            if ($permission.grantedTo.user) {
                Write-Host "  Granted To: $($permission.grantedTo.user.displayName) ($($permission.grantedTo.user.email))"
            }

            Write-Host "-" * 80
        }

        return $response.value
    }
    catch {
        Write-Error "Error listing permissions: $_"
        throw
    }
}

function Remove-Permission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [Parameter(Mandatory = $true)]
        [string]$ItemId,

        [Parameter(Mandatory = $true)]
        [string]$PermissionId
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Deleting permission: $PermissionId" -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId/permissions/$PermissionId"
        Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Delete

        Write-Host "Permission deleted successfully" -ForegroundColor Green
    }
    catch {
        Write-Error "Error deleting permission: $_"
        throw
    }
}

function Get-SitePermissions {
    <#
        .SYNOPSIS
        Lista los permisos de un sitio. Solo lectura.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }
        Write-Host "Listing permissions of site: $SiteId" -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/permissions"
        $response = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Get

        Write-Host "Found $($response.value.Count) site permissions:" -ForegroundColor Green
        Write-Host "-" * 80
        foreach ($permission in $response.value) {
            Write-Host "Permission ID: $($permission.id)" -ForegroundColor Yellow
            Write-Host "  Roles: $($permission.roles -join ', ')"
            Write-Host "-" * 80
        }
        return $response.value
    }
    catch {
        Write-Error "Error listing site permissions: $_"
        throw
    }
}

# Main execution: demo de solo lectura contra el sitio de pruebas book-test.
try {
    Write-Host "=== SharePoint Permission Operations Example ===" -ForegroundColor Yellow

    if (-not $script:GraphHeaders) { Connect-GraphApi }

    $hostname = if ($env:SHAREPOINT_HOSTNAME) { $env:SHAREPOINT_HOSTNAME } else { "olddogsoft1.sharepoint.com" }
    $sitePath = if ($env:SHAREPOINT_SITE_PATH) { $env:SHAREPOINT_SITE_PATH } else { "book-test" }

    # Resolver el sitio por path para obtener su ID.
    $encodedPath = [uri]::EscapeDataString("sites/$sitePath")
    $siteUri = "https://graph.microsoft.com/v1.0/sites/${hostname}:/$encodedPath"
    $site = Invoke-RestMethod -Uri $siteUri -Headers $script:GraphHeaders -Method Get

    Get-SitePermissions -SiteId $site.id

    Write-Host "`nPermission operations completed successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Script execution failed: $_"
    exit 1
}
