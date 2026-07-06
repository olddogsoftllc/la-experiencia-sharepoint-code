#requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Document Operations Example

.DESCRIPTION
    Chapter 04: Documents
    Demonstrates upload, download, and search operations for documents

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

function Send-FileToSharePoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationFileName
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        if (-not (Test-Path -Path $FilePath)) {
            throw "File not found: $FilePath"
        }

        Write-Host "Uploading file: $FilePath" -ForegroundColor Cyan

        $fileContent = [System.IO.File]::ReadAllBytes($FilePath)
        $encodedName = [System.Web.HttpUtility]::UrlEncode($DestinationFileName)

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/root:/$encodedName`:/content"

        $uploadHeaders = @{
            'Authorization' = $script:GraphHeaders['Authorization']
            'Content-Type' = 'application/octet-stream'
        }

        $response = Invoke-RestMethod -Uri $uri -Headers $uploadHeaders -Method Put -Body $fileContent

        Write-Host "File uploaded successfully:" -ForegroundColor Green
        Write-Host "  ID: $($response.id)"
        Write-Host "  Name: $($response.name)"
        Write-Host "  Size: $($response.size) bytes"
        Write-Host "  Web URL: $($response.webUrl)"

        return $response
    }
    catch {
        Write-Error "Error uploading file: $_"
        throw
    }
}

function Get-FileFromSharePoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [Parameter(Mandatory = $true)]
        [string]$ItemId,

        [Parameter(Mandatory = $true)]
        [string]$DownloadPath
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Downloading file to: $DownloadPath" -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId/content"

        $response = Invoke-WebRequest -Uri $uri -Headers $script:GraphHeaders -Method Get -OutFile $DownloadPath

        Write-Host "File downloaded successfully to: $DownloadPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Error downloading file: $_"
        throw
    }
}

function Find-SharePointFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Query
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Searching for files: '$Query'" -ForegroundColor Cyan

        $searchBody = @{
            requests = @(
                @{
                    entityTypes = @("driveItem")
                    query = @{
                        queryString = $Query
                    }
                }
            )
        } | ConvertTo-Json -Depth 3

        $uri = "https://graph.microsoft.com/v1.0/search/query"
        $response = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Post -Body $searchBody

        Write-Host "Search results:" -ForegroundColor Green
        Write-Host "-" * 80

        foreach ($result in $response.value) {
            foreach ($hitContainer in $result.hitsContainers) {
                foreach ($hit in $hitContainer.hits) {
                    $resource = $hit.resource
                    Write-Host "Name: $($resource.name)" -ForegroundColor Yellow
                    Write-Host "  Web URL: $($resource.webUrl)"
                    Write-Host "  Size: $($resource.size) bytes"
                    Write-Host "-" * 80
                }
            }
        }

        return $response
    }
    catch {
        Write-Error "Error searching files: $_"
        throw
    }
}

function Get-DriveFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [Parameter()]
        [string]$FolderPath = ""
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }

        Write-Host "Listing files in drive..." -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root/children"

        if ($FolderPath) {
            $encodedPath = [System.Web.HttpUtility]::UrlEncode($FolderPath)
            $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root:/$encodedPath`:/children"
        }

        $response = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Get

        Write-Host "Found $($response.value.Count) items:" -ForegroundColor Green
        Write-Host "-" * 80

        foreach ($item in $response.value) {
            $itemType = if ($item.folder) { "Folder" } else { "File" }
            Write-Host "$($itemType): $($item.name)" -ForegroundColor Yellow
            Write-Host "  ID: $($item.id)"
            Write-Host "  Size: $($item.size) bytes"
            Write-Host "  Last Modified: $($item.lastModifiedDateTime)"
            Write-Host "-" * 80
        }

        return $response.value
    }
    catch {
        Write-Error "Error listing files: $_"
        throw
    }
}

function Get-FileMetadata {
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

        Write-Host "Fetching metadata for item: $ItemId" -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId"
        $response = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Get

        Write-Host "File metadata:" -ForegroundColor Green
        Write-Host "  Name: $($response.name)"
        Write-Host "  ID: $($response.id)"
        Write-Host "  Size: $($response.size) bytes"
        Write-Host "  Created: $($response.createdDateTime)"
        Write-Host "  Modified: $($response.lastModifiedDateTime)"
        Write-Host "  Web URL: $($response.webUrl)"

        return $response
    }
    catch {
        Write-Error "Error getting file metadata: $_"
        throw
    }
}

function Get-DriveLibraries {
    <#
        .SYNOPSIS
        Lists the document libraries (drives) of a site. Read-only.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId
    )

    try {
        if (-not $script:GraphHeaders) { Connect-GraphApi }
        Write-Host "Listing document libraries of site: $SiteId" -ForegroundColor Cyan

        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives"
        $response = Invoke-RestMethod -Uri $uri -Headers $script:GraphHeaders -Method Get

        Write-Host "Found $($response.value.Count) libraries:" -ForegroundColor Green
        Write-Host "-" * 80
        foreach ($drive in $response.value) {
            Write-Host "Library: $($drive.name)" -ForegroundColor Yellow
            Write-Host "  ID: $($drive.id)"
            Write-Host "  Type: $($drive.driveType)"
            Write-Host "-" * 80
        }
        return $response.value
    }
    catch {
        Write-Error "Error listing libraries: $_"
        throw
    }
}

# Main execution: read-only demo against the book-test test site.
try {
    Write-Host "=== SharePoint Document Operations Example ===" -ForegroundColor Yellow

    if (-not $script:GraphHeaders) { Connect-GraphApi }

    $hostname = if ($env:SHAREPOINT_HOSTNAME) { $env:SHAREPOINT_HOSTNAME } else { "olddogsoft1.sharepoint.com" }
    $sitePath = if ($env:SHAREPOINT_SITE_PATH) { $env:SHAREPOINT_SITE_PATH } else { "book-test" }

    # Resolve the site by path to get its ID.
    $encodedPath = [uri]::EscapeDataString("sites/$sitePath")
    $siteUri = "https://graph.microsoft.com/v1.0/sites/${hostname}:/$encodedPath"
    $site = Invoke-RestMethod -Uri $siteUri -Headers $script:GraphHeaders -Method Get

    Get-DriveLibraries -SiteId $site.id

    Write-Host "`nDocument operations completed successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Script execution failed: $_"
    exit 1
}
