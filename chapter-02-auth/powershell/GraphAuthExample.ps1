#requires -Version 7.0
<#
.SYNOPSIS
    Microsoft Graph Client Credentials Authentication Example

.DESCRIPTION
    Chapter 02: Authentication
    Demonstrates authenticating to Microsoft Graph using client credentials flow in PowerShell

.NOTES
    Requires environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET
#>

[CmdletBinding()]
param()

# Import required modules
#Requires -Modules MSAL.PS

function Get-GraphAccessToken {
    <cmdletbinding()
    param()

    try {
        # Retrieve credentials from environment variables
        $tenantId = $env:TENANT_ID
        $clientId = $env:CLIENT_ID
        $clientSecret = $env:CLIENT_SECRET

        # Validate environment variables
        if (-not $tenantId) {
            throw "TENANT_ID environment variable is required"
        }
        if (-not $clientId) {
            throw "CLIENT_ID environment variable is required"
        }
        if (-not $clientSecret) {
            throw "CLIENT_SECRET environment variable is required"
        }

        Write-Verbose "Requesting access token for Microsoft Graph"

        # Acquire token using MSAL
        $tokenResponse = Get-MsalToken `
            -ClientId $clientId `
            -ClientSecret (ConvertTo-SecureString $clientSecret -AsPlainText -Force) `
            -TenantId $tenantId `
            -Scopes "https://graph.microsoft.com/.default"

        Write-Host "Successfully authenticated to Microsoft Graph" -ForegroundColor Green

        return $tokenResponse.AccessToken
    }
    catch {
        Write-Error "Authentication failed: $_"
        throw
    }
}

function Connect-GraphWithToken {
    <cmdletbinding()
    param()

    try {
        $accessToken = Get-GraphAccessToken

        # Create headers for API calls
        $script:graphHeaders = @{
            'Authorization' = "Bearer $accessToken"
            'Content-Type' = 'application/json'
        }

        Write-Host "Connected to Microsoft Graph API" -ForegroundColor Green
        return $script:graphHeaders
    }
    catch {
        Write-Error "Failed to connect: $_"
        throw
    }
}

function Test-GraphConnection {
    <cmdletbinding()
    param()

    try {
        if (-not $script:graphHeaders) {
            Connect-GraphWithToken
        }

        Write-Verbose "Testing connection to Microsoft Graph"

        # Test connection by retrieving organization details
        $uri = "https://graph.microsoft.com/v1.0/organization"
        $response = Invoke-RestMethod -Uri $uri -Headers $script:graphHeaders -Method Get

        $orgName = $response.value[0].displayName
        Write-Host "Connected to tenant: $orgName" -ForegroundColor Cyan

        return $response
    }
    catch {
        Write-Error "Connection test failed: $_"
        throw
    }
}

# Main execution

try {
    Write-Host "=== Microsoft Graph Authentication Example ===" -ForegroundColor Yellow

    # Connect to Graph
    Connect-GraphWithToken

    # Test the connection
    Test-GraphConnection

    Write-Host "`nAuthentication completed successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Script execution failed: $_"
    exit 1
}
