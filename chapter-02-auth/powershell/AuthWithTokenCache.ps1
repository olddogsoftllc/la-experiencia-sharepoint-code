#requires -Version 7.0
<#
.SYNOPSIS
    Token Cache Authentication Example

.DESCRIPTION
    Chapter 02: Authentication
    Demonstrates persistent token caching for improved performance in PowerShell

.NOTES
    Requires environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET
#>

[CmdletBinding()]
param()

#Requires -Modules MSAL.PS

# Token cache variables
$script:TokenCache = @{}
$script:CacheExpiry = $null
$script:TokenLifetimeMinutes = 55  # Token typically valid for 60 minutes

function Get-CachedAccessToken {
    <cmdletbinding()
    param()

    try {
        # Check if we have a valid cached token
        if ($script:CacheExpiry -and (Get-Date) -lt $script:CacheExpiry) {
            if ($script:TokenCache.ContainsKey('AccessToken')) {
                Write-Verbose "Using cached access token"
                return $script:TokenCache['AccessToken']
            }
        }

        # Need to request a new token
        Write-Verbose "Requesting new access token"

        $tenantId = $env:TENANT_ID
        $clientId = $env:CLIENT_ID
        $clientSecret = $env:CLIENT_SECRET

        if (-not ($tenantId -and $clientId -and $clientSecret)) {
            throw "Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET"
        }

        # Acquire new token
        $tokenResponse = Get-MsalToken `
            -ClientId $clientId `
            -ClientSecret (ConvertTo-SecureString $clientSecret -AsPlainText -Force) `
            -TenantId $tenantId `
            -Scopes "https://graph.microsoft.com/.default"

        # Store token in cache
        $script:TokenCache['AccessToken'] = $tokenResponse.AccessToken
        $script:CacheExpiry = (Get-Date).AddMinutes($script:TokenLifetimeMinutes)

        Write-Host "New access token acquired and cached" -ForegroundColor Green

        return $tokenResponse.AccessToken
    }
    catch {
        Write-Error "Failed to acquire token: $_"
        throw
    }
}

function Clear-TokenCache {
    <cmdletbinding()
    param()

    $script:TokenCache.Clear()
    $script:CacheExpiry = $null
    Write-Host "Token cache cleared" -ForegroundColor Yellow
}

function Get-CacheStatus {
    <cmdletbinding()
    param()

    $status = @{
        HasToken = $script:TokenCache.ContainsKey('AccessToken')
        ExpiryTime = $script:CacheExpiry
        IsValid = $false
    }

    if ($status.HasToken -and $script:CacheExpiry) {
        $status.IsValid = (Get-Date) -lt $script:CacheExpiry
    }

    return $status
}

function Invoke-GraphRequestWithCache {
    <cmdletbinding()
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter()]
        [string]$Method = 'GET'
    )

    try {
        $token = Get-CachedAccessToken

        $headers = @{
            'Authorization' = "Bearer $token"
            'Content-Type' = 'application/json'
        }

        $response = Invoke-RestMethod -Uri $Uri -Headers $headers -Method $Method
        return $response
    }
    catch {
        Write-Error "Request failed: $_"
        throw
    }
}

# Main execution
try {
    Write-Host "=== Token Cache Authentication Example ===" -ForegroundColor Yellow

    # First call - will acquire new token
    Write-Host "`nFirst call (new token):" -ForegroundColor Cyan
    $token1 = Get-CachedAccessToken
    Write-Host "Token acquired: $($token1.Substring(0, 20))..." -ForegroundColor Gray

    # Check cache status
    $status = Get-CacheStatus
    Write-Host "Cache status: HasToken=$($status.HasToken), IsValid=$($status.IsValid)" -ForegroundColor Gray

    # Second call - should use cached token
    Write-Host "`nSecond call (from cache):" -ForegroundColor Cyan
    $token2 = Get-CachedAccessToken
    Write-Host "Token from cache: $($token2.Substring(0, 20))..." -ForegroundColor Gray

    # Verify tokens match
    Write-Host "`nTokens match: $($token1 -eq $token2)" -ForegroundColor Green

    # Clear cache
    Clear-TokenCache
    $status = Get-CacheStatus
    Write-Host "After clear - HasToken: $($status.HasToken)" -ForegroundColor Yellow

    Write-Host "`nToken caching example completed!" -ForegroundColor Green
}
catch {
    Write-Error "Script execution failed: $_"
    exit 1
}
