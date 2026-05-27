#requires -Version 7.0
<#
.SYNOPSIS
    Certificate-Based Authentication Example

.DESCRIPTION
    Chapter 02: Authentication
    Demonstrates secure authentication using X.509 certificates in PowerShell

.NOTES
    Requires environment variables: TENANT_ID, CLIENT_ID
    Also requires: CERTIFICATE_THUMBPRINT or CERTIFICATE_PATH
#>

[CmdletBinding()]
param()

#Requires -Modules MSAL.PS

function Get-CertificateFromStore {
    <cmdletbinding()
    param(
        [Parameter(Mandatory = $true)]
        [string]$Thumbprint
    )

    try {
        Write-Verbose "Searching for certificate with thumbprint: $Thumbprint"

        # Search in current user store
        $cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $Thumbprint }

        if (-not $cert) {
            # Search in local machine store
            $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $Thumbprint }
        }

        if (-not $cert) {
            throw "Certificate with thumbprint '$Thumbprint' not found in certificate store"
        }

        Write-Host "Certificate loaded from certificate store" -ForegroundColor Green
        return $cert
    }
    catch {
        Write-Error "Failed to load certificate from store: $_"
        throw
    }
}

function Get-CertificateFromFile {
    <cmdletbinding()
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter()]
        [SecureString]$Password
    )

    try {
        Write-Verbose "Loading certificate from file: $Path"

        if (-not (Test-Path -Path $Path)) {
            throw "Certificate file not found: $Path"
        }

        $cert = if ($Password) {
            [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Path, $Password)
        } else {
            [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Path)
        }

        Write-Host "Certificate loaded from file" -ForegroundColor Green
        return $cert
    }
    catch {
        Write-Error "Failed to load certificate from file: $_"
        throw
    }
}

function Test-Certificate {
    <cmdletbinding()
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    try {
        Write-Host "`nCertificate Details:" -ForegroundColor Cyan
        Write-Host "  Subject: $($Certificate.Subject)" -ForegroundColor Gray
        Write-Host "  Issuer: $($Certificate.Issuer)" -ForegroundColor Gray
        Write-Host "  Thumbprint: $($Certificate.Thumbprint)" -ForegroundColor Gray
        Write-Host "  Valid From: $($Certificate.NotBefore)" -ForegroundColor Gray
        Write-Host "  Valid Until: $($Certificate.NotAfter)" -ForegroundColor Gray
        Write-Host "  Has Private Key: $($Certificate.HasPrivateKey)" -ForegroundColor Gray

        # Validate certificate
        $now = Get-Date

        if ($Certificate.NotAfter -lt $now) {
            throw "Certificate has expired"
        }

        if ($Certificate.NotBefore -gt $now) {
            throw "Certificate is not yet valid"
        }

        if (-not $Certificate.HasPrivateKey) {
            throw "Certificate does not have a private key"
        }

        Write-Host "`nCertificate validation passed" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Certificate validation failed: $_"
        throw
    }
}

function Get-GraphTokenWithCertificate {
    <cmdletbinding()
    param()

    try {
        $tenantId = $env:TENANT_ID
        $clientId = $env:CLIENT_ID
        $thumbprint = $env:CERTIFICATE_THUMBPRINT
        $certPath = $env:CERTIFICATE_PATH

        if (-not ($tenantId -and $clientId)) {
            throw "TENANT_ID and CLIENT_ID environment variables are required"
        }

        # Load certificate
        $certificate = $null
        if ($thumbprint) {
            $certificate = Get-CertificateFromStore -Thumbprint $thumbprint
        }
        elseif ($certPath) {
            $certPassword = $env:CERTIFICATE_PASSWORD
            $securePassword = if ($certPassword) { ConvertTo-SecureString $certPassword -AsPlainText -Force } else { $null }
            $certificate = Get-CertificateFromFile -Path $certPath -Password $securePassword
        }
        else {
            throw "Please provide CERTIFICATE_THUMBPRINT or CERTIFICATE_PATH environment variable"
        }

        # Validate certificate
        Test-Certificate -Certificate $certificate

        Write-Verbose "Acquiring access token with certificate"

        # Get token using certificate
        $tokenResponse = Get-MsalToken `
            -ClientId $clientId `
            -TenantId $tenantId `
            -Scopes "https://graph.microsoft.com/.default" `
            -ClientCertificate $certificate

        Write-Host "Successfully authenticated using certificate" -ForegroundColor Green

        return $tokenResponse.AccessToken
    }
    catch {
        Write-Error "Certificate authentication failed: $_"
        throw
    }
}

# Main execution
try {
    Write-Host "=== Certificate-Based Authentication Example ===" -ForegroundColor Yellow

    $token = Get-GraphTokenWithCertificate
    Write-Host "`nToken acquired: $($token.Substring(0, 30))..." -ForegroundColor Gray

    Write-Host "`nCertificate authentication completed successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Script execution failed: $_"
    exit 1
}
