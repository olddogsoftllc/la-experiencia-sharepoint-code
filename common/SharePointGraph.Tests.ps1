#requires -Version 7.0
# Pester tests for the common SharePointGraph.psm1 module
# Run:  pwsh -NoProfile -Command "Invoke-Pester -Path ./common/SharePointGraph.Tests.ps1 -Output Detailed"
#
# The tests do not touch the network: they validate mode-selection logic (secret vs cert)
# and environment-variable validation. They do not test the real Graph connection (that
# requires the Microsoft.Graph.Authentication module + real creds).

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot 'SharePointGraph.psm1') -Force
    # Stub of Connect-MgGraph so it does NOT actually connect; captures the mode used.
    function global:Connect-MgGraph {
        param(
            [string] $TenantId,
            [string] $ClientId,
            [string] $CertificateThumbprint,
            [string] $CertificateName,
            [System.Management.Automation.PSCredential] $ClientSecretCredential,
            [switch] $NoWelcome
        )
        if ($CertificateThumbprint) {
            throw "PESTER_CERT_MODE_THUMBPRINT:$CertificateThumbprint"
        } elseif ($CertificateName) {
            throw "PESTER_CERT_MODE_NAME:$CertificateName"
        } else {
            throw "PESTER_SECRET_MODE"
        }
    }
    function global:Get-MgContext { $null }
    function global:Disconnect-MgGraph { }
}

Describe 'Connect-SharePointGraph' {
    BeforeEach {
        foreach ($key in 'TENANT_ID','CLIENT_ID','CLIENT_SECRET','CERTIFICATE_THUMBPRINT','CERTIFICATE_NAME') {
            [Environment]::SetEnvironmentVariable($key, $null)
        }
    }

    It 'throws when TENANT_ID is missing' {
        { Connect-SharePointGraph } | Should -Throw -ExpectedMessage '*TENANT_ID*'
    }

    It 'throws when CLIENT_SECRET is missing in secret mode (with tenant and client present)' {
        [Environment]::SetEnvironmentVariable('TENANT_ID', 'fake-tenant')
        [Environment]::SetEnvironmentVariable('CLIENT_ID', 'fake-client')
        { Connect-SharePointGraph } | Should -Throw -ExpectedMessage '*CLIENT_SECRET*'
    }

    It 'uses secret mode by default (CLIENT_SECRET present, no cert)' {
        [Environment]::SetEnvironmentVariable('TENANT_ID', 'fake-tenant')
        [Environment]::SetEnvironmentVariable('CLIENT_ID', 'fake-client')
        [Environment]::SetEnvironmentVariable('CLIENT_SECRET', 'fake-secret')
        { Connect-SharePointGraph } | Should -Throw -ExpectedMessage '*PESTER_SECRET_MODE*'
    }

    It 'auto-detects cert mode when CERTIFICATE_THUMBPRINT is present' {
        [Environment]::SetEnvironmentVariable('TENANT_ID', 'fake-tenant')
        [Environment]::SetEnvironmentVariable('CLIENT_ID', 'fake-client')
        [Environment]::SetEnvironmentVariable('CLIENT_SECRET', 'fake-secret') # must be ignored
        [Environment]::SetEnvironmentVariable('CERTIFICATE_THUMBPRINT', 'ABC123')
        { Connect-SharePointGraph } | Should -Throw -ExpectedMessage '*PESTER_CERT_MODE_THUMBPRINT:ABC123*'
    }

    It 'auto-detects cert mode when CERTIFICATE_NAME is present (no thumbprint)' {
        [Environment]::SetEnvironmentVariable('TENANT_ID', 'fake-tenant')
        [Environment]::SetEnvironmentVariable('CLIENT_ID', 'fake-client')
        [Environment]::SetEnvironmentVariable('CERTIFICATE_NAME', 'my-cert-subject')
        { Connect-SharePointGraph } | Should -Throw -ExpectedMessage '*PESTER_CERT_MODE_NAME:my-cert-subject*'
    }

    It 'forces cert mode with -UseCertificate even when thumbprint and name are missing' {
        [Environment]::SetEnvironmentVariable('TENANT_ID', 'fake-tenant')
        [Environment]::SetEnvironmentVariable('CLIENT_ID', 'fake-client')
        [Environment]::SetEnvironmentVariable('CLIENT_SECRET', 'fake-secret')
        { Connect-SharePointGraph -UseCertificate } | Should -Throw -ExpectedMessage '*CERTIFICATE_THUMBPRINT*'
    }
}