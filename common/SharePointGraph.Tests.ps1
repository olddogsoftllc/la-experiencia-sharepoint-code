#requires -Version 7.0
# Pester tests para el módulo común SharePointGraph.psm1
# Ejecutar:  pwsh -NoProfile -Command "Invoke-Pester -Path ./common/SharePointGraph.Tests.ps1 -Output Detailed"
#
# Los tests no tocan la red: validan la lógica de selección de modo (secret vs cert) y
# la validación de variables de entorno. No prueban la conexión real a Graph (eso requiere
# el módulo Microsoft.Graph.Authentication + creds reales).

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot 'SharePointGraph.psm1') -Force
    # Stub de Connect-MgGraph para que NO intente conectar de verdad; captura el modo usado.
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

    It 'lanza cuando falta TENANT_ID' {
        { Connect-SharePointGraph } | Should -Throw -ExpectedMessage '*TENANT_ID*'
    }

    It 'lanza cuando falta CLIENT_SECRET en modo secret (con tenant y client presentes)' {
        [Environment]::SetEnvironmentVariable('TENANT_ID', 'fake-tenant')
        [Environment]::SetEnvironmentVariable('CLIENT_ID', 'fake-client')
        { Connect-SharePointGraph } | Should -Throw -ExpectedMessage '*CLIENT_SECRET*'
    }

    It 'usa modo secret por defecto (CLIENT_SECRET presente, sin cert)' {
        [Environment]::SetEnvironmentVariable('TENANT_ID', 'fake-tenant')
        [Environment]::SetEnvironmentVariable('CLIENT_ID', 'fake-client')
        [Environment]::SetEnvironmentVariable('CLIENT_SECRET', 'fake-secret')
        { Connect-SharePointGraph } | Should -Throw -ExpectedMessage '*PESTER_SECRET_MODE*'
    }

    It 'auto-detecta modo cert cuando CERTIFICATE_THUMBPRINT presente' {
        [Environment]::SetEnvironmentVariable('TENANT_ID', 'fake-tenant')
        [Environment]::SetEnvironmentVariable('CLIENT_ID', 'fake-client')
        [Environment]::SetEnvironmentVariable('CLIENT_SECRET', 'fake-secret') # debe ignorarse
        [Environment]::SetEnvironmentVariable('CERTIFICATE_THUMBPRINT', 'ABC123')
        { Connect-SharePointGraph } | Should -Throw -ExpectedMessage '*PESTER_CERT_MODE_THUMBPRINT:ABC123*'
    }

    It 'auto-detecta modo cert cuando CERTIFICATE_NAME presente (sin thumbprint)' {
        [Environment]::SetEnvironmentVariable('TENANT_ID', 'fake-tenant')
        [Environment]::SetEnvironmentVariable('CLIENT_ID', 'fake-client')
        [Environment]::SetEnvironmentVariable('CERTIFICATE_NAME', 'my-cert-subject')
        { Connect-SharePointGraph } | Should -Throw -ExpectedMessage '*PESTER_CERT_MODE_NAME:my-cert-subject*'
    }

    It 'fuerza modo cert con -UseCertificate aunque falten thumbprint y name' {
        [Environment]::SetEnvironmentVariable('TENANT_ID', 'fake-tenant')
        [Environment]::SetEnvironmentVariable('CLIENT_ID', 'fake-client')
        [Environment]::SetEnvironmentVariable('CLIENT_SECRET', 'fake-secret')
        { Connect-SharePointGraph -UseCertificate } | Should -Throw -ExpectedMessage '*CERTIFICATE_THUMBPRINT*'
    }
}