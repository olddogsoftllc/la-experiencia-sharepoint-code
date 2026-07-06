# common/SharePointGraph.psm1
# Autenticación compartida para los ejemplos de La Experiencia SharePoint (PowerShell).
# Reemplaza el token hand-rolled con MSAL.PS que cada script duplicaba.
# Requiere: módulo Microsoft.Graph.Authentication (Install-Module Microsoft.Graph.Authentication -Scope CurrentUser)
#
# Soporta dos modos de autenticación (app-only):
#   - client secret (por defecto): lee TENANT_ID/CLIENT_ID/CLIENT_SECRET.
#   - certificado: lee TENANT_ID/CLIENT_ID + CERTIFICATE_THUMBPRINT (cert en store
#     Cert:\CurrentUser\My o Cert:\LocalMachine\My) o CERTIFICATE_NAME (subject en store).
#     Se usa automáticamente cuando hay CERTIFICATE_THUMBPRINT/CERTIFICATE_NAME; se puede
#     forzar con -UseCertificate. Para un .pfx en disco, impórtalo primero al store:
#       Import-PfxCertificate -FilePath .\mi-cert.pfx -Cert Cert:\CurrentUser\My -Password $ss
#     y usa su Thumbprint. (Práctica estándar de PowerShell; Connect-MgGraph lee del store.)

#Requires -Version 7.0

function Connect-SharePointGraph {
    <#
        .SYNOPSIS
        Conecta a Microsoft Graph con app-only leyendo variables de entorno.

        .DESCRIPTION
        Modo client secret por defecto (TENANT_ID/CLIENT_ID/CLIENT_SECRET). Modo certificado
        si CERTIFICATE_THUMBPRINT o CERTIFICATE_NAME están presentes, o si -UseCertificate se
        especifica. En modo certificado, Connect-MgGraph usa el certificado del store de
        Windows/Linux (CurrentUser\My o LocalMachine\My).

        .PARAMETER UseCertificate
        Fuerza el modo certificado aunque CLIENT_SECRET esté presente.
    #>
    [CmdletBinding()]
    param(
        [switch] $UseCertificate
    )

    $tenantId = [Environment]::GetEnvironmentVariable('TENANT_ID')
    $clientId = [Environment]::GetEnvironmentVariable('CLIENT_ID')

    if ([string]::IsNullOrWhiteSpace($tenantId) -or [string]::IsNullOrWhiteSpace($clientId)) {
        throw "Faltan las variables de entorno TENANT_ID y/o CLIENT_ID."
    }

    $certThumbprint = [Environment]::GetEnvironmentVariable('CERTIFICATE_THUMBPRINT')
    $certName       = [Environment]::GetEnvironmentVariable('CERTIFICATE_NAME')

    # Auto-detección: si hay thumbprint/name de certificado, modo certificado.
    $useCertMode = if ($UseCertificate) { $true }
                   elseif ($certThumbprint -or $certName) { $true }
                   else { $false }

    if ($useCertMode) {
        # Modo certificado (store). Connect-MgGraph acepta -CertificateThumbprint o -CertificateName.
        if ([string]::IsNullOrWhiteSpace($certThumbprint) -and [string]::IsNullOrWhiteSpace($certName)) {
            throw "Modo certificado seleccionado pero faltan CERTIFICATE_THUMBPRINT y CERTIFICATE_NAME."
        }
        $params = @{
            TenantId = $tenantId
            ClientId = $clientId
            NoWelcome = $true
        }
        if (-not [string]::IsNullOrWhiteSpace($certThumbprint)) {
            $params.CertificateThumbprint = $certThumbprint
        } else {
            $params.CertificateName = $certName
        }
        Connect-MgGraph @params
        return
    }

    # Modo client secret.
    $clientSecret = [Environment]::GetEnvironmentVariable('CLIENT_SECRET')
    if ([string]::IsNullOrWhiteSpace($clientSecret)) {
        throw "Falta la variable de entorno CLIENT_SECRET (o define CERTIFICATE_THUMBPRINT/CERTIFICATE_NAME para usar certificado)."
    }

    $secure = ConvertTo-SecureString $clientSecret -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($clientId, $secure)

    Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $credential -NoWelcome
}

function Invoke-GraphRequest {
    <#
        .SYNOPSIS
        Wrapper de Invoke-MgGraphRequest que asegura la conexión y usa el contexto compartido.
        Reemplaza Invoke-RestMethod + token bearer manual en los ejemplos.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Uri,
        [Parameter()] [string] $Method = 'GET',
        [Parameter()] $Body,
        [Parameter()] [string] $ContentType = 'application/json'
    )

    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        Connect-SharePointGraph
    }

    $params = @{ Uri = $Uri; Method = $Method }
    if ($PSBoundParameters.ContainsKey('Body')) { $params.Body = $Body; $params.ContentType = $ContentType }
    Invoke-MgGraphRequest @params
}

function Disconnect-SharePointGraph {
    [CmdletBinding()]
    param()
    if (Get-MgContext -ErrorAction SilentlyContinue) { Disconnect-MgGraph | Out-Null }
}

Export-ModuleMember -Function Connect-SharePointGraph, Invoke-GraphRequest, Disconnect-SharePointGraph