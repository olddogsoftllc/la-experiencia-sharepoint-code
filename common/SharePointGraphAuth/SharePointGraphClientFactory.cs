using Azure.Identity;
using Microsoft.Graph;
using System.Security.Cryptography.X509Certificates;

namespace LaExperiencia.SharePoint.Common;

/// <summary>
/// Fábrica única para crear un <see cref="GraphServiceClient"/> autenticado con client secret
/// (o certificado) a partir de variables de entorno. Reemplaza la auth duplicada en cada capítulo.
/// </summary>
public static class SharePointGraphClientFactory
{
    /// <summary>
    /// Crea un cliente de Graph autenticado con client secret (client credentials, app-only).
    /// Lee TENANT_ID, CLIENT_ID y CLIENT_SECRET del entorno.
    /// </summary>
    public static GraphServiceClient CreateFromSecret()
    {
        var options = new GraphAuthOptions();
        var credential = new ClientSecretCredential(options.TenantId, options.ClientId, options.ClientSecret);
        return new GraphServiceClient(credential);
    }

    /// <summary>
    /// Crea un cliente de Graph autenticado con certificado X.509 (.pfx).
    /// Lee TENANT_ID, CLIENT_ID, CERTIFICATE_PATH y CERTIFICATE_PASSWORD del entorno.
    /// </summary>
    public static GraphServiceClient CreateFromCertificate()
    {
        var options = new GraphAuthOptions();
        var certificate = new X509Certificate2(options.CertificatePath, options.CertificatePassword);
        var credential = new ClientCertificateCredential(options.TenantId, options.ClientId, certificate);
        return new GraphServiceClient(credential);
    }
}