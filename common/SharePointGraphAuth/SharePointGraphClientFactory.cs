using Azure.Identity;
using Microsoft.Graph;
using System.Security.Cryptography.X509Certificates;

namespace LaExperiencia.SharePoint.Common;

/// <summary>
/// Single factory for creating a <see cref="GraphServiceClient"/> authenticated with client secret
/// (or certificate) from environment variables. Replaces the duplicated auth in each chapter.
/// </summary>
public static class SharePointGraphClientFactory
{
    /// <summary>
    /// Creates a Graph client authenticated with client secret (client credentials, app-only).
    /// Reads TENANT_ID, CLIENT_ID and CLIENT_SECRET from the environment.
    /// </summary>
    public static GraphServiceClient CreateFromSecret()
    {
        var options = new GraphAuthOptions();
        var credential = new ClientSecretCredential(options.TenantId, options.ClientId, options.ClientSecret);
        return new GraphServiceClient(credential);
    }

    /// <summary>
    /// Creates a Graph client authenticated with an X.509 certificate (.pfx).
    /// Reads TENANT_ID, CLIENT_ID, CERTIFICATE_PATH and CERTIFICATE_PASSWORD from the environment.
    /// </summary>
    public static GraphServiceClient CreateFromCertificate()
    {
        var options = new GraphAuthOptions();
        var certificate = new X509Certificate2(options.CertificatePath, options.CertificatePassword);
        var credential = new ClientCertificateCredential(options.TenantId, options.ClientId, certificate);
        return new GraphServiceClient(credential);
    }
}