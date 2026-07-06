using System;

namespace LaExperiencia.SharePoint.Common;

/// <summary>
/// Lee y valida la configuración de autenticación de Azure AD desde variables de entorno.
/// Una sola fuente de verdad para todos los ejemplos (C#, y equivalente en Python/JS/PS/Java).
/// </summary>
public sealed class GraphAuthOptions
{
    public string TenantId { get; } = Environment.GetEnvironmentVariable("TENANT_ID") ?? string.Empty;
    public string ClientId { get; } = Environment.GetEnvironmentVariable("CLIENT_ID") ?? string.Empty;
    public string ClientSecret { get; } = Environment.GetEnvironmentVariable("CLIENT_SECRET") ?? string.Empty;
    public string CertificatePath { get; } = Environment.GetEnvironmentVariable("CERTIFICATE_PATH") ?? string.Empty;
    public string CertificateThumbprint { get; } = Environment.GetEnvironmentVariable("CERTIFICATE_THUMBPRINT") ?? string.Empty;
    public string CertificatePassword { get; } = Environment.GetEnvironmentVariable("CERTIFICATE_PASSWORD") ?? string.Empty;

    public GraphAuthOptions()
    {
        if (string.IsNullOrWhiteSpace(TenantId))
            throw new ArgumentException("Falta la variable de entorno TENANT_ID.", nameof(TenantId));
        if (string.IsNullOrWhiteSpace(ClientId))
            throw new ArgumentException("Falta la variable de entorno CLIENT_ID.", nameof(ClientId));
        if (string.IsNullOrWhiteSpace(ClientSecret) && string.IsNullOrWhiteSpace(CertificatePath))
            throw new ArgumentException("Se requiere CLIENT_SECRET o CERTIFICATE_PATH en las variables de entorno.");
    }

    public bool UsesCertificate => !string.IsNullOrWhiteSpace(CertificatePath);
}