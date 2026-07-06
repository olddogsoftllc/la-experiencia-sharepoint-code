using LaExperiencia.SharePoint.Common;
using Microsoft.Graph;
using System.Collections.Generic;
using Xunit;

namespace LaExperiencia.SharePoint.Common.Tests;

public class GraphAuthTests
{
    private static readonly string[] EnvKeys = { "TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "CERTIFICATE_PATH", "CERTIFICATE_THUMBPRINT", "CERTIFICATE_PASSWORD" };

    private static void ClearEnv()
    {
        foreach (var key in EnvKeys) Environment.SetEnvironmentVariable(key, null);
    }

    private static void SetSecretEnv()
    {
        ClearEnv();
        Environment.SetEnvironmentVariable("TENANT_ID", "fake-tenant");
        Environment.SetEnvironmentVariable("CLIENT_ID", "fake-client");
        Environment.SetEnvironmentVariable("CLIENT_SECRET", "fake-secret");
    }

    [Fact]
    public void GraphAuthOptions_ThrowsWhenAllEnvMissing()
    {
        ClearEnv();
        var ex = Assert.Throws<ArgumentException>(() => new GraphAuthOptions());
        Assert.Contains("TENANT_ID", ex.Message);
    }

    [Fact]
    public void GraphAuthOptions_ThrowsWhenSecretAndCertificateMissing()
    {
        ClearEnv();
        Environment.SetEnvironmentVariable("TENANT_ID", "t");
        Environment.SetEnvironmentVariable("CLIENT_ID", "c");
        var ex = Assert.Throws<ArgumentException>(() => new GraphAuthOptions());
        Assert.Contains("CLIENT_SECRET", ex.Message);
    }

    [Fact]
    public void GraphAuthOptions_AcceptsSecretEnv()
    {
        SetSecretEnv();
        var options = new GraphAuthOptions();
        Assert.Equal("fake-tenant", options.TenantId);
        Assert.Equal("fake-client", options.ClientId);
        Assert.False(options.UsesCertificate);
    }

    [Fact]
    public void CreateFromSecret_ReturnsNonNullClient_WithFakeEnv()
    {
        SetSecretEnv();
        // ClientSecretCredential is lazy: no network call at construction.
        var client = SharePointGraphClientFactory.CreateFromSecret();
        Assert.NotNull(client);
        Assert.IsType<GraphServiceClient>(client);
    }

    [Fact]
    public void CreateFromSecret_ThrowsWhenEnvMissing()
    {
        ClearEnv();
        Assert.Throws<ArgumentException>(() => SharePointGraphClientFactory.CreateFromSecret());
    }
}