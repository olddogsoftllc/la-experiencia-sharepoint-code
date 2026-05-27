using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace Chapter02Auth
{
    /// <summary>
    /// Chapter 2: Authentication with Token Caching
    /// Demonstrates persistent token cache for console applications
    /// </summary>
    public class AuthWithTokenCache
    {
        public static async Task Main(string[] args)
        {
            Console.WriteLine("=== Chapter 2: Authentication with Token Cache ===\n");

            string tenantId = Environment.GetEnvironmentVariable("SP_TENANT_ID")!;
            string clientId = Environment.GetEnvironmentVariable("SP_CLIENT_ID")!;
            string clientSecret = Environment.GetEnvironmentVariable("SP_CLIENT_SECRET")!;

            if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
            {
                Console.WriteLine("Error: Missing required environment variables");
                return;
            }

            try
            {
                // Configure token cache
                var storageProperties = new StorageCreationPropertiesBuilder(
                    "cache.dat",
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
                    .Build();

                var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
                var graphClient = new GraphServiceClient(credential);

                // Get user information (test connection)
                var site = await graphClient.Sites.GetAsync();

                Console.WriteLine("✓ Authentication successful with token caching");
                Console.WriteLine($"  Sites found: {site?.Value?.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Error: {ex.Message}");
            }
        }
    }
}
