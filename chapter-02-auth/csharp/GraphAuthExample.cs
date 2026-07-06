using Azure.Identity;
using Microsoft.Graph;

namespace Chapter02Auth
{
    /// <summary>
    /// Chapter 2: Authentication with Microsoft Graph using Client Credentials
    /// The SharePoint Experience
    /// </summary>
    public class GraphAuthExample
    {
        public static async Task Main(string[] args)
        {
            Console.WriteLine("=== Chapter 2: Authentication ===\n");

            // Configuration from environment variables
            string tenantId = Environment.GetEnvironmentVariable("TENANT_ID")!;
            string clientId = Environment.GetEnvironmentVariable("CLIENT_ID")!;
            string clientSecret = Environment.GetEnvironmentVariable("CLIENT_SECRET")!;

            if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
            {
                Console.WriteLine("Error: Missing environment variables.");
                Console.WriteLine("Make sure to set: TENANT_ID, CLIENT_ID, CLIENT_SECRET");
                return;
            }

            try
            {
                // Create client credential
                var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

                // Create Graph client
                var graphClient = new GraphServiceClient(credential);

                // Get organization information
                var org = await graphClient.Organization.GetAsync();

                Console.WriteLine("✓ Successfully connected to Microsoft Graph");
                Console.WriteLine($"  Tenant: {org?.Value?.FirstOrDefault()?.DisplayName}");
                Console.WriteLine($"  ID: {org?.Value?.FirstOrDefault()?.Id}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Authentication error: {ex.Message}");
            }
        }
    }
}
