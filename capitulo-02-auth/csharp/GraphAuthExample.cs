using Azure.Identity;
using Microsoft.Graph;

namespace Capitulo02Auth
{
    /// <summary>
    /// Ejemplo de autenticación con Microsoft Graph usando Client Credentials
    /// Capítulo 2: La Experiencia SharePoint
    /// </summary>
    public class GraphAuthExample
    {
        public static async Task Main(string[] args)
        {
            Console.WriteLine("=== Capítulo 2: Autenticación ===\n");

            // Configuración desde variables de entorno
            string tenantId = Environment.GetEnvironmentVariable("SP_TENANT_ID")!;
            string clientId = Environment.GetEnvironmentVariable("SP_CLIENT_ID")!;
            string clientSecret = Environment.GetEnvironmentVariable("SP_CLIENT_SECRET")!;

            if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
            {
                Console.WriteLine("Error: Faltan variables de entorno.");
                Console.WriteLine("Asegúrate de configurar: SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET");
                return;
            }

            try
            {
                // Crear credencial de cliente
                var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

                // Crear cliente de Graph
                var graphClient = new GraphServiceClient(credential);

                // Obtener información de la organización
                var org = await graphClient.Organization.GetAsync();

                Console.WriteLine("✓ Conexión exitosa a Microsoft Graph");
                Console.WriteLine($"  Tenant: {org?.Value?.FirstOrDefault()?.DisplayName}");
                Console.WriteLine($"  ID: {org?.Value?.FirstOrDefault()?.Id}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Error de autenticación: {ex.Message}");
            }
        }
    }
}
