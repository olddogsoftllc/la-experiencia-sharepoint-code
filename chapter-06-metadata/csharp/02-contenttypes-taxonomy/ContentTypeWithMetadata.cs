// ContentTypeWithMetadata.cs
// Ejemplo: Crear Content Type con columnas de metadatos administrados

using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace ManagedMetadataExamples;

class ContentTypeWithMetadata
{
    private readonly GraphServiceClient _graphClient;

    public ContentTypeWithMetadata(GraphServiceClient graphClient)
    {
        _graphClient = graphClient;
    }

    public async Task CreateDocumentContentTypeAsync(string siteId)
    {
        Console.WriteLine("📋 Creando Content Type con metadatos...");

        try
        {
            // Crear Content Type
            var contentType = new ContentType
            {
                Name = "Documento Corporativo",
                Description = "Documento con metadatos empresariales",
                Group = "Custom Content Types"
            };

            var created = await _graphClient.Sites[siteId]
                .ContentTypes
                .PostAsync(contentType);

            if (created != null)
            {
                Console.WriteLine($"   ✅ Content Type creado: {created.Id}");
            }

            // Agregar columnas de metadatos (ejemplo simplificado)
            // Nota: En producción, primero crear columnas de sitio vinculadas al Term Store

        }
        catch (Exception ex)
        {
            Console.WriteLine($"   ❌ Error: {ex.Message}");
        }
    }
}
