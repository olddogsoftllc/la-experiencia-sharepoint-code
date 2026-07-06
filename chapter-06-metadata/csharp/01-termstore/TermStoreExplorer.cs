// TermStoreExplorer.cs
// Ejemplo: Explorar el Term Store de SharePoint usando Microsoft Graph SDK.
// El cliente de Graph se inyecta por constructor (DI) desde el módulo común.

using LaExperiencia.SharePoint.Common;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace LaExperiencia.SharePoint.Chapter06.Metadata;

class TermStoreExplorer
{
    private readonly GraphServiceClient _graphClient;

    public TermStoreExplorer(GraphServiceClient graphClient)
    {
        _graphClient = graphClient;
    }

    public async Task ExploreAsync()
    {
        Console.WriteLine("=" .Repeat(60));
        Console.WriteLine("Explorando Term Store");
        Console.WriteLine("=".Repeat(60));
        Console.WriteLine();

        try
        {
            // Obtener información del Term Store
            var termStore = await _graphClient.Sites["root"]
                .TermStore
                .GetAsync();

            Console.WriteLine("📚 Term Store Info:");
            Console.WriteLine($"   ID: {termStore?.Id}");
            Console.WriteLine($"   Default Language: {termStore?.DefaultLanguageTag}");
            Console.WriteLine($"   Languages: {string.Join(", ", termStore?.LanguageTags ?? new List<string>())}");
            Console.WriteLine();

            // Listar grupos
            var groups = await _graphClient.Sites["root"]
                .TermStore
                .Groups
                .GetAsync();

            Console.WriteLine($"📂 Grupos encontrados: {groups?.Value?.Count ?? 0}\n");

            foreach (var group in groups?.Value ?? new List<Microsoft.Graph.Models.TermStore.Group>())
            {
                Console.WriteLine($"   📁 {group.DisplayName}");
                Console.WriteLine($"      ID: {group.Id}");
                Console.WriteLine($"      Descripción: {group.Description ?? "N/A"}");

                // Obtener conjuntos de términos
                var sets = await _graphClient.Sites["root"]
                    .TermStore
                    .Groups[group.Id]
                    .Sets
                    .GetAsync();

                foreach (var set in sets?.Value ?? new List<Microsoft.Graph.Models.TermStore.Set>())
                {
                    var setName = set.LocalizedNames?.FirstOrDefault()?.Name ?? set.Id;
                    Console.WriteLine($"         📚 {setName}");
                }
                Console.WriteLine();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Entry point: construye el cliente con el módulo común y explora el Term Store (solo lectura).
    /// </summary>
    public static async Task Main(string[] args)
    {
        try
        {
            var graphClient = SharePointGraphClientFactory.CreateFromSecret();
            var explorer = new TermStoreExplorer(graphClient);
            await explorer.ExploreAsync();
            Console.WriteLine("\nTerm Store exploration completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Environment.Exit(1);
        }
    }
}

// Extension para string.Repeat
public static class StringExtensions
{
    public static string Repeat(this string s, int count)
    {
        return string.Concat(Enumerable.Repeat(s, count));
    }
}
