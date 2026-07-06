// TermStoreExplorer.cs
// Example: Explore the SharePoint Term Store using Microsoft Graph SDK.
// The Graph client is injected via constructor (DI) from the common module.

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
            // Get Term Store information
            var termStore = await _graphClient.Sites["root"]
                .TermStore
                .GetAsync();

            Console.WriteLine("📚 Term Store Info:");
            Console.WriteLine($"   ID: {termStore?.Id}");
            Console.WriteLine($"   Default Language: {termStore?.DefaultLanguageTag}");
            Console.WriteLine($"   Languages: {string.Join(", ", termStore?.LanguageTags ?? new List<string>())}");
            Console.WriteLine();

            // List groups
            var groups = await _graphClient.Sites["root"]
                .TermStore
                .Groups
                .GetAsync();

            Console.WriteLine($"📂 Grupos encontrados: {groups?.Value?.Count ?? 0}\n");

            foreach (var group in groups?.Value ?? new List<Microsoft.Graph.Models.TermStore.Group>())
            {
                Console.WriteLine($"   📁 {group.DisplayName}");
                Console.WriteLine($"      ID: {group.Id}");
                Console.WriteLine($"      Description: {group.Description ?? "N/A"}");

                // Get term sets
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
    /// Entry point: builds the client with the common module and explores the Term Store (read-only).
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

// Extension for string.Repeat
public static class StringExtensions
{
    public static string Repeat(this string s, int count)
    {
        return string.Concat(Enumerable.Repeat(s, count));
    }
}
