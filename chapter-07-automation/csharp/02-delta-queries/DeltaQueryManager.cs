// DeltaQueryManager.cs
// Example of Delta Queries with Microsoft Graph for incremental synchronization
// Requires: Microsoft.Graph, Microsoft.Identity.Client
// Reference: Chapter 7 - Automation and Flows

using LaExperiencia.SharePoint.Common;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace LaExperiencia.SharePoint.Chapter07.Automation
{
    public class DeltaQueryManager
    {
        private readonly GraphServiceClient _graphClient;
        private readonly Dictionary<string, string> _deltaTokens;

        public DeltaQueryManager(GraphServiceClient graphClient)
        {
            _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
            _deltaTokens = new Dictionary<string, string>();
        }

        /// <summary>
        /// Performs an initial delta query to get all items
        /// and stores the delta token for subsequent synchronizations.
        /// </summary>
        public async Task<List<DriveItem>> GetInitialDeltaAsync(string driveId, string? folderId = null)
        {
            Console.WriteLine($"🔄 Ejecutando consulta delta inicial en drive: {driveId}");

            var deltaResponse = await _graphClient.Drives[driveId]
                .Items[folderId ?? "root"]
                .Delta
                .GetAsDeltaGetResponseAsync();

            var results = new List<DriveItem>();

            if (deltaResponse?.Value != null)
            {
                results.AddRange(deltaResponse.Value);
                Console.WriteLine($"   Items obtenidos: {deltaResponse.Value.Count}");
            }

            // Save delta token for next synchronization
            var deltaToken = ExtractDeltaToken(deltaResponse?.OdataDeltaLink);
            if (!string.IsNullOrEmpty(deltaToken))
            {
                _deltaTokens[$"{driveId}_{folderId ?? "root"}"] = deltaToken!;
                Console.WriteLine($"   Token delta guardado: {deltaToken.Substring(0, 50)}...");
            }

            return results;
        }

        /// <summary>
        /// Performs a delta synchronization using the stored token.
        /// Returns only the items that have changed since the last query.
        /// </summary>
        public async Task<List<DriveItem>> GetDeltaChangesAsync(string driveId, string? folderId = null)
        {
            var key = $"{driveId}_{folderId ?? "root"}";

            if (!_deltaTokens.ContainsKey(key))
            {
                Console.WriteLine("⚠️  No existe token delta. Ejecutando consulta inicial...");
                return await GetInitialDeltaAsync(driveId, folderId);
            }

            var deltaToken = _deltaTokens[key];
            Console.WriteLine($"🔄 Sincronizando cambios con token delta...");

            var requestInfo = new Microsoft.Graph.Drives.Item.Items.Item.Delta.DeltaRequestBuilder(
                $"https://graph.microsoft.com/v1.0/drives/{driveId}/items/{folderId ?? "root"}/delta?token={deltaToken}",
                _graphClient.RequestAdapter);

            var deltaResponse = await requestInfo.GetAsDeltaGetResponseAsync();
            var results = new List<DriveItem>();

            if (deltaResponse?.Value != null)
            {
                results.AddRange(deltaResponse.Value);
                Console.WriteLine($"   Cambios detectados: {deltaResponse.Value.Count}");

                foreach (var item in deltaResponse.Value)
                {
                    var changeType = item.Deleted != null ? "🗑️  Eliminado" : "✏️  Modificado";
                    Console.WriteLine($"   {changeType}: {item.Name}");
                }
            }

            // Update delta token
            var newToken = ExtractDeltaToken(deltaResponse?.OdataDeltaLink);
            if (!string.IsNullOrEmpty(newToken))
            {
                _deltaTokens[key] = newToken!;
            }

            return results;
        }

        /// <summary>
        /// Delta query for SharePoint lists (list items).
        /// </summary>
        public async Task<List<Microsoft.Graph.Models.ListItem>> GetListDeltaAsync(string siteId, string listId)
        {
            Console.WriteLine($"🔄 Consulta delta para lista: {listId}");

            // The v5 SDK does not expose list item delta in the fluent path;
            // it is invoked via a raw URL using the RequestAdapter.
            var requestInfo = new Microsoft.Kiota.Abstractions.RequestInformation
            {
                HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
                URI = new Uri($"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items/delta")
            };

            var deltaResponse = await _graphClient.RequestAdapter.SendAsync<
                Microsoft.Graph.Models.ListItemCollectionResponse>(
                requestInfo,
                Microsoft.Graph.Models.ListItemCollectionResponse.CreateFromDiscriminatorValue);

            var results = new List<Microsoft.Graph.Models.ListItem>();

            if (deltaResponse?.Value != null)
            {
                results.AddRange(deltaResponse.Value);
            }

            // ListItemCollectionResponse does not expose OdataDeltaLink; it comes in AdditionalData.
            string? deltaLink = null;
            if (deltaResponse?.AdditionalData != null &&
                deltaResponse.AdditionalData.TryGetValue("@odata.deltaLink", out var dlObj) &&
                dlObj != null)
            {
                deltaLink = dlObj.ToString();
            }

            var deltaToken = ExtractDeltaToken(deltaLink);
            if (!string.IsNullOrEmpty(deltaToken))
            {
                _deltaTokens[$"list_{siteId}_{listId}"] = deltaToken!;
            }

            return results;
        }

        /// <summary>
        /// Extracts the delta token from the delta link URL.
        /// </summary>
        private string? ExtractDeltaToken(string? deltaLink)
        {
            if (string.IsNullOrEmpty(deltaLink))
                return null;

            // Manual parsing of the query string (portable, without depending on System.Web.HttpUtility).
            // The token is usually in the ?token= or &deltaToken= parameter (encoded).
            int queryStart = deltaLink.IndexOf('?');
            if (queryStart < 0) return null;
            string query = deltaLink.Substring(queryStart + 1);

            foreach (string pair in query.Split('&'))
            {
                int eq = pair.IndexOf('=');
                if (eq < 0) continue;
                string key = pair.Substring(0, eq);
                string value = pair.Substring(eq + 1);
                if (key == "token" || key == "deltaToken")
                    return Uri.UnescapeDataString(value);
            }
            return null;
        }

        /// <summary>
        /// Shows the stored delta tokens.
        /// </summary>
        public void ShowStoredTokens()
        {
            Console.WriteLine("\n📋 Tokens Delta Almacenados:");
            foreach (var kvp in _deltaTokens)
            {
                Console.WriteLine($"   {kvp.Key}: {kvp.Value.Substring(0, Math.Min(50, kvp.Value.Length))}...");
            }
        }

        /// <summary>
        /// Entry point: builds the client with the common module and runs an initial
        /// delta query (read-only) over the first library of the book-test test site.
        /// </summary>
        public static async Task Main(string[] args)
        {
            try
            {
                var graphClient = SharePointGraphClientFactory.CreateFromSecret();
                var manager = new DeltaQueryManager(graphClient);

                var hostname = Environment.GetEnvironmentVariable("SHAREPOINT_HOSTNAME") ?? "olddogsoft1.sharepoint.com";
                var sitePath = Environment.GetEnvironmentVariable("SHAREPOINT_SITE_PATH") ?? "book-test";

                // Resolve the site and its first library.
                var site = await graphClient.Sites[$"{hostname}:/sites/{sitePath}"].GetAsync();
                if (site == null)
                {
                    Console.WriteLine("Test site not found.");
                    return;
                }
                var drives = await graphClient.Sites[site.Id].Drives.GetAsync();
                var firstDrive = drives?.Value?.FirstOrDefault();
                if (firstDrive == null)
                {
                    Console.WriteLine("No se encontraron bibliotecas en el sitio de pruebas.");
                    return;
                }

                var driveId = firstDrive.Id ?? throw new InvalidOperationException("La biblioteca no tiene ID.");
                Console.WriteLine($"Ejecutando delta inicial en la biblioteca: {firstDrive.Name}");
                var items = await manager.GetInitialDeltaAsync(driveId, "root");
                Console.WriteLine($"\nItems obtenidos en la delta inicial: {items.Count}");
                manager.ShowStoredTokens();
                Console.WriteLine("\nDelta query example completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
}