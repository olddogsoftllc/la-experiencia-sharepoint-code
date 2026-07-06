// DeltaQueryManager.cs
// Ejemplo de Delta Queries con Microsoft Graph para sincronizacion incremental
// Requiere: Microsoft.Graph, Microsoft.Identity.Client
// Referencia: Capitulo 7 - Automatizacion y Flujos

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
        /// Realiza una consulta delta inicial para obtener todos los items
        /// y almacena el token delta para sincronizaciones posteriores.
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

            // Guardar token delta para proxima sincronizacion
            var deltaToken = ExtractDeltaToken(deltaResponse?.OdataDeltaLink);
            if (!string.IsNullOrEmpty(deltaToken))
            {
                _deltaTokens[$"{driveId}_{folderId ?? "root"}"] = deltaToken!;
                Console.WriteLine($"   Token delta guardado: {deltaToken.Substring(0, 50)}...");
            }

            return results;
        }

        /// <summary>
        /// Realiza una sincronizacion delta usando el token almacenado.
        /// Retorna solo los items que han cambiado desde la ultima consulta.
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

            // Actualizar token delta
            var newToken = ExtractDeltaToken(deltaResponse?.OdataDeltaLink);
            if (!string.IsNullOrEmpty(newToken))
            {
                _deltaTokens[key] = newToken!;
            }

            return results;
        }

        /// <summary>
        /// Consulta delta para listas de SharePoint (items de lista).
        /// </summary>
        public async Task<List<Microsoft.Graph.Models.ListItem>> GetListDeltaAsync(string siteId, string listId)
        {
            Console.WriteLine($"🔄 Consulta delta para lista: {listId}");

            // El SDK v5 no expone delta de items de lista en el fluent path;
            // se invoca vía URL cruda usando el RequestAdapter.
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

            // ListItemCollectionResponse no expone OdataDeltaLink; viene en AdditionalData.
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
        /// Extrae el token delta de la URL de delta link.
        /// </summary>
        private string? ExtractDeltaToken(string? deltaLink)
        {
            if (string.IsNullOrEmpty(deltaLink))
                return null;

            // Parseo manual del query string (portable, sin depender de System.Web.HttpUtility).
            // El token suele estar en el parametro ?token= o &deltaToken= (codificado).
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
        /// Muestra los tokens delta almacenados.
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
        /// Entry point: construye el cliente con el módulo común y ejecuta una consulta delta
        /// inicial (solo lectura) sobre la primera biblioteca del sitio de pruebas book-test.
        /// </summary>
        public static async Task Main(string[] args)
        {
            try
            {
                var graphClient = SharePointGraphClientFactory.CreateFromSecret();
                var manager = new DeltaQueryManager(graphClient);

                var hostname = Environment.GetEnvironmentVariable("SHAREPOINT_HOSTNAME") ?? "olddogsoft1.sharepoint.com";
                var sitePath = Environment.GetEnvironmentVariable("SHAREPOINT_SITE_PATH") ?? "book-test";

                // Resolver el sitio y su primera biblioteca.
                var site = await graphClient.Sites[$"{hostname}:/sites/{sitePath}"].GetAsync();
                if (site == null)
                {
                    Console.WriteLine("No se encontró el sitio de pruebas.");
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