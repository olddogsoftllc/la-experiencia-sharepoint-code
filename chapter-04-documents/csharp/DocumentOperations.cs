/*
 * DocumentOperations.cs
 * Chapter 04: Documents
 *
 * SharePoint Document Operations Example
 * Demonstrates upload, download, and search operations for documents
 */

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace SharePointExperience.Chapter04
{
    /// <summary>
    /// Example class demonstrating SharePoint document operations
    /// </summary>
    public class DocumentOperations
    {
        private readonly GraphServiceClient _graphClient;

        /// <summary>
        /// Initializes a new instance of DocumentOperations
        /// </summary>
        public DocumentOperations()
        {
            string tenantId = Environment.GetEnvironmentVariable("TENANT_ID")
                ?? throw new ArgumentException("TENANT_ID environment variable is required");
            string clientId = Environment.GetEnvironmentVariable("CLIENT_ID")
                ?? throw new ArgumentException("CLIENT_ID environment variable is required");
            string clientSecret = Environment.GetEnvironmentVariable("CLIENT_SECRET")
                ?? throw new ArgumentException("CLIENT_SECRET environment variable is required");

            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            _graphClient = new GraphServiceClient(credential);
        }

        /// <summary>
        /// Uploads a file to a SharePoint document library
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <param name="driveId">The drive ID</param>
        /// <param name="filePath">Local path to the file</param>
        /// <param name="destinationFileName">Name for the file in SharePoint</param>
        /// <returns>The uploaded drive item</returns>
        public async Task<DriveItem> UploadFileAsync(
            string siteId,
            string driveId,
            string filePath,
            string destinationFileName)
        {
            try
            {
                Console.WriteLine($"Uploading file: {filePath}");

                // Read file content
                byte[] fileContent = await File.ReadAllBytesAsync(filePath);
                using var stream = new MemoryStream(fileContent);

                // Upload the file
                var uploadedItem = await _graphClient
                    .Sites[siteId]
                    .Drives[driveId]
                    .Items["root"]
                    .ItemWithPath(destinationFileName)
                    .Content
                    .PutAsync(stream);

                Console.WriteLine($"File uploaded successfully:");
                Console.WriteLine($"  ID: {uploadedItem?.Id}");
                Console.WriteLine($"  Name: {uploadedItem?.Name}");
                Console.WriteLine($"  Size: {uploadedItem?.Size} bytes");
                Console.WriteLine($"  Web URL: {uploadedItem?.WebUrl}");

                return uploadedItem;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error uploading file: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Downloads a file from SharePoint
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <param name="driveId">The drive ID</param>
        /// <param name="itemId">The item ID of the file</param>
        /// <param name="downloadPath">Local path to save the file</param>
        public async Task DownloadFileAsync(
            string siteId,
            string driveId,
            string itemId,
            string downloadPath)
        {
            try
            {
                Console.WriteLine($"Downloading file to: {downloadPath}");

                // Get the file content
                var stream = await _graphClient
                    .Sites[siteId]
                    .Drives[driveId]
                    .Items[itemId]
                    .Content
                    .GetAsync();

                // Save to local file
                using var fileStream = new FileStream(downloadPath, FileMode.Create);
                await stream.CopyToAsync(fileStream);

                Console.WriteLine($"File downloaded successfully to: {downloadPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error downloading file: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Searches for files across SharePoint
        /// </summary>
        /// <param name="query">Search query</param>
        public async Task SearchFilesAsync(string query)
        {
            try
            {
                Console.WriteLine($"Searching for files: '{query}'");

                // Search for drive items
                var searchResults = await _graphClient.Search.Query.PostAsync(new Microsoft.Graph.Search.Query.QueryPostRequestBody
                {
                    Requests = new List<Microsoft.Graph.Search.Models.SearchRequest>
                    {
                        new Microsoft.Graph.Search.Models.SearchRequest
                        {
                            EntityTypes = new List<Microsoft.Graph.Search.Models.EntityType>
                            {
                                Microsoft.Graph.Search.Models.EntityType.DriveItem
                            },
                            Query = new Microsoft.Graph.Search.Models.SearchQuery
                            {
                                QueryString = query
                            }
                        }
                    }
                });

                Console.WriteLine($"Search results:");
                Console.WriteLine(new string('-', 80));

                foreach (var result in searchResults?.Value ?? new List<Microsoft.Graph.Search.Models.SearchResponse>())
                {
                    foreach (var hit in result?.HitsContainers ?? new List<Microsoft.Graph.Search.Models.SearchHitsContainer>())
                    {
                        foreach (var item in hit?.Hits ?? new List<Microsoft.Graph.Search.Models.SearchHit>())
                        {
                            var resource = item?.Resource as DriveItem;
                            if (resource != null)
                            {
                                Console.WriteLine($"Name: {resource.Name}");
                                Console.WriteLine($"  Web URL: {resource.WebUrl}");
                                Console.WriteLine($"  Size: {resource.Size} bytes");
                                Console.WriteLine(new string('-', 80));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error searching files: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Lists files in a specific folder
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <param name="driveId">The drive ID</param>
        /// <param name="folderPath">Path to the folder</param>
        public async Task ListFilesInFolderAsync(string siteId, string driveId, string folderPath)
        {
            try
            {
                Console.WriteLine($"Listing files in folder: {folderPath}");

                var items = await _graphClient
                    .Sites[siteId]
                    .Drives[driveId]
                    .Root
                    .ItemWithPath(folderPath)
                    .Children
                    .GetAsync();

                Console.WriteLine($"Found {items?.Value?.Count ?? 0} items:");
                Console.WriteLine(new string('-', 80));

                foreach (var item in items?.Value ?? new List<DriveItem>())
                {
                    string itemType = item.Folder != null ? "Folder" : "File";
                    Console.WriteLine($"{itemType}: {item.Name}");
                    Console.WriteLine($"  ID: {item.Id}");
                    Console.WriteLine($"  Size: {item.Size} bytes");
                    Console.WriteLine($"  Last Modified: {item.LastModifiedDateTime}");
                    Console.WriteLine(new string('-', 80));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing files: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Gets file metadata
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <param name="driveId">The drive ID</param>
        /// <param name="itemId">The item ID</param>
        /// <returns>The drive item metadata</returns>
        public async Task<DriveItem> GetFileMetadataAsync(string siteId, string driveId, string itemId)
        {
            try
            {
                Console.WriteLine($"Fetching metadata for item: {itemId}");

                var item = await _graphClient
                    .Sites[siteId]
                    .Drives[driveId]
                    .Items[itemId]
                    .GetAsync();

                Console.WriteLine($"File metadata:");
                Console.WriteLine($"  Name: {item?.Name}");
                Console.WriteLine($"  ID: {item?.Id}");
                Console.WriteLine($"  Size: {item?.Size} bytes");
                Console.WriteLine($"  Created: {item?.CreatedDateTime}");
                Console.WriteLine($"  Modified: {item?.LastModifiedDateTime}");
                Console.WriteLine($"  Web URL: {item?.WebUrl}");

                return item;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting file metadata: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Entry point for the example
        /// </summary>
        public static async Task Main(string[] args)
        {
            try
            {
                var docOps = new DocumentOperations();

                // Example: List files in root folder
                // Uncomment and provide actual IDs to test
                // await docOps.ListFilesInFolderAsync("site-id", "drive-id", "");

                Console.WriteLine("\nDocument operations class initialized successfully!");
                Console.WriteLine("Use the methods to perform upload, download, and search operations.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
}
