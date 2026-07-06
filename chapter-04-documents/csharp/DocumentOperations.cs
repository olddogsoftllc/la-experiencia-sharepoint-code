/*
 * DocumentOperations.cs
 * Chapter 04: Documents
 *
 * SharePoint Document Operations Example
 * Demonstrates upload, download, and search operations for documents
 */

using System;
using System.IO;
using System.Threading.Tasks;
using LaExperiencia.SharePoint.Common;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace LaExperiencia.SharePoint.Chapter04.Documents
{
    /// <summary>
    /// Example class demonstrating SharePoint document operations.
    /// The Graph client is injected via constructor (DI) from the common module.
    /// </summary>
    public class DocumentOperations
    {
        private readonly GraphServiceClient _graphClient;

        /// <summary>Creates an instance with an injected Graph client (DI).</summary>
        public DocumentOperations(GraphServiceClient graphClient)
        {
            _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        }

        /// <summary>
        /// Uploads a file to a SharePoint document library
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <param name="driveId">The drive ID</param>
        /// <param name="filePath">Local path to the file</param>
        /// <param name="destinationFileName">Name for the file in SharePoint</param>
        /// <returns>The uploaded drive item</returns>
        public async Task<DriveItem?> UploadFileAsync(
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
                    .Drives[driveId]
                    .Items[itemId]
                    .Content
                    .GetAsync();

                if (stream == null)
                {
                    throw new InvalidOperationException("Graph returned no content for the requested item.");
                }

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
                var searchResults = await _graphClient.Search.Query.PostAsQueryPostResponseAsync(new Microsoft.Graph.Search.Query.QueryPostRequestBody
                {
                    Requests = new System.Collections.Generic.List<Microsoft.Graph.Models.SearchRequest>
                    {
                        new Microsoft.Graph.Models.SearchRequest
                        {
                            EntityTypes = new System.Collections.Generic.List<Microsoft.Graph.Models.EntityType?>
                            {
                                Microsoft.Graph.Models.EntityType.DriveItem
                            },
                            Query = new Microsoft.Graph.Models.SearchQuery
                            {
                                QueryString = query
                            }
                        }
                    }
                });

                Console.WriteLine($"Search results:");
                Console.WriteLine(new string('-', 80));

                foreach (var result in searchResults?.Value ?? new System.Collections.Generic.List<Microsoft.Graph.Models.SearchResponse>())
                {
                    foreach (var hit in result?.HitsContainers ?? new System.Collections.Generic.List<Microsoft.Graph.Models.SearchHitsContainer>())
                    {
                        foreach (var item in hit?.Hits ?? new System.Collections.Generic.List<Microsoft.Graph.Models.SearchHit>())
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
                    .Drives[driveId]
                    .Root
                    .ItemWithPath(folderPath)
                    .Children
                    .GetAsync();

                Console.WriteLine($"Found {items?.Value?.Count ?? 0} items:");
                Console.WriteLine(new string('-', 80));

                foreach (var item in items?.Value ?? new System.Collections.Generic.List<DriveItem>())
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
        public async Task<DriveItem?> GetFileMetadataAsync(string siteId, string driveId, string itemId)
        {
            try
            {
                Console.WriteLine($"Fetching metadata for item: {itemId}");

                var item = await _graphClient
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
        /// <summary>
        /// Lists the document libraries (drives) of a site. Read-only.
        /// </summary>
        public async Task ListDrivesAsync(string siteId)
        {
            try
            {
                Console.WriteLine($"Listing document libraries of site: {siteId}");
                var drives = await _graphClient.Sites[siteId].Drives.GetAsync();

                Console.WriteLine($"Found {drives?.Value?.Count ?? 0} libraries:");
                Console.WriteLine(new string('-', 80));
                foreach (var drive in drives?.Value ?? new System.Collections.Generic.List<Drive>())
                {
                    Console.WriteLine($"Library: {drive.Name}");
                    Console.WriteLine($"  ID: {drive.Id}");
                    Console.WriteLine($"  Type: {drive.DriveType}");
                    Console.WriteLine(new string('-', 80));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing libraries: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Entry point: builds the client with the common module and runs a read-only
        /// demo listing the libraries of the book-test test site.
        /// </summary>
        public static async Task Main(string[] args)
        {
            try
            {
                var graphClient = SharePointGraphClientFactory.CreateFromSecret();
                var docOps = new DocumentOperations(graphClient);

                var hostname = Environment.GetEnvironmentVariable("SHAREPOINT_HOSTNAME") ?? "olddogsoft1.sharepoint.com";
                var sitePath = Environment.GetEnvironmentVariable("SHAREPOINT_SITE_PATH") ?? "book-test";

                // Resolve the site by path to get its ID (the path form does not work for .Drives).
                var site = await graphClient.Sites[$"{hostname}:/sites/{sitePath}"].GetAsync();
                if (site == null)
                {
                    Console.WriteLine("Test site not found.");
                    return;
                }
                var siteId = site.Id ?? throw new InvalidOperationException("El sitio no tiene ID.");
                await docOps.ListDrivesAsync(siteId);

                Console.WriteLine("\nDocument operations completed successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
}
