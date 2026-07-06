/*
 * SiteOperations.cs
 * Chapter 03: Sites
 *
 * SharePoint Site Operations Example
 * Demonstrates listing, creating, and retrieving SharePoint sites.
 *
 * Uses the shared auth module (common/SharePointGraphAuth): the Graph client is
 * injected via constructor (DI), not built inside the class.
 */

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using LaExperiencia.SharePoint.Common;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace LaExperiencia.SharePoint.Chapter03.Sites
{
    /// <summary>
    /// Example class demonstrating SharePoint site operations.
    /// </summary>
    public class SiteOperations
    {
        private readonly GraphServiceClient _graphClient;

        /// <summary>
        /// Creates an instance with an injected Graph client (DI).
        /// </summary>
        public SiteOperations(GraphServiceClient graphClient)
        {
            _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        }

        /// <summary>Lists all sites in the organization.</summary>
        public async Task ListSitesAsync()
        {
            try
            {
                Console.WriteLine("Fetching all sites...");

                var sites = await _graphClient.Sites.GetAsync();

                Console.WriteLine($"\nFound {sites?.Value?.Count ?? 0} sites:");
                Console.WriteLine(new string('-', 80));

                foreach (var site in sites?.Value ?? new List<Site>())
                {
                    Console.WriteLine($"Display Name: {site.DisplayName}");
                    Console.WriteLine($"  ID: {site.Id}");
                    Console.WriteLine($"  Web URL: {site.WebUrl}");
                    Console.WriteLine(new string('-', 80));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing sites: {ex.Message}");
                throw;
            }
        }

        /// <summary>Gets a specific site by hostname and site path (e.g. contoso.sharepoint.com / book-test).</summary>
        public async Task<Site?> GetSiteAsync(string hostname, string sitePath)
        {
            try
            {
                Console.WriteLine($"Fetching site: {hostname}/sites/{sitePath}");

                var site = await _graphClient.Sites[$"{hostname}:/sites/{sitePath}"].GetAsync();

                if (site != null)
                {
                    Console.WriteLine("Site found:");
                    Console.WriteLine($"  Display Name: {site.DisplayName}");
                    Console.WriteLine($"  ID: {site.Id}");
                    Console.WriteLine($"  Web URL: {site.WebUrl}");
                    Console.WriteLine($"  Description: {site.Description}");
                }

                return site;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting site: {ex.Message}");
                throw;
            }
        }

        /// <summary>Gets site by its unique identifier.</summary>
        public async Task<Site?> GetSiteByIdAsync(string siteId)
        {
            try
            {
                Console.WriteLine($"Fetching site by ID: {siteId}");

                var site = await _graphClient.Sites[siteId].GetAsync();

                if (site != null)
                {
                    Console.WriteLine("Site found:");
                    Console.WriteLine($"  Display Name: {site.DisplayName}");
                    Console.WriteLine($"  Web URL: {site.WebUrl}");
                }

                return site;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting site by ID: {ex.Message}");
                throw;
            }
        }

        /// <summary>Searches for sites by keyword.</summary>
        public async Task SearchSitesAsync(string keyword)
        {
            try
            {
                Console.WriteLine($"Searching for sites with keyword: '{keyword}'");

                var sites = await _graphClient.Sites.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Search = keyword;
                });

                Console.WriteLine($"\nFound {sites?.Value?.Count ?? 0} matching sites:");
                Console.WriteLine(new string('-', 80));

                foreach (var site in sites?.Value ?? new List<Site>())
                {
                    Console.WriteLine($"Display Name: {site.DisplayName}");
                    Console.WriteLine($"  Web URL: {site.WebUrl}");
                    Console.WriteLine($"  Description: {site.Description}");
                    Console.WriteLine(new string('-', 80));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error searching sites: {ex.Message}");
                throw;
            }
        }

        /// <summary>Gets the root site of the organization.</summary>
        public async Task<Site?> GetRootSiteAsync()
        {
            try
            {
                Console.WriteLine("Fetching root site...");

                var site = await _graphClient.Sites["root"].GetAsync();

                if (site != null)
                {
                    Console.WriteLine("Root site:");
                    Console.WriteLine($"  Display Name: {site.DisplayName}");
                    Console.WriteLine($"  ID: {site.Id}");
                    Console.WriteLine($"  Web URL: {site.WebUrl}");
                }

                return site;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting root site: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Entry point: builds the Graph client with the common module (client secret) and
        /// runs a read-only demo against the book-test test site.
        /// </summary>
        public static async Task Main(string[] args)
        {
            try
            {
                var graphClient = SharePointGraphClientFactory.CreateFromSecret();
                var siteOps = new SiteOperations(graphClient);

                // Test site (overridable via SHAREPOINT_HOSTNAME / SHAREPOINT_SITE_PATH env vars).
                var hostname = Environment.GetEnvironmentVariable("SHAREPOINT_HOSTNAME") ?? "olddogsoft1.sharepoint.com";
                var sitePath = Environment.GetEnvironmentVariable("SHAREPOINT_SITE_PATH") ?? "book-test";

                await siteOps.GetSiteAsync(hostname, sitePath);
                await siteOps.ListSitesAsync();

                Console.WriteLine("\nSite operations completed successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
}