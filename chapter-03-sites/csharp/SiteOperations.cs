/*
 * SiteOperations.cs
 * Chapter 03: Sites
 *
 * SharePoint Site Operations Example
 * Demonstrates listing, creating, and retrieving SharePoint sites
 */

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace SharePointExperience.Chapter03
{
    /// <summary>
    /// Example class demonstrating SharePoint site operations
    /// </summary>
    public class SiteOperations
    {
        private readonly GraphServiceClient _graphClient;

        /// <summary>
        /// Initializes a new instance of SiteOperations
        /// </summary>
        public SiteOperations()
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
        /// Lists all sites in the organization
        /// </summary>
        /// <returns>List of sites</returns>
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
                    Console.WriteLine($"Name: {site.Name}");
                    Console.WriteLine($"  ID: {site.Id}");
                    Console.WriteLine($"  Web URL: {site.WebUrl}");
                    Console.WriteLine($"  Display Name: {site.DisplayName}");
                    Console.WriteLine(new string('-', 80));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing sites: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Gets a specific site by hostname and site path
        /// </summary>
        /// <param name="hostname">The hostname of the site (e.g., contoso.sharepoint.com)</param>
        /// <param name="sitePath">The site path (e.g., sites/marketing)</param>
        /// <returns>The site object</returns>
        public async Task<Site> GetSiteAsync(string hostname, string sitePath)
        {
            try
            {
                Console.WriteLine($"Fetching site: {hostname}/{sitePath}");

                var site = await _graphClient.Sites[$"{hostname}:/sites/{sitePath}"].GetAsync();

                if (site != null)
                {
                    Console.WriteLine($"Site found:");
                    Console.WriteLine($"  Name: {site.Name}");
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

        /// <summary>
        /// Gets site by its unique identifier
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <returns>The site object</returns>
        public async Task<Site> GetSiteByIdAsync(string siteId)
        {
            try
            {
                Console.WriteLine($"Fetching site by ID: {siteId}");

                var site = await _graphClient.Sites[siteId].GetAsync();

                if (site != null)
                {
                    Console.WriteLine($"Site found:");
                    Console.WriteLine($"  Name: {site.Name}");
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

        /// <summary>
        /// Searches for sites by keyword
        /// </summary>
        /// <param name="keyword">Search keyword</param>
        public async Task SearchSitesAsync(string keyword)
        {
            try
            {
                Console.WriteLine($"Searching for sites with keyword: '{keyword}'");

                // Note: Site search uses the search query parameter
                var sites = await _graphClient.Sites.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Search = keyword;
                });

                Console.WriteLine($"\nFound {sites?.Value?.Count ?? 0} matching sites:");
                Console.WriteLine(new string('-', 80));

                foreach (var site in sites?.Value ?? new List<Site>())
                {
                    Console.WriteLine($"Name: {site.Name}");
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

        /// <summary>
        /// Gets the root site of the organization
        /// </summary>
        /// <returns>The root site</returns>
        public async Task<Site> GetRootSiteAsync()
        {
            try
            {
                Console.WriteLine("Fetching root site...");

                var site = await _graphClient.Sites["root"].GetAsync();

                if (site != null)
                {
                    Console.WriteLine($"Root site:");
                    Console.WriteLine($"  Name: {site.Name}");
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
        /// Entry point for the example
        /// </summary>
        public static async Task Main(string[] args)
        {
            try
            {
                var siteOps = new SiteOperations();

                // Get root site
                await siteOps.GetRootSiteAsync();

                // List all sites
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
