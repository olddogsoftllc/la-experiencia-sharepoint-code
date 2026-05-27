/*
 * PermissionOperations.cs
 * Chapter 05: Permissions
 *
 * SharePoint Permission Operations Example
 * Demonstrates managing sharing links and permissions
 */

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Drives.Item.Items.Item.CreateLink;

namespace SharePointExperience.Chapter05
{
    /// <summary>
    /// Example class demonstrating SharePoint permission operations
    /// </summary>
    public class PermissionOperations
    {
        private readonly GraphServiceClient _graphClient;

        /// <summary>
        /// Initializes a new instance of PermissionOperations
        /// </summary>
        public PermissionOperations()
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
        /// Creates an anonymous sharing link for a file or folder
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <param name="driveId">The drive ID</param>
        /// <param name="itemId">The item ID</param>
        /// <param name="linkType">The link type (view or edit)</param>
        /// <returns>The created permission with sharing link</returns>
        public async Task<Permission> CreateAnonymousLinkAsync(
            string siteId,
            string driveId,
            string itemId,
            string linkType = "view")
        {
            try
            {
                Console.WriteLine($"Creating anonymous sharing link for item: {itemId}");

                var requestBody = new CreateLinkPostRequestBody
                {
                    Type = linkType,
                    Scope = "anonymous"
                };

                var permission = await _graphClient
                    .Sites[siteId]
                    .Drives[driveId]
                    .Items[itemId]
                    .CreateLink
                    .PostAsync(requestBody);

                Console.WriteLine("Sharing link created successfully:");
                Console.WriteLine($"  Link: {permission?.Link?.WebUrl}");
                Console.WriteLine($"  Type: {permission?.Link?.Type}");
                Console.WriteLine($"  Scope: {permission?.Link?.Scope}");

                return permission;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating sharing link: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates an organization sharing link (for users within the organization)
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <param name="driveId">The drive ID</param>
        /// <param name="itemId">The item ID</param>
        /// <param name="linkType">The link type (view or edit)</param>
        /// <returns>The created permission with sharing link</returns>
        public async Task<Permission> CreateOrganizationLinkAsync(
            string siteId,
            string driveId,
            string itemId,
            string linkType = "view")
        {
            try
            {
                Console.WriteLine($"Creating organization sharing link for item: {itemId}");

                var requestBody = new CreateLinkPostRequestBody
                {
                    Type = linkType,
                    Scope = "organization"
                };

                var permission = await _graphClient
                    .Sites[siteId]
                    .Drives[driveId]
                    .Items[itemId]
                    .CreateLink
                    .PostAsync(requestBody);

                Console.WriteLine("Organization sharing link created successfully:");
                Console.WriteLine($"  Link: {permission?.Link?.WebUrl}");
                Console.WriteLine($"  Type: {permission?.Link?.Type}");
                Console.WriteLine($"  Scope: {permission?.Link?.Scope}");

                return permission;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating organization link: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Grants access to a specific user
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <param name="driveId">The drive ID</param>
        /// <param name="itemId">The item ID</param>
        /// <param name="userEmail">The email address of the user</param>
        /// <param name="role">The role (write or read)</param>
        /// <returns>The created permission</returns>
        public async Task<Permission> GrantAccessToUserAsync(
            string siteId,
            string driveId,
            string itemId,
            string userEmail,
            string role = "write")
        {
            try
            {
                Console.WriteLine($"Granting {role} access to {userEmail} for item: {itemId}");

                var requestBody = new InvitePostRequestBody
                {
                    Recipients = new List<DriveRecipient>
                    {
                        new DriveRecipient
                        {
                            Email = userEmail
                        }
                    },
                    Roles = new List<string> { role },
                    SendNotification = true,
                    Message = "You have been granted access to this document."
                };

                var permissions = await _graphClient
                    .Sites[siteId]
                    .Drives[driveId]
                    .Items[itemId]
                    .Invite
                    .PostAsync(requestBody);

                Console.WriteLine("Access granted successfully:");
                if (permissions?.Value > 0)
                {
                    var permission = permissions.Value[0];
                    Console.WriteLine($"  Permission ID: {permission.Id}");
                    Console.WriteLine($"  Roles: {string.Join(", ", permission.Roles ?? new List<string>())}");
                }

                return permissions?.Value?[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error granting access: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Lists all permissions for an item
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <param name="driveId">The drive ID</param>
        /// <param name="itemId">The item ID</param>
        public async Task ListPermissionsAsync(string siteId, string driveId, string itemId)
        {
            try
            {
                Console.WriteLine($"Listing permissions for item: {itemId}");

                var permissions = await _graphClient
                    .Sites[siteId]
                    .Drives[driveId]
                    .Items[itemId]
                    .Permissions
                    .GetAsync();

                Console.WriteLine($"Found {permissions?.Value?.Count ?? 0} permissions:");
                Console.WriteLine(new string('-', 80));

                foreach (var permission in permissions?.Value ?? new List<Permission>())
                {
                    Console.WriteLine($"Permission ID: {permission.Id}");
                    Console.WriteLine($"  Roles: {string.Join(", ", permission.Roles ?? new List<string>())}");

                    if (permission.Link != null)
                    {
                        Console.WriteLine($"  Link Type: {permission.Link.Type}");
                        Console.WriteLine($"  Link Scope: {permission.Link.Scope}");
                        Console.WriteLine($"  Web URL: {permission.Link.WebUrl}");
                    }

                    if (permission.GrantedTo?.User != null)
                    {
                        Console.WriteLine($"  Granted To: {permission.GrantedTo.User.DisplayName} ({permission.GrantedTo.User.Email})");
                    }

                    if (permission.GrantedToIdentities != null)
                    {
                        foreach (var identity in permission.GrantedToIdentities)
                        {
                            Console.WriteLine($"  Granted To: {identity.User?.DisplayName} ({identity.User?.Email})");
                        }
                    }

                    Console.WriteLine(new string('-', 80));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing permissions: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Deletes a permission
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <param name="driveId">The drive ID</param>
        /// <param name="itemId">The item ID</param>
        /// <param name="permissionId">The permission ID to delete</param>
        public async Task DeletePermissionAsync(string siteId, string driveId, string itemId, string permissionId)
        {
            try
            {
                Console.WriteLine($"Deleting permission: {permissionId}");

                await _graphClient
                    .Sites[siteId]
                    .Drives[driveId]
                    .Items[itemId]
                    .Permissions[permissionId]
                    .DeleteAsync();

                Console.WriteLine("Permission deleted successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error deleting permission: {ex.Message}");
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
                var permOps = new PermissionOperations();

                Console.WriteLine("Permission operations class initialized successfully!");
                Console.WriteLine("Use the methods to manage sharing links and permissions.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
}
