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
using LaExperiencia.SharePoint.Common;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Drives.Item.Items.Item.CreateLink;
using Microsoft.Graph.Drives.Item.Items.Item.Invite;

namespace LaExperiencia.SharePoint.Chapter05.Permissions
{
    /// <summary>
    /// Example class demonstrating SharePoint permission operations.
    /// The Graph client is injected via constructor (DI) from the common module.
    /// </summary>
    public class PermissionOperations
    {
        private readonly GraphServiceClient _graphClient;

        /// <summary>Creates an instance with an injected Graph client (DI).</summary>
        public PermissionOperations(GraphServiceClient graphClient)
        {
            _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        }

        /// <summary>
        /// Creates an anonymous sharing link for a file or folder
        /// </summary>
        /// <param name="siteId">The site ID</param>
        /// <param name="driveId">The drive ID</param>
        /// <param name="itemId">The item ID</param>
        /// <param name="linkType">The link type (view or edit)</param>
        /// <returns>The created permission with sharing link</returns>
        public async Task<Permission?> CreateAnonymousLinkAsync(
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
        public async Task<Permission?> CreateOrganizationLinkAsync(
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
        public async Task<Permission?> GrantAccessToUserAsync(
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
                    SendInvitation = true,
                    Message = "You have been granted access to this document."
                };

                var permissions = await _graphClient
                    .Drives[driveId]
                    .Items[itemId]
                    .Invite
                    .PostAsInvitePostResponseAsync(requestBody);

                Console.WriteLine("Access granted successfully:");
                if (permissions?.Value is { Count: > 0 } perms)
                {
                    var permission = perms[0];
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
                        Console.WriteLine($"  Granted To: {permission.GrantedTo.User.DisplayName} ({permission.GrantedTo.User.Id})");
                    }

                    if (permission.GrantedToIdentities != null)
                    {
                        foreach (var identity in permission.GrantedToIdentities)
                        {
                            Console.WriteLine($"  Granted To: {identity.User?.DisplayName} ({identity.User?.Id})");
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
        /// Lists the permissions of a site (read-only).
        /// </summary>
        public async Task ListSitePermissionsAsync(string siteId)
        {
            try
            {
                Console.WriteLine($"Listing permissions of site: {siteId}");
                var permissions = await _graphClient.Sites[siteId].Permissions.GetAsync();

                Console.WriteLine($"Found {permissions?.Value?.Count ?? 0} site permissions:");
                Console.WriteLine(new string('-', 80));
                foreach (var permission in permissions?.Value ?? new System.Collections.Generic.List<Permission>())
                {
                    Console.WriteLine($"Permission ID: {permission.Id}");
                    Console.WriteLine($"  Roles: {string.Join(", ", permission.Roles ?? new System.Collections.Generic.List<string>())}");
                    if (permission.GrantedToIdentities != null)
                    {
                        foreach (var identity in permission.GrantedToIdentities)
                        {
                            Console.WriteLine($"  Granted To: {identity.User?.DisplayName}");
                        }
                    }
                    Console.WriteLine(new string('-', 80));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing site permissions: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Entry point: builds the client with the common module and runs a read-only
        /// demo listing the permissions of the book-test test site.
        /// </summary>
        public static async Task Main(string[] args)
        {
            try
            {
                var graphClient = SharePointGraphClientFactory.CreateFromSecret();
                var permOps = new PermissionOperations(graphClient);

                var hostname = Environment.GetEnvironmentVariable("SHAREPOINT_HOSTNAME") ?? "olddogsoft1.sharepoint.com";
                var sitePath = Environment.GetEnvironmentVariable("SHAREPOINT_SITE_PATH") ?? "book-test";

                // Resolve the site by path to get its ID.
                var site = await graphClient.Sites[$"{hostname}:/sites/{sitePath}"].GetAsync();
                if (site == null)
                {
                    Console.WriteLine("Test site not found.");
                    return;
                }
                var siteId = site.Id ?? throw new InvalidOperationException("El sitio no tiene ID.");

                // List permissions for an ITEM in the test site drive (within the
                // Sites.Selected grant). Listing /sites/{id}/permissions requires admin
                // permissions (Sites.Manage/FullControl), outside the least-privilege scope.
                var drives = await graphClient.Sites[siteId].Drives.GetAsync();
                var firstDrive = drives?.Value?.FirstOrDefault();
                if (firstDrive == null)
                {
                    Console.WriteLine("No se encontraron bibliotecas en el sitio de pruebas.");
                    return;
                }
                var driveId = firstDrive.Id ?? throw new InvalidOperationException("La biblioteca no tiene ID.");
                await permOps.ListPermissionsAsync(siteId, driveId, "root");

                Console.WriteLine("\nPermission operations completed successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
}
