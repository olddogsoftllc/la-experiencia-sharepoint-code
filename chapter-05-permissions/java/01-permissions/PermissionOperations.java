// PermissionOperations.java
// Chapter 05: Permissions
//
// Example of permission and sharing operations in SharePoint with Microsoft Graph
// SDK for Java (v6). The GraphServiceClient is injected via constructor (DI) from
// the common module: demonstration of DRY + dependency injection.

package com.sharepointexperience.chapter05;

import com.laexperiencia.sharepoint.common.GraphServiceClientFactory;
import com.microsoft.graph.drives.item.items.item.createlink.CreateLinkPostRequestBody;
import com.microsoft.graph.drives.item.items.item.invite.InvitePostRequestBody;
import com.microsoft.graph.models.DriveRecipient;
import com.microsoft.graph.models.Permission;
import com.microsoft.graph.models.PermissionCollectionResponse;
import com.microsoft.graph.models.SharingLink;
import com.microsoft.graph.models.Site;
import com.microsoft.graph.serviceclient.GraphServiceClient;

import java.util.ArrayList;
import java.util.List;

/**
 * Example demonstrating management of permissions and sharing links in SharePoint:
 * creation of anonymous/organization links, granting access to users, listing and
 * deleting permissions at the item and site level.
 */
public class PermissionOperations {

    private final GraphServiceClient graphClient;

    /** Creates an instance with an injected Graph client (DI). */
    public PermissionOperations(GraphServiceClient graphClient) {
        if (graphClient == null) {
            throw new IllegalArgumentException("graphClient is required");
        }
        this.graphClient = graphClient;
    }

    /**
     * Creates an anonymous sharing link (anyone with the link) for a file or folder.
     */
    public Permission createAnonymousLink(String driveId, String itemId, String linkType) {
        String type = linkType == null || linkType.isBlank() ? "view" : linkType;
        System.out.println("Creating anonymous link (" + type + ") for item: " + itemId);

        CreateLinkPostRequestBody body = new CreateLinkPostRequestBody();
        body.setType(type);
        body.setScope("anonymous");

        Permission permission = graphClient
                .drives().byDriveId(driveId)
                .items().byDriveItemId(itemId)
                .createLink().post(body);

        printLink(permission);
        return permission;
    }

    /**
     * Creates a sharing link for the whole organization.
     */
    public Permission createOrganizationLink(String driveId, String itemId, String linkType) {
        String type = linkType == null || linkType.isBlank() ? "view" : linkType;
        System.out.println("Creating organization link (" + type + ") for item: " + itemId);

        CreateLinkPostRequestBody body = new CreateLinkPostRequestBody();
        body.setType(type);
        body.setScope("organization");

        Permission permission = graphClient
                .drives().byDriveId(driveId)
                .items().byDriveItemId(itemId)
                .createLink().post(body);

        printLink(permission);
        return permission;
    }

    /**
     * Grants access to a specific user on an item.
     */
    public Permission grantAccessToUser(String driveId, String itemId, String userEmail, String role) {
        String effectiveRole = role == null || role.isBlank() ? "write" : role;
        System.out.println("Concediendo acceso '" + effectiveRole + "' a " + userEmail + " sobre el item: " + itemId);

        DriveRecipient recipient = new DriveRecipient();
        recipient.setEmail(userEmail);

        List<DriveRecipient> recipients = new ArrayList<>();
        recipients.add(recipient);

        List<String> roles = new ArrayList<>();
        roles.add(effectiveRole);

        InvitePostRequestBody body = new InvitePostRequestBody();
        body.setRecipients(recipients);
        body.setRoles(roles);
        body.setSendInvitation(true);
        body.setMessage("You have been granted access to this document.");

        var permissions = graphClient
                .drives().byDriveId(driveId)
                .items().byDriveItemId(itemId)
                .invite().post(body);

        Permission granted = (permissions != null && permissions.getValue() != null && !permissions.getValue().isEmpty())
                ? permissions.getValue().get(0)
                : null;

        if (granted != null) {
            System.out.println("Acceso concedido:");
            System.out.println("  Permission ID: " + granted.getId());
            System.out.println("  Roles: " + granted.getRoles());
        }
        return granted;
    }

    /**
     * Lists the permissions of an item (drive).
     */
    public void listPermissions(String driveId, String itemId) {
        System.out.println("Listando permisos del item: " + itemId);

        PermissionCollectionResponse response = graphClient
                .drives().byDriveId(driveId)
                .items().byDriveItemId(itemId)
                .permissions().get();

        List<Permission> permissions = response != null && response.getValue() != null
                ? response.getValue()
                : new ArrayList<>();

        System.out.println("Encontrados " + permissions.size() + " permisos:");
        System.out.println("-".repeat(80));
        for (Permission p : permissions) {
            printPermission(p);
            System.out.println("-".repeat(80));
        }
    }

    /**
     * Lists site-level permissions (read-only).
     */
    public void listSitePermissions(String siteId) {
        System.out.println("Listando permisos del sitio: " + siteId);

        PermissionCollectionResponse response = graphClient
                .sites().bySiteId(siteId)
                .permissions().get();

        List<Permission> permissions = response != null && response.getValue() != null
                ? response.getValue()
                : new ArrayList<>();

        System.out.println("Encontrados " + permissions.size() + " permisos de sitio:");
        System.out.println("-".repeat(80));
        for (Permission p : permissions) {
            printPermission(p);
            System.out.println("-".repeat(80));
        }
    }

    /**
     * Deletes a specific permission from an item.
     */
    public void deletePermission(String driveId, String itemId, String permissionId) {
        System.out.println("Eliminando permiso: " + permissionId);
        graphClient
                .drives().byDriveId(driveId)
                .items().byDriveItemId(itemId)
                .permissions().byPermissionId(permissionId)
                .delete();
        System.out.println("Permiso eliminado correctamente");
    }

    // ============== Helpers ==============

    private void printLink(Permission permission) {
        if (permission == null) {
            System.out.println("  (sin respuesta)");
            return;
        }
        SharingLink link = permission.getLink();
        if (link != null) {
            System.out.println("  Enlace creado:");
            System.out.println("    Web URL: " + link.getWebUrl());
            System.out.println("    Type: " + link.getType());
            System.out.println("    Scope: " + link.getScope());
        }
    }

    private void printPermission(Permission p) {
        System.out.println("Permission ID: " + p.getId());
        System.out.println("  Roles: " + p.getRoles());
        if (p.getLink() != null) {
            System.out.println("  Link Type: " + p.getLink().getType());
            System.out.println("  Link Scope: " + p.getLink().getScope());
            System.out.println("  Web URL: " + p.getLink().getWebUrl());
        }
        if (p.getGrantedToIdentitiesV2() != null) {
            for (var identity : p.getGrantedToIdentitiesV2()) {
                System.out.println("  Granted To (v2): "
                        + (identity.getApplication() != null ? identity.getApplication().getDisplayName()
                            : (identity.getUser() != null ? identity.getUser().getDisplayName() : "N/A")));
            }
        } else if (p.getGrantedTo() != null && p.getGrantedTo().getUser() != null) {
            System.out.println("  Granted To: " + p.getGrantedTo().getUser().getDisplayName());
        }
    }

    /**
     * Entry point: builds the client with the common module and runs a read-only
     * demo listing the permissions of the book-test test site (resolved by path,
     * overridable with SHAREPOINT_HOSTNAME / SHAREPOINT_SITE_PATH).
     */
    public static void main(String[] args) {
        try {
            // create() auto-detects certificate (if CERTIFICATE_PATH is present) or client secret.
            GraphServiceClient client = GraphServiceClientFactory.create();
            PermissionOperations permOps = new PermissionOperations(client);

            String hostname = envOr("SHAREPOINT_HOSTNAME", "olddogsoft1.sharepoint.com");
            String sitePath = envOr("SHAREPOINT_SITE_PATH", "book-test");

            // Resolver el sitio por path para obtener su ID.
            Site site = client.sites().bySiteId(hostname + ":/sites/" + sitePath).get();
            String siteId = site.getId();
            System.out.println("Site: " + site.getDisplayName() + " (" + siteId + ")\n");

            // List the drives of the site and take the first one; list permissions of the root item.
            // (Listing /sites/{id}/permissions requires admin permissions outside Sites.Selected;
            //  permissions of a drive item are covered by the Sites.Selected grant.)
            var drives = client.sites().bySiteId(siteId).drives().get();
            var firstDrive = (drives != null && drives.getValue() != null && !drives.getValue().isEmpty())
                    ? drives.getValue().get(0) : null;
            if (firstDrive == null) {
                System.out.println("No se encontraron bibliotecas en el sitio de pruebas.");
                return;
            }
            System.out.println("Biblioteca: " + firstDrive.getName() + " (" + firstDrive.getId() + ")\n");
            try {
                permOps.listPermissions(firstDrive.getId(), "root");
            } catch (Exception ex) {
                // El SDK de Graph para Java tiene un quirk de parseo de fechas (OffsetDateTime
                // sin zona) en algunas respuestas; lo reportamos sin abortar el ejemplo.
                System.err.println("⚠️ No se pudieron listar los permisos: " + ex.getMessage());
            }

            System.out.println("\nPermission operations completed successfully!");
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }

    private static String envOr(String name, String defaultValue) {
        String v = System.getenv(name);
        return (v == null || v.isBlank()) ? defaultValue : v;
    }
}