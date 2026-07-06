// PermissionOperations.java
// Chapter 05: Permissions
//
// Ejemplo de operaciones de permisos y sharing en SharePoint con Microsoft Graph
// SDK para Java (v6). El GraphServiceClient se inyecta por constructor (DI) desde
// el módulo común: demostración de DRY + inyección de dependencias.

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
 * Ejemplo que demuestra gestión de permisos y enlaces de uso compartido en
 * SharePoint: creación de enlaces anónimos/organización, concesión de acceso a
 * usuarios, listado y borrado de permisos a nivel de item y de sitio.
 */
public class PermissionOperations {

    private final GraphServiceClient graphClient;

    /** Crea una instancia con un cliente de Graph inyectado (DI). */
    public PermissionOperations(GraphServiceClient graphClient) {
        if (graphClient == null) {
            throw new IllegalArgumentException("graphClient is required");
        }
        this.graphClient = graphClient;
    }

    /**
     * Crea un enlace de uso compartido anónimo (cualquiera con el enlace) para un
     * archivo o carpeta.
     */
    public Permission createAnonymousLink(String driveId, String itemId, String linkType) {
        String type = linkType == null || linkType.isBlank() ? "view" : linkType;
        System.out.println("Creando enlace anónimo (" + type + ") para el item: " + itemId);

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
     * Crea un enlace de uso compartido para toda la organización.
     */
    public Permission createOrganizationLink(String driveId, String itemId, String linkType) {
        String type = linkType == null || linkType.isBlank() ? "view" : linkType;
        System.out.println("Creando enlace de organización (" + type + ") para el item: " + itemId);

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
     * Concede acceso a un usuario concreto sobre un item.
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
     * Lista los permisos de un item (drive).
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
     * Lista los permisos a nivel de sitio (solo lectura).
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
     * Elimina un permiso concreto de un item.
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

    // ============== Auxiliares ==============

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
     * Entry point: construye el cliente con el módulo común y ejecuta una demo de
     * solo lectura listando los permisos del sitio de pruebas book-test (resuelto
     * por path, sobreescribible con SHAREPOINT_HOSTNAME / SHAREPOINT_SITE_PATH).
     */
    public static void main(String[] args) {
        try {
            // create() auto-detecta certificado (si CERTIFICATE_PATH está presente) o client secret.
            GraphServiceClient client = GraphServiceClientFactory.create();
            PermissionOperations permOps = new PermissionOperations(client);

            String hostname = envOr("SHAREPOINT_HOSTNAME", "olddogsoft1.sharepoint.com");
            String sitePath = envOr("SHAREPOINT_SITE_PATH", "book-test");

            // Resolver el sitio por path para obtener su ID.
            Site site = client.sites().bySiteId(hostname + ":/sites/" + sitePath).get();
            String siteId = site.getId();
            System.out.println("Sitio: " + site.getDisplayName() + " (" + siteId + ")\n");

            // Listar los drives del sitio y tomar el primero; listar permisos del item root.
            // (Listar /sites/{id}/permissions requiere permisos de admin fuera de Sites.Selected;
            //  los permisos de un item del drive sí están cubiertos por el grant Sites.Selected.)
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