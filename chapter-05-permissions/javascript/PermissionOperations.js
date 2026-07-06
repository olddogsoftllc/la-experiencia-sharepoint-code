/**
 * PermissionOperations.js
 * Chapter 05: Permissions
 *
 * SharePoint Permission Operations Example
 * Demonstrates managing sharing links and permissions
 *
 * Required environment variables:
 * - TENANT_ID
 * - CLIENT_ID
 * - CLIENT_SECRET
 */

const axios = require('axios');
const { getAccessToken } = require('la-experiencia-sharepoint-code/graphAuth');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * SharePoint Permission Operations class. El access token se inyecta por constructor (DI).
 */
class PermissionOperations {
    /** @param {string} accessToken Bearer token for Microsoft Graph (inyectado). */
    constructor(accessToken) {
        if (!accessToken) throw new Error('Se requiere un access token para PermissionOperations.');
        this.accessToken = accessToken;
    }

    /** @returns {Promise<{Authorization: string, 'Content-Type': string}>} */
    async getHeaders() {
        return {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json',
        };
    }

    /** Lists the permissions of a site. Solo lectura. */
    async listSitePermissions(siteId) {
        try {
            console.log(`Listing permissions of site: ${siteId}`);
            const headers = await this.getHeaders();
            const response = await axios.get(`${GRAPH_BASE}/sites/${siteId}/permissions`, { headers });
            const permissions = response.data.value;
            console.log(`Found ${permissions.length} site permissions:`);
            console.log('-'.repeat(80));
            permissions.forEach((permission) => {
                console.log(`Permission ID: ${permission.id}`);
                console.log(`  Roles: ${(permission.roles || []).join(', ')}`);
                (permission.grantedToIdentities || []).forEach((identity) => {
                    console.log(`  Granted To: ${identity.user?.displayName}`);
                });
                console.log('-'.repeat(80));
            });
            return permissions;
        } catch (error) {
            console.error('Error listing site permissions:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Creates an anonymous sharing link
     */
    async createAnonymousLink(siteId, driveId, itemId, linkType = 'view') {
        try {
            console.log(`Creating anonymous sharing link for item: ${itemId}`);

            const headers = await this.getHeaders();
            const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/createLink`;

            const body = {
                type: linkType,
                scope: 'anonymous'
            };

            const response = await axios.post(url, body, { headers });

            console.log('Sharing link created successfully:');
            console.log(`  Link: ${response.data.link.webUrl}`);
            console.log(`  Type: ${response.data.link.type}`);
            console.log(`  Scope: ${response.data.link.scope}`);

            return response.data;
        } catch (error) {
            console.error('Error creating sharing link:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Creates an organization sharing link
     */
    async createOrganizationLink(siteId, driveId, itemId, linkType = 'view') {
        try {
            console.log(`Creating organization sharing link for item: ${itemId}`);

            const headers = await this.getHeaders();
            const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/createLink`;

            const body = {
                type: linkType,
                scope: 'organization'
            };

            const response = await axios.post(url, body, { headers });

            console.log('Organization sharing link created successfully:');
            console.log(`  Link: ${response.data.link.webUrl}`);
            console.log(`  Type: ${response.data.link.type}`);
            console.log(`  Scope: ${response.data.link.scope}`);

            return response.data;
        } catch (error) {
            console.error('Error creating organization link:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Grants access to a specific user
     */
    async grantAccessToUser(siteId, driveId, itemId, userEmail, role = 'write') {
        try {
            console.log(`Granting ${role} access to ${userEmail} for item: ${itemId}`);

            const headers = await this.getHeaders();
            const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/invite`;

            const body = {
                recipients: [{ email: userEmail }],
                roles: [role],
                sendNotification: true,
                message: 'You have been granted access to this document.'
            };

            const response = await axios.post(url, body, { headers });

            console.log('Access granted successfully:');
            const permission = response.data.value[0];
            console.log(`  Permission ID: ${permission.id}`);
            console.log(`  Roles: ${permission.roles.join(', ')}`);

            return response.data;
        } catch (error) {
            console.error('Error granting access:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Lists all permissions for an item
     */
    async listPermissions(siteId, driveId, itemId) {
        try {
            console.log(`Listing permissions for item: ${itemId}`);

            const headers = await this.getHeaders();
            const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/permissions`;

            const response = await axios.get(url, { headers });
            const permissions = response.data.value || [];

            console.log(`Found ${permissions.length} permissions:`);
            console.log('-'.repeat(80));

            permissions.forEach(permission => {
                console.log(`Permission ID: ${permission.id}`);
                console.log(`  Roles: ${(permission.roles || []).join(', ')}`);

                if (permission.link) {
                    console.log(`  Link Type: ${permission.link.type}`);
                    console.log(`  Link Scope: ${permission.link.scope}`);
                    console.log(`  Web URL: ${permission.link.webUrl}`);
                }

                if (permission.grantedTo?.user) {
                    console.log(`  Granted To: ${permission.grantedTo.user.displayName} (${permission.grantedTo.user.email})`);
                }

                if (permission.grantedToIdentities) {
                    permission.grantedToIdentities.forEach(identity => {
                        console.log(`  Granted To: ${identity.user?.displayName} (${identity.user?.email})`);
                    });
                }

                console.log('-'.repeat(80));
            });

            return permissions;
        } catch (error) {
            console.error('Error listing permissions:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Deletes a permission
     */
    async deletePermission(siteId, driveId, itemId, permissionId) {
        try {
            console.log(`Deleting permission: ${permissionId}`);

            const headers = await this.getHeaders();
            const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/permissions/${permissionId}`;

            await axios.delete(url, { headers });

            console.log('Permission deleted successfully');
        } catch (error) {
            console.error('Error deleting permission:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }
}

/**
 * Main execution function
 */
async function main() {
    try {
        console.log('=== SharePoint Permission Operations Example ===\n');
        const token = await getAccessToken();
        const permOps = new PermissionOperations(token);

        const hostname = process.env.SHAREPOINT_HOSTNAME || 'olddogsoft1.sharepoint.com';
        const sitePath = process.env.SHAREPOINT_SITE_PATH || 'book-test';

        // Resolver el sitio por path para obtener su ID.
        const siteUrl = `${GRAPH_BASE}/sites/${encodeURIComponent(hostname)}:/sites/${encodeURIComponent(sitePath)}`;
        const siteResp = await axios.get(siteUrl, { headers: await permOps.getHeaders() });
        const siteId = siteResp.data.id;

        // List the drives of the site and take the first one; list permissions of the root item.
        // (Listing /sites/{id}/permissions requires admin; permissions of a drive item
        //  are covered by the Sites.Selected grant.)
        const drivesResp = await axios.get(`${GRAPH_BASE}/sites/${siteId}/drives`, { headers: await permOps.getHeaders() });
        const firstDrive = drivesResp.data.value[0];
        await permOps.listPermissions(siteId, firstDrive.id, 'root');

        console.log('\nPermission operations completed successfully!');
    } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
    }
}

module.exports = { PermissionOperations };

if (require.main === module) {
    main();
}
