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
const qs = require('querystring');

/**
 * SharePoint Permission Operations class
 */
class PermissionOperations {
    constructor() {
        this.tenantId = process.env.TENANT_ID;
        this.clientId = process.env.CLIENT_ID;
        this.clientSecret = process.env.CLIENT_SECRET;
        this.accessToken = null;

        this.validateConfig();
    }

    validateConfig() {
        const required = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET'];
        const missing = required.filter(key => !process.env[key]);

        if (missing.length > 0) {
            throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
        }
    }

    /**
     * Gets access token for Microsoft Graph
     */
    async getAccessToken() {
        if (this.accessToken) return this.accessToken;

        const tokenEndpoint = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
        const requestBody = {
            client_id: this.clientId,
            client_secret: this.clientSecret,
            scope: 'https://graph.microsoft.com/.default',
            grant_type: 'client_credentials'
        };

        const response = await axios.post(tokenEndpoint, qs.stringify(requestBody), {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });

        this.accessToken = response.data.access_token;
        return this.accessToken;
    }

    /**
     * Gets authenticated headers
     */
    async getHeaders() {
        const token = await this.getAccessToken();
        return {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        };
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

        const permOps = new PermissionOperations();

        console.log('Permission operations class initialized successfully!');
        console.log('Use the methods to manage sharing links and permissions.');
    } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
    }
}

module.exports = { PermissionOperations };

if (require.main === module) {
    main();
}
