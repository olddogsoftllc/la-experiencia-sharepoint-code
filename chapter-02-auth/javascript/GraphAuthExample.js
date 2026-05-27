/**
 * GraphAuthExample.js
 * Chapter 02: Authentication
 *
 * Microsoft Graph Client Credentials Authentication Example
 * Demonstrates authenticating to Microsoft Graph using client credentials flow
 *
 * Required environment variables:
 * - TENANT_ID
 * - CLIENT_ID
 * - CLIENT_SECRET
 */

const axios = require('axios');
const qs = require('querystring');

/**
 * Authentication configuration class
 */
class GraphAuthConfig {
    constructor() {
        this.tenantId = process.env.TENANT_ID;
        this.clientId = process.env.CLIENT_ID;
        this.clientSecret = process.env.CLIENT_SECRET;

        this.validateConfig();
    }

    /**
     * Validates that all required environment variables are present
     */
    validateConfig() {
        const required = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET'];
        const missing = required.filter(key => !process.env[key]);

        if (missing.length > 0) {
            throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
        }
    }
}

/**
 * Microsoft Graph Authentication handler
 */
class GraphAuthExample {
    constructor() {
        this.config = new GraphAuthConfig();
        this.tokenEndpoint = `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`;
        this.scope = 'https://graph.microsoft.com/.default';
    }

    /**
     * Acquires an access token using client credentials flow
     * @returns {Promise<string>} Access token
     */
    async getAccessToken() {
        try {
            const requestBody = {
                client_id: this.config.clientId,
                client_secret: this.config.clientSecret,
                scope: this.scope,
                grant_type: 'client_credentials'
            };

            const response = await axios.post(
                this.tokenEndpoint,
                qs.stringify(requestBody),
                {
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded'
                    }
                }
            );

            console.log('Successfully authenticated to Microsoft Graph');
            return response.data.access_token;
        } catch (error) {
            console.error('Authentication failed:', error.response?.data?.error_description || error.message);
            throw new Error('Failed to acquire access token');
        }
    }

    /**
     * Creates authenticated API client headers
     * @returns {Promise<Object>} Headers object with authorization
     */
    async getAuthenticatedHeaders() {
        const token = await this.getAccessToken();
        return {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        };
    }

    /**
     * Tests the connection by retrieving organization details
     */
    async testConnection() {
        try {
            const headers = await this.getAuthenticatedHeaders();
            const response = await axios.get('https://graph.microsoft.com/v1.0/organization', { headers });

            const org = response.data.value[0];
            console.log(`Connected to tenant: ${org.displayName}`);
            return response.data;
        } catch (error) {
            console.error('Connection test failed:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }
}

/**
 * Main execution function
 */
async function main() {
    try {
        console.log('=== Microsoft Graph Authentication Example ===');

        const auth = new GraphAuthExample();
        await auth.testConnection();

        console.log('\nAuthentication completed successfully!');
    } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
    }
}

// Export for module usage
module.exports = { GraphAuthExample, GraphAuthConfig };

// Run if executed directly
if (require.main === module) {
    main();
}
