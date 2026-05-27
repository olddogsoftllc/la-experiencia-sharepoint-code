/**
 * SiteOperations.js
 * Chapter 03: Sites
 *
 * SharePoint Site Operations Example
 * Demonstrates listing, creating, and retrieving SharePoint sites
 *
 * Required environment variables:
 * - TENANT_ID
 * - CLIENT_ID
 * - CLIENT_SECRET
 */

const axios = require('axios');
const qs = require('querystring');

/**
 * SharePoint Site Operations class
 */
class SiteOperations {
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
     * Lists all sites in the organization
     */
    async listSites() {
        try {
            console.log('Fetching all sites...\n');

            const headers = await this.getHeaders();
            const response = await axios.get('https://graph.microsoft.com/v1.0/sites', { headers });

            const sites = response.data.value;
            console.log(`Found ${sites.length} sites:`);
            console.log('-'.repeat(80));

            sites.forEach(site => {
                console.log(`Name: ${site.name}`);
                console.log(`  ID: ${site.id}`);
                console.log(`  Web URL: ${site.webUrl}`);
                console.log(`  Display Name: ${site.displayName}`);
                console.log('-'.repeat(80));
            });

            return sites;
        } catch (error) {
            console.error('Error listing sites:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Gets a specific site by hostname and site path
     */
    async getSite(hostname, sitePath) {
        try {
            console.log(`Fetching site: ${hostname}/sites/${sitePath}`);

            const headers = await this.getHeaders();
            const encodedPath = encodeURIComponent(`sites/${sitePath}`);
            const url = `https://graph.microsoft.com/v1.0/sites/${hostname}:${encodedPath}`;

            const response = await axios.get(url, { headers });
            const site = response.data;

            console.log('\nSite found:');
            console.log(`  Name: ${site.name}`);
            console.log(`  ID: ${site.id}`);
            console.log(`  Web URL: ${site.webUrl}`);
            console.log(`  Description: ${site.description}`);

            return site;
        } catch (error) {
            console.error('Error getting site:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Gets site by its unique identifier
     */
    async getSiteById(siteId) {
        try {
            console.log(`Fetching site by ID: ${siteId}`);

            const headers = await this.getHeaders();
            const response = await axios.get(`https://graph.microsoft.com/v1.0/sites/${siteId}`, { headers });
            const site = response.data;

            console.log('\nSite found:');
            console.log(`  Name: ${site.name}`);
            console.log(`  Web URL: ${site.webUrl}`);

            return site;
        } catch (error) {
            console.error('Error getting site by ID:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Gets the root site of the organization
     */
    async getRootSite() {
        try {
            console.log('Fetching root site...');

            const headers = await this.getHeaders();
            const response = await axios.get('https://graph.microsoft.com/v1.0/sites/root', { headers });
            const site = response.data;

            console.log('\nRoot site:');
            console.log(`  Name: ${site.name}`);
            console.log(`  ID: ${site.id}`);
            console.log(`  Web URL: ${site.webUrl}`);

            return site;
        } catch (error) {
            console.error('Error getting root site:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Searches for sites by keyword
     */
    async searchSites(keyword) {
        try {
            console.log(`Searching for sites with keyword: '${keyword}'`);

            const headers = await this.getHeaders();
            const response = await axios.get(
                `https://graph.microsoft.com/v1.0/sites?search=${encodeURIComponent(keyword)}`,
                { headers }
            );

            const sites = response.data.value;
            console.log(`\nFound ${sites.length} matching sites:`);
            console.log('-'.repeat(80));

            sites.forEach(site => {
                console.log(`Name: ${site.name}`);
                console.log(`  Web URL: ${site.webUrl}`);
                console.log(`  Description: ${site.description}`);
                console.log('-'.repeat(80));
            });

            return sites;
        } catch (error) {
            console.error('Error searching sites:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }
}

/**
 * Main execution function
 */
async function main() {
    try {
        console.log('=== SharePoint Site Operations Example ===\n');

        const siteOps = new SiteOperations();

        // Get root site
        await siteOps.getRootSite();

        // List all sites
        await siteOps.listSites();

        console.log('\nSite operations completed successfully!');
    } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
    }
}

module.exports = { SiteOperations };

if (require.main === module) {
    main();
}
