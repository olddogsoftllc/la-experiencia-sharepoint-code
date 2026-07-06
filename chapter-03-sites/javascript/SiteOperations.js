/**
 * SiteOperations.js
 * Chapter 03: Sites
 *
 * SharePoint Site Operations Example
 * Demonstrates listing, creating, and retrieving SharePoint sites.
 *
 * Usa el módulo de auth compartido (common/graphAuth): el access token se inyecta
 * por constructor (DI), no se obtiene dentro de la clase.
 */

const axios = require('axios');
const { getAccessToken } = require('la-experiencia-sharepoint-code/graphAuth');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

class SiteOperations {
    /** @param {string} accessToken Bearer token for Microsoft Graph (inyectado). */
    constructor(accessToken) {
        if (!accessToken) throw new Error('Se requiere un access token para SiteOperations.');
        this.accessToken = accessToken;
    }

    /** @returns {Promise<{Authorization: string, 'Content-Type': string}>} */
    async getHeaders() {
        return {
            Authorization: `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json',
        };
    }

    /** Lists all sites in the organization. */
    async listSites() {
        try {
            console.log('Fetching all sites...\n');
            const headers = await this.getHeaders();
            const response = await axios.get(`${GRAPH_BASE}/sites`, { headers });

            const sites = response.data.value;
            console.log(`Found ${sites.length} sites:`);
            console.log('-'.repeat(80));
            sites.forEach((site) => {
                console.log(`Display Name: ${site.displayName ?? site.name ?? 'N/A'}`);
                console.log(`  ID: ${site.id}`);
                console.log(`  Web URL: ${site.webUrl}`);
                console.log('-'.repeat(80));
            });
            return sites;
        } catch (error) {
            console.error('Error listing sites:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /** Gets a specific site by hostname and site path (e.g. contoso.sharepoint.com / book-test). */
    async getSite(hostname, sitePath) {
        try {
            console.log(`Fetching site: ${hostname}/sites/${sitePath}`);
            const headers = await this.getHeaders();
            // Formato Graph path-based: /sites/{hostname}:/sites/{path}
            const url = `${GRAPH_BASE}/sites/${encodeURIComponent(hostname)}:/sites/${encodeURIComponent(sitePath)}`;
            const response = await axios.get(url, { headers });
            const site = response.data;

            console.log('\nSite found:');
            console.log(`  Display Name: ${site.displayName ?? site.name ?? 'N/A'}`);
            console.log(`  ID: ${site.id}`);
            console.log(`  Web URL: ${site.webUrl}`);
            console.log(`  Description: ${site.description ?? 'N/A'}`);
            return site;
        } catch (error) {
            console.error('Error getting site:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /** Gets site by its unique identifier. */
    async getSiteById(siteId) {
        try {
            console.log(`Fetching site by ID: ${siteId}`);
            const headers = await this.getHeaders();
            const response = await axios.get(`${GRAPH_BASE}/sites/${siteId}`, { headers });
            const site = response.data;
            console.log('\nSite found:');
            console.log(`  Display Name: ${site.displayName ?? site.name ?? 'N/A'}`);
            console.log(`  Web URL: ${site.webUrl}`);
            return site;
        } catch (error) {
            console.error('Error getting site by ID:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /** Gets the root site of the organization. */
    async getRootSite() {
        try {
            console.log('Fetching root site...');
            const headers = await this.getHeaders();
            const response = await axios.get(`${GRAPH_BASE}/sites/root`, { headers });
            const site = response.data;
            console.log('\nRoot site:');
            console.log(`  Display Name: ${site.displayName ?? site.name ?? 'N/A'}`);
            console.log(`  ID: ${site.id}`);
            console.log(`  Web URL: ${site.webUrl}`);
            return site;
        } catch (error) {
            console.error('Error getting root site:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /** Searches for sites by keyword. */
    async searchSites(keyword) {
        try {
            console.log(`Searching for sites with keyword: '${keyword}'`);
            const headers = await this.getHeaders();
            const response = await axios.get(
                `${GRAPH_BASE}/sites?search=${encodeURIComponent(keyword)}`,
                { headers }
            );
            const sites = response.data.value;
            console.log(`\nFound ${sites.length} matching sites:`);
            console.log('-'.repeat(80));
            sites.forEach((site) => {
                console.log(`Display Name: ${site.displayName ?? site.name ?? 'N/A'}`);
                console.log(`  Web URL: ${site.webUrl}`);
                console.log(`  Description: ${site.description ?? 'N/A'}`);
                console.log('-'.repeat(80));
            });
            return sites;
        } catch (error) {
            console.error('Error searching sites:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }
}

/** Construye el token con el módulo común y ejecuta una demo de solo lectura contra book-test. */
async function main() {
    try {
        console.log('=== SharePoint Site Operations Example ===\n');
        const token = await getAccessToken();
        const siteOps = new SiteOperations(token);

        const hostname = process.env.SHAREPOINT_HOSTNAME || 'olddogsoft1.sharepoint.com';
        const sitePath = process.env.SHAREPOINT_SITE_PATH || 'book-test';

        await siteOps.getSite(hostname, sitePath);
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