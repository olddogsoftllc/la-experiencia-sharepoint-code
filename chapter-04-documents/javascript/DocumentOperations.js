/**
 * DocumentOperations.js
 * Chapter 04: Documents
 *
 * SharePoint Document Operations Example
 * Demonstrates upload, download, and search operations for documents
 *
 * Required environment variables:
 * - TENANT_ID
 * - CLIENT_ID
 * - CLIENT_SECRET
 */

const axios = require('axios');
const fs = require('fs');
const qs = require('querystring');
const path = require('path');

/**
 * SharePoint Document Operations class
 */
class DocumentOperations {
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
    async getHeaders(contentType = 'application/json') {
        const token = await this.getAccessToken();
        return {
            'Authorization': `Bearer ${token}`,
            'Content-Type': contentType
        };
    }

    /**
     * Uploads a file to SharePoint
     */
    async uploadFile(siteId, driveId, filePath, destinationFileName) {
        try {
            console.log(`Uploading file: ${filePath}`);

            if (!fs.existsSync(filePath)) {
                throw new Error(`File not found: ${filePath}`);
            }

            const fileContent = fs.readFileSync(filePath);
            const encodedName = encodeURIComponent(destinationFileName);

            const headers = await this.getHeaders('application/octet-stream');
            const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/root:/${encodedName}:/content`;

            const response = await axios.put(url, fileContent, { headers });

            console.log('File uploaded successfully:');
            console.log(`  ID: ${response.data.id}`);
            console.log(`  Name: ${response.data.name}`);
            console.log(`  Size: ${response.data.size} bytes`);
            console.log(`  Web URL: ${response.data.webUrl}`);

            return response.data;
        } catch (error) {
            console.error('Error uploading file:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Downloads a file from SharePoint
     */
    async downloadFile(siteId, driveId, itemId, downloadPath) {
        try {
            console.log(`Downloading file to: ${downloadPath}`);

            const headers = await this.getHeaders();
            const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/content`;

            const response = await axios.get(url, {
                headers,
                responseType: 'arraybuffer'
            });

            fs.writeFileSync(downloadPath, response.data);

            console.log(`File downloaded successfully to: ${downloadPath}`);
        } catch (error) {
            console.error('Error downloading file:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Searches for files across SharePoint
     */
    async searchFiles(query) {
        try {
            console.log(`Searching for files: '${query}'`);

            const headers = await this.getHeaders();
            const searchBody = {
                requests: [{
                    entityTypes: ['driveItem'],
                    query: {
                        queryString: query
                    }
                }]
            };

            const response = await axios.post(
                'https://graph.microsoft.com/v1.0/search/query',
                searchBody,
                { headers }
            );

            console.log('Search results:');
            console.log('-'.repeat(80));

            const results = response.data.value || [];
            results.forEach(result => {
                (result.hitsContainers || []).forEach(container => {
                    (container.hits || []).forEach(hit => {
                        const resource = hit.resource;
                        console.log(`Name: ${resource.name}`);
                        console.log(`  Web URL: ${resource.webUrl}`);
                        console.log(`  Size: ${resource.size} bytes`);
                        console.log('-'.repeat(80));
                    });
                });
            });

            return response.data;
        } catch (error) {
            console.error('Error searching files:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Lists files in a specific folder
     */
    async listFilesInFolder(siteId, driveId, folderPath = '') {
        try {
            console.log(`Listing files in folder: ${folderPath || 'root'}`);

            const headers = await this.getHeaders();

            let url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/children`;
            if (folderPath) {
                const encodedPath = encodeURIComponent(folderPath);
                url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/root:/${encodedPath}:/children`;
            }

            const response = await axios.get(url, { headers });
            const items = response.data.value || [];

            console.log(`Found ${items.length} items:`);
            console.log('-'.repeat(80));

            items.forEach(item => {
                const itemType = item.folder ? 'Folder' : 'File';
                console.log(`${itemType}: ${item.name}`);
                console.log(`  ID: ${item.id}`);
                console.log(`  Size: ${item.size} bytes`);
                console.log(`  Last Modified: ${item.lastModifiedDateTime}`);
                console.log('-'.repeat(80));
            });

            return items;
        } catch (error) {
            console.error('Error listing files:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }

    /**
     * Gets file metadata
     */
    async getFileMetadata(siteId, driveId, itemId) {
        try {
            console.log(`Fetching metadata for item: ${itemId}`);

            const headers = await this.getHeaders();
            const response = await axios.get(
                `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}`,
                { headers }
            );

            const item = response.data;
            console.log('File metadata:');
            console.log(`  Name: ${item.name}`);
            console.log(`  ID: ${item.id}`);
            console.log(`  Size: ${item.size} bytes`);
            console.log(`  Created: ${item.createdDateTime}`);
            console.log(`  Modified: ${item.lastModifiedDateTime}`);
            console.log(`  Web URL: ${item.webUrl}`);

            return item;
        } catch (error) {
            console.error('Error getting file metadata:', error.response?.data?.error?.message || error.message);
            throw error;
        }
    }
}

/**
 * Main execution function
 */
async function main() {
    try {
        console.log('=== SharePoint Document Operations Example ===\n');

        const docOps = new DocumentOperations();

        console.log('Document operations class initialized successfully!');
        console.log('Use the methods to perform upload, download, and search operations.');
    } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
    }
}

module.exports = { DocumentOperations };

if (require.main === module) {
    main();
}
