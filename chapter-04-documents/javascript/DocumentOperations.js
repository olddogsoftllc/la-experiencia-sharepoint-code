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
const path = require('path');
const { getAccessToken } = require('la-experiencia-sharepoint-code/graphAuth');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * SharePoint Document Operations class. El access token se inyecta por constructor (DI).
 */
class DocumentOperations {
    /** @param {string} accessToken Bearer token for Microsoft Graph (inyectado). */
    constructor(accessToken) {
        if (!accessToken) throw new Error('Se requiere un access token para DocumentOperations.');
        this.accessToken = accessToken;
    }

    /** @returns {Promise<{Authorization: string, 'Content-Type': string}>} */
    async getHeaders(contentType = 'application/json') {
        return {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': contentType
        };
    }

    /** Lists the document libraries (drives) of a site. Solo lectura. */
    async listDrives(siteId) {
        try {
            console.log(`Listing document libraries of site: ${siteId}`);
            const headers = await this.getHeaders();
            const response = await axios.get(`${GRAPH_BASE}/sites/${siteId}/drives`, { headers });
            const drives = response.data.value;
            console.log(`Found ${drives.length} libraries:`);
            console.log('-'.repeat(80));
            drives.forEach((drive) => {
                console.log(`Library: ${drive.name}`);
                console.log(`  ID: ${drive.id}`);
                console.log(`  Type: ${drive.driveType}`);
                console.log('-'.repeat(80));
            });
            return drives;
        } catch (error) {
            console.error('Error listing libraries:', error.response?.data?.error?.message || error.message);
            throw error;
        }
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
        const token = await getAccessToken();
        const docOps = new DocumentOperations(token);

        const hostname = process.env.SHAREPOINT_HOSTNAME || 'olddogsoft1.sharepoint.com';
        const sitePath = process.env.SHAREPOINT_SITE_PATH || 'book-test';

        // Resolver el sitio por path para obtener su ID.
        const siteUrl = `${GRAPH_BASE}/sites/${encodeURIComponent(hostname)}:/sites/${encodeURIComponent(sitePath)}`;
        const siteResp = await axios.get(siteUrl, { headers: await docOps.getHeaders() });

        await docOps.listDrives(siteResp.data.id);

        console.log('\nDocument operations completed successfully!');
    } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
    }
}

module.exports = { DocumentOperations };

if (require.main === module) {
    main();
}
