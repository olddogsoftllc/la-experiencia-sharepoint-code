/**
 * deltaQueries.js
 * Example of Delta Queries with Microsoft Graph for incremental sync.
 * Reference: Chapter 7 - Automation and Flows
 *
 * Usage:
 *   node deltaQueries.js
 */

const axios = require('axios');
const { getAccessToken } = require('la-experiencia-sharepoint-code/graphAuth');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

class DeltaQueryManager {
    constructor(accessToken) {
        this.accessToken = accessToken;
        this.baseUrl = 'https://graph.microsoft.com/v1.0';
        this.deltaTokens = new Map();
    }

    _getHeaders() {
        return {
            'Authorization': `Bearer ${this.accessToken}`,
            'Accept': 'application/json'
        };
    }

    /**
     * Realiza consulta delta inicial y almacena token.
     */
    async getInitialDelta(driveId, folderId = 'root') {
        console.log(`🔄 Ejecutando consulta delta inicial en drive: ${driveId}`);

        const url = `${this.baseUrl}/drives/${driveId}/items/${folderId}/delta`;
        const response = await axios.get(url, { headers: this._getHeaders() });

        const items = response.data.value || [];
        console.log(`   Items obtenidos: ${items.length}`);

        // Guardar token delta
        const deltaLink = response.data['@odata.deltaLink'] || '';
        const token = this._extractDeltaToken(deltaLink);
        if (token) {
            const key = `${driveId}_${folderId}`;
            this.deltaTokens.set(key, token);
            console.log(`   Token delta guardado: ${token.substring(0, 50)}...`);
        }

        return items;
    }

    /**
     * Sincroniza cambios usando token delta almacenado.
     */
    async getDeltaChanges(driveId, folderId = 'root') {
        const key = `${driveId}_${folderId}`;

        if (!this.deltaTokens.has(key)) {
            console.log('⚠️  No existe token delta. Ejecutando consulta inicial...');
            return await this.getInitialDelta(driveId, folderId);
        }

        const token = this.deltaTokens.get(key);
        console.log('🔄 Sincronizando cambios con token delta...');

        const url = `${this.baseUrl}/drives/${driveId}/items/${folderId}/delta`;
        const response = await axios.get(url, {
            headers: this._getHeaders(),
            params: { token }
        });

        const items = response.data.value || [];
        console.log(`   Cambios detectados: ${items.length}`);

        items.forEach(item => {
            if (item.deleted) {
                console.log(`   🗑️  Deleted: ${item.name || 'unknown'}`);
            } else {
                console.log(`   ✏️  Modificado: ${item.name || 'unknown'}`);
            }
        });

        // Update delta token
        const newDeltaLink = response.data['@odata.deltaLink'] || '';
        const newToken = this._extractDeltaToken(newDeltaLink);
        if (newToken) {
            this.deltaTokens.set(key, newToken);
            console.log('   Token delta actualizado');
        }

        // Verificar paginacion
        const nextLink = response.data['@odata.nextLink'];
        if (nextLink) {
            console.log(`   📄 Hay mas paginas: ${nextLink}`);
        }

        return items;
    }

    /**
     * Consulta delta para items de una lista de SharePoint.
     */
    async getListDelta(siteId, listId) {
        console.log(`🔄 Consulta delta para lista: ${listId}`);

        const url = `${this.baseUrl}/sites/${siteId}/lists/${listId}/items/delta`;
        const response = await axios.get(url, { headers: this._getHeaders() });

        const items = response.data.value || [];

        const deltaLink = response.data['@odata.deltaLink'] || '';
        const token = this._extractDeltaToken(deltaLink);
        if (token) {
            this.deltaTokens.set(`list_${siteId}_${listId}`, token);
        }

        return items;
    }

    /**
     * Extrae el token delta de la URL.
     */
    _extractDeltaToken(deltaLink) {
        if (!deltaLink) return null;

        const url = new URL(deltaLink);
        return url.searchParams.get('token') || url.searchParams.get('deltaToken');
    }

    /**
     * Shows the stored delta tokens.
     */
    showStoredTokens() {
        console.log('\n📋 Tokens Delta Almacenados:');
        this.deltaTokens.forEach((token, key) => {
            console.log(`   ${key}: ${token.substring(0, 50)}...`);
        });
    }
}

async function main() {
    console.log('╔══════════════════════════════════════════════════════════════╗');
    console.log('║     EJEMPLO: DELTA QUERIES CON MICROSOFT GRAPH             ║');
    console.log('║     Sincronizacion Incremental de Documentos                 ║');
    console.log('╚══════════════════════════════════════════════════════════════╝\n');

    const accessToken = await getAccessToken();
    const manager = new DeltaQueryManager(accessToken);

    const hostname = process.env.SHAREPOINT_HOSTNAME || 'olddogsoft1.sharepoint.com';
    const sitePath = process.env.SHAREPOINT_SITE_PATH || 'book-test';

    try {
        // Resolver el sitio por path y su primera biblioteca.
        const siteUrl = `${GRAPH_BASE}/sites/${encodeURIComponent(hostname)}:/sites/${encodeURIComponent(sitePath)}`;
        const siteResp = await axios.get(siteUrl, { headers: manager._getHeaders() });
        const siteId = siteResp.data.id;
        const drivesResp = await axios.get(`${GRAPH_BASE}/sites/${siteId}/drives`, { headers: manager._getHeaders() });
        const firstDrive = drivesResp.data.value[0];
        console.log(`Biblioteca objetivo: ${firstDrive.name}\n`);

        // Consulta inicial
        console.log('--- Consulta Delta Inicial ---');
        await manager.getInitialDelta(firstDrive.id);

        // Sincronizacion de cambios
        console.log('\n--- Sincronizacion de Cambios ---');
        await manager.getDeltaChanges(firstDrive.id);

        manager.showStoredTokens();
        console.log('\n✅ Example completed');

    } catch (error) {
        console.error(`❌ Error: ${error.message}`);
        if (error.response) {
            console.error(`   Status: ${error.response.status}`);
            console.error(`   Data: ${JSON.stringify(error.response.data, null, 2)}`);
        }
    }
}

module.exports = { DeltaQueryManager };

if (require.main === module) {
    main();
}
