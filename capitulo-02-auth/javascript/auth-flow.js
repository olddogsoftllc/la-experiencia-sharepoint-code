/**
 * Capítulo 2: Autenticación
 * Flujo de Client Credentials
 */

require('dotenv').config();
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const tenantId = process.env.SP_TENANT_ID;
const clientId = process.env.SP_CLIENT_ID;
const clientSecret = process.env.SP_CLIENT_SECRET;

async function main() {
    console.log('=== Capítulo 2: Autenticación ===\n');

    if (!tenantId || !clientId || !clientSecret) {
        console.error('Error: Faltan variables de entorno');
        process.exit(1);
    }

    try {
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const authProvider = new TokenCredentialAuthenticationProvider(credential, {
            scopes: ['https://graph.microsoft.com/.default']
        });

        const client = Client.initWithMiddleware({ authProvider });

        const org = await client.api('/organization').get();

        console.log('✓ Conexión exitosa');
        console.log(`  Tenant: ${org.value[0].displayName}`);
        console.log(`  ID: ${org.value[0].id}`);
    } catch (error) {
        console.error('✗ Error:', error.message);
    }
}

main();
