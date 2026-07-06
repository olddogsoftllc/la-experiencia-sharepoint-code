/**
 * Chapter 2: Authentication
 * Client Credentials flow
 */

require('dotenv').config();
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;

async function main() {
    console.log('=== Chapter 2: Authentication ===\n');

    if (!tenantId || !clientId || !clientSecret) {
        console.error('Error: Missing environment variables');
        process.exit(1);
    }

    try {
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const authProvider = new TokenCredentialAuthenticationProvider(credential, {
            scopes: ['https://graph.microsoft.com/.default']
        });

        const client = Client.initWithMiddleware({ authProvider });

        const org = await client.api('/organization').get();

        console.log('✓ Connection successful');
        console.log(`  Tenant: ${org.value[0].displayName}`);
        console.log(`  ID: ${org.value[0].id}`);
    } catch (error) {
        console.error('✗ Error:', error.message);
    }
}

main();
