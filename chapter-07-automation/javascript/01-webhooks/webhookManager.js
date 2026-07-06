// webhookManager.js
// Ejemplo: Gestionar suscripciones webhook con Microsoft Graph.
// El cliente de Graph se inyecta por constructor (DI) desde el módulo común.

const { getGraphClient } = require('la-experiencia-sharepoint-code/graphAuth');

class WebhookManager {
    constructor(graphClient, notificationUrl) {
        this.graphClient = graphClient;
        this.notificationUrl = notificationUrl;
    }

    /** Lista las suscripciones existentes (solo lectura). */
    async listSubscriptions() {
        try {
            console.log('📋 Listando suscripciones webhook...');
            const response = await this.graphClient.api('/subscriptions').get();
            const subs = response.value || [];
            console.log(`   Suscripciones encontradas: ${subs.length}`);
            subs.forEach((s) => console.log(`   - ID: ${s.id} | Resource: ${s.resource}`));
        } catch (error) {
            console.error(`❌ Error listando suscripciones: ${error.message}`);
        }
    }

    async createDriveSubscription(driveId) {
        console.log(`🔔 Creando suscripción para Drive: ${driveId}`);

        const subscription = {
            resource: `/drives/${driveId}/root`,
            changeType: 'created,updated,deleted',
            notificationUrl: this.notificationUrl,
            clientState: require('crypto').randomUUID(),
            expirationDateTime: new Date(Date.now() + 2 * 24 * 60 * 60 * 1000).toISOString()
        };

        const response = await this.graphClient
            .api('/subscriptions')
            .post(subscription);

        console.log(`   ✅ Suscripción: ${response.id}`);
        return response;
    }

    async renewSubscription(subscriptionId) {
        const patch = {
            expirationDateTime: new Date(Date.now() + 2 * 24 * 60 * 60 * 1000).toISOString()
        };

        return await this.graphClient
            .api(`/subscriptions/${subscriptionId}`)
            .patch(patch);
    }

    async deleteSubscription(subscriptionId) {
        await this.graphClient
            .api(`/subscriptions/${subscriptionId}`)
            .delete();
        console.log(`🗑️ Suscripción ${subscriptionId} eliminada`);
    }
}

async function main() {
    const client = await getGraphClient();
    const manager = new WebhookManager(client, 'https://example.com/webhook');
    await manager.listSubscriptions();
    console.log('\nWebhook manager example completed.');
}

module.exports = { WebhookManager };

if (require.main === module) {
    main().catch((e) => { console.error('Error:', e.message); process.exit(1); });
}
