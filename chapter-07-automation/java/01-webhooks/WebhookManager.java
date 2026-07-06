// WebhookManager.java
// Ejemplo: Gestionar suscripciones webhook con Microsoft Graph SDK para Java (v6).
// El GraphServiceClient se inyecta por constructor (DI) desde el módulo común.

package com.sharepointexperiencia.automation;

import com.laexperiencia.sharepoint.common.GraphServiceClientFactory;
import com.microsoft.graph.models.Subscription;
import com.microsoft.graph.models.SubscriptionCollectionResponse;
import com.microsoft.graph.serviceclient.GraphServiceClient;
import java.time.OffsetDateTime;
import java.util.UUID;

public class WebhookManager {
    private final GraphServiceClient graphClient;
    private final String notificationUrl;

    public WebhookManager(GraphServiceClient graphClient, String notificationUrl) {
        this.graphClient = graphClient;
        this.notificationUrl = notificationUrl;
    }

    /** Lista las suscripciones existentes (solo lectura). */
    public void listSubscriptions() {
        try {
            System.out.println("📋 Listando suscripciones webhook...");
            SubscriptionCollectionResponse resp = graphClient.subscriptions().get();
            int count = resp != null && resp.getValue() != null ? resp.getValue().size() : 0;
            System.out.println("   Suscripciones encontradas: " + count);
            if (resp != null && resp.getValue() != null) {
                for (Subscription s : resp.getValue()) {
                    System.out.println("   - ID: " + s.getId() + " | Resource: " + s.getResource());
                }
            }
        } catch (Exception e) {
            System.err.println("❌ Error listando suscripciones: " + e.getMessage());
        }
    }

    public Subscription createDriveSubscription(String driveId) {
        System.out.println("🔔 Creando suscripción para Drive: " + driveId);

        Subscription subscription = new Subscription();
        subscription.setResource("/drives/" + driveId + "/root");
        subscription.setChangeType("created,updated,deleted");
        subscription.setNotificationUrl(notificationUrl);
        subscription.setClientState(UUID.randomUUID().toString());
        subscription.setExpirationDateTime(OffsetDateTime.now().plusDays(2));

        Subscription created = graphClient.subscriptions().post(subscription);

        System.out.println("   ✅ Suscripción: " + created.getId());
        System.out.println("   Expira: " + created.getExpirationDateTime());

        return created;
    }

    public void renewSubscription(String subscriptionId) {
        Subscription patch = new Subscription();
        patch.setExpirationDateTime(OffsetDateTime.now().plusDays(2));

        Subscription renewed = graphClient.subscriptions().bySubscriptionId(subscriptionId).patch(patch);
        System.out.println("🔄 Renovada. Expira: " + renewed.getExpirationDateTime());
    }

    public void deleteSubscription(String subscriptionId) {
        graphClient.subscriptions().bySubscriptionId(subscriptionId).delete();
        System.out.println("🗑️ Suscripción " + subscriptionId + " eliminada");
    }

    public static void main(String[] args) {
        try {
            // create() auto-detecta certificado (si CERTIFICATE_PATH está presente) o client secret.
            var client = GraphServiceClientFactory.create();
            var manager = new WebhookManager(client, "https://example.com/webhook");
            manager.listSubscriptions();
            System.out.println("\nWebhook manager example completed.");
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }
}
