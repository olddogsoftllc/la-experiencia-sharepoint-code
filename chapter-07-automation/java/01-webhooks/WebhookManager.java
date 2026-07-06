// WebhookManager.java
// Example: Manage webhook subscriptions with Microsoft Graph SDK for Java (v6).
// The GraphServiceClient is injected via constructor (DI) from the common module.

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

    /** Lists existing subscriptions (read-only). */
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
        System.out.println("🔔 Creating subscription for Drive: " + driveId);

        Subscription subscription = new Subscription();
        subscription.setResource("/drives/" + driveId + "/root");
        subscription.setChangeType("created,updated,deleted");
        subscription.setNotificationUrl(notificationUrl);
        subscription.setClientState(UUID.randomUUID().toString());
        subscription.setExpirationDateTime(OffsetDateTime.now().plusDays(2));

        Subscription created = graphClient.subscriptions().post(subscription);

        System.out.println("   ✅ Subscription: " + created.getId());
        System.out.println("   Expires: " + created.getExpirationDateTime());

        return created;
    }

    public void renewSubscription(String subscriptionId) {
        Subscription patch = new Subscription();
        patch.setExpirationDateTime(OffsetDateTime.now().plusDays(2));

        Subscription renewed = graphClient.subscriptions().bySubscriptionId(subscriptionId).patch(patch);
        System.out.println("🔄 Renewed. Expires: " + renewed.getExpirationDateTime());
    }

    public void deleteSubscription(String subscriptionId) {
        graphClient.subscriptions().bySubscriptionId(subscriptionId).delete();
        System.out.println("🗑️ Subscription " + subscriptionId + " deleted");
    }

    public static void main(String[] args) {
        try {
            // create() auto-detects a certificate (if CERTIFICATE_PATH is present) or a client secret.
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
