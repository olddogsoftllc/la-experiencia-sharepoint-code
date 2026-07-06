#!/usr/bin/env python3
# webhook_manager.py
# Example: Manage webhook subscriptions with Microsoft Graph

import uuid
from datetime import datetime, timedelta
from msgraph import GraphServiceClient


class WebhookManager:
    def __init__(self, graph_client: GraphServiceClient, notification_url: str):
        self.graph_client = graph_client
        self.notification_url = notification_url

    async def create_drive_subscription(self, drive_id: str):
        print(f"🔔 Creating subscription for Drive: {drive_id}")

        subscription = {
            "resource": f"/drives/{drive_id}/root",
            "changeType": "created,updated,deleted",
            "notificationUrl": self.notification_url,
            "clientState": str(uuid.uuid4()),
            "expirationDateTime": (datetime.utcnow() + timedelta(days=2)).isoformat()
        }

        result = await self.graph_client.subscriptions.post(subscription)
        print(f"   ✅ Subscription: {result.id}")
        return result

    async def renew_subscription(self, subscription_id: str):
        patch = {
            "expirationDateTime": (datetime.utcnow() + timedelta(days=2)).isoformat()
        }

        result = await self.graph_client.subscriptions.by_subscription_id(
            subscription_id
        ).patch(patch)
        print(f"🔄 Renewed. Expires: {result.expiration_date_time}")
        return result

    async def delete_subscription(self, subscription_id: str):
        await self.graph_client.subscriptions.by_subscription_id(
            subscription_id
        ).delete()
        print(f"🗑️ Subscription {subscription_id} deleted")


# Usage
# manager = WebhookManager(graph_client, "https://tu-endpoint.com/webhook")
# await manager.create_drive_subscription("drive-id")
