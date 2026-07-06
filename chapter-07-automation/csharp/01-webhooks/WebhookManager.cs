// WebhookManager.cs
// Example: Create and manage webhook subscriptions in Microsoft Graph

using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace AutomationExamples;

public class WebhookManager
{
    private readonly GraphServiceClient _graphClient;
    private readonly string _notificationUrl;

    public WebhookManager(GraphServiceClient graphClient, string notificationUrl)
    {
        _graphClient = graphClient;
        _notificationUrl = notificationUrl;
    }

    public async Task<Subscription> CreateDriveSubscriptionAsync(string driveId)
    {
        Console.WriteLine($"🔔 Creating subscription for Drive: {driveId}");

        var subscription = new Subscription
        {
            Resource = $"/drives/{driveId}/root",
            ChangeType = "created,updated,deleted",
            NotificationUrl = _notificationUrl,
            ClientState = Guid.NewGuid().ToString(),
            ExpirationDateTime = DateTime.UtcNow.AddDays(2)
        };

        var created = await _graphClient.Subscriptions.PostAsync(subscription);
        if (created == null) throw new InvalidOperationException("Graph did not return the created subscription.");

        Console.WriteLine($"   ✅ Subscription: {created.Id}");
        Console.WriteLine($"   Expires: {created.ExpirationDateTime}");

        return created;
    }

    public async Task RenewSubscriptionAsync(string subscriptionId)
    {
        var patch = new Subscription
        {
            ExpirationDateTime = DateTime.UtcNow.AddDays(2)
        };

        var renewed = await _graphClient.Subscriptions[subscriptionId].PatchAsync(patch);
        if (renewed != null)
        {
            Console.WriteLine($"🔄 Renewed. New expiration: {renewed.ExpirationDateTime}");
        }
    }

    public async Task DeleteSubscriptionAsync(string subscriptionId)
    {
        await _graphClient.Subscriptions[subscriptionId].DeleteAsync();
        Console.WriteLine($"🗑️ Subscription {subscriptionId} deleted");
    }
}
