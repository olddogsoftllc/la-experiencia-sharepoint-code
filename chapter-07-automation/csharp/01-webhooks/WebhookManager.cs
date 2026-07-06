// WebhookManager.cs
// Ejemplo: Crear y gestionar suscripciones webhook en Microsoft Graph

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
        Console.WriteLine($"🔔 Creando suscripción para Drive: {driveId}");

        var subscription = new Subscription
        {
            Resource = $"/drives/{driveId}/root",
            ChangeType = "created,updated,deleted",
            NotificationUrl = _notificationUrl,
            ClientState = Guid.NewGuid().ToString(),
            ExpirationDateTime = DateTime.UtcNow.AddDays(2)
        };

        var created = await _graphClient.Subscriptions.PostAsync(subscription);
        if (created == null) throw new InvalidOperationException("Graph no devolvió la suscripción creada.");

        Console.WriteLine($"   ✅ Suscripción: {created.Id}");
        Console.WriteLine($"   Expira: {created.ExpirationDateTime}");

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
            Console.WriteLine($"🔄 Renovada. Nueva expiración: {renewed.ExpirationDateTime}");
        }
    }

    public async Task DeleteSubscriptionAsync(string subscriptionId)
    {
        await _graphClient.Subscriptions[subscriptionId].DeleteAsync();
        Console.WriteLine($"🗑️ Suscripción {subscriptionId} eliminada");
    }
}
