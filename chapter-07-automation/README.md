# Chapter 7: Automation and Flows - Code Samples

This directory contains practical examples for implementing automation in SharePoint Online via webhooks, delta queries, and data synchronization.

## 📁 Folder Structure

```
capitulo-07-automatizacion/
├── csharp/
│   ├── 01-webhooks/
│   │   └── WebhookManager.cs
│   └── 02-delta-queries/
├── powershell/
│   └── 01-webhooks/
│       └── Manage-GraphSubscriptions.ps1
├── javascript/
│   └── 01-webhooks/
│       └── webhookManager.js
├── python/
│   └── 01-webhooks/
│       └── webhook_manager.py
└── java/
    └── 01-webhooks/
        └── WebhookManager.java
```

## 🎯 Topics Covered

### 01 - Webhooks and Subscriptions

- Create subscriptions for Drive items
- Automatic subscription renewal
- Remove stale subscriptions
- List active subscriptions
- Process webhook notifications (example in the chapter)

### 02 - Delta Queries (coming soon)

- Incremental synchronization
- Storing the deltaLink
- Handling deletions and renames
- Bidirectional synchronization

## 📋 Prerequisites

### Required Permissions

| Permission | Description |
|---------|-------------|
| `Subscriptions.Read.All` | Read subscriptions |
| `Subscriptions.ReadWrite.All` | Create/modify subscriptions |

### Webhook Endpoint Configuration

Your endpoint must:
1. Be publicly accessible (HTTPS)
2. Respond to the validation handshake (return the token)
3. Respond with 202 Accepted for notifications
4. Process asynchronously

### ngrok for Local Development

```bash
# Instalar ngrok
choco install ngrok

# Exponer tu endpoint local
ngrok http https://localhost:5001

# Usar la URL HTTPS proporcionada
# https://abc123.ngrok.io/api/webhooks
```

## 🚀 Execution

### C#

```bash
cd csharp/01-webhooks
# Compilar y ejecutar como parte de tu aplicación ASP.NET Core
```

### PowerShell

```powershell
Import-Module .\Manage-GraphSubscriptions.ps1
New-DriveSubscription -DriveId "abc123" -NotificationUrl "https://tu-endpoint.com/webhook"
```

### JavaScript

```bash
cd javascript/01-webhooks
npm install @microsoft/microsoft-graph-client @azure/identity
node webhookManager.js
```

### Python

```bash
cd python/01-webhooks
pip install -r requirements.txt
python webhook_manager.py
```

### Java

```bash
cd java/01-webhooks
mvn compile exec:java -Dexec.mainClass="WebhookManager"
```

## 🔐 Security

### Client State Validation

```csharp
if (notification.ClientState != storedClientState) {
    logger.LogWarning("ClientState inválido - ignorando notificación");
    return Unauthorized();
}
```

### Important Headers

- `X-Microsoft-Skype-Chain-ID`: notification ID
- `Content-Type: application/json`

## 📚 References

- [Microsoft Graph Webhooks](https://docs.microsoft.com/en-us/graph/webhooks)
- [Delta Query Overview](https://docs.microsoft.com/en-us/graph/delta-query-overview)
- [Change Notifications Lifecycle](https://docs.microsoft.com/en-us/graph/webhooks-lifecycle)

## ⚠️ Limitations

- Subscriptions expire in a maximum of 3 days
- Automatic renewal is required
- Maximum 10,000 subscriptions per organization
- The endpoint must respond in < 30 seconds

## 💡 Best Practices

1. **Always queue** notifications, do not process synchronously
2. **Implement retry** with exponential backoff
3. **Validate clientState** to prevent spoofing
4. **Respond quickly** with 202 Accepted
5. **Renew subscriptions** before expiration
6. **Clean up** stale subscriptions regularly