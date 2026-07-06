# Capítulo 7: Automatización y Flujos - Ejemplos de Código

Este directorio contiene ejemplos prácticos para implementar automatización en SharePoint Online mediante webhooks, delta queries y sincronización de datos.

## 📁 Estructura de Carpetas

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

## 🎯 Temas Cubiertos

### 01 - Webhooks y Suscripciones

- Crear suscripciones para Drive items
- Renovación automática de suscripciones
- Eliminar suscripciones obsoletas
- Listar suscripciones activas
- Procesar notificaciones webhook (ejemplo en capítulo)

### 02 - Delta Queries (próximamente)

- Sincronización incremental
- Almacenamiento de deltaLink
- Manejo de eliminaciones y renombres
- Sincronización bidireccional

## 📋 Requisitos Previos

### Permisos Requeridos

| Permiso | Descripción |
|---------|-------------|
| `Subscriptions.Read.All` | Leer suscripciones |
| `Subscriptions.ReadWrite.All` | Crear/modificar suscripciones |

### Configuración del Endpoint Webhook

Tu endpoint debe:
1. Ser accesible públicamente (HTTPS)
2. Responder al handshake de validación (devolver el token)
3. Responder con 202 Accepted para notificaciones
4. Procesar de forma asíncrona

### ngrok para Desarrollo Local

```bash
# Instalar ngrok
choco install ngrok

# Exponer tu endpoint local
ngrok http https://localhost:5001

# Usar la URL HTTPS proporcionada
# https://abc123.ngrok.io/api/webhooks
```

## 🚀 Ejecución

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

## 🔐 Seguridad

### Validación del Client State

```csharp
if (notification.ClientState != storedClientState) {
    logger.LogWarning("ClientState inválido - ignorando notificación");
    return Unauthorized();
}
```

### Encabezados Importantes

- `X-Microsoft-Skype-Chain-ID`: ID de la notificación
- `Content-Type: application/json`

## 📚 Referencias

- [Microsoft Graph Webhooks](https://docs.microsoft.com/en-us/graph/webhooks)
- [Delta Query Overview](https://docs.microsoft.com/en-us/graph/delta-query-overview)
- [Change Notifications Lifecycle](https://docs.microsoft.com/en-us/graph/webhooks-lifecycle)

## ⚠️ Limitaciones

- Suscripciones expiran en máximo 3 días
- Se requiere renovación automática
- Máximo 10,000 suscripciones por organización
- El endpoint debe responder en < 30 segundos

## 💡 Mejores Prácticas

1. **Siempre encolar** las notificaciones, no procesar sincrónicamente
2. **Implementar retry** con backoff exponencial
3. **Validar clientState** para evitar spoofing
4. **Responder rápido** con 202 Accepted
5. **Renovar suscripciones** antes de expiración
6. **Limpiar suscripciones** obsoletas regularmente
