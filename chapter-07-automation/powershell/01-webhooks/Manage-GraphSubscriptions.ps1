# Manage-GraphSubscriptions.ps1
# Ejemplo: Gestionar suscripciones webhook de Microsoft Graph

function New-DriveSubscription {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DriveId,

        [Parameter(Mandatory=$true)]
        [string]$NotificationUrl
    )

    Write-Host "🔔 Creando suscripción..." -ForegroundColor Cyan

    $subscription = @{
        resource = "/drives/$DriveId/root"
        changeType = "created,updated,deleted"
        notificationUrl = $NotificationUrl
        clientState = [Guid]::NewGuid().ToString()
        expirationDateTime = (Get-Date).AddDays(2).ToString("o")
    } | ConvertTo-Json

    try {
        $response = Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/subscriptions" `
            -Body $subscription

        Write-Host "   ✅ Creada: $($response.id)" -ForegroundColor Green
        return $response
    }
    catch {
        Write-Error "❌ Error: $($_.Exception.Message)"
    }
}

function Get-GraphSubscriptions {
    Write-Host "📋 Suscripciones activas:" -ForegroundColor Cyan

    $subs = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/subscriptions"

    foreach ($sub in $subs.value) {
        Write-Host "`n🆔 $($sub.id)" -ForegroundColor Yellow
        Write-Host "   Recurso: $($sub.resource)"
        Write-Host "   Expira: $($sub.expirationDateTime)"
    }
}

function Remove-GraphSubscription {
    param([Parameter(Mandatory=$true)][string]$SubscriptionId)

    Invoke-MgGraphRequest -Method DELETE `
        -Uri "https://graph.microsoft.com/v1.0/subscriptions/$SubscriptionId"

    Write-Host "🗑️ Suscripción eliminada" -ForegroundColor Green
}

Export-ModuleMember -Function New-DriveSubscription, Get-GraphSubscriptions, Remove-GraphSubscription
