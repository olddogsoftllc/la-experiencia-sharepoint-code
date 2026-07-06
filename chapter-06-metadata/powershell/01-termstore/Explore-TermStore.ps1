# Explore-TermStore.ps1
# Example: Explore the Term Store using PnP.PowerShell

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl
)

function Explore-TermStore {
    param([string]$Url)

    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "Explorando Term Store" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host

    try {
        Connect-PnPOnline -Url $Url -Interactive

        # Get Term Store
        $termStore = Get-PnPTermStore

        Write-Host "📚 Term Store Info:" -ForegroundColor Yellow
        Write-Host "   ID: $($termStore.Id)"
        Write-Host "   Default Language: $($termStore.DefaultLanguage)"
        Write-Host "   Languages: $($termStore.Languages -join ', ')"
        Write-Host

        # List groups
        $groups = Get-PnPTermGroup

        Write-Host "📂 Grupos encontrados: $($groups.Count)" -ForegroundColor Yellow
        Write-Host

        foreach ($group in $groups) {
            Write-Host "   📁 $($group.Name)" -ForegroundColor Green
            Write-Host "      ID: $($group.Id)"
            Write-Host "      Description: $($group.Description)"

            # Get term sets
            $termSets = Get-PnPTermSet -TermGroup $group.Name

            foreach ($termSet in $termSets) {
                Write-Host "      📚 $($termSet.Name)" -ForegroundColor Cyan
            }
            Write-Host
        }
    }
    catch {
        Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        Disconnect-PnPOnline
    }
}

# Run
Explore-TermStore -Url $SiteUrl
