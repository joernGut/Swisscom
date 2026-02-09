<#
.SYNOPSIS
    Azure, Microsoft 365 und Intune Anmeldescript mit MFA-Unterstützung
.DESCRIPTION
    Installiert benötigte Module (falls nicht vorhanden) und meldet sich interaktiv bei Azure, Microsoft 365 und Intune an
.NOTES
    Autor: PowerShell Script
    Datum: 2026-02-09
    Unterstützt Multi-Faktor-Authentifizierung (MFA)
#>

Write-Host "=== Azure, M365 und Intune Anmeldescript (MFA-fähig) ===" -ForegroundColor Cyan
Write-Host ""

# Funktion zum Prüfen und Installieren von Modulen
function Install-RequiredModule {
    param(
        [string]$ModuleName
    )
    
    Write-Host "Prüfe Modul: $ModuleName..." -ForegroundColor Yellow
    
    if (Get-Module -ListAvailable -Name $ModuleName) {
        Write-Host "  Modul '$ModuleName' ist bereits installiert." -ForegroundColor Green
    } else {
        Write-Host "  Modul '$ModuleName' wird installiert..." -ForegroundColor Yellow
        try {
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            Write-Host "  Modul '$ModuleName' erfolgreich installiert." -ForegroundColor Green
        } catch {
            Write-Host "  FEHLER beim Installieren von '$ModuleName': $_" -ForegroundColor Red
            return $false
        }
    }
    
    # Modul importieren
    Write-Host "  Importiere Modul '$ModuleName'..." -ForegroundColor Yellow
    try {
        Import-Module $ModuleName -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "  Modul '$ModuleName' erfolgreich importiert." -ForegroundColor Green
        return $true
    } catch {
        Write-Host "  FEHLER beim Importieren von '$ModuleName': $_" -ForegroundColor Red
        return $false
    }
}

Write-Host "Schritt 1: Überprüfung und Installation der benötigten Module" -ForegroundColor Cyan
Write-Host ""

# Liste der benötigten Module
$requiredModules = @(
    "Az.Accounts",                  # Azure PowerShell
    "Az.Resources",                 # Azure Ressourcen
    "ExchangeOnlineManagement",     # Exchange Online
    "Microsoft.Graph",              # Microsoft Graph
    "Microsoft.Graph.Intune",       # Intune (Teil von Graph)
    "AzureAD"                       # Azure AD
)

$allModulesOk = $true

foreach ($module in $requiredModules) {
    if (-not (Install-RequiredModule -ModuleName $module)) {
        $allModulesOk = $false
    }
    Write-Host ""
}

if (-not $allModulesOk) {
    Write-Host "WARNUNG: Nicht alle Module konnten installiert werden." -ForegroundColor Yellow
    $continue = Read-Host "Möchten Sie trotzdem fortfahren? (J/N)"
    if ($continue -ne "J" -and $continue -ne "j") {
        Write-Host "Script abgebrochen." -ForegroundColor Red
        exit
    }
}

Write-Host ""
Write-Host "Schritt 2: Interaktive Anmeldung bei Azure, Microsoft 365 und Intune" -ForegroundColor Cyan
Write-Host "Hinweis: Sie werden für jeden Service einzeln zur Anmeldung aufgefordert." -ForegroundColor Yellow
Write-Host "Bitte nutzen Sie bei der Anmeldung Ihre MFA-Methode (z.B. Authenticator-App)." -ForegroundColor Yellow
Write-Host ""

Start-Sleep -Seconds 2

# Azure Anmeldung
Write-Host "1. Melde bei Azure an..." -ForegroundColor Cyan
Write-Host "   Ein Browser-Fenster wird geöffnet..." -ForegroundColor Yellow
try {
    Connect-AzAccount -ErrorAction Stop
    Write-Host "   Erfolgreich bei Azure angemeldet!" -ForegroundColor Green
    
    # Zeige Abonnement-Informationen
    $context = Get-AzContext
    Write-Host "   Benutzer: $($context.Account.Id)" -ForegroundColor Gray
    Write-Host "   Abonnement: $($context.Subscription.Name)" -ForegroundColor Gray
    Write-Host "   Tenant: $($context.Tenant.Id)" -ForegroundColor Gray
} catch {
    Write-Host "   FEHLER bei Azure-Anmeldung: $_" -ForegroundColor Red
}
Write-Host ""
Start-Sleep -Seconds 1

# Exchange Online Anmeldung
Write-Host "2. Melde bei Exchange Online an..." -ForegroundColor Cyan
Write-Host "   Ein Browser-Fenster wird geöffnet..." -ForegroundColor Yellow
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "   Erfolgreich bei Exchange Online angemeldet!" -ForegroundColor Green
    
    # Zeige Organisations-Informationen
    $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
    if ($orgConfig) {
        Write-Host "   Organisation: $($orgConfig.DisplayName)" -ForegroundColor Gray
    }
} catch {
    Write-Host "   FEHLER bei Exchange Online-Anmeldung: $_" -ForegroundColor Red
}
Write-Host ""
Start-Sleep -Seconds 1

# Microsoft Graph Anmeldung (inkl. Intune-Berechtigungen)
Write-Host "3. Melde bei Microsoft Graph (inkl. Intune) an..." -ForegroundColor Cyan
Write-Host "   Ein Browser-Fenster wird geöffnet..." -ForegroundColor Yellow
Write-Host "   Hinweis: Intune-Zugriff erfolgt über Microsoft Graph" -ForegroundColor Yellow
try {
    # Erweiterte Berechtigungen für Graph inkl. Intune
    $graphScopes = @(
        "User.Read.All",
        "Group.Read.All",
        "Directory.Read.All",
        "Organization.Read.All",
        "DeviceManagementApps.ReadWrite.All",           # Intune Apps
        "DeviceManagementConfiguration.ReadWrite.All",  # Intune Konfiguration
        "DeviceManagementManagedDevices.ReadWrite.All", # Intune Geräte
        "DeviceManagementServiceConfig.ReadWrite.All"   # Intune Service-Konfiguration
    )
    
    Connect-MgGraph -Scopes $graphScopes -ErrorAction Stop
    Write-Host "   Erfolgreich bei Microsoft Graph angemeldet!" -ForegroundColor Green
    
    # Zeige Kontext-Informationen
    $mgContext = Get-MgContext
    Write-Host "   Benutzer: $($mgContext.Account)" -ForegroundColor Gray
    Write-Host "   Tenant: $($mgContext.TenantId)" -ForegroundColor Gray
    Write-Host "   Scopes: Intune-Berechtigungen inkludiert" -ForegroundColor Gray
} catch {
    Write-Host "   FEHLER bei Microsoft Graph-Anmeldung: $_" -ForegroundColor Red
}
Write-Host ""
Start-Sleep -Seconds 1

# Azure AD Anmeldung
Write-Host "4. Melde bei Azure AD an..." -ForegroundColor Cyan
Write-Host "   Ein Browser-Fenster wird geöffnet..." -ForegroundColor Yellow
try {
    Connect-AzureAD -ErrorAction Stop
    Write-Host "   Erfolgreich bei Azure AD angemeldet!" -ForegroundColor Green
    
    # Zeige Tenant-Informationen
    $tenantDetail = Get-AzureADTenantDetail
    Write-Host "   Tenant Name: $($tenantDetail.DisplayName)" -ForegroundColor Gray
    Write-Host "   Tenant ID: $($tenantDetail.ObjectId)" -ForegroundColor Gray
} catch {
    Write-Host "   FEHLER bei Azure AD-Anmeldung: $_" -ForegroundColor Red
}
Write-Host ""

# Intune Verbindung testen
Write-Host "5. Teste Intune-Verbindung..." -ForegroundColor Cyan
try {
    # Test Intune-Zugriff über Graph
    $intuneDevices = Get-MgDeviceManagementManagedDevice -Top 1 -ErrorAction Stop
    Write-Host "   Intune-Zugriff erfolgreich verifiziert!" -ForegroundColor Green
    
    # Zeige Intune-Informationen
    $allDevices = Get-MgDeviceManagementManagedDevice -All -ErrorAction SilentlyContinue
    if ($allDevices) {
        Write-Host "   Verwaltete Geräte: $($allDevices.Count)" -ForegroundColor Gray
    }
} catch {
    Write-Host "   WARNUNG: Intune-Zugriff konnte nicht verifiziert werden: $_" -ForegroundColor Yellow
    Write-Host "   Möglicherweise fehlen Intune-Lizenzen oder Berechtigungen" -ForegroundColor Yellow
}
Write-Host ""

# Zusammenfassung
Write-Host "=== Anmeldung abgeschlossen ===" -ForegroundColor Cyan
Write-Host ""

# Prüfe welche Verbindungen erfolgreich waren
$connections = @()

if (Get-AzContext -ErrorAction SilentlyContinue) {
    $connections += "Azure (Az)"
}

try {
    $exoTest = Get-OrganizationConfig -ErrorAction SilentlyContinue
    if ($exoTest) {
        $connections += "Exchange Online"
    }
} catch {}

if (Get-MgContext -ErrorAction SilentlyContinue) {
    $connections += "Microsoft Graph"
}

try {
    $aadTest = Get-AzureADTenantDetail -ErrorAction SilentlyContinue
    if ($aadTest) {
        $connections += "Azure AD"
    }
} catch {}

try {
    $intuneTest = Get-MgDeviceManagementManagedDevice -Top 1 -ErrorAction SilentlyContinue
    if ($intuneTest -or $?) {
        $connections += "Intune (via Graph)"
    }
} catch {}

if ($connections.Count -gt 0) {
    Write-Host "Erfolgreich angemeldet bei:" -ForegroundColor Green
    foreach ($conn in $connections) {
        Write-Host "  ✓ $conn" -ForegroundColor White
    }
} else {
    Write-Host "WARNUNG: Keine erfolgreichen Verbindungen hergestellt!" -ForegroundColor Red
}

Write-Host ""
Write-Host "Sie können nun mit der Arbeit beginnen!" -ForegroundColor Cyan
Write-Host ""

# Optionale Anzeige von Hilfe-Befehlen
Write-Host "Nützliche Befehle zum Testen:" -ForegroundColor Yellow
Write-Host "  # Azure" -ForegroundColor Cyan
Write-Host "  Get-AzSubscription                              # Abonnements anzeigen" -ForegroundColor Gray
Write-Host ""
Write-Host "  # Exchange Online" -ForegroundColor Cyan
Write-Host "  Get-Mailbox -ResultSize 10                      # Postfächer anzeigen" -ForegroundColor Gray
Write-Host ""
Write-Host "  # Microsoft Graph" -ForegroundColor Cyan
Write-Host "  Get-MgUser -Top 10                              # Benutzer anzeigen" -ForegroundColor Gray
Write-Host ""
Write-Host "  # Azure AD" -ForegroundColor Cyan
Write-Host "  Get-AzureADUser -Top 10                         # Benutzer anzeigen" -ForegroundColor Gray
Write-Host ""
Write-Host "  # Intune (Microsoft Graph)" -ForegroundColor Cyan
Write-Host "  Get-MgDeviceManagementManagedDevice -Top 10     # Verwaltete Geräte" -ForegroundColor Gray
Write-Host "  Get-MgDeviceAppManagement                       # Intune Apps" -ForegroundColor Gray
Write-Host "  Get-MgDeviceManagementDeviceConfiguration       # Gerätekonfigurationen" -ForegroundColor Gray
Write-Host "  Get-MgDeviceManagementDeviceCompliancePolicy    # Compliance-Richtlinien" -ForegroundColor Gray
Write-Host ""