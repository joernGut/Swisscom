<#
.SYNOPSIS
    Azure, Microsoft 365 und Intune Anmeldescript mit MFA-Unterstützung
.DESCRIPTION
    Installiert benötigte Module (falls nicht vorhanden) und meldet sich interaktiv bei Azure, Microsoft 365 und Intune an
.NOTES
    Autor: PowerShell Script
    Datum: 2026-02-10
    Unterstützt Multi-Faktor-Authentifizierung (MFA)
    Version: 1.1 - Verbesserte Intune-Kompatibilität
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
    "Az.Accounts",                          # Azure PowerShell
    "Az.Resources",                         # Azure Ressourcen
    "ExchangeOnlineManagement",             # Exchange Online
    "Microsoft.Graph.Authentication",       # Microsoft Graph Authentication
    "Microsoft.Graph.Users",                # Microsoft Graph Users
    "Microsoft.Graph.Groups",               # Microsoft Graph Groups
    "Microsoft.Graph.DeviceManagement",     # Intune Device Management
    "AzureAD"                               # Azure AD
)

# Optionale Beta-Module
$optionalModules = @(
    "Microsoft.Graph.Beta.DeviceManagement" # Beta-Features für Intune (optional)
)

$allModulesOk = $true

foreach ($module in $requiredModules) {
    if (-not (Install-RequiredModule -ModuleName $module)) {
        $allModulesOk = $false
    }
    Write-Host ""
}

# Versuche optionale Module zu installieren
Write-Host "Installiere optionale Module..." -ForegroundColor Cyan
foreach ($module in $optionalModules) {
    Write-Host "Prüfe optionales Modul: $module..." -ForegroundColor Yellow
    try {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            Write-Host "  Optionales Modul '$module' installiert." -ForegroundColor Green
        } else {
            Write-Host "  Optionales Modul '$module' bereits vorhanden." -ForegroundColor Green
        }
    } catch {
        Write-Host "  Info: Optionales Modul '$module' nicht verfügbar (nicht kritisch)" -ForegroundColor Yellow
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
        "DeviceManagementApps.Read.All",                # Intune Apps (Read)
        "DeviceManagementConfiguration.Read.All",       # Intune Konfiguration (Read)
        "DeviceManagementManagedDevices.Read.All",      # Intune Geräte (Read)
        "DeviceManagementServiceConfig.Read.All"        # Intune Service-Konfiguration (Read)
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

# Intune Verbindung testen (mit Fallback-Mechanismen)
Write-Host "5. Teste Intune-Verbindung..." -ForegroundColor Cyan
$intuneWorking = $false
$intuneMethod = ""

try {
    # Methode 1: Versuche Standard-Cmdlet
    Write-Host "   Teste Standard-Cmdlets..." -ForegroundColor Gray
    $intuneDevices = Get-MgDeviceManagementManagedDevice -Top 1 -ErrorAction Stop
    $intuneWorking = $true
    $intuneMethod = "Standard Cmdlets"
    Write-Host "   Intune-Zugriff erfolgreich verifiziert (Standard-Cmdlets)!" -ForegroundColor Green
} catch {
    Write-Host "   Standard-Cmdlets nicht verfügbar, versuche Graph API direkt..." -ForegroundColor Yellow
    
    # Methode 2: Fallback auf direkte Graph API Aufrufe
    try {
        $graphUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$top=1"
        $result = Invoke-MgGraphRequest -Uri $graphUri -Method GET -ErrorAction Stop
        $intuneWorking = $true
        $intuneMethod = "Graph API (Direct)"
        Write-Host "   Intune-Zugriff erfolgreich verifiziert (Graph API)!" -ForegroundColor Green
    } catch {
        Write-Host "   WARNUNG: Intune-Zugriff konnte nicht verifiziert werden: $_" -ForegroundColor Yellow
        Write-Host "   Möglicherweise fehlen Intune-Lizenzen oder Berechtigungen" -ForegroundColor Yellow
    }
}

if ($intuneWorking) {
    Write-Host "   Zugriffsmethode: $intuneMethod" -ForegroundColor Gray
    
    # Zeige Intune-Informationen
    try {
        # Versuche Geräteanzahl zu ermitteln
        if ($intuneMethod -eq "Standard Cmdlets") {
            $allDevices = Get-MgDeviceManagementManagedDevice -All -ErrorAction SilentlyContinue
            if ($allDevices) {
                Write-Host "   Verwaltete Geräte: $($allDevices.Count)" -ForegroundColor Gray
            }
        } else {
            # Verwende Graph API für Geräteanzahl
            $graphUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/`$count"
            $deviceCount = Invoke-MgGraphRequest -Uri $graphUri -Method GET -ErrorAction SilentlyContinue
            if ($deviceCount) {
                Write-Host "   Verwaltete Geräte: $deviceCount" -ForegroundColor Gray
            }
        }
        
        # Teste auch Policy-Zugriff
        try {
            $configUri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations?`$top=1"
            $configTest = Invoke-MgGraphRequest -Uri $configUri -Method GET -ErrorAction Stop
            Write-Host "   Policy-Zugriff: Verfügbar" -ForegroundColor Gray
        } catch {
            Write-Host "   Policy-Zugriff: Eingeschränkt" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "   Hinweis: Erweiterte Intune-Informationen nicht verfügbar" -ForegroundColor Yellow
    }
}

Write-Host ""

# Zusammenfassung
Write-Host "=== Anmeldung abgeschlossen ===" -ForegroundColor Cyan
Write-Host ""

# Prüfe welche Verbindungen erfolgreich waren
$connections = @()

if (Get-AzContext -ErrorAction SilentlyContinue) {
    $connections += "✓ Azure (Az)"
}

try {
    $exoTest = Get-OrganizationConfig -ErrorAction SilentlyContinue
    if ($exoTest) {
        $connections += "✓ Exchange Online"
    }
} catch {}

if (Get-MgContext -ErrorAction SilentlyContinue) {
    $connections += "✓ Microsoft Graph"
}

try {
    $aadTest = Get-AzureADTenantDetail -ErrorAction SilentlyContinue
    if ($aadTest) {
        $connections += "✓ Azure AD"
    }
} catch {}

if ($intuneWorking) {
    $connections += "✓ Intune (via $intuneMethod)"
}

if ($connections.Count -gt 0) {
    Write-Host "Erfolgreich angemeldet bei:" -ForegroundColor Green
    foreach ($conn in $connections) {
        Write-Host "  $conn" -ForegroundColor White
    }
} else {
    Write-Host "WARNUNG: Keine erfolgreichen Verbindungen hergestellt!" -ForegroundColor Red
}

Write-Host ""
Write-Host "Sie können nun mit der Arbeit beginnen!" -ForegroundColor Cyan
Write-Host ""

# Optionale Anzeige von Hilfe-Befehlen
Write-Host "Nützliche Befehle zum Testen:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  # Azure" -ForegroundColor Cyan
Write-Host "  Get-AzSubscription                                      # Abonnements anzeigen" -ForegroundColor Gray
Write-Host "  Get-AzResource | Select-Object -First 10                # Ressourcen anzeigen" -ForegroundColor Gray
Write-Host ""
Write-Host "  # Exchange Online" -ForegroundColor Cyan
Write-Host "  Get-Mailbox -ResultSize 10                              # Postfächer anzeigen" -ForegroundColor Gray
Write-Host "  Get-OrganizationConfig                                  # Org-Konfiguration" -ForegroundColor Gray
Write-Host ""
Write-Host "  # Microsoft Graph" -ForegroundColor Cyan
Write-Host "  Get-MgUser -Top 10                                      # Benutzer anzeigen" -ForegroundColor Gray
Write-Host "  Get-MgGroup -Top 10                                     # Gruppen anzeigen" -ForegroundColor Gray
Write-Host ""
Write-Host "  # Azure AD" -ForegroundColor Cyan
Write-Host "  Get-AzureADUser -Top 10                                 # Benutzer anzeigen" -ForegroundColor Gray
Write-Host "  Get-AzureADTenantDetail                                 # Tenant-Details" -ForegroundColor Gray
Write-Host ""
Write-Host "  # Intune (Microsoft Graph - Standard Cmdlets)" -ForegroundColor Cyan
Write-Host "  Get-MgDeviceManagementManagedDevice -Top 10             # Verwaltete Geräte" -ForegroundColor Gray
Write-Host "  Get-MgDeviceManagementDeviceConfiguration               # Gerätekonfigurationen" -ForegroundColor Gray
Write-Host "  Get-MgDeviceManagementDeviceCompliancePolicy            # Compliance-Richtlinien" -ForegroundColor Gray
Write-Host ""
Write-Host "  # Intune (Graph API - Fallback-Methode)" -ForegroundColor Cyan
Write-Host "  # Verwenden Sie diese, wenn Standard-Cmdlets nicht verfügbar sind:" -ForegroundColor Yellow
Write-Host "  Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices' -Method GET" -ForegroundColor Gray
Write-Host "  Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies' -Method GET" -ForegroundColor Gray
Write-Host ""

# Zeige installierte Graph-Module
Write-Host "Installierte Microsoft.Graph Module:" -ForegroundColor Cyan
$graphModules = Get-Module -ListAvailable -Name "Microsoft.Graph*" | Select-Object Name, Version | Sort-Object Name
if ($graphModules) {
    $graphModules | Format-Table -AutoSize
} else {
    Write-Host "  Keine Microsoft.Graph Module gefunden" -ForegroundColor Yellow
}
Write-Host ""