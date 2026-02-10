<#
.SYNOPSIS
    Azure, Microsoft 365, Intune und SharePoint Online Anmeldescript mit MFA-Unterstützung
.DESCRIPTION
    Installiert benötigte Module (falls nicht vorhanden) und meldet sich interaktiv bei 
    Azure, Microsoft 365, Intune und SharePoint Online an
.NOTES
    Autor: BWS PowerShell Script
    Datum: 2026-02-10
    Unterstützt Multi-Faktor-Authentifizierung (MFA)
    Version: 2.0 - SharePoint Online Support hinzugefügt
#>

Write-Host "=== Azure, M365, Intune und SharePoint Online Anmeldescript ===" -ForegroundColor Cyan
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
    "AzureAD",                              # Azure AD
    "PnP.PowerShell"                        # SharePoint Online (PnP - Modern)
)

# Alternative SharePoint Module
$alternativeModules = @(
    "Microsoft.Online.SharePoint.PowerShell"  # SharePoint Online (Legacy)
)

$allModulesOk = $true

foreach ($module in $requiredModules) {
    if (-not (Install-RequiredModule -ModuleName $module)) {
        $allModulesOk = $false
    }
    Write-Host ""
}

# Versuche alternative SharePoint Module wenn PnP.PowerShell fehlschlägt
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    Write-Host "PnP.PowerShell nicht verfügbar, versuche alternative SharePoint Module..." -ForegroundColor Yellow
    foreach ($module in $alternativeModules) {
        Write-Host "Prüfe alternatives Modul: $module..." -ForegroundColor Yellow
        try {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                Write-Host "  Alternatives Modul '$module' installiert." -ForegroundColor Green
            } else {
                Write-Host "  Alternatives Modul '$module' bereits vorhanden." -ForegroundColor Green
            }
        } catch {
            Write-Host "  Info: Alternatives Modul '$module' nicht verfügbar" -ForegroundColor Yellow
        }
        Write-Host ""
    }
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
Write-Host "Schritt 2: Interaktive Anmeldung bei Azure, Microsoft 365, Intune und SharePoint" -ForegroundColor Cyan
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
Start-Sleep -Seconds 1

# SharePoint Online Anmeldung
Write-Host "5. Melde bei SharePoint Online an..." -ForegroundColor Cyan
Write-Host "   Ein Browser-Fenster wird geöffnet..." -ForegroundColor Yellow

# Ermittle SharePoint Admin-URL aus Tenant
$sharePointConnected = $false
$sharePointMethod = ""

if ($tenantDetail) {
    # Extrahiere Tenant-Name aus Domäne
    $tenantDomains = Get-AzureADDomain -ErrorAction SilentlyContinue
    $onMicrosoftDomain = $tenantDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" -and $_.Name -notlike "*.mail.onmicrosoft.com" } | Select-Object -First 1
    
    if ($onMicrosoftDomain) {
        $tenantName = $onMicrosoftDomain.Name -replace "\.onmicrosoft\.com", ""
        $sharePointAdminUrl = "https://$tenantName-admin.sharepoint.com"
        
        Write-Host "   SharePoint Admin URL: $sharePointAdminUrl" -ForegroundColor Gray
        
        # Versuche mit PnP.PowerShell (bevorzugt)
        if (Get-Module -ListAvailable -Name "PnP.PowerShell") {
            Write-Host "   Verwende PnP.PowerShell..." -ForegroundColor Gray
            try {
                Connect-PnPOnline -Url $sharePointAdminUrl -Interactive -ErrorAction Stop
                $sharePointConnected = $true
                $sharePointMethod = "PnP.PowerShell"
                Write-Host "   Erfolgreich bei SharePoint Online angemeldet (PnP)!" -ForegroundColor Green
            } catch {
                Write-Host "   PnP Verbindung fehlgeschlagen, versuche Legacy-Modul..." -ForegroundColor Yellow
            }
        }
        
        # Fallback auf Microsoft.Online.SharePoint.PowerShell
        if (-not $sharePointConnected -and (Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell")) {
            Write-Host "   Verwende Microsoft.Online.SharePoint.PowerShell..." -ForegroundColor Gray
            try {
                Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
                Connect-SPOService -Url $sharePointAdminUrl -ErrorAction Stop
                $sharePointConnected = $true
                $sharePointMethod = "Microsoft.Online.SharePoint.PowerShell"
                Write-Host "   Erfolgreich bei SharePoint Online angemeldet (Legacy)!" -ForegroundColor Green
            } catch {
                Write-Host "   FEHLER bei SharePoint Online-Anmeldung: $_" -ForegroundColor Red
            }
        }
        
        if (-not $sharePointConnected) {
            Write-Host "   WARNUNG: Konnte nicht mit SharePoint verbinden" -ForegroundColor Yellow
            Write-Host "   Bitte installieren Sie eines der Module:" -ForegroundColor Yellow
            Write-Host "     - PnP.PowerShell (empfohlen)" -ForegroundColor Gray
            Write-Host "     - Microsoft.Online.SharePoint.PowerShell" -ForegroundColor Gray
        } else {
            Write-Host "   Methode: $sharePointMethod" -ForegroundColor Gray
            
            # Teste SharePoint-Zugriff
            try {
                if ($sharePointMethod -eq "PnP.PowerShell") {
                    $tenant = Get-PnPTenant -ErrorAction SilentlyContinue
                    if ($tenant) {
                        Write-Host "   SharePoint Tenant: $($tenant.Title)" -ForegroundColor Gray
                    }
                } else {
                    $tenant = Get-SPOTenant -ErrorAction SilentlyContinue
                    if ($tenant) {
                        Write-Host "   SharePoint Root URL: $($tenant.SharePointUrl)" -ForegroundColor Gray
                    }
                }
            } catch {
                Write-Host "   Hinweis: SharePoint Tenant-Informationen nicht verfügbar" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "   WARNUNG: Konnte Tenant-Domäne nicht ermitteln" -ForegroundColor Yellow
        Write-Host "   Bitte melden Sie sich manuell an:" -ForegroundColor Yellow
        Write-Host "     Connect-PnPOnline -Url https://YOURTENANT-admin.sharepoint.com -Interactive" -ForegroundColor Gray
        Write-Host "   oder" -ForegroundColor Yellow
        Write-Host "     Connect-SPOService -Url https://YOURTENANT-admin.sharepoint.com" -ForegroundColor Gray
    }
} else {
    Write-Host "   WARNUNG: Azure AD Tenant-Details nicht verfügbar" -ForegroundColor Yellow
    Write-Host "   SharePoint-Anmeldung übersprungen" -ForegroundColor Yellow
}

Write-Host ""
Start-Sleep -Seconds 1

# Intune Verbindung testen (mit Fallback-Mechanismen)
Write-Host "6. Teste Intune-Verbindung..." -ForegroundColor Cyan
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

if ($sharePointConnected) {
    $connections += "✓ SharePoint Online (via $sharePointMethod)"
}

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
Write-Host "  # SharePoint Online (PnP.PowerShell)" -ForegroundColor Cyan
Write-Host "  Get-PnPTenant                                           # Tenant-Einstellungen" -ForegroundColor Gray
Write-Host "  Get-PnPSite                                             # Sites anzeigen" -ForegroundColor Gray
Write-Host ""
Write-Host "  # SharePoint Online (Legacy)" -ForegroundColor Cyan
Write-Host "  Get-SPOTenant                                           # Tenant-Einstellungen" -ForegroundColor Gray
Write-Host "  Get-SPOSite                                             # Sites anzeigen" -ForegroundColor Gray
Write-Host ""
Write-Host "  # Intune (Microsoft Graph - Standard Cmdlets)" -ForegroundColor Cyan
Write-Host "  Get-MgDeviceManagementManagedDevice -Top 10             # Verwaltete Geräte" -ForegroundColor Gray
Write-Host "  Get-MgDeviceManagementDeviceConfiguration               # Gerätekonfigurationen" -ForegroundColor Gray
Write-Host "  Get-MgDeviceManagementDeviceCompliancePolicy            # Compliance-Richtlinien" -ForegroundColor Gray
Write-Host ""