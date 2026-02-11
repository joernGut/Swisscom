<#
.SYNOPSIS
    Azure, M365, Intune, SharePoint und Teams Login-Script mit GUI
.DESCRIPTION
    GUI-basiertes oder konsolenbasiertes Anmeldescript für Azure, Microsoft 365, Intune, SharePoint Online und Microsoft Teams
    - Auswahl der gewünschten Dienste per Checkbox oder Parameter
    - SharePoint-URL manuell eingeben
    - Automatische Modul-Installation
    - PowerShell 5.1 und 7 Konsolen-Support
.PARAMETER Console
    Starte im Konsolen-Modus ohne GUI
.PARAMETER Azure
    Melde bei Azure an (nur im Konsolen-Modus)
.PARAMETER Exchange
    Melde bei Exchange Online an (nur im Konsolen-Modus)
.PARAMETER Graph
    Melde bei Microsoft Graph an (nur im Konsolen-Modus)
.PARAMETER AzureAD
    Melde bei Azure AD an (nur im Konsolen-Modus)
.PARAMETER SharePoint
    Melde bei SharePoint Online an (nur im Konsolen-Modus)
.PARAMETER Teams
    Melde bei Microsoft Teams an (nur im Konsolen-Modus)
.PARAMETER SharePointUrl
    SharePoint Admin URL (Standard: auto-detect aus Tenant)
.NOTES
    Version: 1.1.0
    Datum: 2025-02-11
    Autor: BWS PowerShell Script
.EXAMPLE
    .\Azure-M365-Login-GUI.ps1
    Startet die GUI
.EXAMPLE
    .\Azure-M365-Login-GUI.ps1 -Console -Azure -Graph -SharePoint -Teams
    Startet im Konsolen-Modus und meldet bei Azure, Graph, SharePoint und Teams an
#>

param(
    [Parameter(Mandatory=$false)]
    [switch]$Console,
    
    [Parameter(Mandatory=$false)]
    [switch]$Azure,
    
    [Parameter(Mandatory=$false)]
    [switch]$Exchange,
    
    [Parameter(Mandatory=$false)]
    [switch]$Graph,
    
    [Parameter(Mandatory=$false)]
    [switch]$AzureAD,
    
    [Parameter(Mandatory=$false)]
    [switch]$SharePoint,
    
    [Parameter(Mandatory=$false)]
    [switch]$Teams,
    
    [Parameter(Mandatory=$false)]
    [string]$SharePointUrl = ""
)

# Script Version
$script:Version = "1.1.0"

# Requires PowerShell 5.1 or higher
#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================================
# Modul-Installation Funktionen
# ============================================================================

function Install-RequiredModule {
    param(
        [string]$ModuleName,
        [System.Windows.Forms.RichTextBox]$OutputBox
    )
    
    $OutputBox.SelectionColor = [System.Drawing.Color]::DarkCyan
    $OutputBox.AppendText("Prüfe Modul: $ModuleName...`r`n")
    $OutputBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    if (Get-Module -ListAvailable -Name $ModuleName) {
        $OutputBox.SelectionColor = [System.Drawing.Color]::Green
        $OutputBox.AppendText("  ✓ Modul '$ModuleName' bereits installiert`r`n")
    } else {
        $OutputBox.SelectionColor = [System.Drawing.Color]::Orange
        $OutputBox.AppendText("  ⚙ Installiere Modul '$ModuleName'...`r`n")
        $OutputBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            $OutputBox.SelectionColor = [System.Drawing.Color]::Green
            $OutputBox.AppendText("  ✓ Modul '$ModuleName' erfolgreich installiert`r`n")
        } catch {
            $OutputBox.SelectionColor = [System.Drawing.Color]::Red
            $OutputBox.AppendText("  ✗ FEHLER beim Installieren: $_`r`n")
            return $false
        }
    }
    
    # Modul importieren
    try {
        Import-Module $ModuleName -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking
        $OutputBox.SelectionColor = [System.Drawing.Color]::Green
        $OutputBox.AppendText("  ✓ Modul importiert`r`n")
        return $true
    } catch {
        $OutputBox.SelectionColor = [System.Drawing.Color]::Red
        $OutputBox.AppendText("  ✗ Fehler beim Importieren: $_`r`n")
        return $false
    }
}

# ============================================================================
# Console Mode Check
# ============================================================================

if ($Console) {
    # Run in console mode
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Azure & M365 Login Manager" -ForegroundColor Cyan
    Write-Host "  Version: $script:Version" -ForegroundColor Cyan
    Write-Host "  Konsolen-Modus" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Import console mode script
    . "$PSScriptRoot\Azure-M365-Intune-SharePoint-Login.ps1"
    
    # Exit after console mode
    exit
}

# ============================================================================
# GUI erstellen
# ============================================================================

Write-Host "Azure & M365 Login Manager v$script:Version wird gestartet..." -ForegroundColor Cyan

$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure & M365 Login Manager"
$form.Size = New-Object System.Drawing.Size(700, 750)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Header Label
$labelHeader = New-Object System.Windows.Forms.Label
$labelHeader.Location = New-Object System.Drawing.Point(20, 20)
$labelHeader.Size = New-Object System.Drawing.Size(640, 30)
$labelHeader.Text = "Wählen Sie die Dienste aus, bei denen Sie sich anmelden möchten:"
$labelHeader.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelHeader)

# GroupBox für Service-Auswahl
$groupBoxServices = New-Object System.Windows.Forms.GroupBox
$groupBoxServices.Location = New-Object System.Drawing.Point(20, 60)
$groupBoxServices.Size = New-Object System.Drawing.Size(640, 230)
$groupBoxServices.Text = "Dienste"
$form.Controls.Add($groupBoxServices)

# Checkbox: Azure
$chkAzure = New-Object System.Windows.Forms.CheckBox
$chkAzure.Location = New-Object System.Drawing.Point(20, 30)
$chkAzure.Size = New-Object System.Drawing.Size(280, 20)
$chkAzure.Text = "Azure (Az Module)"
$chkAzure.Checked = $true
$groupBoxServices.Controls.Add($chkAzure)

# Checkbox: Exchange Online
$chkExchange = New-Object System.Windows.Forms.CheckBox
$chkExchange.Location = New-Object System.Drawing.Point(20, 60)
$chkExchange.Size = New-Object System.Drawing.Size(280, 20)
$chkExchange.Text = "Exchange Online"
$chkExchange.Checked = $true
$groupBoxServices.Controls.Add($chkExchange)

# Checkbox: Microsoft Graph (Intune)
$chkGraph = New-Object System.Windows.Forms.CheckBox
$chkGraph.Location = New-Object System.Drawing.Point(20, 90)
$chkGraph.Size = New-Object System.Drawing.Size(280, 20)
$chkGraph.Text = "Microsoft Graph (inkl. Intune)"
$chkGraph.Checked = $true
$groupBoxServices.Controls.Add($chkGraph)

# Checkbox: Azure AD
$chkAzureAD = New-Object System.Windows.Forms.CheckBox
$chkAzureAD.Location = New-Object System.Drawing.Point(20, 120)
$chkAzureAD.Size = New-Object System.Drawing.Size(280, 20)
$chkAzureAD.Text = "Azure AD"
$chkAzureAD.Checked = $true
$groupBoxServices.Controls.Add($chkAzureAD)

# Checkbox: SharePoint Online
$chkSharePoint = New-Object System.Windows.Forms.CheckBox
$chkSharePoint.Location = New-Object System.Drawing.Point(20, 150)
$chkSharePoint.Size = New-Object System.Drawing.Size(280, 20)
$chkSharePoint.Text = "SharePoint Online"
$chkSharePoint.Checked = $true
$groupBoxServices.Controls.Add($chkSharePoint)

# SharePoint URL Label
$labelSPUrl = New-Object System.Windows.Forms.Label
$labelSPUrl.Location = New-Object System.Drawing.Point(320, 150)
$labelSPUrl.Size = New-Object System.Drawing.Size(100, 20)
$labelSPUrl.Text = "Admin-URL:"
$groupBoxServices.Controls.Add($labelSPUrl)

# SharePoint URL TextBox
$textSPUrl = New-Object System.Windows.Forms.TextBox
$textSPUrl.Location = New-Object System.Drawing.Point(420, 148)
$textSPUrl.Size = New-Object System.Drawing.Size(200, 20)
$textSPUrl.Text = "https://TENANT-admin.sharepoint.com"
$groupBoxServices.Controls.Add($textSPUrl)

# Checkbox: Microsoft Teams
$chkTeams = New-Object System.Windows.Forms.CheckBox
$chkTeams.Location = New-Object System.Drawing.Point(20, 180)
$chkTeams.Size = New-Object System.Drawing.Size(280, 20)
$chkTeams.Text = "Microsoft Teams"
$chkTeams.Checked = $true
$groupBoxServices.Controls.Add($chkTeams)

# Info Label
$labelInfo = New-Object System.Windows.Forms.Label
$labelInfo.Location = New-Object System.Drawing.Point(20, 300)
$labelInfo.Size = New-Object System.Drawing.Size(640, 40)
$labelInfo.Text = "Hinweis: Fehlende Module werden automatisch installiert. Sie werden für jeden Service zur Anmeldung aufgefordert (MFA-Unterstützung)."
$labelInfo.ForeColor = [System.Drawing.Color]::DarkBlue
$form.Controls.Add($labelInfo)

# Connect Button
$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Location = New-Object System.Drawing.Point(20, 350)
$btnConnect.Size = New-Object System.Drawing.Size(120, 35)
$btnConnect.Text = "Anmelden"
$btnConnect.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnConnect.BackColor = [System.Drawing.Color]::Green
$btnConnect.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnConnect)

# PowerShell 5.1 Console Button
$btnConsolePS5 = New-Object System.Windows.Forms.Button
$btnConsolePS5.Location = New-Object System.Drawing.Point(150, 350)
$btnConsolePS5.Size = New-Object System.Drawing.Size(100, 35)
$btnConsolePS5.Text = "PowerShell 5.1"
$btnConsolePS5.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnConsolePS5.BackColor = [System.Drawing.Color]::DodgerBlue
$btnConsolePS5.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnConsolePS5)

# PowerShell 7 Console Button
$btnConsolePS7 = New-Object System.Windows.Forms.Button
$btnConsolePS7.Location = New-Object System.Drawing.Point(260, 350)
$btnConsolePS7.Size = New-Object System.Drawing.Size(100, 35)
$btnConsolePS7.Text = "PowerShell 7"
$btnConsolePS7.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnConsolePS7.BackColor = [System.Drawing.Color]::MediumPurple
$btnConsolePS7.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnConsolePS7)

# Disconnect Button
$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Location = New-Object System.Drawing.Point(370, 350)
$btnDisconnect.Size = New-Object System.Drawing.Size(140, 35)
$btnDisconnect.Text = "Verbindungen trennen"
$btnDisconnect.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnDisconnect.BackColor = [System.Drawing.Color]::OrangeRed
$btnDisconnect.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnDisconnect)

# Clear Button
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Location = New-Object System.Drawing.Point(520, 350)
$btnClear.Size = New-Object System.Drawing.Size(70, 35)
$btnClear.Text = "Löschen"
$form.Controls.Add($btnClear)

# Close Button
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Point(600, 350)
$btnClose.Size = New-Object System.Drawing.Size(60, 35)
$btnClose.Text = "Schließen"
$form.Controls.Add($btnClose)

# Output RichTextBox
$textOutput = New-Object System.Windows.Forms.RichTextBox
$textOutput.Location = New-Object System.Drawing.Point(20, 400)
$textOutput.Size = New-Object System.Drawing.Size(640, 300)
$textOutput.Multiline = $true
$textOutput.ScrollBars = "Both"
$textOutput.Font = New-Object System.Drawing.Font("Consolas", 9)
$textOutput.ReadOnly = $true
$textOutput.BackColor = [System.Drawing.Color]::Black
$textOutput.ForeColor = [System.Drawing.Color]::LightGray
$form.Controls.Add($textOutput)

# ============================================================================
# Event Handlers
# ============================================================================

$btnConnect.Add_Click({
    $btnConnect.Enabled = $false
    $textOutput.Clear()
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
    $textOutput.AppendText("=== Azure & M365 Login Manager ===`r`n`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    # Sammle ausgewählte Services
    $services = @{
        Azure = $chkAzure.Checked
        Exchange = $chkExchange.Checked
        Graph = $chkGraph.Checked
        AzureAD = $chkAzureAD.Checked
        SharePoint = $chkSharePoint.Checked
        SharePointUrl = $textSPUrl.Text
        Teams = $chkTeams.Checked
    }
    
    # ========================================================================
    # Schritt 1: Module installieren
    # ========================================================================
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
    $textOutput.AppendText("Schritt 1: Module installieren`r`n")
    $textOutput.AppendText("=" * 50 + "`r`n")
    [System.Windows.Forms.Application]::DoEvents()
    
    $modulesToInstall = @()
    
    if ($services.Azure) {
        $modulesToInstall += "Az.Accounts", "Az.Resources"
    }
    if ($services.Exchange) {
        $modulesToInstall += "ExchangeOnlineManagement"
    }
    if ($services.Graph) {
        $modulesToInstall += "Microsoft.Graph.Authentication", "Microsoft.Graph.Users", 
                             "Microsoft.Graph.Groups", "Microsoft.Graph.DeviceManagement"
    }
    if ($services.AzureAD) {
        $modulesToInstall += "AzureAD"
    }
    if ($services.SharePoint) {
        $modulesToInstall += "Microsoft.Online.SharePoint.PowerShell"
    }
    
    if ($services.Teams) {
        $modulesToInstall += "MicrosoftTeams"
    }
    
    # Installiere Module
    $allModulesOk = $true
    foreach ($module in $modulesToInstall) {
        if (-not (Install-RequiredModule -ModuleName $module -OutputBox $textOutput)) {
            $allModulesOk = $false
        }
        $textOutput.AppendText("`r`n")
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    if (-not $allModulesOk) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("`r`n⚠ WARNUNG: Nicht alle Module konnten installiert werden!`r`n")
        $textOutput.AppendText("Einige Anmeldungen könnten fehlschlagen.`r`n`r`n")
    }
    
    # ========================================================================
    # Schritt 2: Anmeldungen durchführen
    # ========================================================================
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
    $textOutput.AppendText("`r`nSchritt 2: Anmeldungen durchführen`r`n")
    $textOutput.AppendText("=" * 50 + "`r`n")
    $textOutput.AppendText("Sie werden für jeden Service zur Anmeldung aufgefordert...`r`n`r`n")
    [System.Windows.Forms.Application]::DoEvents()
    
    Start-Sleep -Seconds 1
    
    $connections = @()
    
    # Azure
    if ($services.Azure) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
        $textOutput.AppendText("1. Azure Anmeldung...`r`n")
        $textOutput.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            Connect-AzAccount -ErrorAction Stop | Out-Null
            $context = Get-AzContext
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("   ✓ Erfolgreich angemeldet`r`n")
            $textOutput.AppendText("   Benutzer: $($context.Account.Id)`r`n")
            $textOutput.AppendText("   Tenant: $($context.Tenant.Id)`r`n`r`n")
            $connections += "✓ Azure"
        } catch {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("   ✗ Fehler: $($_.Exception.Message)`r`n`r`n")
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # Exchange Online
    if ($services.Exchange) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
        $textOutput.AppendText("2. Exchange Online Anmeldung...`r`n")
        $textOutput.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("   ✓ Erfolgreich angemeldet`r`n")
            if ($orgConfig) {
                $textOutput.AppendText("   Organisation: $($orgConfig.DisplayName)`r`n")
            }
            $textOutput.AppendText("`r`n")
            $connections += "✓ Exchange Online"
        } catch {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("   ✗ Fehler: $($_.Exception.Message)`r`n`r`n")
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # Microsoft Graph
    if ($services.Graph) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
        $textOutput.AppendText("3. Microsoft Graph Anmeldung (inkl. Intune)...`r`n")
        $textOutput.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            $graphScopes = @(
                "User.Read.All",
                "Group.Read.All",
                "Directory.Read.All",
                "Organization.Read.All",
                "DeviceManagementApps.Read.All",
                "DeviceManagementConfiguration.Read.All",
                "DeviceManagementManagedDevices.Read.All",
                "DeviceManagementServiceConfig.Read.All"
            )
            
            Connect-MgGraph -Scopes $graphScopes -ErrorAction Stop | Out-Null
            $mgContext = Get-MgContext
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("   ✓ Erfolgreich angemeldet`r`n")
            $textOutput.AppendText("   Benutzer: $($mgContext.Account)`r`n")
            $textOutput.AppendText("   Tenant: $($mgContext.TenantId)`r`n")
            $textOutput.AppendText("   Scopes: Intune-Berechtigungen inkludiert`r`n`r`n")
            $connections += "✓ Microsoft Graph"
        } catch {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("   ✗ Fehler: $($_.Exception.Message)`r`n`r`n")
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # Azure AD
    if ($services.AzureAD) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
        $textOutput.AppendText("4. Azure AD Anmeldung...`r`n")
        $textOutput.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            Connect-AzureAD -ErrorAction Stop | Out-Null
            $tenantDetail = Get-AzureADTenantDetail
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("   ✓ Erfolgreich angemeldet`r`n")
            $textOutput.AppendText("   Tenant: $($tenantDetail.DisplayName)`r`n")
            $textOutput.AppendText("   Tenant ID: $($tenantDetail.ObjectId)`r`n`r`n")
            $connections += "✓ Azure AD"
        } catch {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("   ✗ Fehler: $($_.Exception.Message)`r`n`r`n")
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # SharePoint Online
    if ($services.SharePoint) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
        $textOutput.AppendText("5. SharePoint Online Anmeldung...`r`n")
        $textOutput.AppendText("   URL: $($services.SharePointUrl)`r`n")
        $textOutput.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        # Validiere URL
        if ($services.SharePointUrl -notmatch "https://.+-admin\.sharepoint\.com") {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("   ✗ Ungültige SharePoint Admin URL!`r`n")
            $textOutput.AppendText("   Format: https://TENANT-admin.sharepoint.com`r`n`r`n")
        } else {
            try {
                # Importiere Modul explizit
                Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
                
                # Verbinde zu SharePoint
                Connect-SPOService -Url $services.SharePointUrl -ErrorAction Stop
                
                # Teste Verbindung
                $tenant = Get-SPOTenant -ErrorAction Stop
                
                $textOutput.SelectionColor = [System.Drawing.Color]::Green
                $textOutput.AppendText("   ✓ Erfolgreich angemeldet`r`n")
                if ($tenant.RootSiteUrl) {
                    $textOutput.AppendText("   Root Site: $($tenant.RootSiteUrl)`r`n")
                }
                $textOutput.AppendText("`r`n")
                $connections += "✓ SharePoint Online"
            } catch {
                $textOutput.SelectionColor = [System.Drawing.Color]::Red
                $textOutput.AppendText("   ✗ Fehler: $($_.Exception.Message)`r`n")
                $textOutput.AppendText("   Tipp: Prüfen Sie die URL und Ihre Berechtigungen`r`n`r`n")
            }
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # Microsoft Teams
    if ($services.Teams) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
        $textOutput.AppendText("6. Microsoft Teams Anmeldung...`r`n")
        $textOutput.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            # Importiere Modul explizit
            Import-Module MicrosoftTeams -ErrorAction Stop
            
            # Verbinde zu Teams
            Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
            
            # Teste Verbindung mit Get-CsTeamsClientConfiguration
            $teamsConfig = Get-CsTeamsClientConfiguration -ErrorAction Stop
            
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("   ✓ Erfolgreich angemeldet`r`n")
            $textOutput.AppendText("`r`n")
            $connections += "✓ Microsoft Teams"
        } catch {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("   ✗ Fehler: $($_.Exception.Message)`r`n")
            $textOutput.AppendText("   Tipp: Prüfen Sie Ihre Berechtigungen (Teams Administrator erforderlich)`r`n`r`n")
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # ========================================================================
    # Zusammenfassung
    # ========================================================================
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
    $textOutput.AppendText("`r`n" + "=" * 50 + "`r`n")
    $textOutput.AppendText("Zusammenfassung`r`n")
    $textOutput.AppendText("=" * 50 + "`r`n")
    
    if ($connections.Count -gt 0) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Green
        $textOutput.AppendText("`r`nErfolgreich angemeldet bei:`r`n")
        foreach ($conn in $connections) {
            $textOutput.AppendText("  $conn`r`n")
        }
    } else {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("`r`n⚠ Keine erfolgreichen Verbindungen!`r`n")
    }
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
    $textOutput.AppendText("`r`n✓ Anmeldung abgeschlossen!`r`n")
    $textOutput.AppendText("Sie können nun Ihre Scripts ausführen.`r`n")
    $textOutput.ScrollToCaret()
    
    $btnConnect.Enabled = $true
})

$btnConsolePS5.Add_Click({
    $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
    $textOutput.AppendText("`r`n" + "=" * 50 + "`r`n")
    $textOutput.AppendText("PowerShell 5.1 Konsole wird geöffnet...`r`n")
    $textOutput.AppendText("=" * 50 + "`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    # Prüfe aktive Verbindungen
    $activeConnections = @()
    
    # Azure
    try {
        $azContext = Get-AzContext -ErrorAction SilentlyContinue
        if ($azContext) {
            $activeConnections += "Azure (Subscription: $($azContext.Subscription.Name))"
        }
    } catch {}
    
    # Exchange Online
    try {
        $exoSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
        if ($exoSession) {
            $activeConnections += "Exchange Online"
        }
    } catch {}
    
    # Microsoft Graph
    try {
        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($mgContext) {
            $activeConnections += "Microsoft Graph (Tenant: $($mgContext.TenantId))"
        }
    } catch {}
    
    # Azure AD
    try {
        $aadContext = Get-AzureADCurrentSessionInfo -ErrorAction SilentlyContinue
        if ($aadContext) {
            $activeConnections += "Azure AD (Tenant: $($aadContext.TenantId))"
        }
    } catch {}
    
    # SharePoint Online
    try {
        $spoTenant = Get-SPOTenant -ErrorAction SilentlyContinue
        if ($spoTenant) {
            $activeConnections += "SharePoint Online"
        }
    } catch {
        try {
            $pnpConnection = Get-PnPConnection -ErrorAction SilentlyContinue
            if ($pnpConnection) {
                $activeConnections += "SharePoint Online (PnP)"
            }
        } catch {}
    }
    
    # Microsoft Teams
    try {
        $teamsConfig = Get-CsTeamsClientConfiguration -ErrorAction SilentlyContinue
        if ($teamsConfig) {
            $activeConnections += "Microsoft Teams"
        }
    } catch {}
    
    # Zeige aktive Verbindungen
    if ($activeConnections.Count -gt 0) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Green
        $textOutput.AppendText("`r`nAktive Verbindungen:`r`n")
        foreach ($conn in $activeConnections) {
            $textOutput.AppendText("  ✓ $conn`r`n")
        }
    } else {
        $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
        $textOutput.AppendText("`r`n⚠ Keine aktiven Verbindungen gefunden`r`n")
        $textOutput.AppendText("  Bitte melden Sie sich zuerst an!`r`n")
    }
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
    $textOutput.AppendText("`r`nÖffne PowerShell 5.1 Konsole...`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    # Erstelle temporäres Profil-Script
    $tempProfile = [System.IO.Path]::GetTempFileName() + ".ps1"
    
    $profileContent = @"
# PowerShell 5.1 Konsole mit Azure & M365 Verbindungen
`$Host.UI.RawUI.WindowTitle = "PowerShell 5.1 - Azure & M365"
`$Host.UI.RawUI.BackgroundColor = "DarkBlue"
`$Host.UI.RawUI.ForegroundColor = "White"
Clear-Host

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PowerShell 5.1 Konsole" -ForegroundColor Cyan
Write-Host "  Aktive Verbindungen" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Version: `$(`$PSVersionTable.PSVersion)" -ForegroundColor Gray
Write-Host "Edition: `$(`$PSVersionTable.PSEdition)" -ForegroundColor Gray
Write-Host ""

"@

    # Füge Verbindungsinformationen hinzu
    if ($activeConnections.Count -gt 0) {
        $profileContent += "Write-Host 'Aktive Verbindungen:' -ForegroundColor Green`r`n"
        foreach ($conn in $activeConnections) {
            $profileContent += "Write-Host '  ✓ $conn' -ForegroundColor White`r`n"
        }
    } else {
        $profileContent += "Write-Host '⚠ Keine aktiven Verbindungen' -ForegroundColor Yellow`r`n"
        $profileContent += "Write-Host '  Bitte melden Sie sich zuerst in der GUI an!' -ForegroundColor Yellow`r`n"
    }
    
    $profileContent += @"

Write-Host ""
Write-Host "Module werden geladen..." -ForegroundColor Yellow

# Lade alle installierten Module
`$modulesToLoad = @(
    "Az.Accounts",
    "Az.Resources",
    "ExchangeOnlineManagement",
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.DeviceManagement",
    "AzureAD",
    "Microsoft.Online.SharePoint.PowerShell"
)

foreach (`$module in `$modulesToLoad) {
    if (Get-Module -ListAvailable -Name `$module -ErrorAction SilentlyContinue) {
        try {
            Import-Module `$module -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            Write-Host "  ✓ `$module" -ForegroundColor Green
        } catch {
            Write-Host "  ⚠ `$module (Fehler beim Laden)" -ForegroundColor Yellow
        }
    }
}

Write-Host ""
Write-Host "Nützliche Befehle:" -ForegroundColor Yellow
Write-Host "  Get-AzContext                  # Azure Kontext" -ForegroundColor Gray
Write-Host "  Get-MgContext                  # Graph Kontext" -ForegroundColor Gray
Write-Host "  Get-SPOTenant                  # SharePoint Tenant" -ForegroundColor Gray
Write-Host "  Get-Mailbox -ResultSize 10     # Exchange Mailboxen" -ForegroundColor Gray
Write-Host ""
Write-Host "SharePoint Online Module (nur PS 5.1):" -ForegroundColor Cyan
Write-Host "  Get-SPOTenant                  # ✓ Funktioniert" -ForegroundColor Green
Write-Host "  Get-SPOSite                    # ✓ Funktioniert" -ForegroundColor Green
Write-Host ""
Write-Host "Hinweis: PowerShell 5.1 ist für SharePoint Online empfohlen!" -ForegroundColor Yellow
Write-Host ""
"@

    # Speichere Profil
    $profileContent | Out-File -FilePath $tempProfile -Encoding UTF8
    
    # Starte PowerShell 5.1
    try {
        $ps5Path = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        
        if (Test-Path $ps5Path) {
            Start-Process $ps5Path -ArgumentList "-NoExit", "-NoLogo", "-File", $tempProfile
            
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("`r`n✓ PowerShell 5.1 Konsole geöffnet!`r`n")
            $textOutput.AppendText("`r`nVersion: Windows PowerShell 5.1`r`n")
            $textOutput.AppendText("Alle Verbindungen verfügbar.`r`n")
        } else {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("`r`n✗ PowerShell 5.1 nicht gefunden!`r`n")
            $textOutput.AppendText("Pfad: $ps5Path`r`n")
        }
        
        # Cleanup
        Start-Job -ScriptBlock {
            param($file)
            Start-Sleep -Seconds 5
            if (Test-Path $file) { Remove-Item $file -Force -ErrorAction SilentlyContinue }
        } -ArgumentList $tempProfile | Out-Null
        
    } catch {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("`r`n✗ Fehler: $($_.Exception.Message)`r`n")
    }
    
    $textOutput.ScrollToCaret()
})

$btnConsolePS7.Add_Click({
    $textOutput.SelectionColor = [System.Drawing.Color]::Magenta
    $textOutput.AppendText("`r`n" + "=" * 50 + "`r`n")
    $textOutput.AppendText("PowerShell 7 Konsole wird geöffnet...`r`n")
    $textOutput.AppendText("=" * 50 + "`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    # Prüfe aktive Verbindungen
    $activeConnections = @()
    
    # Azure
    try {
        $azContext = Get-AzContext -ErrorAction SilentlyContinue
        if ($azContext) {
            $activeConnections += "Azure (Subscription: $($azContext.Subscription.Name))"
        }
    } catch {}
    
    # Exchange Online
    try {
        $exoSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
        if ($exoSession) {
            $activeConnections += "Exchange Online"
        }
    } catch {}
    
    # Microsoft Graph
    try {
        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($mgContext) {
            $activeConnections += "Microsoft Graph (Tenant: $($mgContext.TenantId))"
        }
    } catch {}
    
    # Azure AD
    try {
        $aadContext = Get-AzureADCurrentSessionInfo -ErrorAction SilentlyContinue
        if ($aadContext) {
            $activeConnections += "Azure AD (Tenant: $($aadContext.TenantId))"
        }
    } catch {}
    
    # SharePoint Online
    try {
        $spoTenant = Get-SPOTenant -ErrorAction SilentlyContinue
        if ($spoTenant) {
            $activeConnections += "SharePoint Online"
        }
    } catch {
        try {
            $pnpConnection = Get-PnPConnection -ErrorAction SilentlyContinue
            if ($pnpConnection) {
                $activeConnections += "SharePoint Online (PnP)"
            }
        } catch {}
    }
    
    # Zeige aktive Verbindungen
    if ($activeConnections.Count -gt 0) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Green
        $textOutput.AppendText("`r`nAktive Verbindungen:`r`n")
        foreach ($conn in $activeConnections) {
            $textOutput.AppendText("  ✓ $conn`r`n")
        }
    } else {
        $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
        $textOutput.AppendText("`r`n⚠ Keine aktiven Verbindungen gefunden`r`n")
        $textOutput.AppendText("  Bitte melden Sie sich zuerst an!`r`n")
    }
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Magenta
    $textOutput.AppendText("`r`nÖffne PowerShell 7 Konsole...`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    # Erstelle temporäres Profil-Script
    $tempProfile = [System.IO.Path]::GetTempFileName() + ".ps1"
    
    $profileContent = @"
# PowerShell 7 Konsole mit Azure & M365 Verbindungen
`$Host.UI.RawUI.WindowTitle = "PowerShell 7 - Azure & M365"
Clear-Host

Write-Host ""
Write-Host "========================================" -ForegroundColor Magenta
Write-Host "  PowerShell 7 Konsole" -ForegroundColor Magenta
Write-Host "  Aktive Verbindungen" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""
Write-Host "Version: `$(`$PSVersionTable.PSVersion)" -ForegroundColor Gray
Write-Host "Edition: `$(`$PSVersionTable.PSEdition)" -ForegroundColor Gray
Write-Host ""

"@

    # Füge Verbindungsinformationen hinzu
    if ($activeConnections.Count -gt 0) {
        $profileContent += "Write-Host 'Aktive Verbindungen:' -ForegroundColor Green`r`n"
        foreach ($conn in $activeConnections) {
            $profileContent += "Write-Host '  ✓ $conn' -ForegroundColor White`r`n"
        }
    } else {
        $profileContent += "Write-Host '⚠ Keine aktiven Verbindungen' -ForegroundColor Yellow`r`n"
        $profileContent += "Write-Host '  Bitte melden Sie sich zuerst in der GUI an!' -ForegroundColor Yellow`r`n"
    }
    
    $profileContent += @"

Write-Host ""
Write-Host "Module werden geladen..." -ForegroundColor Yellow

# Lade alle installierten Module (außer SharePoint SPO - nicht kompatibel mit PS7)
`$modulesToLoad = @(
    "Az.Accounts",
    "Az.Resources",
    "ExchangeOnlineManagement",
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.DeviceManagement",
    "PnP.PowerShell"  # Verwende PnP statt SPO für PS7
)

foreach (`$module in `$modulesToLoad) {
    if (Get-Module -ListAvailable -Name `$module -ErrorAction SilentlyContinue) {
        try {
            Import-Module `$module -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            Write-Host "  ✓ `$module" -ForegroundColor Green
        } catch {
            Write-Host "  ⚠ `$module (Fehler beim Laden)" -ForegroundColor Yellow
        }
    }
}

Write-Host ""
Write-Host "Nützliche Befehle:" -ForegroundColor Yellow
Write-Host "  Get-AzContext                  # Azure Kontext" -ForegroundColor Gray
Write-Host "  Get-MgContext                  # Graph Kontext" -ForegroundColor Gray
Write-Host "  Get-Mailbox -ResultSize 10     # Exchange Mailboxen" -ForegroundColor Gray
Write-Host ""
Write-Host "SharePoint Online Kompatibilität:" -ForegroundColor Cyan
Write-Host "  Get-SPOTenant                  # ⚠ Eingeschränkt in PS7" -ForegroundColor Yellow
Write-Host "  PnP.PowerShell                 # ✓ Empfohlen für PS7" -ForegroundColor Green
Write-Host ""
Write-Host "Hinweis: Für SharePoint Online verwenden Sie PowerShell 5.1!" -ForegroundColor Yellow
Write-Host "Moderne Module (Az, Graph) funktionieren perfekt in PS7." -ForegroundColor Green
Write-Host ""
"@

    # Speichere Profil
    $profileContent | Out-File -FilePath $tempProfile -Encoding UTF8
    
    # Starte PowerShell 7
    try {
        # Suche nach PowerShell 7 Installation
        $ps7Paths = @(
            "C:\Program Files\PowerShell\7\pwsh.exe",
            "C:\Program Files\PowerShell\7-preview\pwsh.exe",
            "$env:LOCALAPPDATA\Microsoft\PowerShell\7\pwsh.exe"
        )
        
        $ps7Path = $null
        foreach ($path in $ps7Paths) {
            if (Test-Path $path) {
                $ps7Path = $path
                break
            }
        }
        
        # Alternativ: Suche in PATH
        if (-not $ps7Path) {
            $pwshCmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            if ($pwshCmd) {
                $ps7Path = $pwshCmd.Source
            }
        }
        
        if ($ps7Path) {
            Start-Process $ps7Path -ArgumentList "-NoExit", "-NoLogo", "-File", $tempProfile
            
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("`r`n✓ PowerShell 7 Konsole geöffnet!`r`n")
            $textOutput.AppendText("`r`nVersion: PowerShell 7+`r`n")
            $textOutput.AppendText("Pfad: $ps7Path`r`n")
            $textOutput.AppendText("Alle Verbindungen verfügbar.`r`n")
        } else {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("`r`n✗ PowerShell 7 nicht gefunden!`r`n")
            $textOutput.AppendText("`r`nBitte installieren Sie PowerShell 7:`r`n")
            $textOutput.AppendText("  winget install Microsoft.PowerShell`r`n")
            $textOutput.AppendText("  oder herunterladen von: https://aka.ms/powershell`r`n")
        }
        
        # Cleanup
        Start-Job -ScriptBlock {
            param($file)
            Start-Sleep -Seconds 5
            if (Test-Path $file) { Remove-Item $file -Force -ErrorAction SilentlyContinue }
        } -ArgumentList $tempProfile | Out-Null
        
    } catch {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("`r`n✗ Fehler: $($_.Exception.Message)`r`n")
    }
    
    $textOutput.ScrollToCaret()
})

$btnDisconnect.Add_Click({
    $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
    $textOutput.AppendText("`r`n" + "=" * 50 + "`r`n")
    $textOutput.AppendText("PowerShell Konsole wird geöffnet...`r`n")
    $textOutput.AppendText("=" * 50 + "`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    # Prüfe aktive Verbindungen
    $activeConnections = @()
    
    # Azure
    try {
        $azContext = Get-AzContext -ErrorAction SilentlyContinue
        if ($azContext) {
            $activeConnections += "Azure (Subscription: $($azContext.Subscription.Name))"
        }
    } catch {}
    
    # Exchange Online
    try {
        $exoSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
        if ($exoSession) {
            $activeConnections += "Exchange Online"
        }
    } catch {}
    
    # Microsoft Graph
    try {
        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($mgContext) {
            $activeConnections += "Microsoft Graph (Tenant: $($mgContext.TenantId))"
        }
    } catch {}
    
    # Azure AD
    try {
        $aadContext = Get-AzureADCurrentSessionInfo -ErrorAction SilentlyContinue
        if ($aadContext) {
            $activeConnections += "Azure AD (Tenant: $($aadContext.TenantId))"
        }
    } catch {}
    
    # SharePoint Online
    try {
        $spoTenant = Get-SPOTenant -ErrorAction SilentlyContinue
        if ($spoTenant) {
            $activeConnections += "SharePoint Online"
        }
    } catch {
        try {
            $pnpConnection = Get-PnPConnection -ErrorAction SilentlyContinue
            if ($pnpConnection) {
                $activeConnections += "SharePoint Online (PnP)"
            }
        } catch {}
    }
    
    # Zeige aktive Verbindungen
    if ($activeConnections.Count -gt 0) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Green
        $textOutput.AppendText("`r`nAktive Verbindungen:`r`n")
        foreach ($conn in $activeConnections) {
            $textOutput.AppendText("  ✓ $conn`r`n")
        }
    } else {
        $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
        $textOutput.AppendText("`r`n⚠ Keine aktiven Verbindungen gefunden`r`n")
        $textOutput.AppendText("  Bitte melden Sie sich zuerst an!`r`n")
    }
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
    $textOutput.AppendText("`r`nÖffne PowerShell Konsole mit aktuellem Kontext...`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    # Erstelle temporäres Profil-Script für die neue PowerShell-Sitzung
    $tempProfile = [System.IO.Path]::GetTempFileName() + ".ps1"
    
    $profileContent = @"
# Temporäres Profil für PowerShell-Konsole mit aktiven Verbindungen
`$Host.UI.RawUI.WindowTitle = "PowerShell - Verbunden mit Azure & M365"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PowerShell Konsole" -ForegroundColor Cyan
Write-Host "  Aktive Verbindungen" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

"@

    # Füge Verbindungsinformationen hinzu
    if ($activeConnections.Count -gt 0) {
        $profileContent += "Write-Host 'Folgende Verbindungen sind aktiv:' -ForegroundColor Green`r`n"
        foreach ($conn in $activeConnections) {
            $profileContent += "Write-Host '  ✓ $conn' -ForegroundColor White`r`n"
        }
    } else {
        $profileContent += "Write-Host '⚠ Keine aktiven Verbindungen' -ForegroundColor Yellow`r`n"
        $profileContent += "Write-Host '  Bitte melden Sie sich zuerst in der GUI an!' -ForegroundColor Yellow`r`n"
    }
    
    $profileContent += @"

Write-Host ""
Write-Host "Nützliche Befehle:" -ForegroundColor Yellow
Write-Host "  Get-AzContext                  # Azure Kontext anzeigen" -ForegroundColor Gray
Write-Host "  Get-MgContext                  # Graph Kontext anzeigen" -ForegroundColor Gray
Write-Host "  Get-AzureADCurrentSessionInfo  # Azure AD Session" -ForegroundColor Gray
Write-Host "  Get-SPOTenant                  # SharePoint Tenant" -ForegroundColor Gray
Write-Host "  Get-OrganizationConfig         # Exchange Org Config" -ForegroundColor Gray
Write-Host ""
Write-Host "Hinweis: Änderungen in dieser Konsole beeinflussen NICHT die GUI!" -ForegroundColor Yellow
Write-Host "Die Verbindungen bleiben auch nach Schließen dieser Konsole aktiv." -ForegroundColor Yellow
Write-Host ""
"@

    # Speichere Profil
    $profileContent | Out-File -FilePath $tempProfile -Encoding UTF8
    
    # Starte neue PowerShell-Konsole mit dem temporären Profil
    try {
        # Verwende Start-Process für neue Konsole im gleichen Kontext
        Start-Process powershell.exe -ArgumentList "-NoExit", "-NoLogo", "-File", $tempProfile
        
        $textOutput.SelectionColor = [System.Drawing.Color]::Green
        $textOutput.AppendText("`r`n✓ PowerShell Konsole geöffnet!`r`n")
        $textOutput.AppendText("`r`nHinweis: Die Konsole verwendet die gleichen Verbindungen.`r`n")
        $textOutput.AppendText("Alle Module und Verbindungen sind verfügbar.`r`n")
        
        # Cleanup nach kurzer Verzögerung
        Start-Job -ScriptBlock {
            param($file)
            Start-Sleep -Seconds 5
            if (Test-Path $file) {
                Remove-Item $file -Force -ErrorAction SilentlyContinue
            }
        } -ArgumentList $tempProfile | Out-Null
        
    } catch {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("`r`n✗ Fehler beim Öffnen der Konsole: $($_.Exception.Message)`r`n")
    }
    
    $textOutput.ScrollToCaret()
})

$btnDisconnect.Add_Click({
    $btnDisconnect.Enabled = $false
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
    $textOutput.AppendText("`r`n" + "=" * 50 + "`r`n")
    $textOutput.AppendText("Trenne alle Verbindungen...`r`n")
    $textOutput.AppendText("=" * 50 + "`r`n`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    $disconnected = @()
    $errors = @()
    
    # Azure trennen
    try {
        $azContext = Get-AzContext -ErrorAction SilentlyContinue
        if ($azContext) {
            $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
            $textOutput.AppendText("Azure...`r`n")
            Disconnect-AzAccount -ErrorAction Stop | Out-Null
            Clear-AzContext -Scope CurrentUser -Force -ErrorAction SilentlyContinue
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("  ✓ Azure Verbindung getrennt`r`n`r`n")
            $disconnected += "Azure"
        }
    } catch {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("  ✗ Fehler beim Trennen: $($_.Exception.Message)`r`n`r`n")
        $errors += "Azure"
    }
    [System.Windows.Forms.Application]::DoEvents()
    
    # Exchange Online trennen
    try {
        $exoSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
        if ($exoSession) {
            $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
            $textOutput.AppendText("Exchange Online...`r`n")
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("  ✓ Exchange Online Verbindung getrennt`r`n`r`n")
            $disconnected += "Exchange Online"
        }
    } catch {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("  ✗ Fehler beim Trennen: $($_.Exception.Message)`r`n`r`n")
        $errors += "Exchange Online"
    }
    [System.Windows.Forms.Application]::DoEvents()
    
    # Microsoft Graph trennen
    try {
        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($mgContext) {
            $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
            $textOutput.AppendText("Microsoft Graph...`r`n")
            Disconnect-MgGraph -ErrorAction Stop
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("  ✓ Microsoft Graph Verbindung getrennt`r`n`r`n")
            $disconnected += "Microsoft Graph"
        }
    } catch {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("  ✗ Fehler beim Trennen: $($_.Exception.Message)`r`n`r`n")
        $errors += "Microsoft Graph"
    }
    [System.Windows.Forms.Application]::DoEvents()
    
    # Azure AD trennen
    try {
        $aadContext = Get-AzureADCurrentSessionInfo -ErrorAction SilentlyContinue
        if ($aadContext) {
            $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
            $textOutput.AppendText("Azure AD...`r`n")
            Disconnect-AzureAD -ErrorAction Stop
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("  ✓ Azure AD Verbindung getrennt`r`n`r`n")
            $disconnected += "Azure AD"
        }
    } catch {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("  ✗ Fehler beim Trennen: $($_.Exception.Message)`r`n`r`n")
        $errors += "Azure AD"
    }
    [System.Windows.Forms.Application]::DoEvents()
    
    # SharePoint Online trennen
    try {
        # Versuche SPO Cmdlet
        $spoConnected = $false
        try {
            $spoTenant = Get-SPOTenant -ErrorAction SilentlyContinue
            if ($spoTenant) {
                $spoConnected = $true
            }
        } catch {
            $spoConnected = $false
        }
        
        if ($spoConnected) {
            $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
            $textOutput.AppendText("SharePoint Online...`r`n")
            Disconnect-SPOService -ErrorAction Stop
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("  ✓ SharePoint Online Verbindung getrennt`r`n`r`n")
            $disconnected += "SharePoint Online"
        }
    } catch {
        # PnP Verbindung prüfen
        try {
            $pnpConnection = Get-PnPConnection -ErrorAction SilentlyContinue
            if ($pnpConnection) {
                $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
                $textOutput.AppendText("SharePoint Online (PnP)...`r`n")
                Disconnect-PnPOnline -ErrorAction Stop
                $textOutput.SelectionColor = [System.Drawing.Color]::Green
                $textOutput.AppendText("  ✓ SharePoint Online Verbindung getrennt`r`n`r`n")
                $disconnected += "SharePoint Online (PnP)"
            }
        } catch {
            # Ignoriere wenn keine SharePoint-Verbindung besteht
        }
    }
    [System.Windows.Forms.Application]::DoEvents()
    
    # Zusammenfassung
    $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
    $textOutput.AppendText("=" * 50 + "`r`n")
    $textOutput.AppendText("Zusammenfassung`r`n")
    $textOutput.AppendText("=" * 50 + "`r`n`r`n")
    
    if ($disconnected.Count -gt 0) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Green
        $textOutput.AppendText("Getrennte Verbindungen:`r`n")
        foreach ($conn in $disconnected) {
            $textOutput.AppendText("  ✓ $conn`r`n")
        }
    } else {
        $textOutput.SelectionColor = [System.Drawing.Color]::Gray
        $textOutput.AppendText("Keine aktiven Verbindungen gefunden`r`n")
    }
    
    if ($errors.Count -gt 0) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("`r`nFehler beim Trennen:`r`n")
        foreach ($err in $errors) {
            $textOutput.AppendText("  ✗ $err`r`n")
        }
    }
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
    $textOutput.AppendText("`r`n✓ Trennvorgang abgeschlossen!`r`n")
    $textOutput.ScrollToCaret()
    
    $btnDisconnect.Enabled = $true
})

$btnClear.Add_Click({
    $textOutput.Clear()
})

$btnClose.Add_Click({
    $form.Close()
})

# Form anzeigen
[void]$form.ShowDialog()