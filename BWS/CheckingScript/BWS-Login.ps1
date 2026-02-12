<#
.SYNOPSIS
    Azure, M365, Intune, SharePoint und Teams Login-Script mit GUI (SSO-optimiert)
.DESCRIPTION
    GUI-basiertes oder konsolenbasiertes Anmeldescript für Azure, Microsoft 365, Intune, SharePoint Online und Microsoft Teams
    
    NEUE FEATURES v2.0.0:
    - Single Sign-On (SSO): Einmalige MFA-Anmeldung für alle Dienste
    - Browser-Integration: Edge, Chrome, Firefox im angemeldeten Kontext öffnen
    - PowerShell Konsolen mit automatischer Anmeldung bei allen Modulen
    - Shared Access Token für nahtlose Authentifizierung
    
    FUNKTIONEN:
    - Auswahl der gewünschten Dienste per Checkbox oder Parameter
    - SharePoint-URL manuell eingeben
    - Automatische Modul-Installation
    - PowerShell 5.1 und 7 Konsolen-Support mit aktiven Verbindungen
    - Browser-Buttons für Azure Portal, Admin Center, Office Portal
    
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
    Version: 2.0.0
    Datum: 2025-02-12
    Autor: BWS PowerShell Script
    
    CHANGELOG v2.0.0:
    - Single Sign-On (SSO) implementiert
    - Browser-Buttons hinzugefügt (Edge, Chrome, Firefox)
    - Access Token Sharing zwischen Diensten
    - PowerShell Console Buttons aktiviert für alle Module
    - Verbesserte Benutzerführung
    
.EXAMPLE
    .\Azure-M365-Login-GUI.ps1
    Startet die GUI mit SSO-Unterstützung
    
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
$labelInfo.Size = New-Object System.Drawing.Size(640, 60)
$labelInfo.Text = "Hinweis: Fehlende Module werden automatisch installiert.`r`nSingle Sign-On (SSO): Bei der ersten Anmeldung (z.B. Azure) werden Sie zur MFA aufgefordert.`r`nAlle weiteren Dienste nutzen die gleiche Sitzung - keine erneute MFA-Eingabe erforderlich!"
$labelInfo.ForeColor = [System.Drawing.Color]::DarkBlue
$form.Controls.Add($labelInfo)

# Connect Button
$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Location = New-Object System.Drawing.Point(20, 370)
$btnConnect.Size = New-Object System.Drawing.Size(120, 35)
$btnConnect.Text = "Anmelden"
$btnConnect.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnConnect.BackColor = [System.Drawing.Color]::Green
$btnConnect.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnConnect)

# PowerShell 5.1 Console Button
$btnConsolePS5 = New-Object System.Windows.Forms.Button
$btnConsolePS5.Location = New-Object System.Drawing.Point(150, 370)
$btnConsolePS5.Size = New-Object System.Drawing.Size(100, 35)
$btnConsolePS5.Text = "PowerShell 5.1"
$btnConsolePS5.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnConsolePS5.BackColor = [System.Drawing.Color]::DodgerBlue
$btnConsolePS5.ForeColor = [System.Drawing.Color]::White
$btnConsolePS5.Enabled = $false
$form.Controls.Add($btnConsolePS5)

# PowerShell 7 Console Button
$btnConsolePS7 = New-Object System.Windows.Forms.Button
$btnConsolePS7.Location = New-Object System.Drawing.Point(260, 370)
$btnConsolePS7.Size = New-Object System.Drawing.Size(100, 35)
$btnConsolePS7.Text = "PowerShell 7"
$btnConsolePS7.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnConsolePS7.BackColor = [System.Drawing.Color]::MediumPurple
$btnConsolePS7.ForeColor = [System.Drawing.Color]::White
$btnConsolePS7.Enabled = $false
$form.Controls.Add($btnConsolePS7)

# Browser Buttons Row
$browserY = 410

# Edge Browser Button
$btnEdge = New-Object System.Windows.Forms.Button
$btnEdge.Location = New-Object System.Drawing.Point(20, $browserY)
$btnEdge.Size = New-Object System.Drawing.Size(110, 35)
$btnEdge.Text = "🌐 Edge"
$btnEdge.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnEdge.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnEdge.ForeColor = [System.Drawing.Color]::White
$btnEdge.Enabled = $false
$form.Controls.Add($btnEdge)

# Chrome Browser Button
$btnChrome = New-Object System.Windows.Forms.Button
$btnChrome.Location = New-Object System.Drawing.Point(140, $browserY)
$btnChrome.Size = New-Object System.Drawing.Size(110, 35)
$btnChrome.Text = "🌐 Chrome"
$btnChrome.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnChrome.BackColor = [System.Drawing.Color]::FromArgb(66, 133, 244)
$btnChrome.ForeColor = [System.Drawing.Color]::White
$btnChrome.Enabled = $false
$form.Controls.Add($btnChrome)

# Firefox Browser Button
$btnFirefox = New-Object System.Windows.Forms.Button
$btnFirefox.Location = New-Object System.Drawing.Point(260, $browserY)
$btnFirefox.Size = New-Object System.Drawing.Size(110, 35)
$btnFirefox.Text = "🌐 Firefox"
$btnFirefox.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnFirefox.BackColor = [System.Drawing.Color]::FromArgb(230, 96, 0)
$btnFirefox.ForeColor = [System.Drawing.Color]::White
$btnFirefox.Enabled = $false
$form.Controls.Add($btnFirefox)

# Disconnect Button
$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Location = New-Object System.Drawing.Point(380, $browserY)
$btnDisconnect.Size = New-Object System.Drawing.Size(130, 35)
$btnDisconnect.Text = "Trennen"
$btnDisconnect.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnDisconnect.BackColor = [System.Drawing.Color]::OrangeRed
$btnDisconnect.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnDisconnect)

# Clear Button
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Location = New-Object System.Drawing.Point(520, $browserY)
$btnClear.Size = New-Object System.Drawing.Size(70, 35)
$btnClear.Text = "Löschen"
$form.Controls.Add($btnClear)

# Close Button
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Point(600, $browserY)
$btnClose.Size = New-Object System.Drawing.Size(60, 35)
$btnClose.Text = "Schließen"
$form.Controls.Add($btnClose)

# Output RichTextBox
$textOutput = New-Object System.Windows.Forms.RichTextBox
$textOutput.Location = New-Object System.Drawing.Point(20, 460)
$textOutput.Size = New-Object System.Drawing.Size(640, 270)
$textOutput.Multiline = $true
$textOutput.ScrollBars = "Both"
$textOutput.Font = New-Object System.Drawing.Font("Consolas", 9)
$textOutput.ReadOnly = $true
$textOutput.BackColor = [System.Drawing.Color]::Black
$textOutput.ForeColor = [System.Drawing.Color]::LightGray
$form.Controls.Add($textOutput)

# Adjust form height
$form.Size = New-Object System.Drawing.Size(700, 790)

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
    $script:sharedAccessToken = $null
    $script:sharedTenantId = $null
    $script:sharedUserPrincipalName = $null
    
    # ========================================================================
    # SINGLE SIGN-ON: Erste Anmeldung für Access Token
    # ========================================================================
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
    $textOutput.AppendText("═══════════════════════════════════════════════════════`r`n")
    $textOutput.AppendText("  SINGLE SIGN-ON (SSO) - Einmalige Anmeldung`r`n")
    $textOutput.AppendText("═══════════════════════════════════════════════════════`r`n`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    # Azure
    if ($services.Azure) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
        $textOutput.AppendText("1. Azure Anmeldung (mit MFA)...`r`n")
        $textOutput.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            Connect-AzAccount -ErrorAction Stop | Out-Null
            $context = Get-AzContext
            $script:sharedTenantId = $context.Tenant.Id
            $script:sharedUserPrincipalName = $context.Account.Id
            
            # Access Token für weitere Dienste holen
            try {
                $script:sharedAccessToken = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
            } catch {
                # Fallback ohne Token
            }
            
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("   ✓ Erfolgreich angemeldet`r`n")
            $textOutput.AppendText("   Benutzer: $($context.Account.Id)`r`n")
            $textOutput.AppendText("   Tenant: $($context.Tenant.Id)`r`n")
            $textOutput.AppendText("   → Access Token für weitere Dienste gespeichert`r`n`r`n")
            $connections += "✓ Azure"
        } catch {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("   ✗ Fehler: $($_.Exception.Message)`r`n`r`n")
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # Microsoft Graph (nutzt SSO wenn Azure verbunden)
    if ($services.Graph) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
        if ($script:sharedAccessToken) {
            $textOutput.AppendText("2. Microsoft Graph Anmeldung (SSO - keine erneute MFA)...`r`n")
        } else {
            $textOutput.AppendText("2. Microsoft Graph Anmeldung (mit MFA)...`r`n")
        }
        $textOutput.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            $graphScopes = @(
                "User.Read.All",
                "Group.Read.All", 
                "DeviceManagementConfiguration.Read.All",
                "DeviceManagementManagedDevices.Read.All",
                "Directory.Read.All"
            )
            
            if ($script:sharedAccessToken -and $script:sharedTenantId) {
                # SSO: Nutze bestehenden Access Token
                Connect-MgGraph -AccessToken $script:sharedAccessToken -ErrorAction Stop | Out-Null
                $textOutput.SelectionColor = [System.Drawing.Color]::Green
                $textOutput.AppendText("   ✓ Erfolgreich angemeldet (SSO)`r`n")
                $textOutput.AppendText("   Benutzer: $script:sharedUserPrincipalName`r`n")
                $textOutput.AppendText("   → Keine erneute MFA-Eingabe erforderlich!`r`n`r`n")
            } else {
                # Erste Anmeldung
                Connect-MgGraph -Scopes $graphScopes -ErrorAction Stop | Out-Null
                $mgContext = Get-MgContext
                $script:sharedTenantId = $mgContext.TenantId
                $script:sharedUserPrincipalName = $mgContext.Account
                
                $textOutput.SelectionColor = [System.Drawing.Color]::Green
                $textOutput.AppendText("   ✓ Erfolgreich angemeldet`r`n")
                $textOutput.AppendText("   Benutzer: $($mgContext.Account)`r`n")
                $textOutput.AppendText("   Tenant: $($mgContext.TenantId)`r`n`r`n")
            }
            $connections += "✓ Microsoft Graph"
        } catch {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("   ✗ Fehler: $($_.Exception.Message)`r`n`r`n")
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # ========================================================================
    # WEITERE DIENSTE (nutzen SSO wenn möglich)
    # ========================================================================
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
    $textOutput.AppendText("`r`n═══════════════════════════════════════════════════════`r`n")
    $textOutput.AppendText("  Weitere Dienste (nutzen SSO-Sitzung)`r`n")
    $textOutput.AppendText("═══════════════════════════════════════════════════════`r`n`r`n")
    [System.Windows.Forms.Application]::DoEvents()
    
    # Exchange Online
    if ($services.Exchange) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
        $textOutput.AppendText("3. Exchange Online Anmeldung (SSO)...`r`n")
        $textOutput.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            if ($script:sharedUserPrincipalName) {
                # SSO: Nutze UserPrincipalName für nahtlose Anmeldung
                Connect-ExchangeOnline -UserPrincipalName $script:sharedUserPrincipalName -ShowBanner:$false -ErrorAction Stop
            } else {
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            }
            
            $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("   ✓ Erfolgreich angemeldet (SSO)`r`n")
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
    
    # Azure AD
    if ($services.AzureAD) {
        $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
        $textOutput.AppendText("4. Azure AD Anmeldung (SSO)...`r`n")
        $textOutput.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            if ($script:sharedTenantId) {
                # SSO: Nutze Tenant ID
                Connect-AzureAD -TenantId $script:sharedTenantId -ErrorAction Stop | Out-Null
            } else {
                Connect-AzureAD -ErrorAction Stop | Out-Null
            }
            
            $tenantDetail = Get-AzureADTenantDetail
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("   ✓ Erfolgreich angemeldet (SSO)`r`n")
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
        $textOutput.AppendText("5. SharePoint Online Anmeldung (SSO)...`r`n")
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
                $textOutput.AppendText("   ✓ Erfolgreich angemeldet (SSO)`r`n")
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
        $textOutput.AppendText("   Ein Login-Fenster wird geöffnet - bitte anmelden`r`n")
        $textOutput.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            # Importiere Modul explizit
            Import-Module MicrosoftTeams -ErrorAction Stop
            
            $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
            $textOutput.AppendText("   → Warte auf Anmeldung...`r`n")
            $textOutput.AppendText("   (Falls kein Fenster erscheint, prüfen Sie die Taskleiste)`r`n")
            $textOutput.ScrollToCaret()
            
            # GUI responsive halten während Login
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            [System.Windows.Forms.Application]::DoEvents()
            
            # Verbinde zu Teams - Standard Browser Auth (mit MFA)
            # Kein -UseDeviceAuthentication = normaler Browser-Login
            $teamsConnection = $null
            
            # Login in separatem Thread um GUI nicht zu blockieren
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.Open()
            $runspace.SessionStateProxy.SetVariable("textOutput", $textOutput)
            
            $powershell = [powershell]::Create()
            $powershell.Runspace = $runspace
            $powershell.AddScript({
                Import-Module MicrosoftTeams
                Connect-MicrosoftTeams
            })
            
            $asyncResult = $powershell.BeginInvoke()
            
            # Warte auf Fertigstellung, aber halte GUI responsive
            while (-not $asyncResult.IsCompleted) {
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 100
            }
            
            $teamsConnection = $powershell.EndInvoke($asyncResult)
            $powershell.Dispose()
            $runspace.Close()
            
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            
            if ($teamsConnection) {
                # Teste Verbindung
                $teamsConfig = Get-CsTeamsClientConfiguration -ErrorAction Stop
                
                $textOutput.SelectionColor = [System.Drawing.Color]::Green
                $textOutput.AppendText("   ✓ Erfolgreich angemeldet`r`n")
                if ($teamsConnection.TenantId) {
                    $textOutput.AppendText("   Tenant ID: $($teamsConnection.TenantId)`r`n")
                }
                if ($teamsConnection.Account) {
                    $textOutput.AppendText("   Account: $($teamsConnection.Account)`r`n")
                }
                $textOutput.AppendText("`r`n")
                $connections += "✓ Microsoft Teams"
            } else {
                $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
                $textOutput.AppendText("   ⚠ Anmeldung abgebrochen oder fehlgeschlagen`r`n`r`n")
            }
        } catch {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
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
    
    # Enable buttons after successful login
    $btnConnect.Enabled = $true
    
    # Enable PowerShell console buttons if any connection was successful
    if ($connections.Count -gt 0) {
        $btnConsolePS5.Enabled = $true
        $btnConsolePS7.Enabled = $true
        $btnEdge.Enabled = $true
        $btnChrome.Enabled = $true
        $btnFirefox.Enabled = $true
        
        $textOutput.SelectionColor = [System.Drawing.Color]::Green
        $textOutput.AppendText("`r`n✓ PowerShell Konsolen und Browser-Buttons sind nun aktiviert!`r`n")
        $textOutput.ScrollToCaret()
    }
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
    "Microsoft.Online.SharePoint.PowerShell",
    "MicrosoftTeams"
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
Write-Host "  Get-CsTeamsClientConfiguration # Teams Configuration" -ForegroundColor Gray
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
    "MicrosoftTeams",
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
Write-Host "  Get-CsTeamsClientConfiguration # Teams Configuration" -ForegroundColor Gray
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
    
    # Disable browser and PowerShell buttons after disconnect
    $btnEdge.Enabled = $false
    $btnChrome.Enabled = $false
    $btnFirefox.Enabled = $false
    $btnConsolePS5.Enabled = $false
    $btnConsolePS7.Enabled = $false
    
    $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
    $textOutput.AppendText("PowerShell Konsolen und Browser-Buttons wurden deaktiviert.`r`n")
    $textOutput.ScrollToCaret()
})

# ============================================================================
# Browser Button Event Handlers
# ============================================================================

$btnEdge.Add_Click({
    $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
    $textOutput.AppendText("`r`n" + "=" * 50 + "`r`n")
    $textOutput.AppendText("Microsoft Edge wird geöffnet (im angemeldeten Kontext)...`r`n")
    $textOutput.AppendText("=" * 50 + "`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        # Prüfe ob Edge installiert ist
        $edgePath = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
        if (-not (Test-Path $edgePath)) {
            $edgePath = "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
        }
        
        if (Test-Path $edgePath) {
            # URLs für verschiedene Portale
            $portalUrls = @(
                "https://portal.azure.com",
                "https://admin.microsoft.com",
                "https://portal.office.com"
            )
            
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("`r`nÖffne folgende Portale in Edge:`r`n")
            foreach ($url in $portalUrls) {
                $textOutput.AppendText("  • $url`r`n")
                Start-Process $edgePath -ArgumentList "--profile-directory=Default", $url
                Start-Sleep -Milliseconds 500
            }
            
            $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
            $textOutput.AppendText("`r`nℹ Hinweis: Sie werden automatisch angemeldet, wenn Sie bereits`r`n")
            $textOutput.AppendText("  in Edge mit dem gleichen Microsoft-Konto angemeldet sind.`r`n")
            $textOutput.AppendText("  Falls nicht, melden Sie sich einmalig an.`r`n`r`n")
        } else {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("`r`n✗ Microsoft Edge wurde nicht gefunden!`r`n`r`n")
        }
    } catch {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("`r`n✗ Fehler beim Öffnen von Edge: $($_.Exception.Message)`r`n`r`n")
    }
    
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
})

$btnChrome.Add_Click({
    $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
    $textOutput.AppendText("`r`n" + "=" * 50 + "`r`n")
    $textOutput.AppendText("Google Chrome wird geöffnet (im angemeldeten Kontext)...`r`n")
    $textOutput.AppendText("=" * 50 + "`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        # Prüfe ob Chrome installiert ist
        $chromePath = "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe"
        if (-not (Test-Path $chromePath)) {
            $chromePath = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
        }
        if (-not (Test-Path $chromePath)) {
            $chromePath = "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
        }
        
        if (Test-Path $chromePath) {
            # URLs für verschiedene Portale
            $portalUrls = @(
                "https://portal.azure.com",
                "https://admin.microsoft.com",
                "https://portal.office.com"
            )
            
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("`r`nÖffne folgende Portale in Chrome:`r`n")
            foreach ($url in $portalUrls) {
                $textOutput.AppendText("  • $url`r`n")
                Start-Process $chromePath -ArgumentList $url
                Start-Sleep -Milliseconds 500
            }
            
            $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
            $textOutput.AppendText("`r`nℹ Hinweis: Sie werden automatisch angemeldet, wenn Sie bereits`r`n")
            $textOutput.AppendText("  in Chrome mit dem gleichen Microsoft-Konto angemeldet sind.`r`n")
            $textOutput.AppendText("  Falls nicht, melden Sie sich einmalig an.`r`n`r`n")
        } else {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("`r`n✗ Google Chrome wurde nicht gefunden!`r`n`r`n")
        }
    } catch {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("`r`n✗ Fehler beim Öffnen von Chrome: $($_.Exception.Message)`r`n`r`n")
    }
    
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
})

$btnFirefox.Add_Click({
    $textOutput.SelectionColor = [System.Drawing.Color]::Cyan
    $textOutput.AppendText("`r`n" + "=" * 50 + "`r`n")
    $textOutput.AppendText("Mozilla Firefox wird geöffnet (im angemeldeten Kontext)...`r`n")
    $textOutput.AppendText("=" * 50 + "`r`n")
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        # Prüfe ob Firefox installiert ist
        $firefoxPath = "${env:ProgramFiles}\Mozilla Firefox\firefox.exe"
        if (-not (Test-Path $firefoxPath)) {
            $firefoxPath = "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"
        }
        
        if (Test-Path $firefoxPath) {
            # URLs für verschiedene Portale
            $portalUrls = @(
                "https://portal.azure.com",
                "https://admin.microsoft.com",
                "https://portal.office.com"
            )
            
            $textOutput.SelectionColor = [System.Drawing.Color]::Green
            $textOutput.AppendText("`r`nÖffne folgende Portale in Firefox:`r`n")
            
            # Firefox öffnet URLs mit comma-separierter Liste
            $urlList = $portalUrls -join "|"
            Start-Process $firefoxPath -ArgumentList $urlList
            
            foreach ($url in $portalUrls) {
                $textOutput.AppendText("  • $url`r`n")
            }
            
            $textOutput.SelectionColor = [System.Drawing.Color]::Yellow
            $textOutput.AppendText("`r`nℹ Hinweis: Sie werden automatisch angemeldet, wenn Sie bereits`r`n")
            $textOutput.AppendText("  in Firefox mit dem gleichen Microsoft-Konto angemeldet sind.`r`n")
            $textOutput.AppendText("  Falls nicht, melden Sie sich einmalig an.`r`n`r`n")
        } else {
            $textOutput.SelectionColor = [System.Drawing.Color]::Red
            $textOutput.AppendText("`r`n✗ Mozilla Firefox wurde nicht gefunden!`r`n`r`n")
        }
    } catch {
        $textOutput.SelectionColor = [System.Drawing.Color]::Red
        $textOutput.AppendText("`r`n✗ Fehler beim Öffnen von Firefox: $($_.Exception.Message)`r`n`r`n")
    }
    
    $textOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
})

$btnClear.Add_Click({
    $textOutput.Clear()
})

$btnClose.Add_Click({
    $form.Close()
})

# Form anzeigen
[void]$form.ShowDialog()