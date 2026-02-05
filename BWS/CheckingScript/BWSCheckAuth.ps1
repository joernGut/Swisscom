<# 
.BWSCheckAuth.ps1
Zweck:
- Stellt Verbindungen zu Microsoft Graph, Exchange Online, Microsoft Teams und SharePoint Online (SPO) her.
- Kann unabhängig ausgeführt werden (nur Authentifizierung).
- Kann vom Main Script dot-sourced werden (liefert Funktionen Connect-/Disconnect-).

Unterstützte Modi:
- Interactive: Benutzer-Login per Browser
- DeviceCode: Gerätecode-Login (falls vom Modul unterstützt)
- AppOnlyCertificate: Enterprise App / Zertifikat (ClientId + Thumbprint + TenantId/Org)

Hinweis:
- Dieses Script installiert Module NICHT automatisch, kann aber optional mit -AllowInstall versuchen zu installieren.
#>

[CmdletBinding()]
param(
    [ValidateSet('Interactive','DeviceCode','AppOnlyCertificate')]
    [string]$Mode = 'Interactive',

    # Microsoft Entra TenantId (GUID) – empfohlen
    [string]$TenantId,

    # Für Interactive: UPN optional (nur für Anzeige/Logik)
    [string]$AccountId,

    # Für AppOnlyCertificate:
    [string]$ClientId,
    [string]$CertificateThumbprint,

    # Für Exchange Online AppOnly: i.d.R. Primary Domain (contoso.onmicrosoft.com oder contoso.com)
    [string]$Organization,

    # Graph Scopes (bei Interactive/DeviceCode)
    [string[]]$Scopes = @('User.Read'),

    # Service-Schalter
    [switch]$ConnectGraph = $true,
    [switch]$ConnectExchange,
    [switch]$ConnectTeams,
    [switch]$ConnectSharePoint,

    # SPO Admin URL, z.B. https://<tenant>-admin.sharepoint.com
    [string]$SharePointAdminUrl,

    # Optional: versucht Module zu installieren/aktualisieren
    [switch]$AllowInstall,

    # Optional: keine Banner
    [switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------
# Helpers: Module Management
# ---------------------------
function Ensure-Module {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [string]$MinimumVersion,

        [switch]$AllowInstall
    )

    $loaded = Get-Module -Name $Name -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $loaded) {
        if (-not $AllowInstall) {
            throw "Modul '$Name' ist nicht installiert. Installiere es z.B. mit: Install-Module $Name -Scope CurrentUser"
        }

        if (-not $Quiet) { Write-Host "Installiere Modul '$Name'..." -ForegroundColor Cyan }
        $params = @{ Name = $Name; Scope = 'CurrentUser'; Force = $true; AllowClobber = $true }
        if ($MinimumVersion) { $params.MinimumVersion = $MinimumVersion }
        Install-Module @params
    }
    else {
        if ($MinimumVersion -and ([version]$loaded.Version -lt [version]$MinimumVersion)) {
            if (-not $AllowInstall) {
                throw "Modul '$Name' ist vorhanden (Version $($loaded.Version)), aber Minimum ist $MinimumVersion. Update z.B.: Update-Module $Name"
            }
            if (-not $Quiet) { Write-Host "Aktualisiere Modul '$Name' (aktuell $($loaded.Version))..." -ForegroundColor Cyan }
            Update-Module -Name $Name -Force
        }
    }
}

function Test-CommandParameter {
    param(
        [Parameter(Mandatory)][string]$Command,
        [Parameter(Mandatory)][string]$ParameterName
    )
    $cmd = Get-Command $Command -ErrorAction SilentlyContinue
    if (-not $cmd) { return $false }
    return $cmd.Parameters.ContainsKey($ParameterName)
}

# ---------------------------
# Auth Context (global script)
# ---------------------------
$script:BwsAuthContext = [ordered]@{
    Mode                 = $Mode
    TenantId             = $TenantId
    AccountId            = $AccountId
    ClientId             = $ClientId
    CertificateThumbprint= $CertificateThumbprint
    Organization         = $Organization
    SharePointAdminUrl   = $SharePointAdminUrl
    GraphScopes          = $Scopes
    Connected            = [ordered]@{
        Graph      = $false
        Exchange   = $false
        Teams      = $false
        SharePoint = $false
    }
    TimestampUtc         = (Get-Date).ToUniversalTime().ToString("s") + "Z"
}

function Get-BwsAuthContext {
    [CmdletBinding()]
    param()
    return [pscustomobject]$script:BwsAuthContext
}

# ---------------------------
# Connect / Disconnect
# ---------------------------
function Connect-BwsServices {
    [CmdletBinding()]
    param(
        [ValidateSet('Interactive','DeviceCode','AppOnlyCertificate')]
        [string]$Mode = $script:BwsAuthContext.Mode,

        [string]$TenantId = $script:BwsAuthContext.TenantId,
        [string]$AccountId = $script:BwsAuthContext.AccountId,

        [string]$ClientId = $script:BwsAuthContext.ClientId,
        [string]$CertificateThumbprint = $script:BwsAuthContext.CertificateThumbprint,
        [string]$Organization = $script:BwsAuthContext.Organization,

        [string[]]$Scopes = $script:BwsAuthContext.GraphScopes,

        [switch]$ConnectGraph = $true,
        [switch]$ConnectExchange,
        [switch]$ConnectTeams,
        [switch]$ConnectSharePoint,

        [string]$SharePointAdminUrl = $script:BwsAuthContext.SharePointAdminUrl,

        [switch]$AllowInstall,
        [switch]$Quiet
    )

    # Graph
    if ($ConnectGraph) {
        Ensure-Module -Name 'Microsoft.Graph.Authentication' -MinimumVersion '2.0.0' -AllowInstall:$AllowInstall

        if (-not $Quiet) { Write-Host "Verbinde Microsoft Graph ($Mode)..." -ForegroundColor Cyan }

        $connectParams = @{}
        if ($TenantId) { $connectParams.TenantId = $TenantId }
        if (Test-CommandParameter -Command 'Connect-MgGraph' -ParameterName 'NoWelcome') { $connectParams.NoWelcome = $true }

        if ($Mode -eq 'AppOnlyCertificate') {
            if (-not $ClientId -or -not $CertificateThumbprint) {
                throw "Für AppOnlyCertificate brauchst du -ClientId und -CertificateThumbprint (und i.d.R. -TenantId)."
            }
            $connectParams.ClientId = $ClientId
            $connectParams.CertificateThumbprint = $CertificateThumbprint
            Connect-MgGraph @connectParams | Out-Null
        }
        else {
            if (-not $Scopes -or $Scopes.Count -eq 0) { $Scopes = @('User.Read') }
            $connectParams.Scopes = $Scopes

            if ($Mode -eq 'DeviceCode') {
                if (Test-CommandParameter -Command 'Connect-MgGraph' -ParameterName 'UseDeviceCode') {
                    $connectParams.UseDeviceCode = $true
                }
                else {
                    throw "Dein Microsoft.Graph.Authentication unterstützt -UseDeviceCode nicht. Update empfohlen: Update-Module Microsoft.Graph.Authentication"
                }
            }
            Connect-MgGraph @connectParams | Out-Null
        }

        $script:BwsAuthContext.Connected.Graph = $true
    }

    # Exchange Online
    if ($ConnectExchange) {
        Ensure-Module -Name 'ExchangeOnlineManagement' -MinimumVersion '3.0.0' -AllowInstall:$AllowInstall

        if (-not $Quiet) { Write-Host "Verbinde Exchange Online ($Mode)..." -ForegroundColor Cyan }

        if ($Mode -eq 'AppOnlyCertificate') {
            if (-not $Organization) { throw "Für EXO AppOnly brauchst du -Organization (z.B. contoso.onmicrosoft.com)." }
            Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false | Out-Null
        }
        else {
            # Interaktiv / DeviceCode -> EXO nutzt interaktiv (DeviceCode gibt’s je nach Version nicht als Standard)
            if ($AccountId) {
                Connect-ExchangeOnline -UserPrincipalName $AccountId -ShowBanner:$false | Out-Null
            }
            else {
                Connect-ExchangeOnline -ShowBanner:$false | Out-Null
            }
        }

        $script:BwsAuthContext.Connected.Exchange = $true
    }

    # Teams
    if ($ConnectTeams) {
        Ensure-Module -Name 'MicrosoftTeams' -MinimumVersion '5.0.0' -AllowInstall:$AllowInstall

        if (-not $Quiet) { Write-Host "Verbinde Microsoft Teams ($Mode)..." -ForegroundColor Cyan }

        if ($Mode -eq 'AppOnlyCertificate') {
            if (-not $TenantId) { throw "Für Teams AppOnly brauchst du -TenantId." }
            # Parameter heißen je nach Version leicht anders; wir versuchen robust:
            $params = @{}
            if (Test-CommandParameter -Command 'Connect-MicrosoftTeams' -ParameterName 'TenantId') { $params.TenantId = $TenantId }
            if (Test-CommandParameter -Command 'Connect-MicrosoftTeams' -ParameterName 'ApplicationId') { $params.ApplicationId = $ClientId }
            if (Test-CommandParameter -Command 'Connect-MicrosoftTeams' -ParameterName 'CertificateThumbprint') { $params.CertificateThumbprint = $CertificateThumbprint }
            Connect-MicrosoftTeams @params | Out-Null
        }
        else {
            $params = @{}
            if ($TenantId -and (Test-CommandParameter -Command 'Connect-MicrosoftTeams' -ParameterName 'TenantId')) { $params.TenantId = $TenantId }
            Connect-MicrosoftTeams @params | Out-Null
        }

        $script:BwsAuthContext.Connected.Teams = $true
    }

    # SharePoint Online (SPO Mgmt Shell)
    if ($ConnectSharePoint) {
        if (-not $SharePointAdminUrl) {
            throw "Für SharePoint (SPO) brauchst du -SharePointAdminUrl, z.B. https://<tenant>-admin.sharepoint.com"
        }

        Ensure-Module -Name 'Microsoft.Online.SharePoint.PowerShell' -MinimumVersion '16.0.0' -AllowInstall:$AllowInstall

        if (-not $Quiet) { Write-Host "Verbinde SharePoint Online (SPO)..." -ForegroundColor Cyan }
        Connect-SPOService -Url $SharePointAdminUrl

        $script:BwsAuthContext.Connected.SharePoint = $true
    }

    if (-not $Quiet) {
        Write-Host "Verbindungen hergestellt: Graph=$($script:BwsAuthContext.Connected.Graph), EXO=$($script:BwsAuthContext.Connected.Exchange), Teams=$($script:BwsAuthContext.Connected.Teams), SPO=$($script:BwsAuthContext.Connected.SharePoint)" -ForegroundColor Green
    }

    return Get-BwsAuthContext
}

function Disconnect-BwsServices {
    [CmdletBinding()]
    param(
        [switch]$Graph,
        [switch]$Exchange,
        [switch]$Teams
        # SPO hat i.d.R. kein echtes Disconnect
    )

    if ($Graph -and (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue)) {
        Disconnect-MgGraph | Out-Null
        $script:BwsAuthContext.Connected.Graph = $false
    }

    if ($Exchange -and (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        $script:BwsAuthContext.Connected.Exchange = $false
    }

    if ($Teams -and (Get-Command Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue)) {
        Disconnect-MicrosoftTeams | Out-Null
        $script:BwsAuthContext.Connected.Teams = $false
    }

    return Get-BwsAuthContext
}

# ---------------------------
# Entry Point (nur wenn direkt ausgeführt)
# ---------------------------
function Invoke-BwsAuthEntryPoint {
    [CmdletBinding()]
    param()

    Connect-BwsServices @PSBoundParameters | Out-Null
    Get-BwsAuthContext
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-BwsAuthEntryPoint
}
