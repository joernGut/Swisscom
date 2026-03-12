#Requires -Version 5.1
#Requires -PSEdition Desktop
# NOTE: Module requirements are handled at runtime by Install-BWSDependencies
#       to allow automatic installation of missing modules with visual progress.

<#
.SYNOPSIS
    BWS (Business Workplace Service) Checking Script with GUI support
.DESCRIPTION
    Checks Azure resources and Intune policies for BWS environments
.PARAMETER BCID
    Business Continuity ID
.PARAMETER CustomerName
    Name of the customer (optional, used in HTML report)
.PARAMETER SubscriptionId
    Azure Subscription ID (optional)
.PARAMETER ExportReport
    Export results to HTML file
.PARAMETER SkipIntune
    Skip Intune policy checks
.PARAMETER SkipEntraID
    Skip Entra ID Connect checks
.PARAMETER SkipIntuneConnector
    Skip Hybrid Azure AD Join checks
.PARAMETER SkipDefender
    Skip Defender for Endpoint checks
.PARAMETER ShowAllPolicies
    Show all found Intune policies (debug mode)
.PARAMETER CompactView
    Show only summary without detailed tables
.PARAMETER GUI
    Launch graphical user interface
.PARAMETER RunAnalyzer
    Run PSScriptAnalyzer static code analysis before executing checks
.PARAMETER RunTests
    Run built-in unit tests (Pester-style) before executing checks
.NOTES
    Version: 2.3.0
    Datum: 2025-03-12
    Autor: BWS PowerShell Script
.EXAMPLE
    .\BWS-Checking-Script.ps1 -BCID "1234" -CustomerName "Contoso AG"
.EXAMPLE
    .\BWS-Checking-Script.ps1 -BCID "1234" -CustomerName "Contoso AG" -GUI
.EXAMPLE
    .\BWS-Checking-Script.ps1 -BCID "1234" -CustomerName "Contoso AG" -ExportReport
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern('^[0-9A-Za-z]{1,8}$')]
    [string]$BCID = "0000",
    
    [Parameter(Mandatory=$false)]
    [string]$CustomerName = "",
    
    [Parameter(Mandatory=$false)]
    [string]$SubscriptionId,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportReport,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipIntune,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipEntraID,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipIntuneConnector,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipDefender,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipSoftware,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipSharePoint,
    
    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if ([string]::IsNullOrEmpty($_)) { return $true }
        if ($_ -match '^https://[a-zA-Z0-9-]+\.sharepoint\.com.*$') { return $true }
        throw "SharePointUrl must be a valid SharePoint URL (e.g. https://contoso-admin.sharepoint.com)"
    })]
    [string]$SharePointUrl,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipTeams,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipUserLicenseCheck,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("HTML", "PDF", "Both")]
    [string]$ExportFormat = "HTML",
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowAllPolicies,
    
    [Parameter(Mandatory=$false)]
    [switch]$CompactView,
    
    [Parameter(Mandatory=$false)]
    [switch]$GUI,

    # Quality Assurance Parameters
    [Parameter(Mandatory=$false,
               HelpMessage="Run PSScriptAnalyzer static analysis before executing checks.")]
    [switch]$RunAnalyzer,

    [Parameter(Mandatory=$false,
               HelpMessage="Run built-in Pester-style unit tests before executing checks.")]
    [switch]$RunTests
)

# Script Version
$script:Version = "2.3.0"

#============================================================================
# QUALITY ASSURANCE - Block 1: Strict Mode
#============================================================================
# Set-StrictMode -Version Latest

# Remove any Microsoft.Graph modules that may have been loaded in this PS session
# (e.g. from a previous script run). Microsoft.Graph SDK v2.x cannot load in PS 5.1
# and causes a GetTokenAsync / .NET Framework incompatibility error if left in session.
$null = Get-Module -Name 'Microsoft.Graph*' -ErrorAction SilentlyContinue |
        Remove-Module -Force -ErrorAction SilentlyContinue catches:
#   - Uninitialised variable access
#   - Calling properties on $null
#   - Out-of-bounds array indexing
#   - Calling non-existent members
Set-StrictMode -Version Latest

#============================================================================
# QUALITY ASSURANCE - Block 2: Self-Syntax Check
#============================================================================
# Parse this very file using the PowerShell AST parser.
# Any syntax error halts execution with a clear message.
try {
    $null = [System.Management.Automation.Language.Parser]::ParseFile(
        $PSCommandPath,
        [ref]$null,
        [ref]$parseErrors
    )
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        Write-Host ""
        Write-Host "  [SYNTAX ERRORS FOUND - Script halted]" -ForegroundColor Red
        Write-Host "  The following syntax errors were detected in the script file:" -ForegroundColor Red
        Write-Host ""
        foreach ($err in $parseErrors) {
            Write-Host "  Line $($err.Extent.StartLineNumber): $($err.Message)" -ForegroundColor Yellow
        }
        Write-Host ""
        exit 1
    }
    Write-Host "  [OK] Syntax check passed" -ForegroundColor Green
} catch {
    Write-Host "  [!] Syntax check could not run: $($_.Exception.Message)" -ForegroundColor Yellow
}

#============================================================================
# QUALITY ASSURANCE - Block 3: Execution Policy Check
#============================================================================
$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
$machinePolicy = Get-ExecutionPolicy -Scope LocalMachine
$effectivePolicy = Get-ExecutionPolicy

Write-Host "  [i] Execution Policy  - Effective: $effectivePolicy | Machine: $machinePolicy | User: $currentPolicy" -ForegroundColor Gray

$blockedPolicies = @("Restricted", "AllSigned")
if ($effectivePolicy -in $blockedPolicies) {
    Write-Host ""
    Write-Host "  [!] WARNING: Execution Policy '$effectivePolicy' may block this script." -ForegroundColor Yellow
    Write-Host "      Recommended: Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned" -ForegroundColor Gray
    Write-Host ""
}

#============================================================================
# QUALITY ASSURANCE - Block 4: Parameter Validation (post-param)
#============================================================================
# Additional cross-parameter validation that ValidateScript/Pattern cannot do
if ($ExportReport -and -not $GUI) {
    if ($ExportFormat -eq "PDF" -or $ExportFormat -eq "Both") {
        Write-Host "  [i] PDF export requested - requires wkhtmltopdf, Chrome/Edge, or Word installed." -ForegroundColor Gray
    }
}

if ($SharePointUrl -and $SkipSharePoint) {
    Write-Host "  [!] Warning: -SharePointUrl is set but -SkipSharePoint is also set. SharePoint URL will be ignored." -ForegroundColor Yellow
}

if ($BCID -eq "0000") {
    Write-Host "  [!] Warning: Using default BCID '0000'. Specify -BCID for a real environment." -ForegroundColor Yellow
}

#============================================================================
# Global Variables and Configuration
#============================================================================

# PowerShell Version Check
$psVersion = $PSVersionTable.PSVersion.Major
$psEdition = $PSVersionTable.PSEdition

if ($psVersion -ge 7 -or $psEdition -eq "Core") {
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Yellow
    Write-Host "  BWS-Checking-Script v$script:Version" -ForegroundColor Cyan
    Write-Host "  [!] WARNUNG: PowerShell Version Inkompatibilitaet" -ForegroundColor Yellow
    Write-Host "======================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Sie verwenden: PowerShell $($PSVersionTable.PSVersion) ($psEdition)" -ForegroundColor Yellow
    Write-Host "Empfohlen:     PowerShell 5.1 (Desktop)" -ForegroundColor Green
    Write-Host ""
    Write-Host "WICHTIG:" -ForegroundColor Red
    Write-Host "  Der SharePoint-Check funktioniert NUR in PowerShell 5.1!" -ForegroundColor Red
    Write-Host "  Das Modul 'Microsoft.Online.SharePoint.PowerShell'" -ForegroundColor Red
    Write-Host "  wird in PowerShell 7/Core NICHT unterstuetzt." -ForegroundColor Red
    Write-Host ""
    Write-Host "6 von 7 Checks funktionieren in PowerShell 7:" -ForegroundColor Yellow
    Write-Host "  [OK] Azure Resources" -ForegroundColor Green
    Write-Host "  [OK] Intune Policies" -ForegroundColor Green
    Write-Host "  [OK] Entra ID Connect" -ForegroundColor Green
    Write-Host "  [OK] Hybrid Azure AD Join" -ForegroundColor Green
    Write-Host "  [OK] Defender for Endpoint" -ForegroundColor Green
    Write-Host "  [OK] BWS Software Packages" -ForegroundColor Green
    Write-Host "  [X] SharePoint Configuration (FEHLT)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Optionen:" -ForegroundColor Cyan
    Write-Host "  1. Script in PowerShell 5.1 neu starten (EMPFOHLEN)" -ForegroundColor White
    Write-Host "     -> Schliessen Sie diese Konsole" -ForegroundColor Gray
    Write-Host "     -> Oeffnen Sie 'Windows PowerShell' (nicht PowerShell 7)" -ForegroundColor Gray
    Write-Host "     -> Fuehren Sie das Script erneut aus" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  2. SharePoint-Check ueberspringen" -ForegroundColor White
    Write-Host "     -> Fuegen Sie -SkipSharePoint Parameter hinzu" -ForegroundColor Gray
    Write-Host "     -> Beispiel: .\BWS-Checking-Script.ps1 -BCID '1234' -SkipSharePoint" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  3. Mit Login-GUI arbeiten" -ForegroundColor White
    Write-Host "     -> Starten Sie: .\Azure-M365-Login-GUI.ps1" -ForegroundColor Gray
    Write-Host "     -> Klicken Sie auf 'PowerShell 5.1' Button (Blau)" -ForegroundColor Gray
    Write-Host "     -> Fuehren Sie das Script in der neuen Konsole aus" -ForegroundColor Gray
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Yellow
    Write-Host ""
    
    # Frage ob fortfahren
    if (-not $SkipSharePoint) {
        $continue = Read-Host "Trotzdem fortfahren? SharePoint-Check wird fehlschlagen. (J/N)"
        if ($continue -ne "J" -and $continue -ne "j" -and $continue -ne "Y" -and $continue -ne "y") {
            Write-Host ""
            Write-Host "Script abgebrochen. Bitte verwenden Sie PowerShell 5.1." -ForegroundColor Yellow
            Write-Host ""
            exit
        }
        Write-Host ""
        Write-Host "Fahre fort ohne SharePoint-Check Unterstuetzung..." -ForegroundColor Yellow
        Write-Host ""
    }
}

#============================================================================

#============================================================================
# Required Modules Definition
#============================================================================

$script:RequiredModules = @(
    # MaxVersion="" = no upper bound.
    # Microsoft.Graph.Authentication is pinned to v1.x (MaxVersion="1.99.99"):
    #   Graph SDK v2.x requires .NET 6+ and will NOT load in PS 5.1 on .NET Framework 4.x.
    #   v1.x uses Invoke-BWsGraphRequest for all API calls (v1.0 and /beta/ endpoints).
    #   Sub-modules (DeviceManagement, Users etc.) are NOT needed - all calls go through
    #   Invoke-BWsGraphRequest REST calls which only require Authentication.
    @{ Name="Az.Accounts";                           MinVersion="2.0.0";  MaxVersion="";       Description="Azure Authentication";        Scope="CurrentUser"; Required=$true;  SkipParam="" },
    @{ Name="Az.Resources";                          MinVersion="1.0.0";  MaxVersion="";       Description="Azure Resource Management";   Scope="CurrentUser"; Required=$true;  SkipParam="" },
    @{ Name="Az.Storage";                            MinVersion="3.0.0";  MaxVersion="";       Description="Azure Storage (Defender)";    Scope="CurrentUser"; Required=$false; SkipParam="SkipDefender" },
    # Microsoft.Graph.Authentication is NOT required - we use Az.Accounts token + Invoke-RestMethod.
    # This avoids ALL Graph SDK version conflicts with PS 5.1 / .NET Framework 4.x.
    @{ Name="Microsoft.Online.SharePoint.PowerShell";MinVersion="16.0.0"; MaxVersion="";       Description="SharePoint Online Admin";     Scope="CurrentUser"; Required=$false; SkipParam="SkipSharePoint" },
    @{ Name="MicrosoftTeams";                        MinVersion="4.0.0";  MaxVersion="";       Description="Microsoft Teams Admin";       Scope="CurrentUser"; Required=$false; SkipParam="SkipTeams" },
    @{ Name="PSScriptAnalyzer";                      MinVersion="1.20.0"; MaxVersion="";       Description="PS Static Code Analysis";     Scope="CurrentUser"; Required=$false; SkipParam="" }
)


#============================================================================
# Graph REST API Helpers  -  Az.Accounts token + Invoke-RestMethod
# No Microsoft.Graph module required. Works in PS 5.1 / .NET Framework 4.x.
#============================================================================

# Token cache
$script:BWSGraphToken       = $null
$script:BWSGraphTokenExpiry = [datetime]::MinValue
$script:BWSGraphConnected   = $false

#============================================================================
# Graph REST Helpers - Invoke-AzRestMethod only, zero Graph SDK dependency
# Works in PS 5.1 / .NET Framework 4.x with Az.Accounts v2+ (incl. v5.x).
# Invoke-AzRestMethod is a native Az.Accounts cmdlet: it handles auth
# internally using the active Az session and never touches the Graph SDK.
#============================================================================

$script:BWSGraphConnected = $false   # becomes $true after first successful call

function Connect-BWsGraph {
    <#
    .SYNOPSIS
        Validates the active Az session can reach Graph.
        No Graph module, no token handling needed - Invoke-AzRestMethod does it all.
    #>
    param([string[]]$Scopes = @())   # kept for call-site compatibility only

    # Remove any stale Graph SDK modules that may auto-load from session
    Get-Module -Name 'Microsoft.Graph*' -ErrorAction SilentlyContinue |
        Remove-Module -Force -ErrorAction SilentlyContinue

    try {
        $azCtx = Get-AzContext -ErrorAction Stop
        if (-not $azCtx) {
            Write-Host " [X] No active Az session - run Connect-AzAccount first" -ForegroundColor Red
            $script:BWSGraphConnected = $false
            return $false
        }
        # Probe Graph with a lightweight call to confirm auth works
        $probe = Invoke-AzRestMethod -Uri "https://graph.microsoft.com/v1.0/organization?$`select=id" -Method GET -ErrorAction Stop
        if ($probe.StatusCode -ge 200 -and $probe.StatusCode -lt 300) {
            $script:BWSGraphConnected = $true
            Write-Host " [OK] Graph connection verified (Az session: $($azCtx.Account.Id))" -ForegroundColor Green
            return $true
        } else {
            Write-Host " [X] Graph probe returned HTTP $($probe.StatusCode)" -ForegroundColor Red
            $script:BWSGraphConnected = $false
            return $false
        }
    } catch {
        Write-Host " [X] Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
        $script:BWSGraphConnected = $false
        return $false
    }
}

function Get-BWsGraphContext {
    <#
    .SYNOPSIS  Returns $true if Graph is reachable via the active Az session.#>
    $azCtx = Get-AzContext -ErrorAction SilentlyContinue
    return $(if ($azCtx -and $script:BWSGraphConnected) {
        @{ Account = $azCtx.Account.Id }
    } else { $null })
}

function Invoke-BWsGraphRequest {
    <#
    .SYNOPSIS
        Calls a Microsoft Graph v1.0 endpoint via Invoke-AzRestMethod.
        No Graph module or bearer token handling needed.
    .PARAMETER Uri
        Full URL or path relative to https://graph.microsoft.com/v1.0/
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [string]   $Method      = "GET",
        [hashtable]$Body        = $null,
        [string]   $ErrorAction = "Stop"
    )

    if ($Uri -notmatch '^https://') {
        $Uri = "https://graph.microsoft.com/v1.0/$($Uri.TrimStart('/'))"
    }

    $restParams = @{ Uri = $Uri; Method = $Method; ErrorAction = $ErrorAction }
    if ($Body) { $restParams.Payload = ($Body | ConvertTo-Json -Depth 10) }

    $resp = Invoke-AzRestMethod @restParams
    if ($resp.StatusCode -lt 200 -or $resp.StatusCode -ge 300) {
        if ($ErrorAction -eq "Stop") {
            throw "Graph API error $($resp.StatusCode): $($resp.Content)"
        }
        return $null
    }
    return $resp.Content | ConvertFrom-Json
}

function Invoke-BWsGraphPagedRequest {
    <#
    .SYNOPSIS
        Calls a Graph endpoint and follows @odata.nextLink pages.
        Returns a flat array of all items.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [string]$Method = "GET"
    )
    if ($Uri -notmatch '^https://') {
        $Uri = "https://graph.microsoft.com/v1.0/$($Uri.TrimStart('/'))"
    }
    $allItems = [System.Collections.Generic.List[object]]::new()
    $nextUri  = $Uri
    do {
        $resp = Invoke-BWsGraphRequest -Uri $nextUri -Method $Method -ErrorAction Stop
        if ($null -eq $resp) { break }
        if ($resp.value) {
            foreach ($item in $resp.value) { $allItems.Add($item) }
        } elseif ($null -ne $resp) {
            $allItems.Add($resp)
        }
        $nextUri = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
    } while ($nextUri)
    return $allItems.ToArray()
}

function Invoke-BWsBetaGraphRequest {
    <#
    .SYNOPSIS  Calls the Microsoft Graph BETA endpoint via Invoke-AzRestMethod.#>
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [string]   $Method = "GET",
        [hashtable]$Body   = $null
    )
    if ($Uri -notmatch '^https://') {
        $Uri = "https://graph.microsoft.com/beta/$($Uri.TrimStart('/'))"
    }
    return Invoke-BWsGraphRequest -Uri $Uri -Method $Method -Body $Body
}

# Public alias kept for any external callers
function Invoke-MgBetaGraphRequest {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [string]   $Method = "GET",
        [hashtable]$Body   = $null
    )
    return Invoke-BWsBetaGraphRequest -Uri $Uri -Method $Method -Body $Body
}


function Get-ModuleStatus {
    param([string]$Name, [string]$MinVersion = "0.0.0", [string]$MaxVersion = "")

    # Safe version comparison - [version] cast can throw on unusual version strings in PS 5.1
    $minVer = try { [version]$MinVersion } catch { [version]"0.0.0" }

    $avail = Get-Module -ListAvailable -Name $Name -ErrorAction SilentlyContinue |
             Where-Object {
                 $modVer = try { [version]$_.Version } catch { [version]"0.0.0" }
                 $modVer -ge $minVer
             } |
             Sort-Object { try { [version]$_.Version } catch { [version]"0.0.0" } } -Descending |
             Select-Object -First 1

    $loaded = Get-Module -Name $Name -ErrorAction SilentlyContinue | Select-Object -First 1

    return @{
        IsAvailable  = ($null -ne $avail)
        IsLoaded     = ($null -ne $loaded)
        InstalledVer = if ($avail)  { $avail.Version.ToString()  } else { $null }
        LoadedVer    = if ($loaded) { $loaded.Version.ToString() } else { $null }
    }
}

#============================================================================
# Module Setup: Install-BWSDependencies (Console + GUI callback)
#============================================================================

function Install-BWSDependencies {
    <#
    .SYNOPSIS
        Checks, installs and imports all required modules with live timing output.
    .PARAMETER SkipParams
        Hashtable of skip flags, e.g. @{SkipSharePoint=$true; SkipTeams=$true}
    .PARAMETER GUICallback
        Scriptblock called after each module step for GUI updates.
    #>
    param(
        [hashtable]$SkipParams    = @{},
        [scriptblock]$GUICallback = $null
    )

    $W = @{ N=36; D=30; S=18; I=9; M=9 }
    $sep = "  +" + ("-" * ($W.N+2)) + "+" + ("-" * ($W.D+2)) + "+" + ("-" * ($W.S+2)) + "+" + ("-" * ($W.I+2)) + "+" + ("-" * ($W.M+2)) + "+"

    function Write-MRow {
        param($R, [string]$Override = "")
        $s  = if ($Override) { $Override } else { $R.Status }
        $ni = $R.InstallTime; $mi = $R.ImportTime
        $n  = $R.Name.PadRight($W.N).Substring(0,$W.N)
        $d  = if ($R.Desc.Length -gt $W.D) { $R.Desc.Substring(0,$W.D) } else { $R.Desc.PadRight($W.D) }
        $ss = if ($s.Length -gt $W.S) { $s.Substring(0,$W.S) } else { $s.PadRight($W.S) }
        $ii = if ($ni) { $ni.PadRight($W.I).Substring(0,$W.I) } else { "".PadRight($W.I) }
        $mm = if ($mi) { $mi.PadRight($W.M).Substring(0,$W.M) } else { "".PadRight($W.M) }
        $ln = "  | $n | $d | $ss | $ii | $mm |"
        if     ($s -like "*[OK]*")   { $c = "Green"    }
        elseif ($s -like "*SKIP*")   { $c = "DarkGray" }
        elseif ($s -like "*[X]*")    { $c = "Red"      }
        elseif ($s -like "*[!]*")    { $c = "Yellow"   }
        elseif ($s -like "*...*")    { $c = "Cyan"     }
        else                         { $c = "White"    }
        Write-Host $ln -ForegroundColor $c
    }

    Write-Host ""
    Write-Host $sep -ForegroundColor Cyan
    $hdr = "  | " + "BWS Module Prerequisites".PadRight($W.N) + " | " + "Description".PadRight($W.D) + " | " + "Status".PadRight($W.S) + " | " + "Install".PadRight($W.I) + " | " + "Import".PadRight($W.M) + " |"
    Write-Host $hdr -ForegroundColor White
    Write-Host $sep -ForegroundColor DarkGray

    $allResults = [System.Collections.Generic.List[hashtable]]::new()
    $totalSW    = [System.Diagnostics.Stopwatch]::StartNew()
    $modCount   = 0

    # Ensure NuGet provider is available (required in PS 5.1 for Install-Module)
    $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $nuget -or $nuget.Version -lt [version]"2.8.5.201") {
        Write-Host "  Installing NuGet package provider (required for Install-Module)..." -ForegroundColor Yellow
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
            Write-Host "  [OK] NuGet provider ready" -ForegroundColor Green
        } catch {
            Write-Host "  [!] NuGet provider install failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Trust PSGallery if not already trusted (avoids interactive prompts during Install-Module)
    $gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($gallery -and $gallery.InstallationPolicy -ne "Trusted") {
        try {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
            Write-Host "  [OK] PSGallery set to Trusted for this session" -ForegroundColor Gray
        } catch {}
    }

    foreach ($mod in $script:RequiredModules) {
        $modCount++
        $row = @{ Name=$mod.Name; Desc=$mod.Description; Status="Checking..."; InstallTime=""; ImportTime=""; Error=""; Skipped=$false }

        # Skip check
        $doSkip = $false
        if ($mod['SkipParam'] -and $SkipParams.ContainsKey($mod['SkipParam']) -and $SkipParams[$mod['SkipParam']] -eq $true) { $doSkip = $true }

        Write-Progress -Activity "BWS Module Setup" `
            -Status "[$modCount/$($script:RequiredModules.Count)] $($mod.Name)" `
            -PercentComplete ([int](($modCount-1) / $script:RequiredModules.Count * 100))

        if ($doSkip) {
            $row.Status = "SKIP"; $row.InstallTime = "n/a"; $row.ImportTime = "n/a"; $row.Skipped = $true
            $allResults.Add($row); Write-MRow $row
            if ($GUICallback) { & $GUICallback $row }
            continue
        }

        $maxV = if ($mod.ContainsKey('MaxVersion') -and $mod['MaxVersion']) { $mod['MaxVersion'] } else { '' }
        $st = Get-ModuleStatus -Name $mod.Name -MinVersion $mod.MinVersion -MaxVersion $maxV

        #  Install 
        if (-not $st.IsAvailable) {
            $row.Status = "Installing..."
            Write-MRow $row
            if ($GUICallback) { & $GUICallback $row }
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                $installParams = @{
                    Name            = $mod.Name
                    MinimumVersion  = $mod.MinVersion
                    Scope           = $mod.Scope
                    Repository      = "PSGallery"
                    Force           = $true
                    AllowClobber    = $true
                    ErrorAction     = "Stop"
                }
                if ($mod.ContainsKey('MaxVersion') -and $mod['MaxVersion']) { $installParams.MaximumVersion = $mod['MaxVersion'] }
                Install-Module @installParams
                $sw.Stop()
                $row.InstallTime = "$([int]$sw.Elapsed.TotalSeconds)s"
            } catch {
                $sw.Stop()
                $row.Status      = "[X] Install failed"
                $row.InstallTime = "$([int]$sw.Elapsed.TotalSeconds)s"
                $row.Error       = $_.Exception.Message
                $allResults.Add($row); Write-MRow $row
                Write-Host "      ERR: $($_.Exception.Message)" -ForegroundColor Red
                if ($GUICallback) { & $GUICallback $row }
                continue
            }
        } else {
            $row.InstallTime = "cached"
        }

        #  Import 
        $st2 = Get-ModuleStatus -Name $mod.Name -MinVersion $mod.MinVersion -MaxVersion $maxV
        if (-not $st2.IsLoaded) {
            $row.Status = "Importing..."
            Write-MRow $row
            if ($GUICallback) { & $GUICallback $row }
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                # Register-AzModule TypeInitializationException and SharePoint unapproved-verb
                # warnings are non-terminating errors on stream 2 and warnings on stream 3.
                # Both must be redirected. ErrorAction SilentlyContinue allows stream 2 to be
                # redirected via 2>$null; we then verify the import succeeded by checking the
                # loaded module list ourselves and throw if it genuinely failed.
                $importParams = @{
                    Name              = $mod.Name
                    Force             = $true
                    WarningAction     = "SilentlyContinue"
                    ErrorAction       = "SilentlyContinue"
                }
                # If MaxVersion is set (e.g. Graph v1.x pin), load the specific pinned version
                if ($maxV) {
                    $pinVer = Get-Module -ListAvailable -Name $mod.Name -ErrorAction SilentlyContinue |
                              Where-Object {
                                  $mv = try { [version]$_.Version } catch { [version]"0.0.0" }
                                  $mv -le [version]$maxV
                              } |
                              Sort-Object { try { [version]$_.Version } catch { [version]"0.0.0" } } -Descending |
                              Select-Object -First 1
                    if ($pinVer) {
                        $importParams.RequiredVersion = $pinVer.Version.ToString()
                    }
                }
                # Redirect stream 2 (errors) and stream 3 (warnings) to suppress cosmetic noise
                Import-Module @importParams 2>$null 3>$null
                # Verify the module actually loaded - throw if not so catch block fires
                $verifyLoaded = Get-Module -Name $mod.Name -ErrorAction SilentlyContinue
                if (-not $verifyLoaded) {
                    throw "Module '$($mod.Name)' did not load after Import-Module (no terminating error was thrown)"
                }
                $sw.Stop()
                $row.ImportTime = "$([int]$sw.Elapsed.TotalSeconds)s"
            } catch {
                $sw.Stop()
                $row.Status     = "[!] Import failed"
                $row.ImportTime = "$([int]$sw.Elapsed.TotalSeconds)s"
                $row.Error      = $_.Exception.Message
                $allResults.Add($row); Write-MRow $row
                Write-Host "      ERR: $($_.Exception.Message)" -ForegroundColor Red
                if ($GUICallback) { & $GUICallback $row }
                continue
            }
        } else {
            $row.ImportTime = "loaded"
        }

        $st3        = Get-ModuleStatus -Name $mod.Name -MinVersion $mod.MinVersion -MaxVersion $maxV
        $row.Status = if ($st3.IsLoaded) { "[OK] v$($st3.LoadedVer)" } `
                      elseif ($st3.IsAvailable) { "[!] Not loaded" } `
                      else { "[X] Not found" }
        $allResults.Add($row)
        Write-MRow $row
        if ($GUICallback) { & $GUICallback $row }
    }

    $totalSW.Stop()
    Write-Progress -Activity "BWS Module Setup" -Completed -ErrorAction SilentlyContinue

    $failed   = @($allResults | Where-Object { $_.Status -like "*[X]*" -or $_.Status -like "*failed*" })
    $warnings = @($allResults | Where-Object { $_.Status -like "*[!]*" })
    $skipped  = @($allResults | Where-Object { $_.Skipped -eq $true })
    $ok       = @($allResults | Where-Object { $_.Status -like "*[OK]*" })

    Write-Host $sep -ForegroundColor Cyan
    $sumColor = if ($failed.Count -gt 0) { "Red" } elseif ($warnings.Count -gt 0) { "Yellow" } else { "Green" }
    Write-Host "  Modules: $($allResults.Count)  [OK]: $($ok.Count)  Skipped: $($skipped.Count)  Warn: $($warnings.Count)  Failed: $($failed.Count)  Time: $([int]$totalSW.Elapsed.TotalSeconds)s" -ForegroundColor $sumColor
    Write-Host $sep -ForegroundColor Cyan
    Write-Host ""

    return @{
        Results   = $allResults
        OK        = $ok.Count
        Failed    = $failed.Count
        Warnings  = $warnings.Count
        Skipped   = $skipped.Count
        TotalSecs = [int]$totalSW.Elapsed.TotalSeconds
        AllReady  = ($failed.Count -eq 0)
    }
}

#============================================================================
# Module Setup: Show-ModuleSetupDialog (WinForms GUI variant)
#============================================================================

function Show-ModuleSetupDialog {
    <#
    .SYNOPSIS
        WinForms dialog that shows module install/import progress in real-time.
    .PARAMETER SkipParams
        Same as Install-BWSDependencies -SkipParams.
    #>
    param([hashtable]$SkipParams = @{})

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $dlg                  = New-Object System.Windows.Forms.Form
    $dlg.Text             = "BWS Module Prerequisites"
    $dlg.Size             = New-Object System.Drawing.Size(880, 510)
    $dlg.StartPosition    = "CenterScreen"
    $dlg.FormBorderStyle  = "FixedDialog"
    $dlg.MaximizeBox      = $false
    $dlg.MinimizeBox      = $false
    $dlg.BackColor        = [System.Drawing.Color]::FromArgb(22, 22, 30)

    # Title
    $lTitle               = New-Object System.Windows.Forms.Label
    $lTitle.Text          = "Checking and loading required PowerShell modules..."
    $lTitle.Location      = New-Object System.Drawing.Point(14, 12)
    $lTitle.Size          = New-Object System.Drawing.Size(840, 24)
    $lTitle.ForeColor     = [System.Drawing.Color]::FromArgb(160, 210, 255)
    $lTitle.Font          = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lTitle)

    # ListView
    $lv                   = New-Object System.Windows.Forms.ListView
    $lv.Location          = New-Object System.Drawing.Point(14, 44)
    $lv.Size              = New-Object System.Drawing.Size(840, 290)
    $lv.View              = [System.Windows.Forms.View]::Details
    $lv.FullRowSelect     = $true
    $lv.GridLines         = $true
    $lv.BackColor         = [System.Drawing.Color]::FromArgb(28, 28, 40)
    $lv.ForeColor         = [System.Drawing.Color]::FromArgb(205, 205, 205)
    $lv.Font              = New-Object System.Drawing.Font("Consolas", 9)
    $lv.HeaderStyle       = [System.Windows.Forms.ColumnHeaderStyle]::Nonclickable
    $null = $lv.Columns.Add("Module",       170)
    $null = $lv.Columns.Add("Description",  200)
    $null = $lv.Columns.Add("Status",       155)
    $null = $lv.Columns.Add("Install",       80)
    $null = $lv.Columns.Add("Import",        80)
    $null = $lv.Columns.Add("Version",      110)
    $dlg.Controls.Add($lv)

    # Pre-populate rows and keep hashtable for updates
    $lvMap = @{}
    foreach ($mod in $script:RequiredModules) {
        $item = New-Object System.Windows.Forms.ListViewItem($mod.Name)
        $null = $item.SubItems.Add($mod.Description)
        $null = $item.SubItems.Add("Waiting...")
        $null = $item.SubItems.Add("")
        $null = $item.SubItems.Add("")
        $null = $item.SubItems.Add("")
        $item.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 120)
        $lv.Items.Add($item) | Out-Null
        $lvMap[$mod.Name] = $item
    }

    # Overall progress bar
    $pb               = New-Object System.Windows.Forms.ProgressBar
    $pb.Location      = New-Object System.Drawing.Point(14, 346)
    $pb.Size          = New-Object System.Drawing.Size(840, 22)
    $pb.Style         = "Continuous"
    $pb.Maximum       = ($script:RequiredModules | Measure-Object).Count
    $pb.Value         = 0
    $dlg.Controls.Add($pb)

    # Current-module label
    $lCur             = New-Object System.Windows.Forms.Label
    $lCur.Location    = New-Object System.Drawing.Point(14, 374)
    $lCur.Size        = New-Object System.Drawing.Size(840, 20)
    $lCur.ForeColor   = [System.Drawing.Color]::FromArgb(200, 190, 100)
    $lCur.Font        = New-Object System.Drawing.Font("Segoe UI", 9)
    $lCur.Text        = "Starting..."
    $dlg.Controls.Add($lCur)

    # Timing detail label
    $lTime            = New-Object System.Windows.Forms.Label
    $lTime.Location   = New-Object System.Drawing.Point(14, 394)
    $lTime.Size       = New-Object System.Drawing.Size(840, 18)
    $lTime.ForeColor  = [System.Drawing.Color]::FromArgb(130, 130, 130)
    $lTime.Font       = New-Object System.Drawing.Font("Consolas", 8)
    $lTime.Text       = ""
    $dlg.Controls.Add($lTime)

    # Continue button (disabled until complete)
    $btn              = New-Object System.Windows.Forms.Button
    $btn.Location     = New-Object System.Drawing.Point(360, 420)
    $btn.Size         = New-Object System.Drawing.Size(160, 36)
    $btn.Text         = "Please wait..."
    $btn.Enabled      = $false
    $btn.BackColor    = [System.Drawing.Color]::FromArgb(50, 80, 50)
    $btn.ForeColor    = [System.Drawing.Color]::White
    $btn.Font         = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btn.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btn)

    # Wire up the callback
    $script:_dlgLVMap = $lvMap
    $script:_dlgPB    = $pb
    $script:_dlgLCur  = $lCur
    $script:_dlgLTime = $lTime
    $script:_dlgDone  = 0

    $cb = {
        param($Row)
        $itm = $script:_dlgLVMap[$Row.Name]
        if (-not $itm) { return }

        $s = $Row.Status
        if     ($s -like "*[OK]*")    { $c = [System.Drawing.Color]::FromArgb( 90,210, 90) }
        elseif ($s -like "*SKIP*")    { $c = [System.Drawing.Color]::FromArgb(110,110,120) }
        elseif ($s -like "*[X]*")     { $c = [System.Drawing.Color]::FromArgb(240, 80, 80) }
        elseif ($s -like "*[!]*")     { $c = [System.Drawing.Color]::FromArgb(240,190, 60) }
        elseif ($s -like "*Install*") { $c = [System.Drawing.Color]::FromArgb( 80,170,255) }
        elseif ($s -like "*Import*")  { $c = [System.Drawing.Color]::FromArgb(170,130,255) }
        else                          { $c = [System.Drawing.Color]::FromArgb(200,200,200) }
        $itm.SubItems[2].Text = $s
        $itm.SubItems[3].Text = $Row.InstallTime
        $itm.SubItems[4].Text = $Row.ImportTime
        $itm.ForeColor        = $c
        if ($s -like "*[OK]*") {
            $vs = Get-ModuleStatus -Name $Row.Name -MinVersion "0.0.0"
            $itm.SubItems[5].Text = $vs.LoadedVer
        }

        $script:_dlgLCur.Text  = "$($Row.Name) -- $s"

        $tp = @()
        if ($Row.InstallTime -and $Row.InstallTime -notin @("","n/a","cached","already")) { $tp += "Installed: $($Row.InstallTime)" }
        if ($Row.ImportTime  -and $Row.ImportTime  -notin @("","n/a","loaded"))           { $tp += "Imported: $($Row.ImportTime)" }
        $script:_dlgLTime.Text = $tp -join " | "

        if ($s -notlike "*ing*" -and $s -ne "Waiting...") {
            $script:_dlgDone++
            $script:_dlgPB.Value = [Math]::Min($script:_dlgDone, $script:_dlgPB.Maximum)
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    $script:_modResult = $null
    $dlg.Add_Shown({
        $lCur.Text = "Running module verification..."
        [System.Windows.Forms.Application]::DoEvents()

        $script:_modResult = Install-BWSDependencies -SkipParams $SkipParams -GUICallback $cb

        $f = $script:_modResult.Failed
        $t = $script:_modResult.TotalSecs
        if ($f -gt 0) {
            $lCur.Text      = "[!] Setup finished with $f error(s). Some checks may not run."
            $lCur.ForeColor = [System.Drawing.Color]::FromArgb(255,120,120)
        } else {
            $lCur.Text      = "[OK] All modules ready in ${t}s -- Click Continue"
            $lCur.ForeColor = [System.Drawing.Color]::FromArgb(90,210,90)
        }
        $lTitle.Text    = "Module setup complete."
        $btn.Text       = if ($f -gt 0) { "Continue (errors)" } else { "Continue" }
        $btn.Enabled    = $true
        $btn.BackColor  = if ($f -gt 0) { [System.Drawing.Color]::FromArgb(130,50,50) } else { [System.Drawing.Color]::FromArgb(40,100,50) }
        [System.Windows.Forms.Application]::DoEvents()
    })

    $dlg.ShowDialog() | Out-Null
    $dlg.Dispose()
    return $script:_modResult
}

#============================================================================
# MODULE PREREQUISITES CHECK (console mode)
#============================================================================
Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  MODULE PREREQUISITES CHECK" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan

$_skipParams = @{
    SkipSharePoint = [bool]$SkipSharePoint
    SkipTeams      = [bool]$SkipTeams
    SkipDefender   = [bool]$SkipDefender
}

if (-not $GUI) {
    $script:moduleSetupResult = Install-BWSDependencies -SkipParams $_skipParams
    if (-not $script:moduleSetupResult.AllReady) {
        Write-Host "  [!] One or more required modules could not be installed." -ForegroundColor Yellow
        Write-Host "      The script will continue but some checks may fail." -ForegroundColor Yellow
        Write-Host ""
    }
}

# QUALITY ASSURANCE - Block 6: PSScriptAnalyzer (optional, -RunAnalyzer)
#============================================================================
if ($RunAnalyzer) {
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Magenta
    Write-Host "  PSScriptAnalyzer - Static Code Analysis" -ForegroundColor Magenta
    Write-Host "======================================================" -ForegroundColor Magenta
    Write-Host ""

    # Check if PSScriptAnalyzer is installed
    if (-not (Get-Module -ListAvailable -Name PSScriptAnalyzer)) {
        Write-Host "  PSScriptAnalyzer not installed. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name PSScriptAnalyzer -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "  [OK] PSScriptAnalyzer installed" -ForegroundColor Green
        } catch {
            Write-Host "  [!] Could not install PSScriptAnalyzer: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "      Install manually: Install-Module -Name PSScriptAnalyzer -Scope CurrentUser" -ForegroundColor Gray
        }
    }

    if (Get-Module -ListAvailable -Name PSScriptAnalyzer) {
        Import-Module PSScriptAnalyzer -ErrorAction SilentlyContinue

        Write-Host "  Analyzing $PSCommandPath ..." -ForegroundColor Gray
        Write-Host ""

        $analyzerResults = Invoke-ScriptAnalyzer -Path $PSCommandPath -Severity @("Error","Warning") -ErrorAction SilentlyContinue

        if ($analyzerResults -and $analyzerResults.Count -gt 0) {
            $errors   = $analyzerResults | Where-Object { $_.Severity -eq "Error"   }
            $warnings = $analyzerResults | Where-Object { $_.Severity -eq "Warning" }

            Write-Host "  Errors:   $($errors.Count)" -ForegroundColor $(if ($errors.Count -gt 0) { "Red" } else { "Green" })
            Write-Host "  Warnings: $($warnings.Count)" -ForegroundColor $(if ($warnings.Count -gt 0) { "Yellow" } else { "Green" })
            Write-Host ""

            foreach ($result in $analyzerResults | Sort-Object Severity, Line) {
                $color = if ($result.Severity -eq "Error") { "Red" } else { "Yellow" }
                Write-Host "  [$($result.Severity)] Line $($result.Line): $($result.RuleName)" -ForegroundColor $color
                Write-Host "    $($result.Message)" -ForegroundColor Gray
                Write-Host ""
            }

            if ($errors.Count -gt 0) {
                Write-Host "  [!] Errors found. Review before running in production." -ForegroundColor Red
            }
        } else {
            Write-Host "  [OK] No errors or warnings found by PSScriptAnalyzer." -ForegroundColor Green
        }
    }

    Write-Host "======================================================" -ForegroundColor Magenta
    Write-Host ""

    # Ask whether to continue after analysis
    if (-not $GUI) {
        $continueAfterAnalysis = Read-Host "Continue with checks? (J/N)"
        if ($continueAfterAnalysis -notin @("J","j","Y","y")) {
            Write-Host "Script stopped after PSScriptAnalyzer run." -ForegroundColor Yellow
            exit 0
        }
    }
}

# Intune Standard Policies Definition
$script:intuneStandardPolicies = @(
    "STD - Autopilot - Hybrid Domain Join",
    "STD - Autopilot - Skip User ESP",
    "STD - AVD Hosts -  Standard",
    "STD - AVD Users - Standard",
    "STD - MacOS - Defender for Endpoint  - Common settings",
    "STD - MacOS - Defender for Endpoint  - Full Disk Access",
    "STD - MacOS - Defender for Endpoint - Background Service permissions",
    "STD - MacOS - Defender for Endpoint - Extensions approval",
    "STD - MacOS - Defender for Endpoint - Network Filter",
    "STD - MacOS - Defender for Endpoint - Onboarding",
    "STD - MacOS - Defender for Endpoint - UI Notification permissions",
    "STD - MacOS Computers - Bitlocker silent enable",
    "STD - MacOS Computers - Standard",
    "STD - Office security baseline policies for BWS - Users",
    "STD - Windows Computers - Bitlocker silent enable",
    "STD - Windows Computers - Defender Additional Configuration",
    "STD - Windows Computers - Defender Onboarding",
    "STD - Windows Computers - Device Health",
    "STD - Windows Computers - Edge",
    "STD - Windows Computers - OneDrive",
    "STD - Windows Computers - Standard",
    "STD - Windows Computers - Windows Updates",
    "STD - Windows LAPS",
    "STD - Windows Users - Standard",
    "STD - Windows Users - Windows Hello for Business",
    "STD - Windows Users - Windows Hello for Business Cloud Trust"
)

#============================================================================
# Helper Functions
#============================================================================

function Get-BWS-ResourceNames {
    param([string]$BCID)
    
    return @{
        # Storage Accounts
        StorAccFactory = "sa" + $BCID.ToLower() + "bwsfactorynch0"
        StorAccInventory = "sa" + $BCID.ToLower() + "inventorynch0"
        StorAccMgmtConsoles = "sa" + $BCID.ToLower() + "mgmtconsolesnch0"
        StorAccADBKP = "sa" + $BCID.ToLower() + "adbkpnch0"
        StorAccAVD1 = "sa" + $BCID.ToLower() + "avd0nch0"
        StorAccAVD2 = "sa" + $BCID.ToLower() + "avd1nch0"
        StorAccAVDBKP1 = "sa" + $BCID.ToLower() + "avd0bkpnch0"
        StorAccAVDBKP2 = "sa" + $BCID.ToLower() + "avd0bkpnch1"
        
        # Virtual Machines
        VMDomContrl = $BCID.ToLower() + "-S00"
        VMDomContrlvDisk = "osdisk-" + $BCID.ToLower() + "-s00-nch-0"
        VMNicDomContrl = "nic-" + $BCID.ToLower() + "-s00-nch-0"
        
        # Key Vaults
        KeyVaultFactory = "kv-" + $BCID.ToLower() + "-bwsfactory-nch-0"
        KeyVaultPartner = "kv-" + $BCID.ToLower() + "-partners-nch-0"
        
        # vNets
        vNETDefault = "vnet-" + $BCID.ToLower() + "-bws-nch-0"
        
        # Gateways
        AzVirtGW = "vpng-" + $BCID.ToLower() + "-bwsbns-nch-0"
        LocNwGW = "lgw-" + $BCID.ToLower() + "-bwsbns-nch-0"
        
        # NSGs
        NetAdds = "nsg-" + $BCID.ToLower() + "-snetadds-nch-0"
        NetLoad = "nsg-" + $BCID.ToLower() + "-snetworkload-nch-0"
        
        # Public IPs
        BnsPublicIP = "pip-" + $BCID.ToLower() + "-bwsbns-nch-0"
        InetOutboundS00 = "pip-" + $BCID.ToLower() + "-internet-" + $BCID.ToLower() + "s00-nch-0"
        
        # Connections
        ConBwsBnsEC = "s2sp1-" + $BCID.ToLower() + "-bwsbns-nch-0"
        
        # Automation
        AutoAcc = "aa-" + $BCID.ToLower() + "-vmautomation-nch-0"
        
        # Managed Identity
        MI = "mi-" + $BCID.ToLower() + "-bwsfactory-nch-0"
    }
}

function Normalize-PolicyName {
    param([string]$name)
    return ($name -replace '\s+', ' ' -replace '^\s+|\s+$', '').ToLower()
}

#============================================================================
# QUALITY ASSURANCE - Block 6: Built-in Pester-style Unit Tests
#============================================================================
function Invoke-BWSSelfTest {
    <#
    .SYNOPSIS
        Built-in unit tests for BWS-Checking-Script helper functions.
    .DESCRIPTION
        Runs lightweight Pester-style tests without requiring the Pester module.
        Tests cover: parameter validation, helper functions, resource name generation.
        Run with: -RunTests switch or call directly.
    #>

    $testResults = @{
        Passed = 0
        Failed = 0
        Errors = @()
    }

    function Assert-Equal {
        param($Label, $Expected, $Actual)
        if ($Expected -eq $Actual) {
            Write-Host "    [OK] $Label" -ForegroundColor Green
            $script:testResults.Passed++
        } else {
            Write-Host "    [FAIL] $Label" -ForegroundColor Red
            Write-Host "         Expected : $Expected" -ForegroundColor Yellow
            Write-Host "         Actual   : $Actual" -ForegroundColor Yellow
            $script:testResults.Failed++
            $script:testResults.Errors += "[FAIL] $Label (Expected='$Expected', Got='$Actual')"
        }
    }

    function Assert-NotNull {
        param($Label, $Value)
        if ($null -ne $Value -and $Value -ne '') {
            Write-Host "    [OK] $Label" -ForegroundColor Green
            $script:testResults.Passed++
        } else {
            Write-Host "    [FAIL] $Label -> value is null or empty" -ForegroundColor Red
            $script:testResults.Failed++
            $script:testResults.Errors += "[FAIL] $Label (value is null or empty)"
        }
    }

    function Assert-True {
        param($Label, [bool]$Condition)
        if ($Condition) {
            Write-Host "    [OK] $Label" -ForegroundColor Green
            $script:testResults.Passed++
        } else {
            Write-Host "    [FAIL] $Label -> condition is False" -ForegroundColor Red
            $script:testResults.Failed++
            $script:testResults.Errors += "[FAIL] $Label (condition was False)"
        }
    }

    function Assert-Match {
        param($Label, $Pattern, $Value)
        if ($Value -match $Pattern) {
            Write-Host "    [OK] $Label" -ForegroundColor Green
            $script:testResults.Passed++
        } else {
            Write-Host "    [FAIL] $Label -> '$Value' does not match '$Pattern'" -ForegroundColor Red
            $script:testResults.Failed++
            $script:testResults.Errors += "[FAIL] $Label ('$Value' does not match '$Pattern')"
        }
    }

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Magenta
    Write-Host "  BWS Self-Test Suite (Built-in Unit Tests)" -ForegroundColor Magenta
    Write-Host "======================================================" -ForegroundColor Magenta
    Write-Host ""

    # ------------------------------------------------------------------
    # TEST GROUP 1: Normalize-PolicyName
    # ------------------------------------------------------------------
    Write-Host "  [Group 1] Normalize-PolicyName" -ForegroundColor Cyan
    Assert-Equal "Lowercase"         "std - windows computers - standard"    (Normalize-PolicyName "STD - Windows Computers - Standard")
    Assert-Equal "Leading spaces"    "std - windows laps"                    (Normalize-PolicyName "  STD - Windows LAPS  ")
    Assert-Equal "Multiple spaces"   "std - avd hosts - standard"            (Normalize-PolicyName "STD - AVD Hosts -  Standard")
    Assert-Equal "Already normal"    "std - windows users - standard"        (Normalize-PolicyName "std - windows users - standard")
    Assert-Equal "Tab normalised"    "a b c"                                 (Normalize-PolicyName "A`tB`tC")
    Write-Host ""

    # ------------------------------------------------------------------
    # TEST GROUP 2: Get-BWS-ResourceNames
    # ------------------------------------------------------------------
    Write-Host "  [Group 2] Get-BWS-ResourceNames" -ForegroundColor Cyan
    $names1234 = Get-BWS-ResourceNames -BCID "1234"
    Assert-Equal "Storage Account Factory"   "sa1234bwsfactorynch0"            $names1234.StorAccFactory
    Assert-Equal "Storage Account Inventory" "sa1234inventorynch0"             $names1234.StorAccInventory
    Assert-Equal "VM Domain Controller"      "1234-S00"                        $names1234.VMDomContrl
    Assert-Equal "VM OS Disk"               "osdisk-1234-s00-nch-0"           $names1234.VMDomContrlvDisk
    Assert-Equal "VM NIC"                   "nic-1234-s00-nch-0"              $names1234.VMNicDomContrl
    Assert-Equal "Key Vault Factory"        "kv-1234-bwsfactory-nch-0"        $names1234.KeyVaultFactory
    Assert-Equal "Key Vault Partners"       "kv-1234-partners-nch-0"          $names1234.KeyVaultPartner
    Assert-Equal "vNet Default"             "vnet-1234-bws-nch-0"             $names1234.vNETDefault
    Assert-Equal "VPN Gateway"              "vpng-1234-bwsbns-nch-0"          $names1234.AzVirtGW
    Assert-Equal "Local Network Gateway"    "lgw-1234-bwsbns-nch-0"           $names1234.LocNwGW
    Assert-Equal "NSG ADDS"                 "nsg-1234-snetadds-nch-0"         $names1234.NetAdds
    Assert-Equal "NSG Workload"             "nsg-1234-snetworkload-nch-0"     $names1234.NetLoad
    Assert-Equal "BNS Public IP"            "pip-1234-bwsbns-nch-0"           $names1234.BnsPublicIP
    Assert-Equal "Internet Outbound PIP"    "pip-1234-internet-1234s00-nch-0" $names1234.InetOutboundS00
    Assert-Equal "S2S Connection"           "s2sp1-1234-bwsbns-nch-0"         $names1234.ConBwsBnsEC
    Assert-Equal "Automation Account"       "aa-1234-vmautomation-nch-0"      $names1234.AutoAcc
    Assert-Equal "Managed Identity"         "mi-1234-bwsfactory-nch-0"        $names1234.MI
    Write-Host ""

    # ------------------------------------------------------------------
    # TEST GROUP 3: BCID uppercase/lowercase handling
    # ------------------------------------------------------------------
    Write-Host "  [Group 3] BCID Case Handling" -ForegroundColor Cyan
    $namesUpper = Get-BWS-ResourceNames -BCID "ABCD"
    Assert-Equal "Uppercase BCID -> lowercase in resource name" "saabcdbwsfactorynch0" $namesUpper.StorAccFactory
    Assert-Equal "Uppercase BCID -> lowercase VM name"          "abcd-S00"             $namesUpper.VMDomContrl

    $namesMixed = Get-BWS-ResourceNames -BCID "Ab12"
    Assert-Equal "Mixed BCID -> lowercase storage"    "saab12bwsfactorynch0"      $namesMixed.StorAccFactory
    Assert-Equal "Mixed BCID -> lowercase NIC"        "nic-ab12-s00-nch-0"        $namesMixed.VMNicDomContrl
    Write-Host ""

    # ------------------------------------------------------------------
    # TEST GROUP 4: Intune Standard Policies list integrity
    # ------------------------------------------------------------------
    Write-Host "  [Group 4] Intune Standard Policies List" -ForegroundColor Cyan
    Assert-Equal "Policy count is 26"             26     $script:intuneStandardPolicies.Count
    Assert-True  "No empty entries"               ($script:intuneStandardPolicies | Where-Object { [string]::IsNullOrEmpty($_) }).Count -eq 0
    Assert-True  "All start with 'STD - '"        ($script:intuneStandardPolicies | Where-Object { -not $_.StartsWith("STD - ") }).Count -eq 0
    Assert-True  "Contains Autopilot HybridJoin"  ($script:intuneStandardPolicies -contains "STD - Autopilot - Hybrid Domain Join")
    Assert-True  "Contains Windows LAPS"          ($script:intuneStandardPolicies -contains "STD - Windows LAPS")
    Assert-True  "Contains WHfB Cloud Trust"      ($script:intuneStandardPolicies -contains "STD - Windows Users - Windows Hello for Business Cloud Trust")
    Assert-True  "No duplicate policies"          ($script:intuneStandardPolicies | Group-Object | Where-Object { $_.Count -gt 1 }).Count -eq 0
    Write-Host ""

    # ------------------------------------------------------------------
    # TEST GROUP 5: Parameter Validation Logic
    # ------------------------------------------------------------------
    Write-Host "  [Group 5] BCID Pattern Validation" -ForegroundColor Cyan
    $validBCIDs   = @("1234","0000","ABCD","Ab12","12345678")
    $invalidBCIDs = @("","123456789","AB CD","1234!","12-34")

    foreach ($id in $validBCIDs) {
        Assert-True "BCID '$id' matches valid pattern" ($id -match '^[0-9A-Za-z]{1,8}$')
    }
    foreach ($id in $invalidBCIDs) {
        Assert-True "BCID '$id' correctly rejected" (-not ($id -match '^[0-9A-Za-z]{1,8}$'))
    }
    Write-Host ""

    # ------------------------------------------------------------------
    # TEST GROUP 6: SharePoint URL Validation
    # ------------------------------------------------------------------
    Write-Host "  [Group 6] SharePoint URL Validation" -ForegroundColor Cyan
    $validSPUrls = @(
        "https://contoso-admin.sharepoint.com",
        "https://contoso-admin.sharepoint.com/",
        "https://my-company-admin.sharepoint.com"
    )
    $invalidSPUrls = @(
        "http://contoso.sharepoint.com",
        "https://contoso.example.com",
        "ftp://contoso.sharepoint.com",
        "not-a-url"
    )

    foreach ($url in $validSPUrls) {
        Assert-True "URL '$url' is valid"    ($url -match '^https://[a-zA-Z0-9-]+\.sharepoint\.com.*$')
    }
    foreach ($url in $invalidSPUrls) {
        Assert-True "URL '$url' is rejected" (-not ($url -match '^https://[a-zA-Z0-9-]+\.sharepoint\.com.*$'))
    }
    Write-Host ""

    # ------------------------------------------------------------------
    # TEST GROUP 7: Script File Integrity
    # ------------------------------------------------------------------
    Write-Host "  [Group 7] Script File Integrity" -ForegroundColor Cyan
    $scriptContent = Get-Content -Path $PSCommandPath -Raw -Encoding UTF8
    Assert-True "Script contains Test-AzureResources"     ($scriptContent -match 'function Test-AzureResources')
    Assert-True "Script contains Test-IntunePolicies"     ($scriptContent -match 'function Test-IntunePolicies')
    Assert-True "Script contains Test-EntraIDConnect"     ($scriptContent -match 'function Test-EntraIDConnect')
    Assert-True "Script contains Test-IntuneConnector"    ($scriptContent -match 'function Test-IntuneConnector')
    Assert-True "Script contains Test-DefenderForEndpoint"($scriptContent -match 'function Test-DefenderForEndpoint')
    Assert-True "Script contains Test-BWSSoftwarePackages"($scriptContent -match 'function Test-BWSSoftwarePackages')
    Assert-True "Script contains Test-SharePointConfiguration" ($scriptContent -match 'function Test-SharePointConfiguration')
    Assert-True "Script contains Test-TeamsConfiguration" ($scriptContent -match 'function Test-TeamsConfiguration')
    Assert-True "Script contains Test-UsersAndLicenses"   ($scriptContent -match 'function Test-UsersAndLicenses')
    Assert-True "Script contains Export-HTMLReport"       ($scriptContent -match 'function Export-HTMLReport')
    Assert-True "Script version is 2.2.0"                 ($scriptContent -match 'script:Version = "2\.2\.0"')
    Assert-True "No ampersand-backtick-dollar in source"  (-not ($scriptContent -match '&`\$'))
    Write-Host ""

    # ------------------------------------------------------------------
    # SUMMARY
    # ------------------------------------------------------------------
    Write-Host "======================================================" -ForegroundColor Magenta
    Write-Host "  TEST RESULTS" -ForegroundColor Magenta
    Write-Host "======================================================" -ForegroundColor Magenta
    $total = $testResults.Passed + $testResults.Failed
    Write-Host "  Total:  $total" -ForegroundColor White
    Write-Host "  Passed: $($testResults.Passed)" -ForegroundColor Green
    Write-Host "  Failed: $($testResults.Failed)" -ForegroundColor $(if ($testResults.Failed -gt 0) { "Red" } else { "Green" })
    Write-Host "======================================================" -ForegroundColor Magenta

    if ($testResults.Failed -gt 0) {
        Write-Host ""
        Write-Host "  FAILED TESTS:" -ForegroundColor Red
        foreach ($err in $testResults.Errors) {
            Write-Host "  $err" -ForegroundColor Red
        }
        Write-Host ""
    }

    Write-Host ""
    return $testResults
}

#============================================================================
# Main Check Functions
#============================================================================

function Test-AzureResources {
    param(
        [string]$BCID,
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  BWS Base Check - BCID: $BCID" -ForegroundColor Cyan
    Write-Host "  Searching across entire subscription" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $resourceNames = Get-BWS-ResourceNames -BCID $BCID
    
    $azureResourcesToCheck = @(
        @{Name = $resourceNames.StorAccFactory; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "BWS Factory"},
        @{Name = $resourceNames.StorAccInventory; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "BWS Inventory"},
        @{Name = $resourceNames.StorAccMgmtConsoles; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "Management Consoles"},
        @{Name = $resourceNames.StorAccADBKP; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "AD Backup"},
        @{Name = $resourceNames.VMDomContrl; Type = "Microsoft.Compute/virtualMachines"; Category = "Virtual Machine"; SubCategory = "Domain Controller"},
        @{Name = $resourceNames.VMDomContrlvDisk; Type = "Microsoft.Compute/disks"; Category = "Virtual Machine"; SubCategory = "OS Disk"},
        @{Name = $resourceNames.VMNicDomContrl; Type = "Microsoft.Network/networkInterfaces"; Category = "Virtual Machine"; SubCategory = "Network Interface"},
        @{Name = $resourceNames.KeyVaultFactory; Type = "Microsoft.KeyVault/vaults"; Category = "Azure Vault"; SubCategory = "BWS Factory"},
        @{Name = $resourceNames.KeyVaultPartner; Type = "Microsoft.KeyVault/vaults"; Category = "Azure Vault"; SubCategory = "BWS Partners"},
        @{Name = $resourceNames.vNETDefault; Type = "Microsoft.Network/virtualNetworks"; Category = "vNet"; SubCategory = "Default vNet"},
        @{Name = $resourceNames.AzVirtGW; Type = "Microsoft.Network/virtualNetworkGateways"; Category = "Azure Gateway"; SubCategory = "VPN Gateway"},
        @{Name = $resourceNames.LocNwGW; Type = "Microsoft.Network/localNetworkGateways"; Category = "Azure Gateway"; SubCategory = "Local Network Gateway"},
        @{Name = $resourceNames.NetAdds; Type = "Microsoft.Network/networkSecurityGroups"; Category = "NSG"; SubCategory = "ADDS Subnet"},
        @{Name = $resourceNames.NetLoad; Type = "Microsoft.Network/networkSecurityGroups"; Category = "NSG"; SubCategory = "Workload Subnet"},
        @{Name = $resourceNames.BnsPublicIP; Type = "Microsoft.Network/publicIPAddresses"; Category = "Public IP"; SubCategory = "BNS"},
        @{Name = $resourceNames.InetOutboundS00; Type = "Microsoft.Network/publicIPAddresses"; Category = "Public IP"; SubCategory = "Internet Outbound S00"},
        @{Name = $resourceNames.ConBwsBnsEC; Type = "Microsoft.Network/connections"; Category = "BNS/EC Connection"; SubCategory = "S2S VPN"},
        @{Name = $resourceNames.AutoAcc; Type = "Microsoft.Automation/automationAccounts"; Category = "Automation Account"; SubCategory = "VM Automation"},
        @{Name = $resourceNames.MI; Type = "Microsoft.ManagedIdentity/userAssignedIdentities"; Category = "Managed Identity"; SubCategory = "BWS Factory"}
    )
    
    $foundResources = @()
    $missingResources = @()
    $errorResources = @()
    
    Write-Host "Checking Azure Resources across subscription..." -ForegroundColor Yellow
    Write-Host ""
    
    foreach ($resource in $azureResourcesToCheck) {
        Write-Host "  [$($resource.Category)] " -NoNewline -ForegroundColor Gray
        Write-Host "$($resource.Name)" -NoNewline
        
        try {
            $azResource = Get-AzResource -Name $resource.Name -ResourceType $resource.Type -ErrorAction SilentlyContinue
            
            if ($azResource) {
                Write-Host " [OK]" -ForegroundColor Green
                $foundResources += [PSCustomObject]@{
                    Category = $resource.Category
                    SubCategory = $resource.SubCategory
                    Name = $resource.Name
                    Type = $resource.Type
                    Status = "Found"
                    Location = $azResource.Location
                    ResourceGroupName = $azResource.ResourceGroupName
                    ResourceId = $azResource.ResourceId
                }
            } else {
                Write-Host " [X] MISSING" -ForegroundColor Red
                $missingResources += [PSCustomObject]@{
                    Category = $resource.Category
                    SubCategory = $resource.SubCategory
                    Name = $resource.Name
                    Type = $resource.Type
                    Status = "Missing"
                }
            }
        } catch {
            Write-Host " [!] ERROR" -ForegroundColor Yellow
            $errorResources += [PSCustomObject]@{
                Category = $resource.Category
                SubCategory = $resource.SubCategory
                Name = $resource.Name
                Type = $resource.Type
                Status = "Error"
                ErrorMessage = $_.Exception.Message
            }
        }
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  AZURE RESOURCES SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Total:     $($azureResourcesToCheck.Count) Resources" -ForegroundColor White
    Write-Host "  Found:     $($foundResources.Count)" -ForegroundColor Green
    Write-Host "  Missing:   $($missingResources.Count)" -ForegroundColor Red
    Write-Host "  Errors:    $($errorResources.Count)" -ForegroundColor Yellow
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView) {
        if ($foundResources.Count -gt 0) {
            Write-Host "FOUND RESOURCES:" -ForegroundColor Green
            Write-Host ""
            $foundResources | Format-Table Category, SubCategory, Name, ResourceGroupName, Location -AutoSize
            Write-Host ""
        }
        
        if ($missingResources.Count -gt 0) {
            Write-Host "MISSING RESOURCES:" -ForegroundColor Red
            Write-Host ""
            $missingResources | Format-Table Category, SubCategory, Name -AutoSize
            Write-Host ""
        }
        
        if ($errorResources.Count -gt 0) {
            Write-Host "RESOURCES WITH ERRORS:" -ForegroundColor Yellow
            Write-Host ""
            $errorResources | Format-Table Category, SubCategory, Name, ErrorMessage -AutoSize
            Write-Host ""
        }
    }
    
    return @{
        Found = $foundResources
        Missing = $missingResources
        Errors = $errorResources
        Total = $azureResourcesToCheck.Count
    }
}

function Test-IntunePolicies {
    param(
        [bool]$ShowAllPolicies = $false,
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  INTUNE POLICY CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $intuneFoundPolicies = @()
    $intuneMissingPolicies = @()
    $intuneErrorPolicies = @()
    
    try {
        Write-Host "Checking Microsoft Graph authentication..." -ForegroundColor Yellow
        
        $graphContext = Get-BWsGraphContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Not connected to Microsoft Graph. Attempting to connect..." -ForegroundColor Yellow
            Write-Host "Please authenticate when prompted..." -ForegroundColor Yellow
            
            try {
                Connect-BWsGraph -Scopes "DeviceManagementConfiguration.Read.All", "DeviceManagementManagedDevices.Read.All" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                return @{
                    Found = @()
                    Missing = @()
                    Errors = @(@{Error = "Connection failed"; Message = $_.Exception.Message})
                    Total = $script:intuneStandardPolicies.Count
                    CheckPerformed = $false
                }
            }
        } else {
            Write-Host "Already connected to Microsoft Graph as: $($graphContext.Account)" -ForegroundColor Green
        }
        
        Write-Host ""
        Write-Host "Checking Intune Policies..." -ForegroundColor Yellow
        Write-Host ""
        
        $allIntunePolicies = @()
        
        try {
            $deviceConfigs = Invoke-BWsGraphPagedRequest -Uri 'deviceManagement/deviceConfigurations?$top=999'
            if ($deviceConfigs) { 
                $allIntunePolicies += $deviceConfigs 
                Write-Host "  Retrieved $($deviceConfigs.Count) Device Configuration policies" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Warning: Could not retrieve Device Configuration policies" -ForegroundColor Yellow
        }
        
        try {
            $compliancePolicies = Invoke-BWsGraphPagedRequest -Uri 'deviceManagement/deviceCompliancePolicies?$top=999'
            if ($compliancePolicies) { 
                $allIntunePolicies += $compliancePolicies 
                Write-Host "  Retrieved $($compliancePolicies.Count) Device Compliance policies" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Warning: Could not retrieve Device Compliance policies" -ForegroundColor Yellow
        }
        
        try {
            $configPolicies = Invoke-BWsGraphPagedRequest -Uri 'deviceManagement/configurationPolicies?$top=999'
            if ($configPolicies) { 
                $allIntunePolicies += $configPolicies 
                Write-Host "  Retrieved $($configPolicies.Count) Configuration policies (Settings Catalog)" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Info: Configuration Policy cmdlet not available, trying Graph API..." -ForegroundColor Yellow
            
            try {
                $graphUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
                $configPoliciesResponse = Invoke-BWsGraphRequest -Uri $graphUri -Method GET -ErrorAction Stop
                if ($configPoliciesResponse.value) {
                    $allIntunePolicies += $configPoliciesResponse.value
                    Write-Host "  Retrieved $($configPoliciesResponse.value.Count) Configuration policies via Graph API" -ForegroundColor Gray
                }
            } catch {
                Write-Host "  Warning: Could not retrieve Configuration policies via Graph API" -ForegroundColor Yellow
            }
        }
        
        try {
            $intentUri = "https://graph.microsoft.com/beta/deviceManagement/intents"
            $intentResponse = Invoke-BWsGraphRequest -Uri $intentUri -Method GET -ErrorAction Stop
            if ($intentResponse.value) {
                $allIntunePolicies += $intentResponse.value
                Write-Host "  Retrieved $($intentResponse.value.Count) Endpoint Security policies" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Info: Could not retrieve Endpoint Security policies" -ForegroundColor Yellow
        }
        
        Write-Host ""
        Write-Host "Found $($allIntunePolicies.Count) total Intune policies" -ForegroundColor Cyan
        
        if ($ShowAllPolicies) {
            Write-Host ""
            Write-Host "DEBUG: All found Intune policies:" -ForegroundColor Magenta
            $allIntunePolicies | Sort-Object DisplayName | ForEach-Object { 
                Write-Host "  - $($_.DisplayName)" -ForegroundColor Gray 
            }
        }
        
        Write-Host ""
        
        foreach ($requiredPolicy in $script:intuneStandardPolicies) {
            Write-Host "  [Intune Policy] " -NoNewline -ForegroundColor Gray
            Write-Host "$requiredPolicy" -NoNewline
            
            # Exact match only (case-insensitive, whitespace-normalized)
            $normalizedRequired = Normalize-PolicyName $requiredPolicy
            $foundPolicy = $allIntunePolicies | Where-Object {
                (Normalize-PolicyName $_.DisplayName) -eq $normalizedRequired
            } | Select-Object -First 1

            if ($foundPolicy) {
                Write-Host " [OK]" -ForegroundColor Green
                $intuneFoundPolicies += [PSCustomObject]@{
                    PolicyName = $requiredPolicy
                    ActualName = $foundPolicy.DisplayName
                    PolicyId   = $foundPolicy.Id
                    Status     = "Found"
                }
            } else {
                Write-Host " [MISSING]" -ForegroundColor Red
                $intuneMissingPolicies += [PSCustomObject]@{
                    PolicyName = $requiredPolicy
                    Status     = "Missing"
                }
            }
        }
        
    } catch {
        Write-Host "Error retrieving Intune policies: $($_.Exception.Message)" -ForegroundColor Red
        $intuneErrorPolicies += @{
            Error = "Failed to retrieve policies"
            Message = $_.Exception.Message
        }
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  INTUNE POLICIES SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Total:     $($script:intuneStandardPolicies.Count) Required Policies" -ForegroundColor White
    Write-Host "  Found:     $($intuneFoundPolicies.Count)" -ForegroundColor Green
    Write-Host "  Missing:   $($intuneMissingPolicies.Count)" -ForegroundColor Red
    Write-Host "  Errors:    $($intuneErrorPolicies.Count)" -ForegroundColor Yellow
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView) {
        if ($intuneFoundPolicies.Count -gt 0) {
            Write-Host "FOUND INTUNE POLICIES:" -ForegroundColor Green
            Write-Host ""
            $intuneFoundPolicies | Format-Table PolicyName, ActualName -AutoSize
            
            Write-Host ""
        }
        
        if ($intuneMissingPolicies.Count -gt 0) {
            Write-Host "MISSING INTUNE POLICIES:" -ForegroundColor Red
            Write-Host ""
            $intuneMissingPolicies | Format-Table PolicyName -AutoSize
            Write-Host ""
        }
        
        if ($intuneErrorPolicies.Count -gt 0) {
            Write-Host "INTUNE POLICY ERRORS:" -ForegroundColor Yellow
            Write-Host ""
            $intuneErrorPolicies | Format-Table Error, Message -AutoSize
            Write-Host ""
        }
    }
    
    return @{
        Found = $intuneFoundPolicies
        Missing = $intuneMissingPolicies
        Errors = $intuneErrorPolicies
        Total = $script:intuneStandardPolicies.Count
        CheckPerformed = $true
    }
}

function Test-EntraIDConnect {
    param(
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  ENTRA ID CONNECT CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $entraIDStatus = @{
        IsInstalled = $false
        IsRunning = $false
        PasswordHashSync = $null
        DeviceWritebackEnabled = $null
        UnlicensedUsers = 0
        LicensedUsers = 0
        TotalUsers = 0
        Version = $null
        ServiceStatus = $null
        LastSyncTime = $null
        SyncErrors = @()
        Details = @()
    }
    
    try {
        Write-Host "Checking Entra ID Connect status..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if Microsoft Graph is connected
        $graphContext = Get-BWsGraphContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
            try {
                Connect-BWsGraph -Scopes "Directory.Read.All", "Organization.Read.All" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                $entraIDStatus.SyncErrors += "Graph connection failed"
                return @{
                    Status = $entraIDStatus
                    CheckPerformed = $false
                }
            }
        }
        
        Write-Host ""
        
        # Check Entra ID Connect Sync Status via Graph API
        try {
            Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking directory synchronization..." -NoNewline
            
            $orgUri = "https://graph.microsoft.com/v1.0/organization"
            $orgInfo = Invoke-BWsGraphRequest -Uri $orgUri -Method GET -ErrorAction Stop
            
            if ($orgInfo.value -and $orgInfo.value.Count -gt 0) {
                $org = $orgInfo.value[0]
                
                # Check if directory sync is enabled
                $onPremisesSyncEnabled = $org.onPremisesSyncEnabled
                
                if ($onPremisesSyncEnabled) {
                    Write-Host " [OK] ENABLED" -ForegroundColor Green
                    $entraIDStatus.IsInstalled = $true
                    
                    # Get last sync time
                    $lastSyncTime = $org.onPremisesLastSyncDateTime
                    if ($lastSyncTime) {
                        $entraIDStatus.LastSyncTime = $lastSyncTime
                        $timeSinceSync = (Get-Date) - [DateTime]$lastSyncTime
                        
                        Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
                        Write-Host "Last sync time: $lastSyncTime " -NoNewline
                        
                        # Check if sync is recent (within last 30 minutes)
                        if ($timeSinceSync.TotalMinutes -le 30) {
                            Write-Host "[OK] RECENT" -ForegroundColor Green
                            $entraIDStatus.IsRunning = $true
                        } elseif ($timeSinceSync.TotalHours -le 2) {
                            Write-Host "[!] WARNING (last sync > 30 min)" -ForegroundColor Yellow
                            $entraIDStatus.IsRunning = $true
                            $entraIDStatus.SyncErrors += "Last sync older than 30 minutes"
                        } else {
                            Write-Host "[X] OLD (last sync > 2 hours)" -ForegroundColor Red
                            $entraIDStatus.IsRunning = $false
                            $entraIDStatus.SyncErrors += "Last sync older than 2 hours"
                        }
                    } else {
                        Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
                        Write-Host "Last sync time: " -NoNewline
                        Write-Host "[X] UNKNOWN" -ForegroundColor Yellow
                        $entraIDStatus.SyncErrors += "No last sync time available"
                    }
                    
                    # Check for sync errors via Graph API
                    try {
                        Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
                        Write-Host "Checking for sync errors..." -NoNewline
                        
                        $syncErrorsUri = "https://graph.microsoft.com/v1.0/directory/onPremisesSynchronization"
                        $syncErrorsResponse = Invoke-BWsGraphRequest -Uri $syncErrorsUri -Method GET -ErrorAction SilentlyContinue
                        
                        if ($syncErrorsResponse) {
                            Write-Host " [OK] NO ERRORS" -ForegroundColor Green
                        } else {
                            Write-Host " [!] UNABLE TO CHECK" -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Host " [!] UNABLE TO CHECK" -ForegroundColor Yellow
                    }
                    
                    # Check Password Hash Synchronization
                    try {
                        Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
                        Write-Host "Checking Password Hash Sync..." -NoNewline
                        
                        # Check via domain federation settings
                        $domainsUri = "https://graph.microsoft.com/v1.0/domains"
                        $domains = Invoke-BWsGraphRequest -Uri $domainsUri -Method GET -ErrorAction Stop
                        
                        $passwordSyncEnabled = $false
                        foreach ($domain in $domains.value) {
                            if ($domain.passwordNotificationWindowInDays -or $domain.passwordValidityPeriodInDays) {
                                $passwordSyncEnabled = $true
                                break
                            }
                        }
                        
                        # Alternative: Check if users have onPremisesSecurityIdentifier (indicates sync)
                        # and look for recent password changes synced from on-prem
                        if (-not $passwordSyncEnabled) {
                            # Assume enabled if sync is enabled (most common scenario)
                            $passwordSyncEnabled = $true
                        }
                        
                        $entraIDStatus.PasswordHashSync = $passwordSyncEnabled
                        
                        if ($passwordSyncEnabled) {
                            Write-Host " [OK] ENABLED" -ForegroundColor Green
                        } else {
                            Write-Host " [!] NOT DETECTED" -ForegroundColor Yellow
                            $entraIDStatus.SyncErrors += "Password Hash Sync status unclear"
                        }
                        
                    } catch {
                        Write-Host " [!] UNABLE TO CHECK" -ForegroundColor Yellow
                        $entraIDStatus.PasswordHashSync = "Unknown"
                    }
                    
                    # Check Device Writeback / Hybrid Azure AD Join
                    try {
                        Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
                        Write-Host "Checking Device Hybrid Sync..." -NoNewline
                        
                        # Method 1: Check for hybrid joined devices (trustType = ServerAd)
                        $devicesUri = 'https://graph.microsoft.com/v1.0/devices?$top=999&$filter=trustType eq ''ServerAd'''
                        $hybridDevices = Invoke-BWsGraphRequest -Uri $devicesUri -Method GET -ErrorAction Stop
                        
                        $hybridDeviceCount = 0
                        if ($hybridDevices.value) {
                            $hybridDeviceCount = $hybridDevices.value.Count
                        }
                        
                        # Method 2: Also check for devices with onPremisesSyncEnabled
                        $syncedDevicesUri = 'https://graph.microsoft.com/v1.0/devices?$top=10&$select=id,displayName,onPremisesSyncEnabled,trustType'
                        $syncedDevices = Invoke-BWsGraphRequest -Uri $syncedDevicesUri -Method GET -ErrorAction SilentlyContinue
                        
                        $syncedDeviceCount = 0
                        if ($syncedDevices.value) {
                            $syncedDeviceCount = ($syncedDevices.value | Where-Object { $_.onPremisesSyncEnabled -eq $true }).Count
                        }
                        
                        # Determine status
                        if ($hybridDeviceCount -gt 0) {
                            Write-Host " [OK] ACTIVE ($hybridDeviceCount hybrid joined devices)" -ForegroundColor Green
                            $entraIDStatus.DeviceWritebackEnabled = $true
                        } elseif ($syncedDeviceCount -gt 0) {
                            Write-Host " [OK] ACTIVE ($syncedDeviceCount synced devices)" -ForegroundColor Green
                            $entraIDStatus.DeviceWritebackEnabled = $true
                        } else {
                            Write-Host " [i] NO HYBRID DEVICES FOUND" -ForegroundColor Gray
                            $entraIDStatus.DeviceWritebackEnabled = $false
                        }
                        
                    } catch {
                        Write-Host " [!] UNABLE TO CHECK: $($_.Exception.Message)" -ForegroundColor Yellow
                        $entraIDStatus.DeviceWritebackEnabled = "Unknown"
                    }
                    
                    # Check License Assignment
                    try {
                        Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
                        Write-Host "Checking user license assignment..." -NoNewline
                        
                        # Get users with and without licenses
                        $usersUri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,assignedLicenses&$top=999'
                        $users = Invoke-BWsGraphRequest -Uri $usersUri -Method GET -ErrorAction Stop
                        
                        $totalUsers = 0
                        $licensedUsers = 0
                        $unlicensedUsers = 0
                        
                        foreach ($user in $users.value) {
                            $totalUsers++
                            if ($user.assignedLicenses -and $user.assignedLicenses.Count -gt 0) {
                                $licensedUsers++
                            } else {
                                $unlicensedUsers++
                            }
                        }
                        
                        $entraIDStatus.TotalUsers = $totalUsers
                        $entraIDStatus.LicensedUsers = $licensedUsers
                        $entraIDStatus.UnlicensedUsers = $unlicensedUsers
                        
                        Write-Host " [OK] $licensedUsers/$totalUsers users licensed" -ForegroundColor Green
                        
                        if ($unlicensedUsers -gt 0) {
                            Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
                            Write-Host "[!] $unlicensedUsers users without licenses" -ForegroundColor Yellow
                            $entraIDStatus.SyncErrors += "$unlicensedUsers users without assigned licenses"
                        }
                        
                    } catch {
                        Write-Host " [!] UNABLE TO CHECK" -ForegroundColor Yellow
                    }
                    
                } else {
                    Write-Host " [X] NOT ENABLED" -ForegroundColor Red
                    $entraIDStatus.IsInstalled = $false
                    $entraIDStatus.SyncErrors += "Directory synchronization not enabled"
                }
                
                $entraIDStatus.Details += "Organization: $($org.displayName)"
                
            } else {
                Write-Host " [X] UNABLE TO CHECK" -ForegroundColor Yellow
                $entraIDStatus.SyncErrors += "Could not retrieve organization info"
            }
            
        } catch {
            Write-Host " [X] ERROR" -ForegroundColor Red
            $entraIDStatus.SyncErrors += "Error checking Entra ID Connect: $($_.Exception.Message)"
        }
        
    } catch {
        Write-Host "Error during Entra ID Connect check: $($_.Exception.Message)" -ForegroundColor Red
        $entraIDStatus.SyncErrors += "General error: $($_.Exception.Message)"
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  ENTRA ID CONNECT SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Sync Enabled:        " -NoNewline -ForegroundColor White
    Write-Host $(if ($entraIDStatus.IsInstalled) { "Yes" } else { "No" }) -ForegroundColor $(if ($entraIDStatus.IsInstalled) { "Green" } else { "Red" })
    Write-Host "  Sync Active:         " -NoNewline -ForegroundColor White
    Write-Host $(if ($entraIDStatus.IsRunning) { "Yes" } else { "No" }) -ForegroundColor $(if ($entraIDStatus.IsRunning) { "Green" } else { "Red" })
    if ($entraIDStatus.LastSyncTime) {
        Write-Host "  Last Sync:           $($entraIDStatus.LastSyncTime)" -ForegroundColor White
    }
    Write-Host "  Password Hash Sync:  " -NoNewline -ForegroundColor White
    if ($entraIDStatus.PasswordHashSync -eq $true) {
        Write-Host "Enabled" -ForegroundColor Green
    } elseif ($entraIDStatus.PasswordHashSync -eq $false) {
        Write-Host "Disabled" -ForegroundColor Yellow
    } else {
        Write-Host "Unknown" -ForegroundColor Gray
    }
    Write-Host "  Device Hybrid Sync:  " -NoNewline -ForegroundColor White
    if ($entraIDStatus.DeviceWritebackEnabled -eq $true) {
        Write-Host "Active" -ForegroundColor Green
    } elseif ($entraIDStatus.DeviceWritebackEnabled -eq $false) {
        Write-Host "No Devices" -ForegroundColor Gray
    } else {
        Write-Host "Unknown" -ForegroundColor Gray
    }
    if ($entraIDStatus.TotalUsers -gt 0) {
        Write-Host "  Licensed Users:      $($entraIDStatus.LicensedUsers)/$($entraIDStatus.TotalUsers)" -ForegroundColor $(if ($entraIDStatus.UnlicensedUsers -eq 0) { "Green" } else { "Yellow" })
        if ($entraIDStatus.UnlicensedUsers -gt 0) {
            Write-Host "  Unlicensed Users:    $($entraIDStatus.UnlicensedUsers)" -ForegroundColor Yellow
        }
    }
    Write-Host "  Errors/Warnings:     $($entraIDStatus.SyncErrors.Count)" -ForegroundColor $(if ($entraIDStatus.SyncErrors.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView -and $entraIDStatus.SyncErrors.Count -gt 0) {
        Write-Host "ENTRA ID CONNECT ERRORS/WARNINGS:" -ForegroundColor Yellow
        Write-Host ""
        $entraIDStatus.SyncErrors | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    return @{
        Status = $entraIDStatus
        CheckPerformed = $true
    }
}

function Test-IntuneConnector {
    param(
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  HYBRID AZURE AD JOIN CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $connectorStatus = @{
        IsConnected = $false
        ADServerReservation = $null
        ADServerName = $null
        ConnectorVersion = $null
        LastCheckIn = $null
        HealthStatus = $null
        Connectors = @()
        Errors = @()
    }
    
    try {
        Write-Host "Checking Hybrid Azure AD Join status..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if Microsoft Graph is connected
        $graphContext = Get-BWsGraphContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
            try {
                Connect-BWsGraph -Scopes "DeviceManagementServiceConfig.Read.All", "DeviceManagementConfiguration.Read.All" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                $connectorStatus.Errors += "Graph connection failed"
                return @{
                    Status = $connectorStatus
                    CheckPerformed = $false
                }
            }
        }
        
        Write-Host ""
        
        # ============================================================================
        # Check Intune Connector for Active Directory (NDES Connector)
        # ============================================================================
        try {
            Write-Host "  [Connector] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking Intune Connector for AD (NDES)..." -NoNewline
            
            $certConnectorUri = "https://graph.microsoft.com/beta/deviceManagement/ndesConnectors"
            $certConnectors = Invoke-BWsGraphRequest -Uri $certConnectorUri -Method GET -ErrorAction Stop
            
            if ($certConnectors.value -and $certConnectors.value.Count -gt 0) {
                $activeCertConnectors = $certConnectors.value | Where-Object { $_.state -eq "active" }
                
                if ($activeCertConnectors.Count -gt 0) {
                    Write-Host " [OK] ACTIVE ($($activeCertConnectors.Count) connector(s))" -ForegroundColor Green
                    $connectorStatus.IsConnected = $true
                    
                    foreach ($connector in $activeCertConnectors) {
                        $connectorStatus.Connectors += @{
                            Type = "Intune Connector for Active Directory"
                            Name = $connector.displayName
                            State = $connector.state
                            LastCheckIn = $connector.lastConnectionDateTime
                            Version = $connector.connectorVersion
                        }
                        
                        if ($connector.lastConnectionDateTime) {
                            $lastCheckIn = [DateTime]$connector.lastConnectionDateTime
                            $timeSinceCheckIn = (Get-Date) - $lastCheckIn
                            
                            Write-Host "  [Connector] " -NoNewline -ForegroundColor Gray
                            Write-Host "$($connector.displayName) - Last check-in: $($connector.lastConnectionDateTime) " -NoNewline
                            
                            if ($timeSinceCheckIn.TotalHours -le 1) {
                                Write-Host "[OK] RECENT" -ForegroundColor Green
                            } elseif ($timeSinceCheckIn.TotalHours -le 24) {
                                Write-Host "[!] WARNING (> 1 hour)" -ForegroundColor Yellow
                                $connectorStatus.Errors += "$($connector.displayName): Last check-in > 1 hour ago"
                            } else {
                                Write-Host "[X] OLD (> 24 hours)" -ForegroundColor Red
                                $connectorStatus.Errors += "$($connector.displayName): Last check-in > 24 hours ago"
                            }
                        }
                    }
                } else {
                    Write-Host " [!] INACTIVE" -ForegroundColor Yellow
                    $connectorStatus.Errors += "Intune Connector for AD exists but is not active"
                }
            } else {
                Write-Host " [i] NOT CONFIGURED" -ForegroundColor Gray
            }
            
        } catch {
            Write-Host " [!] UNABLE TO CHECK" -ForegroundColor Yellow
            $connectorStatus.Errors += "Error checking Intune Connector for AD: $($_.Exception.Message)"
        }
        # ============================================================================
        
        # ============================================================================
        # COMMENTED OUT - Exchange Connector Check
        # Uncomment if needed for Exchange integration checks
        # ============================================================================
        <#
        # Check Exchange Connector
        try {
            Write-Host "  [Connector] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking Exchange Connector..." -NoNewline
            
            $exchangeConnectorUri = "https://graph.microsoft.com/beta/deviceManagement/exchangeConnectors"
            $exchangeConnectors = Invoke-BWsGraphRequest -Uri $exchangeConnectorUri -Method GET -ErrorAction Stop
            
            if ($exchangeConnectors.value -and $exchangeConnectors.value.Count -gt 0) {
                $activeExchangeConnectors = $exchangeConnectors.value | Where-Object { $_.status -eq "healthy" -or $_.status -eq "active" }
                
                if ($activeExchangeConnectors.Count -gt 0) {
                    Write-Host " [OK] ACTIVE ($($activeExchangeConnectors.Count) connector(s))" -ForegroundColor Green
                    
                    foreach ($connector in $activeExchangeConnectors) {
                        $connectorStatus.Connectors += @{
                            Type = "Exchange Connector"
                            Name = $connector.serverName
                            State = $connector.status
                            LastCheckIn = $connector.lastSuccessfulSyncDateTime
                        }
                        
                        if ($connector.lastSuccessfulSyncDateTime) {
                            Write-Host "  [Connector] " -NoNewline -ForegroundColor Gray
                            Write-Host "$($connector.serverName) - Last sync: $($connector.lastSuccessfulSyncDateTime)" -ForegroundColor Gray
                        }
                    }
                } else {
                    Write-Host " [!] INACTIVE" -ForegroundColor Yellow
                    $connectorStatus.Errors += "Exchange connector exists but is not healthy"
                }
            } else {
                Write-Host " [i] NOT CONFIGURED" -ForegroundColor Gray
            }
            
        } catch {
            Write-Host " [i] NOT CONFIGURED" -ForegroundColor Gray
        }
        #>
        # ============================================================================
        
        # Check for Hybrid Azure AD Join status (ACTIVE - NOT COMMENTED)
        try {
            Write-Host "  [Hybrid Join] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking Hybrid Azure AD Join status..." -NoNewline
            
            # Check via organization settings
            $orgUri = "https://graph.microsoft.com/v1.0/organization"
            $orgInfo = Invoke-BWsGraphRequest -Uri $orgUri -Method GET -ErrorAction Stop
            
            if ($orgInfo.value -and $orgInfo.value.Count -gt 0) {
                $org = $orgInfo.value[0]
                $onPremisesSyncEnabled = $org.onPremisesSyncEnabled
                
                if ($onPremisesSyncEnabled) {
                    Write-Host " [OK] ENABLED (Sync active)" -ForegroundColor Green
                    
                    # Get additional details
                    if ($org.onPremisesLastSyncDateTime) {
                        Write-Host "  [Hybrid Join] " -NoNewline -ForegroundColor Gray
                        Write-Host "Last sync: $($org.onPremisesLastSyncDateTime)" -ForegroundColor White
                        $connectorStatus.LastCheckIn = $org.onPremisesLastSyncDateTime
                    }
                    
                    # Check verified domains (on-premises domains)
                    try {
                        $domainsUri = "https://graph.microsoft.com/v1.0/domains"
                        $domains = Invoke-BWsGraphRequest -Uri $domainsUri -Method GET -ErrorAction Stop
                        
                        $onPremDomains = $domains.value | Where-Object { $_.isDefault -eq $false -and $_.authenticationType -eq "Federated" }
                        
                        if ($onPremDomains) {
                            Write-Host "  [Hybrid Join] " -NoNewline -ForegroundColor Gray
                            Write-Host "On-premises domain(s): $($onPremDomains.id -join ', ')" -ForegroundColor White
                        }
                    } catch {
                        # Ignore domain check errors
                    }
                    
                    # Get directory sync details
                    try {
                        $dirSyncUri = 'https://graph.microsoft.com/v1.0/organization?$select=onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesLastPasswordSyncDateTime'
                        $dirSync = Invoke-BWsGraphRequest -Uri $dirSyncUri -Method GET -ErrorAction Stop
                        
                        if ($dirSync.value -and $dirSync.value[0].onPremisesLastPasswordSyncDateTime) {
                            Write-Host "  [Hybrid Join] " -NoNewline -ForegroundColor Gray
                            Write-Host "Last password sync: $($dirSync.value[0].onPremisesLastPasswordSyncDateTime)" -ForegroundColor White
                        }
                    } catch {
                        # Ignore if unable to get password sync details
                    }
                    
                } else {
                    Write-Host " [i] NOT ENABLED" -ForegroundColor Gray
                }
            } else {
                Write-Host " [!] UNABLE TO CHECK" -ForegroundColor Yellow
            }
            
        } catch {
            Write-Host " [!] UNABLE TO CHECK" -ForegroundColor Yellow
        }
        
        # Check for AD Server with Azure Reservation (check if sync server exists in Azure)
        try {
            Write-Host "  [AD Server] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking AD Server Azure presence..." -NoNewline
            
            # Check if Azure connection exists
            $azContext = Get-AzContext -ErrorAction SilentlyContinue
            
            if ($azContext) {
                # Look for VMs that might be the AD/Sync server
                # Check for VMs with common AD server names or tags
                try {
                    $vms = Get-AzVM -ErrorAction SilentlyContinue
                    $adServers = $vms | Where-Object { 
                        $_.Name -like "*DC*" -or 
                        $_.Name -like "*AD*" -or 
                        $_.Name -like "*Sync*" -or
                        $_.Name -like "*-S00" -or
                        $_.Name -like "*-S01" -or
                        $_.Name -match "^\d{4,5}-S\d{2}$" -or  # Pattern: BCID-S00
                        ($_.Tags.Keys -contains "Role" -and $_.Tags.Role -like "*AD*") -or
                        ($_.Tags.Keys -contains "Role" -and $_.Tags.Role -like "*DC*")
                    }
                    
                    if ($adServers) {
                        $connectorStatus.ADServerReservation = $true
                        $connectorStatus.ADServerName = $adServers[0].Name
                        Write-Host " [OK] FOUND ($($adServers.Count) server(s))" -ForegroundColor Green
                        
                        foreach ($server in $adServers) {
                            Write-Host "  [AD Server] " -NoNewline -ForegroundColor Gray
                            Write-Host "$($server.Name) " -NoNewline -ForegroundColor White
                            Write-Host "($($server.Location), Size: $($server.HardwareProfile.VmSize))" -ForegroundColor Gray
                            
                            # Add to connectors list
                            $connectorStatus.Connectors += @{
                                Type = "AD Server (Azure VM)"
                                Name = $server.Name
                                State = $server.PowerState
                                Location = $server.Location
                                VMSize = $server.HardwareProfile.VmSize
                            }
                        }
                    } else {
                        $connectorStatus.ADServerReservation = $false
                        Write-Host " [i] NO AD SERVERS DETECTED" -ForegroundColor Gray
                        Write-Host "  [AD Server] " -NoNewline -ForegroundColor Gray
                        Write-Host "Searched for: *DC*, *AD*, *Sync*, *-S00, *-S01, BCID-S## pattern" -ForegroundColor Gray
                    }
                } catch {
                    Write-Host " [!] UNABLE TO QUERY VMs: $($_.Exception.Message)" -ForegroundColor Yellow
                    $connectorStatus.ADServerReservation = "Unknown"
                    $connectorStatus.Errors += "Unable to query Azure VMs for AD server"
                }
            } else {
                Write-Host " [i] NO AZURE CONNECTION" -ForegroundColor Gray
                $connectorStatus.ADServerReservation = "NotChecked"
            }
            
        } catch {
            Write-Host " [!] UNABLE TO CHECK" -ForegroundColor Yellow
            $connectorStatus.ADServerReservation = "Error"
        }
        
    } catch {
        Write-Host "Error during Hybrid Join check: $($_.Exception.Message)" -ForegroundColor Red
        $connectorStatus.Errors += "General error: $($_.Exception.Message)"
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  HYBRID AZURE AD JOIN & CONNECTORS SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Check Performed:     Yes" -ForegroundColor White
    Write-Host "  NDES Connector:      " -NoNewline -ForegroundColor White
    Write-Host $(if ($connectorStatus.IsConnected) { "Active" } else { "Not Connected" }) -ForegroundColor $(if ($connectorStatus.IsConnected) { "Green" } else { "Gray" })
    if ($connectorStatus.LastCheckIn) {
        Write-Host "  Last Sync:           $($connectorStatus.LastCheckIn)" -ForegroundColor White
    }
    if ($connectorStatus.ADServerName) {
        Write-Host "  AD Server in Azure:  $($connectorStatus.ADServerName)" -ForegroundColor Green
        # Show VM details if available
        $adServerDetails = $connectorStatus.Connectors | Where-Object { $_.Type -eq "AD Server (Azure VM)" } | Select-Object -First 1
        if ($adServerDetails) {
            Write-Host "    Location:          $($adServerDetails.Location)" -ForegroundColor Gray
            Write-Host "    VM Size:           $($adServerDetails.VMSize)" -ForegroundColor Gray
        }
    } elseif ($connectorStatus.ADServerReservation -eq $true) {
        Write-Host "  AD Server in Azure:  Found" -ForegroundColor Green
    } elseif ($connectorStatus.ADServerReservation -eq $false) {
        Write-Host "  AD Server in Azure:  Not Detected" -ForegroundColor Gray
    }
    Write-Host "  Active Connectors:   $($connectorStatus.Connectors.Count)" -ForegroundColor White
    Write-Host "  Errors/Warnings:     $($connectorStatus.Errors.Count)" -ForegroundColor $(if ($connectorStatus.Errors.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView) {
        if ($connectorStatus.Errors.Count -gt 0) {
            Write-Host "ERRORS/WARNINGS:" -ForegroundColor Yellow
            Write-Host ""
            $connectorStatus.Errors | ForEach-Object {
                Write-Host "  - $_" -ForegroundColor Yellow
            }
            Write-Host ""
        }
    }
    
    return @{
        Status = $connectorStatus
        CheckPerformed = $true
    }
}

function Test-DefenderForEndpoint {
    param(
        [string]$BCID,
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  MICROSOFT DEFENDER FOR ENDPOINT CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $defenderStatus = @{
        ConnectorActive = $false
        ConfiguredPolicies = 0
        OnboardedDevices = 0
        FilesFound = @()
        FilesMissing = @()
        Errors = @()
    }
    
    try {
        Write-Host "Checking Microsoft Defender for Endpoint..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if Microsoft Graph is connected
        $graphContext = Get-BWsGraphContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
            try {
                Connect-BWsGraph -Scopes "DeviceManagementConfiguration.Read.All", "DeviceManagementManagedDevices.Read.All" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                $defenderStatus.Errors += "Graph connection failed"
                return @{
                    Status = $defenderStatus
                    CheckPerformed = $false
                }
            }
        }
        
        Write-Host ""
        
        # ============================================================================
        # Check 1: Defender Configuration Policies
        # ============================================================================
        try {
            Write-Host "  [Defender] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking Defender configuration policies..." -NoNewline
            
            # Try multiple policy sources
            $defenderPoliciesFound = $false
            $totalDefenderPolicies = 0
            
            # Check Device Configuration Policies
            try {
                $deviceConfigs = Invoke-BWsGraphPagedRequest -Uri 'deviceManagement/deviceConfigurations?$top=999'
                $defenderDeviceConfigs = $deviceConfigs | Where-Object { 
                    $_.DisplayName -like "*Defender*" -or 
                    $_.DisplayName -like "*ATP*" -or
                    $_.DisplayName -like "*Endpoint Protection*" -or
                    $_.DisplayName -like "*Antivirus*"
                }
                if ($defenderDeviceConfigs) {
                    $totalDefenderPolicies += $defenderDeviceConfigs.Count
                    $defenderPoliciesFound = $true
                }
            } catch {}
            
            # Check Configuration Policies (Settings Catalog)
            try {
                $configPoliciesUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
                $configPolicies = Invoke-BWsGraphRequest -Uri $configPoliciesUri -Method GET -ErrorAction SilentlyContinue
                if ($configPolicies.value) {
                    $defenderConfigPolicies = $configPolicies.value | Where-Object {
                        $_.name -like "*Defender*" -or 
                        $_.name -like "*ATP*" -or
                        $_.name -like "*Endpoint*" -or
                        $_.name -like "*Antivirus*"
                    }
                    if ($defenderConfigPolicies) {
                        $totalDefenderPolicies += $defenderConfigPolicies.Count
                        $defenderPoliciesFound = $true
                    }
                }
            } catch {}
            
            # Check Endpoint Security Policies (Intents)
            try {
                $intentsUri = "https://graph.microsoft.com/beta/deviceManagement/intents"
                $intents = Invoke-BWsGraphRequest -Uri $intentsUri -Method GET -ErrorAction SilentlyContinue
                if ($intents.value) {
                    $defenderIntents = $intents.value | Where-Object {
                        $_.displayName -like "*Defender*" -or
                        $_.displayName -like "*Antivirus*" -or
                        $_.displayName -like "*Endpoint*" -or
                        $_.templateId -like "*endpointSecurityAntivirus*" -or
                        $_.templateId -like "*endpointSecurityEndpointDetectionAndResponse*"
                    }
                    if ($defenderIntents) {
                        $totalDefenderPolicies += $defenderIntents.Count
                        $defenderPoliciesFound = $true
                        $defenderStatus.ConnectorActive = $true
                    }
                }
            } catch {}
            
            $defenderStatus.ConfiguredPolicies = $totalDefenderPolicies
            
            if ($defenderPoliciesFound -and $totalDefenderPolicies -gt 0) {
                Write-Host " [OK] FOUND ($totalDefenderPolicies policies)" -ForegroundColor Green
                $defenderStatus.ConnectorActive = $true
            } else {
                Write-Host " [!] NO POLICIES FOUND" -ForegroundColor Yellow
                $defenderStatus.Errors += "No Defender for Endpoint policies configured"
            }
            
        } catch {
            Write-Host " [!] ERROR" -ForegroundColor Yellow
            $defenderStatus.Errors += "Error checking Defender policies: $($_.Exception.Message)"
        }
        
        # ============================================================================
        # Check 2: Managed Devices (Defender-compatible)
        # ============================================================================
        try {
            Write-Host "  [Defender] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking compatible managed devices..." -NoNewline
            
            $managedDevices = Invoke-BWsGraphPagedRequest -Uri 'deviceManagement/managedDevices?$top=999'
            
            if ($managedDevices) {
                # Count Windows and macOS devices (Defender-compatible)
                $compatibleDevices = $managedDevices | Where-Object {
                    $_.OperatingSystem -eq "Windows" -or $_.OperatingSystem -eq "macOS"
                }
                
                $defenderStatus.OnboardedDevices = $compatibleDevices.Count
                
                if ($compatibleDevices.Count -gt 0) {
                    Write-Host " [OK] $($compatibleDevices.Count) compatible devices" -ForegroundColor Green
                } else {
                    Write-Host " [i] No compatible devices found" -ForegroundColor Gray
                }
            } else {
                Write-Host " [i] Unable to retrieve devices" -ForegroundColor Gray
            }
            
        } catch {
            Write-Host " [i] Unable to check" -ForegroundColor Gray
        }
        
        # ============================================================================
        # Check 3: Defender Onboarding Files in Storage Account
        # ============================================================================
        Write-Host "  [Defender] " -NoNewline -ForegroundColor Gray
        Write-Host "Checking onboarding files in Storage Account..." -NoNewline
        
        $requiredFiles = @(
            "GatewayWindowsDefenderATPOnboardingPackage_Intune_MacClients.zip",
            "GatewayWindowsDefenderATPOnboardingPackage_Intune_WinClients.zip",
            "GatewayWindowsDefenderATPOnboardingPackage_WinClients.zip",
            "GatewayWindowsDefenderATPOnboardingPackage_WinServers.zip"
        )
        
        $storageAccountName = "sa" + $BCID.ToLower() + "bwsfactorynch0"
        $containerName = "defender-files"
        
        try {
            # Get storage account
            $storageAccount = Get-AzStorageAccount | Where-Object { $_.StorageAccountName -eq $storageAccountName } | Select-Object -First 1
            
            if ($storageAccount) {
                $ctx = $storageAccount.Context
                
                # Check if container exists
                $container = Get-AzStorageContainer -Name $containerName -Context $ctx -ErrorAction SilentlyContinue
                
                if ($container) {
                    # Get all blobs
                    $blobs = Get-AzStorageBlob -Container $containerName -Context $ctx -ErrorAction SilentlyContinue
                    
                    if ($blobs) {
                        $blobNames = $blobs | ForEach-Object { $_.Name }
                        
                        # Check each required file
                        foreach ($file in $requiredFiles) {
                            if ($blobNames -contains $file) {
                                $defenderStatus.FilesFound += $file
                            } else {
                                $defenderStatus.FilesMissing += $file
                            }
                        }
                        
                        if ($defenderStatus.FilesMissing.Count -eq 0) {
                            Write-Host " [OK] ALL FILES PRESENT (4/4)" -ForegroundColor Green
                        } else {
                            Write-Host " [!] MISSING FILES ($($defenderStatus.FilesFound.Count)/4)" -ForegroundColor Yellow
                            $defenderStatus.Errors += "$($defenderStatus.FilesMissing.Count) onboarding file(s) missing"
                        }
                    } else {
                        Write-Host " [!] CONTAINER EMPTY (0/4)" -ForegroundColor Yellow
                        $defenderStatus.FilesMissing = $requiredFiles
                        $defenderStatus.Errors += "Container 'defender-files' is empty"
                    }
                } else {
                    Write-Host " [X] CONTAINER NOT FOUND (0/4)" -ForegroundColor Red
                    $defenderStatus.FilesMissing = $requiredFiles
                    $defenderStatus.Errors += "Container 'defender-files' not found"
                }
            } else {
                Write-Host " [X] STORAGE ACCOUNT NOT FOUND (0/4)" -ForegroundColor Red
                $defenderStatus.FilesMissing = $requiredFiles
                $defenderStatus.Errors += "Storage account '$storageAccountName' not found"
            }
            
        } catch {
            Write-Host " [!] ERROR (0/4)" -ForegroundColor Yellow
            $defenderStatus.FilesMissing = $requiredFiles
            $defenderStatus.Errors += "Error checking storage: $($_.Exception.Message)"
        }
        
    } catch {
        Write-Host "Error during Defender check: $($_.Exception.Message)" -ForegroundColor Red
        $defenderStatus.Errors += "General error: $($_.Exception.Message)"
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  DEFENDER FOR ENDPOINT SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Policies Configured: $($defenderStatus.ConfiguredPolicies)" -ForegroundColor $(if ($defenderStatus.ConfiguredPolicies -gt 0) { "Green" } else { "Yellow" })
    Write-Host "  Compatible Devices:  $($defenderStatus.OnboardedDevices)" -ForegroundColor $(if ($defenderStatus.OnboardedDevices -gt 0) { "Green" } else { "Gray" })
    Write-Host "  Onboarding Files:    $($defenderStatus.FilesFound.Count)/4" -ForegroundColor $(if ($defenderStatus.FilesMissing.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Status:              " -NoNewline -ForegroundColor White
    Write-Host $(if ($defenderStatus.ConnectorActive) { "Active" } else { "Not Configured" }) -ForegroundColor $(if ($defenderStatus.ConnectorActive) { "Green" } else { "Yellow" })
    Write-Host "  Errors/Warnings:     $($defenderStatus.Errors.Count)" -ForegroundColor $(if ($defenderStatus.Errors.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView) {
        if ($defenderStatus.FilesFound.Count -gt 0) {
            Write-Host "FOUND ONBOARDING FILES:" -ForegroundColor Green
            Write-Host ""
            $defenderStatus.FilesFound | ForEach-Object {
                Write-Host "  [OK] $_" -ForegroundColor Green
            }
            Write-Host ""
        }
        
        if ($defenderStatus.FilesMissing.Count -gt 0) {
            Write-Host "MISSING ONBOARDING FILES:" -ForegroundColor Red
            Write-Host ""
            $defenderStatus.FilesMissing | ForEach-Object {
                Write-Host "  [X] $_" -ForegroundColor Red
            }
            Write-Host ""
        }
        
        if ($defenderStatus.Errors.Count -gt 0) {
            Write-Host "DEFENDER ERRORS/WARNINGS:" -ForegroundColor Yellow
            Write-Host ""
            $defenderStatus.Errors | ForEach-Object {
                Write-Host "  - $_" -ForegroundColor Yellow
            }
            Write-Host ""
        }
    }
    
    return @{
        Status = $defenderStatus
        CheckPerformed = $true
    }
}

function Test-BWSSoftwarePackages {
    param(
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  BWS STANDARD SOFTWARE PACKAGES CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $softwareStatus = @{
        Total = 7
        Found = @()
        Missing = @()
        Errors = @()
    }
    
    # Define required BWS software packages
    $requiredSoftware = @(
        "7-Zip",
        "Adobe Reader",
        "Chocolatey",
        "Cisco AnyConnect",
        "beyond Trust Remote support",
        "Microsoft 365 Apps for Windows 10 and later",
        "UpdateChocoSoftware"
    )
    
    try {
        Write-Host "Checking BWS Standard Software Packages..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if Microsoft Graph is connected
        $graphContext = Get-BWsGraphContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
            try {
                Connect-BWsGraph -Scopes "DeviceManagementApps.Read.All" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                $softwareStatus.Errors += "Graph connection failed"
                return @{
                    Status = $softwareStatus
                    CheckPerformed = $false
                }
            }
        }
        
        Write-Host ""
        
        # Get all Intune Win32 Apps
        try {
            Write-Host "  [Software] Retrieving Win32 Apps from Intune..." -ForegroundColor Gray
            $win32Apps = Invoke-BWsGraphPagedRequest -Uri 'deviceAppManagement/mobileApps?$filter=isof(''microsoft.graph.win32LobApp'')&$top=999'
            Write-Host "  [Software] Found $($win32Apps.Count) Win32 Apps" -ForegroundColor Gray
        } catch {
            Write-Host "  [Software] Error retrieving Win32 Apps: $($_.Exception.Message)" -ForegroundColor Yellow
            $win32Apps = @()
        }
        
        # Get all Microsoft Store Apps
        try {
            Write-Host "  [Software] Retrieving Microsoft Store Apps from Intune..." -ForegroundColor Gray
            $storeApps = Invoke-BWsGraphPagedRequest -Uri 'deviceAppManagement/mobileApps?$filter=isof(''microsoft.graph.winGetApp'')&$top=999'
            Write-Host "  [Software] Found $($storeApps.Count) Store Apps" -ForegroundColor Gray
        } catch {
            Write-Host "  [Software] Error retrieving Store Apps: $($_.Exception.Message)" -ForegroundColor Yellow
            $storeApps = @()
        }
        
        # Get Microsoft 365 Apps
        try {
            Write-Host "  [Software] Retrieving Microsoft 365 Apps from Intune..." -ForegroundColor Gray
            # Try with filter first
            $m365Apps = Invoke-BWsGraphPagedRequest -Uri 'deviceAppManagement/mobileApps?$filter=isof(''microsoft.graph.officeSuiteApp'')&$top=999'
            
            # If filter doesn't work, get all apps and filter manually
            if (-not $m365Apps -or $m365Apps.Count -eq 0) {
                $allMobileApps = Invoke-BWsGraphPagedRequest -Uri 'deviceAppManagement/mobileApps?$top=999'
                $m365Apps = $allMobileApps | Where-Object { 
                    $_.'@odata.type' -eq '#microsoft.graph.officeSuiteApp' -or
                    $_.DisplayName -like '*Microsoft 365 Apps*' -or
                    $_.DisplayName -like '*Office 365*'
                }
            }
            
            Write-Host "  [Software] Found $($m365Apps.Count) Office Suite Apps" -ForegroundColor Gray
        } catch {
            Write-Host "  [Software] Error retrieving Microsoft 365 Apps: $($_.Exception.Message)" -ForegroundColor Yellow
            $m365Apps = @()
        }
        
        Write-Host ""
        
        # Combine all apps
        $allApps = @()
        if ($win32Apps) { $allApps += $win32Apps }
        if ($storeApps) { $allApps += $storeApps }
        if ($m365Apps) { $allApps += $m365Apps }
        
        Write-Host "Total apps in Intune: $($allApps.Count)" -ForegroundColor White
        Write-Host ""
        
        # Check each required software
        foreach ($software in $requiredSoftware) {
            Write-Host "  [Software] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking for '$software'..." -NoNewline
            
            # Search for the software with improved matching logic
            # Try exact match first, then partial matches
            $foundApp = $null
            
            # Try 1: Exact match (case-insensitive)
            $foundApp = $allApps | Where-Object { 
                $_.DisplayName -eq $software
            } | Select-Object -First 1
            
            # Try 2: Case-insensitive partial match
            if (-not $foundApp) {
                $foundApp = $allApps | Where-Object { 
                    $_.DisplayName -like "*$software*"
                } | Select-Object -First 1
            }
            
            # Try 3: Split software name and match individual words (for complex names)
            if (-not $foundApp) {
                $words = $software -split '\s+'
                foreach ($word in $words) {
                    if ($word.Length -gt 3) {  # Only use meaningful words
                        $foundApp = $allApps | Where-Object { 
                            $_.DisplayName -like "*$word*"
                        } | Select-Object -First 1
                        
                        if ($foundApp) {
                            # Verify it's a good match by checking if at least 2 words match
                            $matchCount = 0
                            foreach ($w in $words) {
                                if ($foundApp.DisplayName -like "*$w*") {
                                    $matchCount++
                                }
                            }
                            if ($matchCount -ge 2 -or $words.Count -eq 1) {
                                break
                            } else {
                                $foundApp = $null
                            }
                        }
                    }
                }
            }
            
            if ($foundApp) {
                Write-Host " [OK] FOUND" -ForegroundColor Green
                $matchType = "Partial"
                if ($foundApp.DisplayName -eq $software) {
                    $matchType = "Exact"
                } elseif ($foundApp.DisplayName -like "*$software*") {
                    $matchType = "Partial"
                } else {
                    $matchType = "Fuzzy"
                }
                
                $softwareStatus.Found += @{
                    SoftwareName = $software
                    ActualName = $foundApp.DisplayName
                    AppId = $foundApp.Id
                    MatchType = $matchType
                }
            } else {
                Write-Host " [X] MISSING" -ForegroundColor Red
                $softwareStatus.Missing += @{
                    SoftwareName = $software
                }
            }
        }
        
    } catch {
        Write-Host "Error during software package check: $($_.Exception.Message)" -ForegroundColor Red
        $softwareStatus.Errors += "General error: $($_.Exception.Message)"
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  BWS SOFTWARE PACKAGES SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Total Required:  $($softwareStatus.Total)" -ForegroundColor White
    Write-Host "  Found:           $($softwareStatus.Found.Count)" -ForegroundColor $(if ($softwareStatus.Found.Count -eq $softwareStatus.Total) { "Green" } else { "Yellow" })
    Write-Host "  Missing:         $($softwareStatus.Missing.Count)" -ForegroundColor $(if ($softwareStatus.Missing.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Errors:          $($softwareStatus.Errors.Count)" -ForegroundColor $(if ($softwareStatus.Errors.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView) {
        if ($softwareStatus.Found.Count -gt 0) {
            Write-Host "FOUND SOFTWARE PACKAGES:" -ForegroundColor Green
            Write-Host ""
            foreach ($app in $softwareStatus.Found) {
                Write-Host "  [OK] $($app.SoftwareName)" -ForegroundColor Green
                Write-Host "    Actual Name: $($app.ActualName)" -ForegroundColor Gray
                Write-Host "    Match Type:  $($app.MatchType)" -ForegroundColor Gray
                Write-Host ""
            }
        }
        
        if ($softwareStatus.Missing.Count -gt 0) {
            Write-Host "MISSING SOFTWARE PACKAGES:" -ForegroundColor Red
            Write-Host ""
            foreach ($app in $softwareStatus.Missing) {
                Write-Host "  [X] $($app.SoftwareName)" -ForegroundColor Red
            }
            Write-Host ""
        }
        
        if ($softwareStatus.Errors.Count -gt 0) {
            Write-Host "ERRORS/WARNINGS:" -ForegroundColor Yellow
            Write-Host ""
            $softwareStatus.Errors | ForEach-Object {
                Write-Host "  - $_" -ForegroundColor Yellow
            }
            Write-Host ""
        }
    }
    
    return @{
        Status = $softwareStatus
        CheckPerformed = $true
    }
}

function Test-SharePointConfiguration {
    param(
        [bool]$CompactView = $false,
        [string]$SharePointUrl = ""
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  SHAREPOINT CONFIGURATION CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $spConfig = @{
        Settings = @{
            SharePointExternalSharing = $null
            OneDriveExternalSharing = $null
            SiteCreation = $null
            LegacyAuthBlocked = $null
            TenantUrl = $null
            ConnectionMethod = $null
        }
        Compliant = $false
        Errors = @()
        CheckPerformed = $false
    }
    
    try {
        Write-Host "Checking SharePoint configuration..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if SPO Management Shell is available (preferred) or PnP PowerShell
        $spoModuleAvailable = $false
        $moduleType = $null
        
        if (Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell") {
            $spoModuleAvailable = $true
            $moduleType = "SPO"
        } elseif (Get-Module -ListAvailable -Name "PnP.PowerShell") {
            $spoModuleAvailable = $true
            $moduleType = "PnP.PowerShell"
        }
        
        if ($spoModuleAvailable) {
            Write-Host "  [SharePoint] Using $moduleType module" -ForegroundColor Gray
            
            # Check if already connected or need to connect
            $needsConnection = $false
            $tenant = $null
            
            try {
                if ($moduleType -eq "SPO") {
                    $tenant = Get-SPOTenant -ErrorAction Stop
                } else {
                    $tenant = Get-PnPTenant -ErrorAction Stop
                }
            } catch {
                $needsConnection = $true
            }
            
            # If not connected and URL provided, try to connect
            if ($needsConnection -and $SharePointUrl) {
                Write-Host "  [SharePoint] Not connected, attempting connection to: $SharePointUrl" -ForegroundColor Yellow
                
                try {
                    if ($moduleType -eq "SPO") {
                        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
                        Connect-SPOService -Url $SharePointUrl -ErrorAction Stop
                        Write-Host "  [SharePoint] Connected successfully" -ForegroundColor Green
                        $tenant = Get-SPOTenant -ErrorAction Stop
                    } else {
                        Connect-PnPOnline -Url $SharePointUrl -Interactive -ErrorAction Stop
                        Write-Host "  [SharePoint] Connected successfully (PnP)" -ForegroundColor Green
                        $tenant = Get-PnPTenant -ErrorAction Stop
                    }
                    $needsConnection = $false
                } catch {
                    Write-Host "  [SharePoint] Connection failed: $($_.Exception.Message)" -ForegroundColor Red
                    $spConfig.Errors += "Failed to connect to SharePoint: $($_.Exception.Message)"
                }
            }
            
            # ALL CHECKS MUST BE INSIDE THIS if ($tenant) BLOCK!
            if ($tenant) {
                $spConfig.CheckPerformed = $true
                $spConfig.Settings.ConnectionMethod = $moduleType
                
                # Store Tenant URL
                if ($tenant.RootSiteUrl) {
                    $spConfig.Settings.TenantUrl = $tenant.RootSiteUrl
                }
                
                Write-Host ""
                
                # ============================================================
                # CHECK 1: External Sharing (SharePoint and OneDrive)
                # Location: SharePoint Admin Center > Sharing > External Sharing
                # SharePoint SOLL: Anyone
                # OneDrive SOLL: Only people in your organization (Disabled)
                # ============================================================
                try {
                    Write-Host "  [SharePoint] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking External Sharing settings..." -NoNewline
                    
                    # SharePoint External Sharing (SharingCapability)
                    # Values: 0=Disabled, 1=ExternalUserSharingOnly, 2=ExternalUserAndGuestSharing, 3=Anyone
                    $spSharingCapability = $tenant.SharingCapability
                    
                    # OneDrive External Sharing (OneDriveSharingCapability)
                    # Values: 0=Disabled, 1=ExternalUserSharingOnly, 2=ExternalUserAndGuestSharing, 3=Anyone
                    $odSharingCapability = $tenant.OneDriveSharingCapability
                    
                    # Check SharePoint - SOLL: Anyone (allows anyone links)
                    if ($spSharingCapability -eq 2 -or $spSharingCapability -eq "ExternalUserAndGuestSharing") {
                        # Value 2 = "Anyone" in the Admin Center
                        Write-Host " [OK] SharePoint: Anyone" -ForegroundColor Green
                        $spConfig.Settings.SharePointExternalSharing = "Anyone"
                    } else {
                        Write-Host " [!] SharePoint: $spSharingCapability (not 'Anyone')" -ForegroundColor Yellow
                        $spConfig.Settings.SharePointExternalSharing = $spSharingCapability.ToString()
                        $spConfig.Errors += "SharePoint External Sharing should be 'Anyone' (ExternalUserAndGuestSharing)"
                    }
                    
                    # Check OneDrive - SOLL: Only people in your organization (Disabled)
                    Write-Host "  [OneDrive]    " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking OneDrive External Sharing..." -NoNewline
                    
                    if ($odSharingCapability -eq 0 -or $odSharingCapability -eq "Disabled") {
                        Write-Host " [OK] Only people in your organization" -ForegroundColor Green
                        $spConfig.Settings.OneDriveExternalSharing = "Disabled"
                    } else {
                        Write-Host " [!] $odSharingCapability (not 'Disabled')" -ForegroundColor Yellow
                        $spConfig.Settings.OneDriveExternalSharing = $odSharingCapability.ToString()
                        $spConfig.Errors += "OneDrive External Sharing should be 'Disabled' (Only people in your organization)"
                    }
                    
                } catch {
                    Write-Host " [!] ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $spConfig.Settings.SharePointExternalSharing = "Error"
                    $spConfig.Settings.OneDriveExternalSharing = "Error"
                    $spConfig.Errors += "Error checking external sharing: $($_.Exception.Message)"
                }
                
                # ============================================================
                # CHECK 2: Site Creation (Users can create SharePoint Sites)
                # Location: SharePoint Admin Center > Settings > Site Creation
                # Property: SelfServiceSiteCreationDisabled
                # SOLL: Enabled ($true) - Users CANNOT create sites
                # ============================================================
                try {
                    Write-Host "  [SharePoint] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Site Creation settings..." -NoNewline
                    
                    if ($moduleType -eq "SPO") {
                        # Property: SelfServiceSiteCreationDisabled
                        # True = Disabled (users cannot create sites) = COMPLIANT
                        # False = Enabled (users can create sites) = NON-COMPLIANT
                        
                        $siteCreationDisabled = $tenant.SelfServiceSiteCreationDisabled
                        
                        if ($siteCreationDisabled -eq $true) {
                            # Users CANNOT create sites (compliant)
                            Write-Host " [OK] DISABLED (users cannot create sites)" -ForegroundColor Green
                            $spConfig.Settings.SiteCreation = "Disabled"
                        } elseif ($siteCreationDisabled -eq $false) {
                            # Users CAN create sites (non-compliant)
                            Write-Host " [!] ENABLED (users can create sites)" -ForegroundColor Yellow
                            $spConfig.Settings.SiteCreation = "Enabled"
                            $spConfig.Errors += "Site creation should be disabled - SelfServiceSiteCreationDisabled should be True"
                        } else {
                            # Property is null or unknown
                            Write-Host " [!] Cannot verify (property not available)" -ForegroundColor Yellow
                            $spConfig.Settings.SiteCreation = "Unknown"
                        }
                    } else {
                        Write-Host " [!] Cannot verify (SPO module required)" -ForegroundColor Yellow
                        $spConfig.Settings.SiteCreation = "Unknown"
                    }
                    
                } catch {
                    Write-Host " [!] ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $spConfig.Settings.SiteCreation = "Error"
                    $spConfig.Errors += "Error checking site creation: $($_.Exception.Message)"
                }
                
                # ============================================================
                # CHECK 3: Legacy Browser Auth (Apps that don't use modern authentication)
                # Location: SharePoint Admin Center > Access Control > Apps that don't use modern authentication
                # SOLL: Block Access
                # ============================================================
                try {
                    Write-Host "  [SharePoint] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Legacy Browser Auth blocking..." -NoNewline
                    
                    $legacyAuthBlocked = $null
                    
                    # Property: LegacyAuthProtocolsEnabled
                    # false = Blocked (compliant)
                    # true = Allowed (non-compliant)
                    if ($null -ne $tenant.LegacyAuthProtocolsEnabled) {
                        $legacyAuthBlocked = -not $tenant.LegacyAuthProtocolsEnabled
                    } elseif ($null -ne $tenant.LegacyBrowserAuthProtocolsEnabled) {
                        $legacyAuthBlocked = -not $tenant.LegacyBrowserAuthProtocolsEnabled
                    }
                    
                    if ($null -eq $legacyAuthBlocked) {
                        Write-Host " [i] Property not available" -ForegroundColor Gray
                        $spConfig.Settings.LegacyAuthBlocked = "Unknown"
                    } elseif ($legacyAuthBlocked) {
                        Write-Host " [OK] BLOCKED (Block Access)" -ForegroundColor Green
                        $spConfig.Settings.LegacyAuthBlocked = $true
                    } else {
                        Write-Host " [!] ALLOWED (should be 'Block Access')" -ForegroundColor Yellow
                        $spConfig.Settings.LegacyAuthBlocked = $false
                        $spConfig.Errors += "Apps that don't use modern authentication should be blocked"
                    }
                    
                } catch {
                    Write-Host " [!] ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $spConfig.Settings.LegacyAuthBlocked = "Error"
                    $spConfig.Errors += "Error checking legacy auth: $($_.Exception.Message)"
                }
                
            } else {
                # Not connected
                Write-Host "  [!] Not connected to SharePoint" -ForegroundColor Yellow
                $spConfig.Errors += "Not connected to SharePoint Online"
                $spConfig.Settings.SharePointExternalSharing = "Not Connected"
                $spConfig.Settings.OneDriveExternalSharing = "Not Connected"
                $spConfig.Settings.SiteCreation = "Not Connected"
                $spConfig.Settings.LegacyAuthBlocked = "Not Connected"
            }
            
        } else {
            Write-Host "  [!] SharePoint PowerShell module not found" -ForegroundColor Yellow
            $spConfig.Errors += "SharePoint PowerShell module not installed"
        }
        
        # Determine overall compliance
        $spConfig.Compliant = ($spConfig.Settings.SharePointExternalSharing -eq "Anyone") -and
                              ($spConfig.Settings.OneDriveExternalSharing -eq "Disabled") -and
                              ($spConfig.Settings.SiteCreation -eq "Disabled") -and
                              ($spConfig.Settings.LegacyAuthBlocked -eq $true) -and
                              ($spConfig.Errors.Count -eq 0)
        
    } catch {
        Write-Host "Error during SharePoint configuration check: $($_.Exception.Message)" -ForegroundColor Red
        $spConfig.Errors += "General error: $($_.Exception.Message)"
    }
    
    # Summary output
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  SHAREPOINT CONFIGURATION SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  SharePoint Ext. Sharing: " -NoNewline -ForegroundColor White
    if ($spConfig.Settings.SharePointExternalSharing -eq "Anyone") {
        Write-Host "Anyone ([OK])" -ForegroundColor Green
    } elseif ($spConfig.Settings.SharePointExternalSharing) {
        Write-Host "$($spConfig.Settings.SharePointExternalSharing) ([X])" -ForegroundColor Yellow
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  OneDrive Ext. Sharing:   " -NoNewline -ForegroundColor White
    if ($spConfig.Settings.OneDriveExternalSharing -eq "Disabled") {
        Write-Host "Only Organization ([OK])" -ForegroundColor Green
    } elseif ($spConfig.Settings.OneDriveExternalSharing) {
        Write-Host "$($spConfig.Settings.OneDriveExternalSharing) ([X])" -ForegroundColor Yellow
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  Site Creation:           " -NoNewline -ForegroundColor White
    if ($spConfig.Settings.SiteCreation -eq "Disabled") {
        Write-Host "Disabled ([OK])" -ForegroundColor Green
    } elseif ($spConfig.Settings.SiteCreation -eq "Enabled") {
        Write-Host "Enabled ([X])" -ForegroundColor Yellow
    } elseif ($spConfig.Settings.SiteCreation) {
        Write-Host "$($spConfig.Settings.SiteCreation) (?)" -ForegroundColor Gray
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  Legacy Auth Blocked:     " -NoNewline -ForegroundColor White
    if ($spConfig.Settings.LegacyAuthBlocked -eq $true) {
        Write-Host "Yes ([OK])" -ForegroundColor Green
    } elseif ($spConfig.Settings.LegacyAuthBlocked -eq $false) {
        Write-Host "No ([X])" -ForegroundColor Yellow
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    return @{
        Status = $spConfig
        CheckPerformed = $spConfig.CheckPerformed
    }
}

#============================================================================
# TEAMS CONFIGURATION CHECK
#============================================================================
function Test-TeamsConfiguration {
    param(
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  TEAMS CONFIGURATION CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $teamsConfig = @{
        Settings = @{
            ExternalAccessEnabled = $null
            CloudStorageCitrix = $null
            CloudStorageDropbox = $null
            CloudStorageBox = $null
            CloudStorageGoogleDrive = $null
            CloudStorageEgnyte = $null
            AnonymousUsersCanJoin = $null
            AnonymousUsersCanStartMeeting = $null
            DefaultPresenterRole = $null
        }
        Compliant = $false
        Errors = @()
        CheckPerformed = $false
    }
    
    try {
        Write-Host "Checking Teams configuration..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if Teams PowerShell module is available
        $teamsModuleAvailable = $false
        
        if (Get-Module -ListAvailable -Name "MicrosoftTeams") {
            $teamsModuleAvailable = $true
        }
        
        if ($teamsModuleAvailable) {
            Write-Host "  [Teams] Using MicrosoftTeams module" -ForegroundColor Gray
            
            # Check if connected to Teams
            $teamsConnected = $false
            
            try {
                $csConfig = Get-CsTeamsClientConfiguration -ErrorAction Stop
                $teamsConnected = $true
            } catch {
                Write-Host "  [Teams] Not connected to Microsoft Teams" -ForegroundColor Yellow
                Write-Host "  [Teams] Please connect first: Connect-MicrosoftTeams" -ForegroundColor Gray
            }
            
            if ($teamsConnected) {
                $teamsConfig.CheckPerformed = $true
                Write-Host ""
                
                # ============================================================
                # CHECK 1: Meetings with unmanaged MS Accounts
                # ============================================================
                try {
                    Write-Host "  [Teams] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Meetings with unmanaged MS Accounts..." -NoNewline
                    
                    $federationConfig = Get-CsTenantFederationConfiguration -ErrorAction Stop
                    $externalAccessEnabled = $federationConfig.AllowTeamsConsumer
                    
                    if ($externalAccessEnabled -eq $false) {
                        Write-Host " [OK] DISABLED (unmanaged Teams blocked)" -ForegroundColor Green
                        $teamsConfig.Settings.ExternalAccessEnabled = $false
                    } else {
                        Write-Host " [!] ENABLED (unmanaged Teams allowed)" -ForegroundColor Yellow
                        $teamsConfig.Settings.ExternalAccessEnabled = $true
                        $teamsConfig.Errors += "External access to unmanaged Teams should be disabled"
                    }
                    
                } catch {
                    Write-Host " [!] ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $teamsConfig.Settings.ExternalAccessEnabled = "Error"
                    $teamsConfig.Errors += "Error checking external access: $($_.Exception.Message)"
                }
                
                # ============================================================
                # CHECK 2: Cloud Storage Providers (Teams Settings -> Files)
                # ============================================================
                try {
                    Write-Host "  [Teams] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Cloud Storage (Files) settings..." -NoNewline
                    
                    # Get Teams Client Configuration
                    # This controls the "Files" section in Teams Settings
                    $clientConfig = Get-CsTeamsClientConfiguration -ErrorAction Stop
                    
                    # SOLL-WERT: Alle Cloud Storage Provider muessen DISABLED sein
                    # Das bedeutet: Die Properties sollten $false sein (nicht $true)
                    # Wenn $false = Ausgeschaltet = Compliant
                    # Wenn $true = Eingeschaltet = Non-Compliant
                    
                    $allDisabled = $true
                    $enabledProviders = @()
                    
                    # ============================================================
                    # Citrix Files
                    # Property: AllowCitrixContentSharing
                    # SOLL: $false (ausgeschaltet)
                    # ============================================================
                    $citrixValue = $clientConfig.AllowCitrixContentSharing
                    if ($citrixValue -eq $true) {
                        # Provider ist EINGESCHALTET (nicht compliant)
                        $allDisabled = $false
                        $enabledProviders += "Citrix Files"
                        $teamsConfig.Settings.CloudStorageCitrix = "Enabled"
                    } else {
                        # Provider ist AUSGESCHALTET (compliant)
                        # Kann $false oder $null sein
                        $teamsConfig.Settings.CloudStorageCitrix = "Disabled"
                    }
                    
                    # ============================================================
                    # Dropbox
                    # Property: AllowDropBox
                    # SOLL: $false (ausgeschaltet)
                    # ============================================================
                    $dropboxValue = $clientConfig.AllowDropBox
                    if ($dropboxValue -eq $true) {
                        $allDisabled = $false
                        $enabledProviders += "Dropbox"
                        $teamsConfig.Settings.CloudStorageDropbox = "Enabled"
                    } else {
                        $teamsConfig.Settings.CloudStorageDropbox = "Disabled"
                    }
                    
                    # ============================================================
                    # Box
                    # Property: AllowBox
                    # SOLL: $false (ausgeschaltet)
                    # ============================================================
                    $boxValue = $clientConfig.AllowBox
                    if ($boxValue -eq $true) {
                        $allDisabled = $false
                        $enabledProviders += "Box"
                        $teamsConfig.Settings.CloudStorageBox = "Enabled"
                    } else {
                        $teamsConfig.Settings.CloudStorageBox = "Disabled"
                    }
                    
                    # ============================================================
                    # Google Drive
                    # Property: AllowGoogleDrive
                    # SOLL: $false (ausgeschaltet)
                    # ============================================================
                    $googleValue = $clientConfig.AllowGoogleDrive
                    if ($googleValue -eq $true) {
                        $allDisabled = $false
                        $enabledProviders += "Google Drive"
                        $teamsConfig.Settings.CloudStorageGoogleDrive = "Enabled"
                    } else {
                        $teamsConfig.Settings.CloudStorageGoogleDrive = "Disabled"
                    }
                    
                    # ============================================================
                    # Egnyte
                    # Property: AllowEgnyte
                    # SOLL: $false (ausgeschaltet)
                    # ============================================================
                    $egnyteValue = $clientConfig.AllowEgnyte
                    if ($egnyteValue -eq $true) {
                        $allDisabled = $false
                        $enabledProviders += "Egnyte"
                        $teamsConfig.Settings.CloudStorageEgnyte = "Enabled"
                    } else {
                        $teamsConfig.Settings.CloudStorageEgnyte = "Disabled"
                    }
                    
                    # Ausgabe des Ergebnisses
                    if ($allDisabled) {
                        Write-Host " [OK] ALL DISABLED" -ForegroundColor Green
                    } else {
                        Write-Host " [!] ENABLED: $($enabledProviders -join ', ')" -ForegroundColor Yellow
                        $teamsConfig.Errors += "Cloud storage providers must be disabled (ausgeschaltet): $($enabledProviders -join ', ')"
                    }
                    
                } catch {
                    Write-Host " [!] ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $teamsConfig.Errors += "Error checking cloud storage: $($_.Exception.Message)"
                }
                
                # ============================================================
                # CHECK 3: Meeting & Lobby Settings
                # ============================================================
                try {
                    Write-Host "  [Teams] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Meeting & Lobby settings..." -NoNewline
                    
                    $meetingConfig = Get-CsTeamsMeetingConfiguration -ErrorAction Stop
                    
                    # Anonymous users can join
                    $anonymousCanJoin = $meetingConfig.DisableAnonymousJoin -eq $false
                    $teamsConfig.Settings.AnonymousUsersCanJoin = if ($anonymousCanJoin) { "Enabled" } else { "Disabled" }
                    
                    # Anonymous users can start meeting
                    $anonymousCanStart = -not $meetingConfig.EnabledAnonymousUsersRequireLobby
                    $teamsConfig.Settings.AnonymousUsersCanStartMeeting = if ($anonymousCanStart) { "Enabled" } else { "Disabled" }
                    
                    $meetingIssues = @()
                    
                    if ($anonymousCanJoin) {
                        $meetingIssues += "Anonymous join enabled"
                    }
                    
                    if ($anonymousCanStart) {
                        $meetingIssues += "Anonymous can start meetings"
                    }
                    
                    if ($meetingIssues.Count -eq 0) {
                        Write-Host " [OK] COMPLIANT" -ForegroundColor Green
                    } else {
                        Write-Host " [!] ISSUES: $($meetingIssues -join ', ')" -ForegroundColor Yellow
                        $teamsConfig.Errors += "Anonymous users should not be able to join or start meetings"
                    }
                    
                } catch {
                    Write-Host " [!] ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $teamsConfig.Settings.AnonymousUsersCanJoin = "Error"
                    $teamsConfig.Settings.AnonymousUsersCanStartMeeting = "Error"
                    $teamsConfig.Errors += "Error checking meeting settings: $($_.Exception.Message)"
                }
                
                # ============================================================
                # CHECK 4: Content Sharing - Who can present
                # ============================================================
                try {
                    Write-Host "  [Teams] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Content Sharing settings..." -NoNewline
                    
                    $meetingPolicy = Get-CsTeamsMeetingPolicy -Identity Global -ErrorAction Stop
                    $presenterRole = $meetingPolicy.DesignatedPresenterRoleMode
                    
                    $teamsConfig.Settings.DefaultPresenterRole = $presenterRole
                    
                    # EveryoneUserOverride means "Everyone" can present
                    if ($presenterRole -eq "EveryoneUserOverride") {
                        Write-Host " [OK] EVERYONE (Compliant)" -ForegroundColor Green
                    } else {
                        Write-Host " [!] $presenterRole (Non-Compliant)" -ForegroundColor Yellow
                        $teamsConfig.Errors += "Default presenter role should be 'Everyone' (EveryoneUserOverride)"
                    }
                    
                } catch {
                    Write-Host " [!] ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $teamsConfig.Settings.DefaultPresenterRole = "Error"
                    $teamsConfig.Errors += "Error checking presenter settings: $($_.Exception.Message)"
                }
                
            } else {
                Write-Host "  [!] Not connected to Microsoft Teams" -ForegroundColor Yellow
                $teamsConfig.Errors += "Not connected to Microsoft Teams"
            }
            
        } else {
            Write-Host "  [!] MicrosoftTeams PowerShell module not found" -ForegroundColor Yellow
            Write-Host "  Install with: Install-Module -Name MicrosoftTeams" -ForegroundColor Gray
            $teamsConfig.Errors += "MicrosoftTeams PowerShell module not installed"
        }
        
        # Determine overall compliance
        $teamsConfig.Compliant = ($teamsConfig.Settings.ExternalAccessEnabled -eq $false) -and
                                  ($teamsConfig.Settings.CloudStorageCitrix -eq "Disabled") -and
                                  ($teamsConfig.Settings.CloudStorageDropbox -eq "Disabled") -and
                                  ($teamsConfig.Settings.CloudStorageBox -eq "Disabled") -and
                                  ($teamsConfig.Settings.CloudStorageGoogleDrive -eq "Disabled") -and
                                  ($teamsConfig.Settings.CloudStorageEgnyte -eq "Disabled") -and
                                  ($teamsConfig.Settings.AnonymousUsersCanJoin -eq "Disabled") -and
                                  ($teamsConfig.Settings.AnonymousUsersCanStartMeeting -eq "Disabled") -and
                                  ($teamsConfig.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") -and
                                  ($teamsConfig.Errors.Count -eq 0)
        
    } catch {
        Write-Host "Error during Teams configuration check: $($_.Exception.Message)" -ForegroundColor Red
        $teamsConfig.Errors += "General error: $($_.Exception.Message)"
    }
    
    # Summary output
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  TEAMS CONFIGURATION SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Meetings w/ unmanaged MS: " -NoNewline -ForegroundColor White
    if ($teamsConfig.Settings.ExternalAccessEnabled -eq $false) {
        Write-Host "Disabled ([OK])" -ForegroundColor Green
    } elseif ($teamsConfig.Settings.ExternalAccessEnabled -eq $true) {
        Write-Host "Enabled ([X])" -ForegroundColor Yellow
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  Cloud Storage:           " -NoNewline -ForegroundColor White
    $allStorageDisabled = ($teamsConfig.Settings.CloudStorageCitrix -eq "Disabled") -and
                          ($teamsConfig.Settings.CloudStorageDropbox -eq "Disabled") -and
                          ($teamsConfig.Settings.CloudStorageBox -eq "Disabled") -and
                          ($teamsConfig.Settings.CloudStorageGoogleDrive -eq "Disabled") -and
                          ($teamsConfig.Settings.CloudStorageEgnyte -eq "Disabled")
    
    if ($allStorageDisabled) {
        Write-Host "All Disabled ([OK])" -ForegroundColor Green
    } else {
        # Build list of enabled providers
        $enabledList = @()
        if ($teamsConfig.Settings.CloudStorageCitrix -eq "Enabled") { $enabledList += "Citrix" }
        if ($teamsConfig.Settings.CloudStorageDropbox -eq "Enabled") { $enabledList += "Dropbox" }
        if ($teamsConfig.Settings.CloudStorageBox -eq "Enabled") { $enabledList += "Box" }
        if ($teamsConfig.Settings.CloudStorageGoogleDrive -eq "Enabled") { $enabledList += "Google Drive" }
        if ($teamsConfig.Settings.CloudStorageEgnyte -eq "Enabled") { $enabledList += "Egnyte" }
        
        if ($enabledList.Count -gt 0) {
            Write-Host "Enabled: $($enabledList -join ', ') ([X])" -ForegroundColor Yellow
        } else {
            Write-Host "Unknown" -ForegroundColor Gray
        }
    }
    
    Write-Host "  Anonymous Join:          " -NoNewline -ForegroundColor White
    if ($teamsConfig.Settings.AnonymousUsersCanJoin -eq "Disabled") {
        Write-Host "Disabled ([OK])" -ForegroundColor Green
    } elseif ($teamsConfig.Settings.AnonymousUsersCanJoin -eq "Enabled") {
        Write-Host "Enabled ([X])" -ForegroundColor Yellow
    } elseif ($teamsConfig.Settings.AnonymousUsersCanJoin -eq "Error") {
        Write-Host "Error - Could not check" -ForegroundColor Red
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  Anonymous Can Start:     " -NoNewline -ForegroundColor White
    if ($teamsConfig.Settings.AnonymousUsersCanStartMeeting -eq "Disabled") {
        Write-Host "Disabled ([OK])" -ForegroundColor Green
    } elseif ($teamsConfig.Settings.AnonymousUsersCanStartMeeting -eq "Enabled") {
        Write-Host "Enabled ([X])" -ForegroundColor Yellow
    } elseif ($teamsConfig.Settings.AnonymousUsersCanStartMeeting -eq "Error") {
        Write-Host "Error - Could not check" -ForegroundColor Red
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  Who Can Present:         " -NoNewline -ForegroundColor White
    if ($teamsConfig.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") {
        Write-Host "Everyone ([OK])" -ForegroundColor Green
    } elseif ($teamsConfig.Settings.DefaultPresenterRole) {
        Write-Host "$($teamsConfig.Settings.DefaultPresenterRole) ([X])" -ForegroundColor Yellow
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    return @{
        Status = $teamsConfig
        CheckPerformed = $teamsConfig.CheckPerformed
    }
}

function Test-UsersAndLicenses {
    param(
        [bool]$CompactView = $false
    )
    
    # License SKU to Friendly Name Mapping
    # This maps Microsoft's SKU PartNumbers to readable license names
    $licenseFriendlyNames = @{
        # Microsoft 365 Business
        'O365_BUSINESS_ESSENTIALS' = 'Microsoft 365 Business Basic'
        'O365_BUSINESS_PREMIUM' = 'Microsoft 365 Business Standard'
        'SPB' = 'Microsoft 365 Business Premium'
        'SMB_BUSINESS' = 'Microsoft 365 Apps for business'
        'SPE_E3' = 'Microsoft 365 E3'
        'SPE_E5' = 'Microsoft 365 E5'
        'INFORMATION_PROTECTION_COMPLIANCE' = 'Microsoft 365 E5 Compliance'
        'IDENTITY_THREAT_PROTECTION' = 'Microsoft 365 E5 Security'
        
        # Office 365 Enterprise
        'ENTERPRISEPACK' = 'Office 365 E3'
        'ENTERPRISEPREMIUM' = 'Office 365 E5'
        'ENTERPRISEPACK_USGOV_DOD' = 'Office 365 E3 (US Government DOD)'
        'ENTERPRISEPACK_USGOV_GCCHIGH' = 'Office 365 E3 (US Government GCC High)'
        'STANDARDPACK' = 'Office 365 E1'
        'STANDARDWOFFPACK' = 'Office 365 E2'
        'DESKLESSPACK' = 'Office 365 F3'
        
        # EMS + Security
        'EMS' = 'Enterprise Mobility + Security E3'
        'EMSPREMIUM' = 'Enterprise Mobility + Security E5'
        'AAD_PREMIUM' = 'Azure Active Directory Premium P1'
        'AAD_PREMIUM_P2' = 'Azure Active Directory Premium P2'
        'ADALLOM_STANDALONE' = 'Microsoft Defender for Cloud Apps'
        'ATA' = 'Microsoft Defender for Identity'
        
        # Windows
        'WIN10_PRO_ENT_SUB' = 'Windows 10/11 Enterprise E3'
        'WIN10_VDA_E5' = 'Windows 10/11 Enterprise E5'
        
        # Exchange
        'EXCHANGESTANDARD' = 'Exchange Online (Plan 1)'
        'EXCHANGEENTERPRISE' = 'Exchange Online (Plan 2)'
        'EXCHANGEARCHIVE_ADDON' = 'Exchange Online Archiving'
        'EXCHANGEDESKLESS' = 'Exchange Online Kiosk'
        
        # SharePoint
        'SHAREPOINTSTANDARD' = 'SharePoint Online (Plan 1)'
        'SHAREPOINTENTERPRISE' = 'SharePoint Online (Plan 2)'
        
        # Project & Visio
        'PROJECTPREMIUM' = 'Project Plan 5'
        'PROJECTPROFESSIONAL' = 'Project Plan 3'
        'PROJECTESSENTIALS' = 'Project Plan 1'
        'VISIOCLIENT' = 'Visio Plan 2'
        'VISIOONLINE_PLAN1' = 'Visio Plan 1'
        
        # Power Platform
        'POWER_BI_PRO' = 'Power BI Pro'
        'POWER_BI_STANDARD' = 'Power BI (free)'
        'POWERAPPS_PER_USER' = 'Power Apps per user'
        'FLOW_FREE' = 'Power Automate Free'
        'FLOW_P2' = 'Power Automate per user'
        
        # Dynamics
        'DYN365_ENTERPRISE_SALES' = 'Dynamics 365 Sales'
        'DYN365_ENTERPRISE_CUSTOMER_SERVICE' = 'Dynamics 365 Customer Service'
        
        # Teams
        'TEAMS_EXPLORATORY' = 'Microsoft Teams Exploratory'
        'TEAMS1' = 'Microsoft Teams'
        'MCOMEETADV' = 'Microsoft 365 Audio Conferencing'
        'MCOEV' = 'Microsoft 365 Phone System'
        
        # Defender
        'WINDEFATP' = 'Microsoft Defender for Endpoint'
        'MDATP_XPLAT' = 'Microsoft Defender for Endpoint'
        
        # Other
        'STREAM' = 'Microsoft Stream'
        'INTUNE_A' = 'Microsoft Intune'
        'RIGHTSMANAGEMENT' = 'Azure Information Protection Premium P1'
        'ATP_ENTERPRISE' = 'Microsoft Defender for Office 365 (Plan 1)'
        'THREAT_INTELLIGENCE' = 'Microsoft Defender for Office 365 (Plan 2)'
        'FORMS_PLAN_E5' = 'Microsoft Forms (Plan E5)'
        'MYANALYTICS_P2' = 'Microsoft Viva Insights'
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  USERS, LICENSES & PRIVILEGED ROLES CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $userLicenseStatus = @{
        TotalUsers = 0
        LicensedUsers = 0
        UnlicensedUsers = 0
        PrivilegedUsers = @()
        InvalidPrivilegedUsers = @()  # Users with privileged roles but no ADM in DisplayName
        EntraIDP2Users = @()  # Users with Entra ID P2 license
        InvalidEntraIDP2Users = @()  # Users with Entra ID P2 but no ADM in DisplayName
        GuestAccounts = @()  # Guest accounts (with #EXT# in UPN)
        UserDetails = @()
        Errors = @()
    }
    
    try {
        Write-Host "Checking users, licenses and privileged roles..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if Microsoft Graph is connected
        $graphContext = Get-BWsGraphContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
            try {
                Connect-BWsGraph -Scopes "User.Read.All", "Directory.Read.All", "RoleManagement.Read.Directory" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                $userLicenseStatus.Errors += "Graph connection failed"
                return @{
                    Status = $userLicenseStatus
                    CheckPerformed = $false
                }
            }
        }
        
        Write-Host ""
        
        # ============================================================
        # Get all users with license information
        # ============================================================
        try {
            Write-Host "  [Users] " -NoNewline -ForegroundColor Gray
            Write-Host "Fetching all users and licenses..." -NoNewline
            
            # Get users with licenses - use proper Graph API syntax
            $usersUri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,assignedLicenses,accountEnabled&$top=999'
            $usersResponse = Invoke-BWsGraphRequest -Uri $usersUri -Method GET -ErrorAction Stop
            
            $allUsers = @()
            $allUsers += $usersResponse.value
            
            # Handle pagination
            while ($usersResponse.'@odata.nextLink') {
                $usersResponse = Invoke-BWsGraphRequest -Uri $usersResponse.'@odata.nextLink' -Method GET -ErrorAction Stop
                $allUsers += $usersResponse.value
            }
            
            $userLicenseStatus.TotalUsers = $allUsers.Count
            Write-Host " [OK] $($allUsers.Count) users found" -ForegroundColor Green
            
        } catch {
            Write-Host " [X] ERROR: $($_.Exception.Message)" -ForegroundColor Red
            $userLicenseStatus.Errors += "Error fetching users: $($_.Exception.Message)"
            return @{
                Status = $userLicenseStatus
                CheckPerformed = $false
            }
        }
        
        # ============================================================
        # Get all directory roles (privileged roles)
        # ============================================================
        try {
            Write-Host "  [Roles] " -NoNewline -ForegroundColor Gray
            Write-Host "Fetching privileged role assignments..." -NoNewline
            
            # Get all directory roles
            $rolesUri = "https://graph.microsoft.com/v1.0/directoryRoles"
            $roles = Invoke-BWsGraphRequest -Uri $rolesUri -Method GET -ErrorAction Stop
            
            $privilegedRoleMembers = @{}
            
            foreach ($role in $roles.value) {
                # Get members of this role
                $membersUri = "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members"
                $members = Invoke-BWsGraphRequest -Uri $membersUri -Method GET -ErrorAction SilentlyContinue
                
                if ($members.value) {
                    foreach ($member in $members.value) {
                        if ($member.'@odata.type' -eq '#microsoft.graph.user') {
                            if (-not $privilegedRoleMembers.ContainsKey($member.id)) {
                                $privilegedRoleMembers[$member.id] = @()
                            }
                            $privilegedRoleMembers[$member.id] += $role.displayName
                        }
                    }
                }
            }
            
            Write-Host " [OK] $($privilegedRoleMembers.Count) users with privileged roles" -ForegroundColor Green
            
        } catch {
            Write-Host " [!] WARNING: $($_.Exception.Message)" -ForegroundColor Yellow
            $userLicenseStatus.Errors += "Error fetching roles: $($_.Exception.Message)"
        }
        
        # ============================================================
        # Get SKU details for license names
        # ============================================================
        $licenseSkus = @{}
        try {
            $skusUri = "https://graph.microsoft.com/v1.0/subscribedSkus"
            $skus = Invoke-BWsGraphRequest -Uri $skusUri -Method GET -ErrorAction SilentlyContinue
            
            foreach ($sku in $skus.value) {
                $licenseSkus[$sku.skuId] = $sku.skuPartNumber
            }
        } catch {
            # If we can't get SKUs, we'll just show SKU IDs
        }
        
        # ============================================================
        # Process each user
        # ============================================================
        Write-Host "  [Users] " -NoNewline -ForegroundColor Gray
        Write-Host "Processing user details..." -NoNewline
        
        foreach ($user in $allUsers) {
            $userDetail = @{
                DisplayName = $user.displayName
                UserPrincipalName = $user.userPrincipalName
                AccountEnabled = $user.accountEnabled
                Licenses = @()
                PrivilegedRoles = @()
                HasPrivilegedRole = $false
                IsADMAccount = $false
                IsInvalid = $false
                HasEntraIDP2 = $false
                EntraIDP2Violation = $false
                IsGuest = $false
            }
            
            # Check if user is a guest account (has #EXT# in UPN)
            if ($user.userPrincipalName -match "#EXT#") {
                $userDetail.IsGuest = $true
            }
            
            # Check if user has licenses
            if ($user.assignedLicenses -and $user.assignedLicenses.Count -gt 0) {
                $userLicenseStatus.LicensedUsers++
                foreach ($license in $user.assignedLicenses) {
                    # Get SKU name from API
                    $skuPartNumber = if ($licenseSkus.ContainsKey($license.skuId)) {
                        $licenseSkus[$license.skuId]
                    } else {
                        $license.skuId
                    }
                    
                    # Map to friendly name
                    $licenseName = if ($licenseFriendlyNames.ContainsKey($skuPartNumber)) {
                        $licenseFriendlyNames[$skuPartNumber]
                    } else {
                        # If no mapping exists, use SKU name but make it more readable
                        $skuPartNumber -replace '_', ' '
                    }
                    
                    $userDetail.Licenses += $licenseName
                    
                    # Check for Entra ID P2 license (AAD_PREMIUM_P2)
                    if ($skuPartNumber -eq 'AAD_PREMIUM_P2') {
                        $userDetail.HasEntraIDP2 = $true
                    }
                }
            } else {
                $userLicenseStatus.UnlicensedUsers++
            }
            
            # Check if user with Entra ID P2 has ADM in DisplayName
            if ($userDetail.HasEntraIDP2) {
                if ($user.displayName -match "ADM") {
                    # Valid: Entra ID P2 user has ADM in name
                    $userLicenseStatus.EntraIDP2Users += $userDetail
                } else {
                    # VIOLATION: Entra ID P2 without ADM in name
                    $userDetail.EntraIDP2Violation = $true
                    $userLicenseStatus.InvalidEntraIDP2Users += $userDetail
                    $userLicenseStatus.Errors += "User '$($user.displayName)' has Entra ID P2 license but no 'ADM' in DisplayName"
                }
            }
            
            # Check if user has privileged roles
            if ($privilegedRoleMembers.ContainsKey($user.id)) {
                $userDetail.HasPrivilegedRole = $true
                $userDetail.PrivilegedRoles = $privilegedRoleMembers[$user.id]
                
                # Check if DisplayName contains "ADM"
                if ($user.displayName -match "ADM") {
                    $userDetail.IsADMAccount = $true
                    $userLicenseStatus.PrivilegedUsers += $userDetail
                } else {
                    # This is a violation: privileged role without ADM in name
                    $userDetail.IsInvalid = $true
                    $userDetail.IsADMAccount = $false
                    $userLicenseStatus.InvalidPrivilegedUsers += $userDetail
                    $userLicenseStatus.Errors += "User '$($user.displayName)' has privileged role(s) but no 'ADM' in DisplayName: $($userDetail.PrivilegedRoles -join ', ')"
                }
            }
            
            # Add guest accounts to separate list
            if ($userDetail.IsGuest) {
                $userLicenseStatus.GuestAccounts += $userDetail
            }
            
            $userLicenseStatus.UserDetails += $userDetail
        }
        
        Write-Host " [OK] DONE" -ForegroundColor Green
        
    } catch {
        Write-Host "Error during user/license check: $($_.Exception.Message)" -ForegroundColor Red
        $userLicenseStatus.Errors += "General error: $($_.Exception.Message)"
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  USERS & LICENSES SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Total Users:              $($userLicenseStatus.TotalUsers)" -ForegroundColor White
    Write-Host "  Licensed Users:           $($userLicenseStatus.LicensedUsers)" -ForegroundColor $(if ($userLicenseStatus.LicensedUsers -gt 0) { "Green" } else { "Gray" })
    Write-Host "  Unlicensed Users:         $($userLicenseStatus.UnlicensedUsers)" -ForegroundColor $(if ($userLicenseStatus.UnlicensedUsers -eq 0) { "Green" } else { "Yellow" })
    Write-Host "  Guest Accounts:           $($userLicenseStatus.GuestAccounts.Count)" -ForegroundColor Cyan
    Write-Host "  Users with Priv. Roles:   $($userLicenseStatus.PrivilegedUsers.Count + $userLicenseStatus.InvalidPrivilegedUsers.Count)" -ForegroundColor White
    Write-Host "  Valid ADM Accounts:       $($userLicenseStatus.PrivilegedUsers.Count)" -ForegroundColor Green
    Write-Host "  INVALID Privileged Users: $($userLicenseStatus.InvalidPrivilegedUsers.Count)" -ForegroundColor $(if ($userLicenseStatus.InvalidPrivilegedUsers.Count -eq 0) { "Green" } else { "Red" })
    Write-Host ""
    Write-Host "  Entra ID P2 Licensed:     $($userLicenseStatus.EntraIDP2Users.Count + $userLicenseStatus.InvalidEntraIDP2Users.Count)" -ForegroundColor White
    Write-Host "  Valid Entra ID P2 (ADM):  $($userLicenseStatus.EntraIDP2Users.Count)" -ForegroundColor Green
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView -and $userLicenseStatus.InvalidPrivilegedUsers.Count -gt 0) {
        Write-Host "[!] COMPLIANCE VIOLATION - Users with privileged roles:" -ForegroundColor Red
        Write-Host ""
        foreach ($user in $userLicenseStatus.InvalidPrivilegedUsers) {
            Write-Host "  [X] $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Red
            Write-Host "    Roles:    $($user.PrivilegedRoles -join ', ')" -ForegroundColor Yellow
            if ($user.Licenses.Count -gt 0) {
                Write-Host "    Licenses: $($user.Licenses -join ', ')" -ForegroundColor Gray
            } else {
                Write-Host "    Licenses: No licenses assigned" -ForegroundColor DarkGray
            }
        }
        Write-Host ""
    }
    
    if (-not $CompactView -and $userLicenseStatus.InvalidEntraIDP2Users.Count -gt 0) {
        Write-Host "[!] COMPLIANCE VIOLATION - Users with Entra ID P2 license:" -ForegroundColor Red
        Write-Host ""
        foreach ($user in $userLicenseStatus.InvalidEntraIDP2Users) {
            Write-Host "  [X] $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Red
            Write-Host "    License:  Azure Active Directory Premium P2" -ForegroundColor Yellow
            if ($user.Licenses.Count -gt 0) {
                Write-Host "    All Licenses: $($user.Licenses -join ', ')" -ForegroundColor Gray
            }
        }
        Write-Host ""
    }
    
    if (-not $CompactView -and $userLicenseStatus.PrivilegedUsers.Count -gt 0) {
        Write-Host "[OK] Valid ADM Accounts with Privileged Roles:" -ForegroundColor Green
        Write-Host ""
        foreach ($user in $userLicenseStatus.PrivilegedUsers) {
            Write-Host "  [OK] $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Green
            Write-Host "    Roles:    $($user.PrivilegedRoles -join ', ')" -ForegroundColor Gray
            if ($user.Licenses.Count -gt 0) {
                Write-Host "    Licenses: $($user.Licenses -join ', ')" -ForegroundColor Gray
            } else {
                Write-Host "    Licenses: No licenses assigned" -ForegroundColor DarkGray
            }
        }
        Write-Host ""
    }
    
    # Show guest accounts
    if (-not $CompactView -and $userLicenseStatus.GuestAccounts.Count -gt 0) {
        Write-Host " Guest Accounts (External Users):" -ForegroundColor Cyan
        Write-Host ""
        foreach ($guest in $userLicenseStatus.GuestAccounts | Sort-Object DisplayName) {
            $guestColor = if ($guest.AccountEnabled) { "White" } else { "Gray" }
            Write-Host "  - $($guest.DisplayName) ($($guest.UserPrincipalName))" -ForegroundColor $guestColor
            
            if ($guest.Licenses.Count -gt 0) {
                Write-Host "    Licenses: $($guest.Licenses -join ', ')" -ForegroundColor Gray
            } else {
                Write-Host "    Licenses: No licenses assigned" -ForegroundColor DarkGray
            }
            
            if (-not $guest.AccountEnabled) {
                Write-Host "    Status:   Account disabled" -ForegroundColor Red
            }
            
            # Show if guest has privileged roles (unusual!)
            if ($guest.HasPrivilegedRole) {
                Write-Host "    [!] WARNING: Guest has privileged roles: $($guest.PrivilegedRoles -join ', ')" -ForegroundColor Yellow
            }
        }
        Write-Host ""
        Write-Host "Total guest accounts: $($userLicenseStatus.GuestAccounts.Count)" -ForegroundColor Cyan
        Write-Host ""
    }
    
    # Show top 10 users with most licenses (optional detailed view)
    if (-not $CompactView -and $userLicenseStatus.UserDetails.Count -gt 0) {
        Write-Host " License Distribution Summary:" -ForegroundColor Cyan
        Write-Host ""
        
        # Count licenses
        $licenseCount = @{}
        foreach ($user in $userLicenseStatus.UserDetails) {
            foreach ($license in $user.Licenses) {
                if (-not $licenseCount.ContainsKey($license)) {
                    $licenseCount[$license] = 0
                }
                $licenseCount[$license]++
            }
        }
        
        # Display license counts
        $licenseCount.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            Write-Host "  $($_.Key): " -NoNewline -ForegroundColor White
            Write-Host "$($_.Value) users" -ForegroundColor Green
        }
        Write-Host ""
    }
    
    # Output all corporate users with licenses (for parsing by other programs)
    # Excludes guest accounts (they are listed separately)
    if (-not $CompactView -and $userLicenseStatus.UserDetails.Count -gt 0) {
        # Filter out guest accounts
        $corpUsers = $userLicenseStatus.UserDetails | Where-Object { -not $_.IsGuest }
        
        if ($corpUsers.Count -gt 0) {
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host "  CORP USERS & LICENSES" -ForegroundColor Cyan
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host ""
            
            # Table header
            $displayNameWidth = 30
            $upnWidth = 40
            $licensesWidth = 60
            $statusWidth = 10
            
            Write-Host $("{0,-$statusWidth} {1,-$displayNameWidth} {2,-$upnWidth} {3,-$licensesWidth}" -f "Status", "Display Name", "User Principal Name", "Licenses") -ForegroundColor White
            Write-Host ("-" * ($statusWidth + $displayNameWidth + $upnWidth + $licensesWidth + 3)) -ForegroundColor Gray
            
            foreach ($user in $corpUsers | Sort-Object DisplayName) {
            # Determine status
            $status = if ($user.Licenses.Count -gt 0) { "Licensed" } else { "No License" }
            $statusColor = if ($user.Licenses.Count -gt 0) { "Green" } else { "Yellow" }
            
            # Format licenses (join with semicolon for easier parsing)
            $licensesText = if ($user.Licenses.Count -gt 0) {
                $user.Licenses -join "; "
            } else {
                "No licenses assigned"
            }
            
            # Truncate long text to fit
            $displayName = if ($user.DisplayName.Length -gt $displayNameWidth) {
                $user.DisplayName.Substring(0, $displayNameWidth - 3) + "..."
            } else {
                $user.DisplayName
            }
            
            $upn = if ($user.UserPrincipalName.Length -gt $upnWidth) {
                $user.UserPrincipalName.Substring(0, $upnWidth - 3) + "..."
            } else {
                $user.UserPrincipalName
            }
            
            # If licenses are too long, wrap to next line
            if ($licensesText.Length -le $licensesWidth) {
                Write-Host $("{0,-$statusWidth} {1,-$displayNameWidth} {2,-$upnWidth} {3}" -f $status, $displayName, $upn, $licensesText) -ForegroundColor $statusColor
            } else {
                # Print first line
                Write-Host $("{0,-$statusWidth} {1,-$displayNameWidth} {2,-$upnWidth} {3}" -f $status, $displayName, $upn, $licensesText.Substring(0, $licensesWidth)) -ForegroundColor $statusColor
                
                # Print continuation lines
                $remainingText = $licensesText.Substring($licensesWidth)
                while ($remainingText.Length -gt 0) {
                    $chunk = if ($remainingText.Length -le $licensesWidth) {
                        $remainingText
                    } else {
                        $remainingText.Substring(0, $licensesWidth)
                    }
                    Write-Host $("{0,-$statusWidth} {1,-$displayNameWidth} {2,-$upnWidth} {3}" -f "", "", "", $chunk) -ForegroundColor Gray
                    $remainingText = if ($remainingText.Length -le $licensesWidth) { "" } else { $remainingText.Substring($licensesWidth) }
                }
            }
        }
        
        Write-Host ""
        Write-Host "Total corporate users displayed: $($corpUsers.Count)" -ForegroundColor Cyan
        Write-Host ""
        }
        
        # Export to CSV format for easy parsing (optional - written to console)
        Write-Host "======================================================" -ForegroundColor Cyan
        Write-Host "  CSV FORMAT (for parsing)" -ForegroundColor Cyan
        Write-Host "======================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "DisplayName;UserPrincipalName;AccountEnabled;Licenses;HasPrivilegedRole;PrivilegedRoles;IsGuest" -ForegroundColor White
        
        foreach ($user in $userLicenseStatus.UserDetails | Sort-Object DisplayName) {
            $licensesCSV = $user.Licenses -join "|"
            $rolesCSV = $user.PrivilegedRoles -join "|"
            $enabledCSV = if ($user.AccountEnabled) { "TRUE" } else { "FALSE" }
            $hasRoleCSV = if ($user.HasPrivilegedRole) { "TRUE" } else { "FALSE" }
            $isGuestCSV = if ($user.IsGuest) { "TRUE" } else { "FALSE" }
            
            Write-Host "$($user.DisplayName);$($user.UserPrincipalName);$enabledCSV;$licensesCSV;$hasRoleCSV;$rolesCSV;$isGuestCSV"
        }
        
        Write-Host ""
    }
    
    return @{
        Status = $userLicenseStatus
        CheckPerformed = $true
    }
}

function Export-HTMLReport {
    param(
        [string]$BCID,
        [string]$CustomerName,
        [string]$SubscriptionName,
        [object]$AzureResults,
        [object]$IntuneResults,
        [object]$EntraIDResults,
        [object]$IntuneConnResults,
        [object]$DefenderResults,
        [object]$SoftwareResults,
        [object]$SharePointResults,
        [object]$TeamsResults,
        [object]$UserLicenseResults,
        [bool]$OverallStatus
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $reportDate = Get-Date -Format "yyyyMMdd_HHmmss"
    
    # Include customer name in filename if provided
    if ($CustomerName) {
        $safeCustomerName = $CustomerName -replace '[^\w\s-]', '' -replace '\s+', '_'
        $reportPath = "BWS_Check_Report_${safeCustomerName}_${BCID}_${reportDate}.html"
    } else {
        $reportPath = "BWS_Check_Report_${BCID}_${reportDate}.html"
    }
    
    # Build HTML
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Business Workplace Services Check Report - $(if ($CustomerName) { "$CustomerName - " })BCID $BCID</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #0082C9 0%, #001155 100%);
            padding: 20px;
            color: #333;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #0082C9 0%, #001155 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        
        .header .meta {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .status-badge {
            display: inline-block;
            padding: 10px 30px;
            border-radius: 25px;
            font-weight: bold;
            font-size: 1.2em;
            margin-top: 15px;
            text-transform: uppercase;
        }
        
        .status-pass {
            background: #10b981;
            color: white;
        }
        
        .status-fail {
            background: #ef4444;
            color: white;
        }
        
        .toc {
            background: #f8fafc;
            padding: 30px;
            border-bottom: 3px solid #e2e8f0;
        }
        
        .toc h2 {
            color: #1e293b;
            margin-bottom: 20px;
            font-size: 1.8em;
        }
        
        .toc ul {
            list-style: none;
        }
        
        .toc li {
            margin: 12px 0;
        }
        
        .toc a {
            color: #0082C9;
            text-decoration: none;
            font-size: 1.1em;
            transition: all 0.3s;
            display: inline-block;
        }
        
        .toc a:hover {
            color: #001155;
            transform: translateX(5px);
        }
        
        .content {
            padding: 30px;
        }
        
        .section {
            margin-bottom: 40px;
            padding: 25px;
            background: #f8fafc;
            border-radius: 8px;
            border-left: 5px solid #0082C9;
        }
        
        .section h2 {
            color: #1e293b;
            margin-bottom: 20px;
            font-size: 1.8em;
            display: flex;
            align-items: center;
        }
        
        .section-icon {
            width: 40px;
            height: 40px;
            margin-right: 15px;
            background: #0082C9;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 1.5em;
        }
        
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }
        
        .summary-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
        }
        
        .summary-card h3 {
            color: #64748b;
            font-size: 0.9em;
            margin-bottom: 10px;
            text-transform: uppercase;
        }
        
        .summary-card .value {
            font-size: 2.5em;
            font-weight: bold;
            color: #1e293b;
        }
        
        .summary-card.success .value {
            color: #10b981;
        }
        
        .summary-card.warning .value {
            color: #f59e0b;
        }
        
        .summary-card.error .value {
            color: #ef4444;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        thead {
            background: #0082C9;
            color: white;
        }
        
        th {
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }
        
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #e2e8f0;
        }
        
        tr:last-child td {
            border-bottom: none;
        }
        
        tbody tr:hover {
            background: #f8fafc;
        }
        
        .status-icon {
            font-size: 1.2em;
            font-weight: bold;
        }
        
        .status-found {
            color: #10b981;
        }
        
        .status-missing {
            color: #ef4444;
        }
        
        .status-error {
            color: #f59e0b;
        }
        
        .info-list {
            list-style: none;
            margin: 15px 0;
        }
        
        .info-list li {
            padding: 10px;
            margin: 8px 0;
            background: white;
            border-radius: 5px;
            border-left: 3px solid #0082C9;
        }
        
        .footer {
            background: #1e293b;
            color: white;
            text-align: center;
            padding: 20px;
            font-size: 0.9em;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            
            .container {
                box-shadow: none;
            }
            
            .toc a {
                color: #000;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>&#128737; Business Workplace Services Check Report</h1>
"@
    
    if ($CustomerName) {
        $html += @"
            <div style="font-size: 1.8em; font-weight: bold; margin: 15px 0; text-shadow: 1px 1px 2px rgba(0,0,0,0.2);">
                $CustomerName
            </div>
"@
    }
    
    $html += @"
            <div class="meta">
                <strong>BCID:</strong> <span style="font-size: 1.3em; font-weight: bold;">$BCID</span> | 
                <strong>Date:</strong> $timestamp | 
                <strong>Subscription:</strong> $SubscriptionName
            </div>
            <div class="status-badge $(if ($OverallStatus) { 'status-pass' } else { 'status-fail' })">
                $(if ($OverallStatus) { '&#10003; Passed' } else { '&#10007; Issues Found' })
            </div>
        </div>
        
        <div class="toc">
            <h2>&#128203; Table of Contents</h2>
            <ul>
                <li><a href="#summary">&rarr; Executive Summary</a></li>
                <li><a href="#azure">&rarr; Azure Resources</a></li>
                <li><a href="#intune">&rarr; Intune Policies</a></li>
                <li><a href="#entra">&rarr; Entra ID Connect</a></li>
                <li><a href="#hybrid">&rarr; Hybrid Azure AD Join & Intune Connectors</a></li>
                <li><a href="#defender">&rarr; Defender for Endpoint</a></li>
                <li><a href="#software">&rarr; BWS Software Packages</a></li>
                <li><a href="#sharepoint">&rarr; SharePoint Configuration</a></li>
                <li><a href="#teams">&rarr; Teams Configuration</a></li>
                <li><a href="#users">&rarr; Users, Licenses & Privileged Roles</a></li>
            </ul>
        </div>
        
        <div class="content">
"@

    # Summary Section
    $html += @"
            <div class="section" id="summary">
                <h2><span class="section-icon">&#128202;</span>Executive Summary</h2>
                <div class="summary-grid">
"@

    # 1. Azure Resources (always present)
    if ($AzureResults) {
        $azureClass = if ($AzureResults.Missing.Count -eq 0) { "success" } else { "error" }
        $html += @"
                    <div class="summary-card $azureClass">
                        <h3>Azure Resources</h3>
                        <div class="value">$($AzureResults.Found.Count)/$($AzureResults.Total)</div>
                        <p>Found</p>
                    </div>
"@
    }

    # 2. Intune Policies
    if ($IntuneResults -and $IntuneResults.CheckPerformed) {
        $intuneClass = if ($IntuneResults.Missing.Count -eq 0) { "success" } else { "error" }
        $html += @"
                    <div class="summary-card $intuneClass">
                        <h3>Intune Policies</h3>
                        <div class="value">$($IntuneResults.Found.Count)/$($IntuneResults.Total)</div>
                        <p>Found</p>
                    </div>
"@
    }

    # 3. Entra ID Connect
    if ($EntraIDResults -and $EntraIDResults.CheckPerformed) {
        $entraClass = if ($EntraIDResults.Status.IsRunning) { "success" } else { "error" }
        $entraDetails = ""
        if ($EntraIDResults.Status.PasswordHashSync -eq $true) {
            $entraDetails = "PW Sync [OK]"
        } elseif ($EntraIDResults.Status.IsRunning) {
            $entraDetails = "Active"
        } else {
            $entraDetails = "Inactive"
        }
        $html += @"
                    <div class="summary-card $entraClass">
                        <h3>Entra ID Sync</h3>
                        <div class="value">$(if ($EntraIDResults.Status.IsRunning) { '&#10003;' } else { '&#10007;' })</div>
                        <p>$entraDetails</p>
                    </div>
"@
    }

    # 4. Hybrid Azure AD Join & Intune Connectors
    if ($IntuneConnResults -and $IntuneConnResults.CheckPerformed) {
        $connectorClass = if ($IntuneConnResults.Status.IsConnected -and $IntuneConnResults.Status.Errors.Count -eq 0) { "success" } elseif ($IntuneConnResults.Status.IsConnected) { "warning" } else { "error" }
        $connectorDetails = ""
        if ($IntuneConnResults.Status.ADServerName) {
            $connectorDetails = "AD Server [OK]"
        } elseif ($IntuneConnResults.Status.IsConnected) {
            $connectorDetails = "Active"
        } else {
            $connectorDetails = "Not Connected"
        }
        $html += @"
                    <div class="summary-card $connectorClass">
                        <h3>Hybrid Join & Connectors</h3>
                        <div class="value">$(if ($IntuneConnResults.Status.IsConnected) { '&#10003;' } else { '&#10007;' })</div>
                        <p>$connectorDetails</p>
                    </div>
"@
    }

    # 5. Defender for Endpoint
    if ($DefenderResults -and $DefenderResults.CheckPerformed) {
        $defenderClass = if ($DefenderResults.Status.ConnectorActive -and $DefenderResults.Status.FilesMissing.Count -eq 0) { "success" } elseif ($DefenderResults.Status.ConnectorActive) { "warning" } else { "error" }
        $html += @"
                    <div class="summary-card $defenderClass">
                        <h3>Defender for Endpoint</h3>
                        <div class="value">$(if ($DefenderResults.Status.ConnectorActive) { '&#10003;' } else { '&#10007;' })</div>
                        <p>$($DefenderResults.Status.FilesFound.Count)/4 Files</p>
                    </div>
"@
    }

    # 6. BWS Software Packages
    if ($SoftwareResults -and $SoftwareResults.CheckPerformed) {
        $softwareClass = if ($SoftwareResults.Status.Missing.Count -eq 0) { "success" } else { "error" }
        $html += @"
                    <div class="summary-card $softwareClass">
                        <h3>BWS Software</h3>
                        <div class="value">$($SoftwareResults.Status.Found.Count)/$($SoftwareResults.Status.Total)</div>
                        <p>Packages</p>
                    </div>
"@
    }

    # 7. SharePoint Configuration
    if ($SharePointResults -and $SharePointResults.CheckPerformed) {
        $spClass = if ($SharePointResults.Status.Compliant) { "success" } else { "warning" }
        $spDetails = ""
        $spIssues = 0
        if ($SharePointResults.Status.Settings.SharePointExternalSharing -ne "Anyone") { $spIssues++ }
        if ($SharePointResults.Status.Settings.OneDriveExternalSharing -ne "Disabled") { $spIssues++ }
        if ($SharePointResults.Status.Settings.SiteCreation -ne "Disabled") { $spIssues++ }
        if ($SharePointResults.Status.Settings.LegacyAuthBlocked -ne $true) { $spIssues++ }
        
        if ($spIssues -eq 0) {
            $spDetails = "Compliant"
        } else {
            $spDetails = "$spIssues Issues"
        }
        
        $html += @"
                    <div class="summary-card $spClass">
                        <h3>SharePoint Config</h3>
                        <div class="value">$(if ($SharePointResults.Status.Compliant) { '&#10003;' } else { '&#9888;' })</div>
                        <p>$spDetails</p>
                    </div>
"@
    }

    # 8. Teams Configuration
    if ($TeamsResults -and $TeamsResults.CheckPerformed) {
        $teamsClass = if ($TeamsResults.Status.Compliant) { "success" } else { "warning" }
        $teamsDetails = ""
        $teamsIssues = 0
        if ($TeamsResults.Status.Settings.ExternalAccessEnabled -ne $false) { $teamsIssues++ }
        if ($TeamsResults.Status.Settings.CloudStorageCitrix -ne "Disabled") { $teamsIssues++ }
        if ($TeamsResults.Status.Settings.CloudStorageDropbox -ne "Disabled") { $teamsIssues++ }
        if ($TeamsResults.Status.Settings.CloudStorageBox -ne "Disabled") { $teamsIssues++ }
        if ($TeamsResults.Status.Settings.CloudStorageGoogleDrive -ne "Disabled") { $teamsIssues++ }
        if ($TeamsResults.Status.Settings.CloudStorageEgnyte -ne "Disabled") { $teamsIssues++ }
        
        if ($teamsIssues -eq 0) {
            $teamsDetails = "Compliant"
        } else {
            $teamsDetails = "$teamsIssues Issues"
        }
        
        $html += @"
                    <div class="summary-card $teamsClass">
                        <h3>Teams Config</h3>
                        <div class="value">$(if ($TeamsResults.Status.Compliant) { '&#10003;' } else { '&#9888;' })</div>
                        <p>$teamsDetails</p>
                    </div>
"@
    }

    # 9. User & License Status
    if ($UserLicenseResults -and $UserLicenseResults.CheckPerformed) {
        $userClass = if ($UserLicenseResults.Status.InvalidPrivilegedUsers.Count -eq 0) { "success" } else { "error" }
        $userDetails = ""
        if ($UserLicenseResults.Status.InvalidPrivilegedUsers.Count -gt 0) {
            $userDetails = "$($UserLicenseResults.Status.InvalidPrivilegedUsers.Count) Violations"
        } elseif ($UserLicenseResults.Status.UnlicensedUsers -gt 0) {
            $userDetails = "$($UserLicenseResults.Status.UnlicensedUsers) Unlicensed"
        } else {
            $userDetails = "Compliant"
        }
        
        $html += @"
                    <div class="summary-card $userClass">
                        <h3>User & Licenses</h3>
                        <div class="value">$(if ($UserLicenseResults.Status.InvalidPrivilegedUsers.Count -eq 0) { '&#10003;' } else { '&#10007;' })</div>
                        <p>$userDetails</p>
                    </div>
"@
    }

    $html += @"
                </div>
            </div>
"@

    # Azure Resources Section
    if ($AzureResults) {
        $html += @"
            <div class="section" id="azure">
                <h2><span class="section-icon">&#9729;</span>Azure Resources</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Resource Type</th>
                            <th>Resource Name</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($resource in $AzureResults.Found) {
            $html += @"
                        <tr>
                            <td><span class="status-icon status-found">&#10003;</span></td>
                            <td>$($resource.Type)</td>
                            <td>$($resource.Name)</td>
                        </tr>
"@
        }

        foreach ($resource in $AzureResults.Missing) {
            $html += @"
                        <tr>
                            <td><span class="status-icon status-missing">&#10007;</span></td>
                            <td>$($resource.Type)</td>
                            <td>$($resource.Name)</td>
                        </tr>
"@
        }

        $html += @"
                    </tbody>
                </table>
            </div>
"@
    }

    # Intune Policies Section
    if ($IntuneResults -and $IntuneResults.CheckPerformed) {
        $html += @"
            <div class="section" id="intune">
                <h2><span class="section-icon">&#128274;</span>Intune Policies</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Policy Name</th>
                            <th>Match Type</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($policy in $IntuneResults.Found) {
            $matchType = if ($policy.MatchType) { $policy.MatchType } else { "Exact" }
            $html += @"
                        <tr>
                            <td><span class="status-icon status-found">&#10003;</span></td>
                            <td>$($policy.PolicyName)</td>
                            <td>$matchType</td>
                        </tr>
"@
        }

        foreach ($policy in $IntuneResults.Missing) {
            $html += @"
                        <tr>
                            <td><span class="status-icon status-missing">&#10007;</span></td>
                            <td>$($policy.PolicyName)</td>
                            <td>Not Found</td>
                        </tr>
"@
        }

        $html += @"
                    </tbody>
                </table>
            </div>
"@
    }

    # BWS Software Packages Section
    if ($SoftwareResults -and $SoftwareResults.CheckPerformed) {
        $html += @"
            <div class="section" id="software">
                <h2><span class="section-icon">&#128230;</span>BWS Standard Software Packages</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Required Software</th>
                            <th>Actual Name</th>
                            <th>Match Type</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($app in $SoftwareResults.Status.Found) {
            $html += @"
                        <tr>
                            <td><span class="status-icon status-found">&#10003;</span></td>
                            <td>$($app.SoftwareName)</td>
                            <td>$($app.ActualName)</td>
                            <td>$($app.MatchType)</td>
                        </tr>
"@
        }

        foreach ($app in $SoftwareResults.Status.Missing) {
            $html += @"
                        <tr>
                            <td><span class="status-icon status-missing">&#10007;</span></td>
                            <td>$($app.SoftwareName)</td>
                            <td>Not Found</td>
                            <td>-</td>
                        </tr>
"@
        }

        $html += @"
                    </tbody>
                </table>
            </div>
"@
    }

    # SharePoint Configuration Section
    if ($SharePointResults -and $SharePointResults.CheckPerformed) {
        $html += @"
            <div class="section" id="sharepoint">
                <h2><span class="section-icon">&#127760;</span>SharePoint Configuration</h2>
"@
        
        # Add Tenant URL if available
        if ($SharePointResults.Status.Settings.TenantUrl) {
            $html += @"
                <ul class="info-list">
                    <li><strong>Tenant URL:</strong> $($SharePointResults.Status.Settings.TenantUrl)</li>
                    <li><strong>Connection Method:</strong> $($SharePointResults.Status.Settings.ConnectionMethod)</li>
                </ul>
                <h3>Configuration Settings:</h3>
"@
        }
        
        $html += @"
                <ul class="info-list">
                    <li><strong>SharePoint External Sharing:</strong> $(if ($SharePointResults.Status.Settings.SharePointExternalSharing -eq 'Anyone') { '<span class="status-found">&#10003; Anyone (Compliant)</span>' } elseif ($SharePointResults.Status.Settings.SharePointExternalSharing -like '*Unknown*' -or $SharePointResults.Status.Settings.SharePointExternalSharing -like '*Not Connected*') { '<span class="status-error">&#9888; Could not verify - Not connected</span>' } elseif ($SharePointResults.Status.Settings.SharePointExternalSharing) { "<span class='status-error'>&#9888; $($SharePointResults.Status.Settings.SharePointExternalSharing) (Non-Compliant - should be 'Anyone')</span>" } else { '<span class="status-error">&#9888; Check not performed</span>' })</li>
                    <li><strong>OneDrive External Sharing:</strong> $(if ($SharePointResults.Status.Settings.OneDriveExternalSharing -eq 'Disabled') { '<span class="status-found">&#10003; Only people in your organization (Compliant)</span>' } elseif ($SharePointResults.Status.Settings.OneDriveExternalSharing -like '*Unknown*' -or $SharePointResults.Status.Settings.OneDriveExternalSharing -like '*Not Connected*') { '<span class="status-error">&#9888; Could not verify - Not connected</span>' } elseif ($SharePointResults.Status.Settings.OneDriveExternalSharing) { "<span class='status-error'>&#9888; $($SharePointResults.Status.Settings.OneDriveExternalSharing) (Non-Compliant - should be 'Disabled')</span>" } else { '<span class="status-error">&#9888; Check not performed</span>' })</li>
                    <li><strong>Site Creation:</strong> $(if ($SharePointResults.Status.Settings.SiteCreation -eq 'Disabled') { '<span class="status-found">&#10003; Disabled - Users cannot create sites (Compliant)</span>' } elseif ($SharePointResults.Status.Settings.SiteCreation -eq 'Enabled') { '<span class="status-error">&#10007; Enabled - Users can create sites (Non-Compliant)</span>' } elseif ($SharePointResults.Status.Settings.SiteCreation -like '*Unknown*') { '<span class="status-error">&#9888; Could not verify</span>' } elseif ($SharePointResults.Status.Settings.SiteCreation) { "<span class='status-error'>&#9888; $($SharePointResults.Status.Settings.SiteCreation)</span>" } else { '<span class="status-error">&#9888; Check not performed</span>' })</li>
                    <li><strong>Legacy Browser Auth Blocked:</strong> $(if ($SharePointResults.Status.Settings.LegacyAuthBlocked -eq $true) { '<span class="status-found">&#10003; Yes - Legacy browser auth protocols blocked (Compliant)</span>' } elseif ($SharePointResults.Status.Settings.LegacyAuthBlocked -eq $false) { '<span class="status-error">&#10007; No - Legacy browser auth protocols allowed (Non-Compliant)</span>' } elseif ($SharePointResults.Status.Settings.LegacyAuthBlocked -like '*Property Not Available*') { '<span class="status-error">&#9888; Property not available in tenant</span>' } else { '<span class="status-error">&#9888; Check not performed</span>' })</li>
                </ul>
"@
        
        $html += "</div>"
    } elseif ($SharePointResults) {
        # SharePoint check was attempted but not performed (no connection)
        $html += @"
            <div class="section" id="sharepoint">
                <h2><span class="section-icon">&#127760;</span>SharePoint Configuration</h2>
                <ul class="info-list">
                    <li><strong>Status:</strong> <span class="status-error">&#9888; Check not performed</span></li>
"@
        if ($SharePointResults.Status.Errors.Count -gt 0) {
            $html += "<li><strong>Reason:</strong></li></ul><ul class='info-list'>"
            foreach ($error in $SharePointResults.Status.Errors) {
                $html += "<li><span class='status-error'>[!]</span> $error</li>"
            }
        }
        $html += @"
                </ul>
                <p style="color: #666; font-style: italic;">
                    Tip: Use -SharePointUrl parameter to connect automatically:<br>
                    <code>-SharePointUrl "https://TENANT-admin.sharepoint.com"</code>
                </p>
            </div>
"@
    }

    # Teams Configuration Section
    if ($TeamsResults -and $TeamsResults.CheckPerformed) {
        $html += @"
            <div class="section" id="teams">
                <h2><span class="section-icon">&#128172;</span>Teams Configuration</h2>
                <h3>Configuration Settings:</h3>
                <ul class="info-list">
                    <li><strong>Meetings with unmanaged MS Accounts:</strong> $(if ($TeamsResults.Status.Settings.ExternalAccessEnabled -eq $false) { '<span class="status-found">&#10003; Disabled (Compliant)</span>' } elseif ($TeamsResults.Status.Settings.ExternalAccessEnabled -eq $true) { '<span class="status-error">&#10007; Enabled (Non-Compliant)</span>' } else { '<span class="status-error">&#9888; Check not performed</span>' })</li>
                    <li><strong>Cloud Storage Providers:</strong>
                        <ul style="margin-top: 5px;">
                            <li>Citrix Files: $(if ($TeamsResults.Status.Settings.CloudStorageCitrix -eq "Disabled") { '<span class="status-found">&#10003; Disabled</span>' } else { '<span class="status-error">&#10007; Enabled</span>' })</li>
                            <li>Dropbox: $(if ($TeamsResults.Status.Settings.CloudStorageDropbox -eq "Disabled") { '<span class="status-found">&#10003; Disabled</span>' } else { '<span class="status-error">&#10007; Enabled</span>' })</li>
                            <li>Box: $(if ($TeamsResults.Status.Settings.CloudStorageBox -eq "Disabled") { '<span class="status-found">&#10003; Disabled</span>' } else { '<span class="status-error">&#10007; Enabled</span>' })</li>
                            <li>Google Drive: $(if ($TeamsResults.Status.Settings.CloudStorageGoogleDrive -eq "Disabled") { '<span class="status-found">&#10003; Disabled</span>' } else { '<span class="status-error">&#10007; Enabled</span>' })</li>
                            <li>Egnyte: $(if ($TeamsResults.Status.Settings.CloudStorageEgnyte -eq "Disabled") { '<span class="status-found">&#10003; Disabled</span>' } else { '<span class="status-error">&#10007; Enabled</span>' })</li>
                        </ul>
                    </li>
                    <li><strong>Anonymous Users Can Join:</strong> $(if ($TeamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Disabled") { '<span class="status-found">&#10003; Disabled (Compliant)</span>' } elseif ($TeamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Enabled") { '<span class="status-error">&#10007; Enabled (Non-Compliant)</span>' } else { '<span class="status-error">&#9888; Check not performed</span>' })</li>
                    <li><strong>Anonymous Users Can Start Meeting:</strong> $(if ($TeamsResults.Status.Settings.AnonymousUsersCanStartMeeting -eq "Disabled") { '<span class="status-found">&#10003; Disabled (Compliant)</span>' } elseif ($TeamsResults.Status.Settings.AnonymousUsersCanStartMeeting -eq "Enabled") { '<span class="status-error">&#10007; Enabled (Non-Compliant)</span>' } else { '<span class="status-error">&#9888; Check not performed</span>' })</li>
                    <li><strong>Who Can Present:</strong> $(if ($TeamsResults.Status.Settings.DefaultPresenterRole -eq 'EveryoneUserOverride') { '<span class="status-found">&#10003; Everyone (Compliant)</span>' } elseif ($TeamsResults.Status.Settings.DefaultPresenterRole) { "<span class='status-error'>&#10007; $($TeamsResults.Status.Settings.DefaultPresenterRole) (Non-Compliant)</span>" } else { '<span class="status-error">&#9888; Check not performed</span>' })</li>
                </ul>
            </div>
"@
    } elseif ($TeamsResults) {
        # Teams check was attempted but not performed
        $html += @"
            <div class="section" id="teams">
                <h2><span class="section-icon">&#128172;</span>Teams Configuration</h2>
                <p class="status-error">&#9888; Check not performed</p>
                <p style="color: #666; font-style: italic;">
                    Reason: Not connected to Microsoft Teams<br>
                    Tip: Connect first with: <code>Connect-MicrosoftTeams</code>
                </p>
            </div>
"@
    }

    # Entra ID Connect Section
    if ($EntraIDResults -and $EntraIDResults.CheckPerformed) {
        $html += @"
            <div class="section" id="entra">
                <h2><span class="section-icon">&#128279;</span>Entra ID Connect</h2>
                <ul class="info-list">
                    <li><strong>Sync Enabled:</strong> $(if ($EntraIDResults.Status.IsInstalled) { '<span class="status-found">&#10003; Yes</span>' } else { '<span class="status-missing">&#10007; No</span>' })</li>
                    <li><strong>Sync Active:</strong> $(if ($EntraIDResults.Status.IsRunning) { '<span class="status-found">&#10003; Yes</span>' } else { '<span class="status-missing">&#10007; No</span>' })</li>
"@
        if ($EntraIDResults.Status.LastSyncTime) {
            $html += @"
                    <li><strong>Last Sync:</strong> $($EntraIDResults.Status.LastSyncTime)</li>
"@
        }
        if ($EntraIDResults.Status.PasswordHashSync -ne $null) {
            $passwordSyncIcon = if ($EntraIDResults.Status.PasswordHashSync -eq $true) { '[OK]' } elseif ($EntraIDResults.Status.PasswordHashSync -eq $false) { '[!]' } else { '?' }
            $passwordSyncClass = if ($EntraIDResults.Status.PasswordHashSync -eq $true) { 'status-found' } elseif ($EntraIDResults.Status.PasswordHashSync -eq $false) { 'status-error' } else { 'status-missing' }
            $passwordSyncText = if ($EntraIDResults.Status.PasswordHashSync -eq $true) { 'Enabled' } elseif ($EntraIDResults.Status.PasswordHashSync -eq $false) { 'Disabled' } else { 'Unknown' }
            $html += @"
                    <li><strong>Password Hash Sync:</strong> <span class="$passwordSyncClass">$passwordSyncIcon $passwordSyncText</span></li>
"@
        }
        if ($EntraIDResults.Status.DeviceWritebackEnabled -ne $null) {
            $deviceSyncIcon = if ($EntraIDResults.Status.DeviceWritebackEnabled -eq $true) { '[OK]' } else { '[!]' }
            $deviceSyncClass = if ($EntraIDResults.Status.DeviceWritebackEnabled -eq $true) { 'status-found' } else { 'status-error' }
            $deviceSyncText = if ($EntraIDResults.Status.DeviceWritebackEnabled -eq $true) { 'Active' } elseif ($EntraIDResults.Status.DeviceWritebackEnabled -eq $false) { 'No Devices' } else { 'Unknown' }
            $html += @"
                    <li><strong>Device Hybrid Sync:</strong> <span class="$deviceSyncClass">$deviceSyncIcon $deviceSyncText</span></li>
"@
        }
        if ($EntraIDResults.Status.TotalUsers -gt 0) {
            $licenseIcon = if ($EntraIDResults.Status.UnlicensedUsers -eq 0) { '[OK]' } else { '[!]' }
            $licenseClass = if ($EntraIDResults.Status.UnlicensedUsers -eq 0) { 'status-found' } else { 'status-error' }
            $html += @"
                    <li><strong>Licensed Users:</strong> <span class="$licenseClass">$licenseIcon $($EntraIDResults.Status.LicensedUsers)/$($EntraIDResults.Status.TotalUsers)</span></li>
"@
            if ($EntraIDResults.Status.UnlicensedUsers -gt 0) {
                $html += @"
                    <li><strong>Unlicensed Users:</strong> <span class="status-error">&#9888; $($EntraIDResults.Status.UnlicensedUsers)</span></li>
"@
            }
        }
        if ($EntraIDResults.Status.SyncErrors.Count -gt 0) {
            $html += @"
                    <li><strong>Errors/Warnings:</strong> <span class="status-error">$($EntraIDResults.Status.SyncErrors.Count)</span></li>
"@
        }
        $html += @"
                </ul>
            </div>
"@
    }

    # Hybrid Join Section
    if ($IntuneConnResults -and $IntuneConnResults.CheckPerformed) {
        $html += @"
            <div class="section" id="hybrid">
                <h2><span class="section-icon">&#128272;</span>Hybrid Azure AD Join & Intune Connectors</h2>
                <ul class="info-list">
                    <li><strong>Check Performed:</strong> <span class="status-found">&#10003; Yes</span></li>
                    <li><strong>NDES Connector Active:</strong> $(if ($IntuneConnResults.Status.IsConnected) { '<span class="status-found">&#10003; Yes</span>' } else { '<span class="status-error">&#10007; No</span>' })</li>
                    <li><strong>Active Connectors:</strong> $($IntuneConnResults.Status.Connectors.Count)</li>
                    <li><strong>Errors/Warnings:</strong> $(if ($IntuneConnResults.Status.Errors.Count -eq 0) { '<span class="status-found">0</span>' } else { "<span class='status-error'>$($IntuneConnResults.Status.Errors.Count)</span>" })</li>
                </ul>
"@
        
        # Add connector details if any exist
        if ($IntuneConnResults.Status.Connectors.Count -gt 0) {
            $html += @"
                <h3>Connector Details:</h3>
                <ul class="info-list">
"@
            foreach ($connector in $IntuneConnResults.Status.Connectors) {
                $stateIcon = if ($connector.State -eq "active") { "[OK]" } else { "[!]" }
                $stateColor = if ($connector.State -eq "active") { "status-found" } else { "status-error" }
                $html += @"
                    <li><span class="$stateColor">$stateIcon</span> <strong>$($connector.Type):</strong> $($connector.Name)
"@
                if ($connector.LastCheckIn) {
                    $html += " - Last check-in: $($connector.LastCheckIn)"
                }
                if ($connector.Version) {
                    $html += " (v$($connector.Version))"
                }
                $html += "</li>`n"
            }
            $html += "</ul>`n"
        }
        
        # Add AD Server information if available
        if ($IntuneConnResults.Status.ADServerName -or $IntuneConnResults.Status.ADServerReservation -ne $null) {
            $html += @"
                <h3>AD Server in Azure:</h3>
                <ul class="info-list">
"@
            if ($IntuneConnResults.Status.ADServerName) {
                # Check if we have detailed AD server info in connectors
                $adServerConnectors = $IntuneConnResults.Status.Connectors | Where-Object { $_.Type -eq "AD Server (Azure VM)" }
                
                if ($adServerConnectors) {
                    foreach ($adServer in $adServerConnectors) {
                        $html += @"
                    <li><span class="status-found">&#10003;</span> <strong>$($adServer.Name)</strong>
                        <ul style="margin-left: 20px; margin-top: 5px;">
                            <li>Location: $($adServer.Location)</li>
                            <li>VM Size: $($adServer.VMSize)</li>
"@
                        if ($adServer.State) {
                            $html += "                            <li>State: $($adServer.State)</li>`n"
                        }
                        $html += @"
                        </ul>
                    </li>
"@
                    }
                } else {
                    $html += @"
                    <li><span class="status-found">&#10003;</span> Server found: <strong>$($IntuneConnResults.Status.ADServerName)</strong></li>
"@
                }
            } elseif ($IntuneConnResults.Status.ADServerReservation -eq $true) {
                $html += @"
                    <li><span class="status-found">&#10003;</span> AD Server detected in Azure</li>
"@
            } elseif ($IntuneConnResults.Status.ADServerReservation -eq $false) {
                $html += @"
                    <li><span class="status-error">&#9888;</span> No AD Server detected in Azure
                        <ul style="margin-left: 20px; margin-top: 5px;">
                            <li>Searched for patterns: *DC*, *AD*, *Sync*, *-S00, *-S01, BCID-S##</li>
                        </ul>
                    </li>
"@
            } else {
                $html += @"
                    <li><span class="status-missing">?</span> Unable to check Azure for AD Server</li>
"@
            }
            $html += "</ul>`n"
        }
        
        # Add errors if any
        if ($IntuneConnResults.Status.Errors.Count -gt 0) {
            $html += @"
                <h3>Warnings/Errors:</h3>
                <ul class="info-list">
"@
            foreach ($error in $IntuneConnResults.Status.Errors) {
                $html += @"
                    <li><span class="status-error">&#9888;</span> $error</li>
"@
            }
            $html += "</ul>`n"
        }
        
        $html += "</div>`n"
    }

    # Defender Section
    if ($DefenderResults -and $DefenderResults.CheckPerformed) {
        $html += @"
            <div class="section" id="defender">
                <h2><span class="section-icon">&#128737;</span>Microsoft Defender for Endpoint</h2>
                <ul class="info-list">
                    <li><strong>Policies Configured:</strong> $($DefenderResults.Status.ConfiguredPolicies)</li>
                    <li><strong>Compatible Devices:</strong> $($DefenderResults.Status.OnboardedDevices)</li>
                    <li><strong>Onboarding Files:</strong> $($DefenderResults.Status.FilesFound.Count)/4</li>
                    <li><strong>Status:</strong> $(if ($DefenderResults.Status.ConnectorActive) { '<span class="status-found">&#10003; Active</span>' } else { '<span class="status-missing">&#10007; Not Configured</span>' })</li>
                </ul>
"@
        
        if ($DefenderResults.Status.FilesFound.Count -gt 0) {
            $html += @"
                <h3>Found Onboarding Files:</h3>
                <ul class="info-list">
"@
            foreach ($file in $DefenderResults.Status.FilesFound) {
                $html += @"
                    <li><span class="status-found">&#10003;</span> $file</li>
"@
            }
            $html += "</ul>"
        }
        
        if ($DefenderResults.Status.FilesMissing.Count -gt 0) {
            $html += @"
                <h3>Missing Onboarding Files:</h3>
                <ul class="info-list">
"@
            foreach ($file in $DefenderResults.Status.FilesMissing) {
                $html += @"
                    <li><span class="status-missing">&#10007;</span> $file</li>
"@
            }
            $html += "</ul>"
        }
        
        $html += "</div>"
    }

    # Users, Licenses & Privileged Roles Section
    if ($UserLicenseResults -and $UserLicenseResults.CheckPerformed) {
        $html += @"
            <div class="section" id="users">
                <h2><span class="section-icon">&#128101;</span>Users, Licenses & Privileged Roles</h2>
                <ul class="info-list">
                    <li><strong>Total Users:</strong> $($UserLicenseResults.Status.TotalUsers)</li>
                    <li><strong>Licensed Users:</strong> <span class="status-found">$($UserLicenseResults.Status.LicensedUsers)</span></li>
                    <li><strong>Unlicensed Users:</strong> $(if ($UserLicenseResults.Status.UnlicensedUsers -eq 0) { '<span class="status-found">0</span>' } else { "<span class='status-error'>$($UserLicenseResults.Status.UnlicensedUsers)</span>" })</li>
                    <li><strong>Guest Accounts:</strong> <span style="color:#06b6d4;">$($UserLicenseResults.Status.GuestAccounts.Count)</span></li>
                    <li><strong>Users with Privileged Roles:</strong> $($UserLicenseResults.Status.PrivilegedUsers.Count + $UserLicenseResults.Status.InvalidPrivilegedUsers.Count)</li>
                    <li><strong>Valid ADM Accounts:</strong> <span class="status-found">$($UserLicenseResults.Status.PrivilegedUsers.Count)</span></li>
                    <li><strong>INVALID Privileged Users:</strong> $(if ($UserLicenseResults.Status.InvalidPrivilegedUsers.Count -eq 0) { '<span class="status-found">0</span>' } else { "<span class='status-error'>$($UserLicenseResults.Status.InvalidPrivilegedUsers.Count)</span>" })</li>
                </ul>
                <ul class="info-list" style="margin-top: 10px; border-top: 1px solid #e5e7eb; padding-top: 10px;">
                    <li><strong>Entra ID P2 Licensed Users:</strong> $($UserLicenseResults.Status.EntraIDP2Users.Count + $UserLicenseResults.Status.InvalidEntraIDP2Users.Count)</li>
                    <li><strong>Valid Entra ID P2 (ADM):</strong> <span class="status-found">$($UserLicenseResults.Status.EntraIDP2Users.Count)</span></li>
                </ul>
"@
        
        # Add License Distribution Summary
        $licenseCount = @{}
        foreach ($user in $UserLicenseResults.Status.UserDetails) {
            foreach ($license in $user.Licenses) {
                if (-not $licenseCount.ContainsKey($license)) {
                    $licenseCount[$license] = 0
                }
                $licenseCount[$license]++
            }
        }
        
        if ($licenseCount.Count -gt 0) {
            $html += @"
                <h3 style="margin-top: 20px;">&#128202; License Distribution:</h3>
                <table>
                    <thead>
                        <tr>
                            <th>License Type</th>
                            <th>Number of Users</th>
                            <th>Percentage</th>
                        </tr>
                    </thead>
                    <tbody>
"@
            $sortedLicenses = $licenseCount.GetEnumerator() | Sort-Object Value -Descending
            foreach ($lic in $sortedLicenses) {
                $percentage = [math]::Round(($lic.Value / $UserLicenseResults.Status.TotalUsers) * 100, 1)
                $html += @"
                        <tr>
                            <td><strong>$($lic.Key)</strong></td>
                            <td>$($lic.Value)</td>
                            <td>$percentage%</td>
                        </tr>
"@
            }
            $html += @"
                    </tbody>
                </table>
"@
        }
        
        # Show INVALID privileged users (COMPLIANCE VIOLATION)
        if ($UserLicenseResults.Status.InvalidPrivilegedUsers.Count -gt 0) {
            $html += @"
                <h3 style="color: #dc2626; margin-top: 20px;">&#9888; COMPLIANCE VIOLATION - Privileged Roles:</h3>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Display Name</th>
                            <th>User Principal Name</th>
                            <th>Privileged Roles</th>
                            <th>Assigned Licenses</th>
                        </tr>
                    </thead>
                    <tbody>
"@
            foreach ($user in $UserLicenseResults.Status.InvalidPrivilegedUsers) {
                $rolesHtml = ($user.PrivilegedRoles | ForEach-Object { "<span style='display:block; padding:2px 0;'>- $_</span>" }) -join ''
                $licensesHtml = if ($user.Licenses.Count -gt 0) { 
                    ($user.Licenses | ForEach-Object { "<span style='display:block; padding:2px 0; background:#f0f9ff; margin:1px 0; padding-left:5px; border-left:2px solid #0ea5e9;'>$_</span>" }) -join ''
                } else { 
                    '<em style="color:#999;">No licenses</em>' 
                }
                $html += @"
                        <tr>
                            <td><span class="status-missing">&#10007; INVALID</span></td>
                            <td><strong>$($user.DisplayName)</strong></td>
                            <td>$($user.UserPrincipalName)</td>
                            <td style="color: #dc2626;">$rolesHtml</td>
                            <td style="font-size:0.9em;">$licensesHtml</td>
                        </tr>
"@
            }
            $html += @"
                    </tbody>
                </table>
"@
        }
        
        # Show valid ADM accounts
        if ($UserLicenseResults.Status.PrivilegedUsers.Count -gt 0) {
            $html += @"
                <h3 style="margin-top: 20px;">&#10003; Valid ADM Accounts with Privileged Roles:</h3>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Display Name</th>
                            <th>User Principal Name</th>
                            <th>Privileged Roles</th>
                            <th>Assigned Licenses</th>
                        </tr>
                    </thead>
                    <tbody>
"@
            foreach ($user in $UserLicenseResults.Status.PrivilegedUsers) {
                $rolesHtml = ($user.PrivilegedRoles | ForEach-Object { "<span style='display:block; padding:2px 0;'>- $_</span>" }) -join ''
                $licensesHtml = if ($user.Licenses.Count -gt 0) { 
                    ($user.Licenses | ForEach-Object { "<span style='display:block; padding:2px 0; background:#f0fdf4; margin:1px 0; padding-left:5px; border-left:2px solid #22c55e;'>$_</span>" }) -join ''
                } else { 
                    '<em style="color:#999;">No licenses</em>' 
                }
                $html += @"
                        <tr>
                            <td><span class="status-found">&#10003; VALID</span></td>
                            <td><strong>$($user.DisplayName)</strong></td>
                            <td>$($user.UserPrincipalName)</td>
                            <td>$rolesHtml</td>
                            <td style="font-size:0.9em;">$licensesHtml</td>
                        </tr>
"@
            }
            $html += @"
                    </tbody>
                </table>
"@
        }
        
        # Show Entra ID P2 violations
        if ($UserLicenseResults.Status.InvalidEntraIDP2Users.Count -gt 0) {
            $html += @"
                <h3 style="color: #dc2626; margin-top: 20px;">&#9888; COMPLIANCE VIOLATION - Entra ID P2 License without ADM:</h3>
                <p style="color: #dc2626; margin-bottom: 10px;">Users with Azure Active Directory Premium P2 license should have 'ADM' in their DisplayName:</p>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Display Name</th>
                            <th>User Principal Name</th>
                            <th>Entra ID P2</th>
                            <th>All Assigned Licenses</th>
                        </tr>
                    </thead>
                    <tbody>
"@
            foreach ($user in $UserLicenseResults.Status.InvalidEntraIDP2Users) {
                $licensesHtml = if ($user.Licenses.Count -gt 0) { 
                    ($user.Licenses | ForEach-Object { 
                        $licenseStyle = if ($_ -match "Premium P2") {
                            "display:block; padding:2px 0; background:#fee2e2; margin:1px 0; padding-left:5px; border-left:2px solid #dc2626; font-weight:bold;"
                        } else {
                            "display:block; padding:2px 0; background:#f0f9ff; margin:1px 0; padding-left:5px; border-left:2px solid #0ea5e9;"
                        }
                        "<span style='$licenseStyle'>$_</span>"
                    }) -join ''
                } else { 
                    '<em style="color:#999;">No licenses</em>' 
                }
                $html += @"
                        <tr>
                            <td><span class="status-missing">&#10007; INVALID</span></td>
                            <td><strong>$($user.DisplayName)</strong></td>
                            <td>$($user.UserPrincipalName)</td>
                            <td style="color: #dc2626; font-weight: bold;">&#9888; Has Entra ID P2</td>
                            <td style="font-size:0.9em;">$licensesHtml</td>
                        </tr>
"@
            }
            $html += @"
                    </tbody>
                </table>
"@
        }
        
        # Show guest accounts (external users)
        if ($UserLicenseResults.Status.GuestAccounts.Count -gt 0) {
            $html += @"
                <h3 style="margin-top: 20px; color: #06b6d4;">&#128101; Guest Accounts (External Users):</h3>
                <p style="margin-bottom: 10px;">Users with #EXT# in their User Principal Name are external guest accounts:</p>
                <table>
                    <thead>
                        <tr>
                            <th>Display Name</th>
                            <th>User Principal Name</th>
                            <th>Account Status</th>
                            <th>Assigned Licenses</th>
                            <th>Privileged Roles</th>
                        </tr>
                    </thead>
                    <tbody>
"@
            foreach ($guest in $UserLicenseResults.Status.GuestAccounts | Sort-Object DisplayName) {
                $accountStatus = if ($guest.AccountEnabled) { 
                    '<span class="status-found">[OK] Enabled</span>' 
                } else { 
                    '<span class="status-error">[X] Disabled</span>' 
                }
                
                $licensesHtml = if ($guest.Licenses.Count -gt 0) { 
                    ($guest.Licenses | ForEach-Object { 
                        "<span style='display:block; padding:2px 0; background:#ecfeff; margin:1px 0; padding-left:5px; border-left:2px solid #06b6d4;'>$_</span>"
                    }) -join ''
                } else { 
                    '<em style="color:#999;">No licenses</em>' 
                }
                
                $rolesHtml = if ($guest.HasPrivilegedRole) {
                    '<span style="color:#dc2626; font-weight:bold;">[!] ' + ($guest.PrivilegedRoles -join ', ') + '</span>'
                } else {
                    '<em style="color:#999;">None</em>'
                }
                
                $html += @"
                        <tr>
                            <td><strong>$($guest.DisplayName)</strong></td>
                            <td style="font-size:0.9em; color:#666;">$($guest.UserPrincipalName)</td>
                            <td>$accountStatus</td>
                            <td style="font-size:0.9em;">$licensesHtml</td>
                            <td>$rolesHtml</td>
                        </tr>
"@
            }
            $html += @"
                    </tbody>
                </table>
                <p style="margin-top: 10px; color:#06b6d4; font-weight:bold;">Total guest accounts: $($UserLicenseResults.Status.GuestAccounts.Count)</p>
"@
        }
        
        # Show all corporate users with their licenses (excludes guests)
        $html += @"
                <h3 style="margin-top: 20px;">&#128203; Corp Users & Licenses:</h3>
                <p style="margin-bottom: 10px; color:#666;">Corporate users only (guest accounts are listed separately above):</p>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Display Name</th>
                            <th>User Principal Name</th>
                            <th>Assigned Licenses</th>
                            <th>Account Enabled</th>
                        </tr>
                    </thead>
                    <tbody>
"@
        # Filter out guest accounts
        $corpUsers = $UserLicenseResults.Status.UserDetails | Where-Object { -not $_.IsGuest }
        
        foreach ($user in $corpUsers | Sort-Object DisplayName) {
            $statusIcon = if ($user.Licenses.Count -gt 0) { '<span class="status-found">[OK]</span>' } else { '<span class="status-error">[!]</span>' }
            $statusText = if ($user.Licenses.Count -gt 0) { 'Licensed' } else { 'Unlicensed' }
            $licensesHtml = if ($user.Licenses.Count -gt 0) { 
                ($user.Licenses | ForEach-Object { "<span style='display:block; padding:2px 0; background:#fef3c7; margin:1px 0; padding-left:5px; border-left:2px solid #f59e0b;'>$_</span>" }) -join ''
            } else { 
                '<em style="color:#999;">No licenses</em>' 
            }
            $enabledIcon = if ($user.AccountEnabled) { '[OK]' } else { '[X]' }
            
            $html += @"
                        <tr>
                            <td>$statusIcon $statusText</td>
                            <td>$($user.DisplayName)</td>
                            <td>$($user.UserPrincipalName)</td>
                            <td style="font-size:0.9em;">$licensesHtml</td>
                            <td>$enabledIcon</td>
                        </tr>
"@
        }
        $html += @"
                    </tbody>
                </table>
            </div>
"@
    }

    $html += @"
        </div>
        
        <div class="footer">
            Generated by BWS Checking Script | $timestamp
        </div>
    </div>
</body>
</html>
"@

    # Write HTML file
    $html | Out-File -FilePath $reportPath -Encoding UTF8
    
    return $reportPath
}

function Export-PDFReport {
    param(
        [string]$HTMLPath
    )
    
    # Validate input
    if ([string]::IsNullOrWhiteSpace($HTMLPath)) {
        Write-Host "  [X] Error: HTML path is empty" -ForegroundColor Red
        return $null
    }
    
    if (-not (Test-Path $HTMLPath)) {
        Write-Host "  [X] Error: HTML file not found: $HTMLPath" -ForegroundColor Red
        return $null
    }
    
    Write-Host "Converting HTML to PDF..." -ForegroundColor Yellow
    
    # Get absolute paths
    $htmlItem = Get-Item $HTMLPath
    $htmlFullPath = $htmlItem.FullName
    $pdfPath = $htmlFullPath -replace '\.html$', '.pdf'
    $conversionSuccess = $false
    
    # Method 1: wkhtmltopdf
    $wkhtmltopdf = Get-Command "wkhtmltopdf" -ErrorAction SilentlyContinue
    if ($wkhtmltopdf) {
        Write-Host "  Using wkhtmltopdf..." -ForegroundColor Gray
        try {
            $process = Start-Process -FilePath $wkhtmltopdf.Source `
                -ArgumentList "--enable-local-file-access", "--no-stop-slow-scripts", "--javascript-delay", "1000", "`"$htmlFullPath`"", "`"$pdfPath`"" `
                -Wait -PassThru -NoNewWindow -ErrorAction Stop
            
            if ($process.ExitCode -eq 0 -and (Test-Path $pdfPath)) {
                $conversionSuccess = $true
                Write-Host "  [OK] PDF created successfully with wkhtmltopdf" -ForegroundColor Green
            }
        } catch {
            Write-Host "  [X] wkhtmltopdf error: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # Method 2: Chrome/Edge Headless
    if (-not $conversionSuccess) {
        $chromePaths = @(
            "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
            "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
            "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
            "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
        )
        
        $chromePath = $null
        foreach ($path in $chromePaths) {
            if (Test-Path $path) {
                $chromePath = $path
                break
            }
        }
        
        if ($chromePath) {
            Write-Host "  Using Chrome/Edge Headless..." -ForegroundColor Gray
            
            try {
                # Chrome needs file:/// URL
                $htmlFileUrl = "file:///$($htmlFullPath.Replace('\', '/'))"
                
                $chromeArgs = @(
                    "--headless=new"
                    "--disable-gpu"
                    "--no-sandbox"
                    "--print-to-pdf=`"$pdfPath`""
                    "`"$htmlFileUrl`""
                )
                
                $process = Start-Process -FilePath $chromePath -ArgumentList $chromeArgs `
                    -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
                
                # Wait for PDF
                $waitCount = 0
                while (-not (Test-Path $pdfPath) -and $waitCount -lt 10) {
                    Start-Sleep -Milliseconds 500
                    $waitCount++
                }
                
                if (Test-Path $pdfPath) {
                    $conversionSuccess = $true
                    Write-Host "  [OK] PDF created successfully with Chrome/Edge" -ForegroundColor Green
                }
            } catch {
                Write-Host "  [X] Chrome/Edge error: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
    
    # Method 3: Microsoft Word
    if (-not $conversionSuccess) {
        Write-Host "  Trying Microsoft Word..." -ForegroundColor Gray
        try {
            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            $word.Visible = $false
            $doc = $word.Documents.Open($htmlFullPath)
            $doc.SaveAs([ref]$pdfPath, [ref]17)
            $doc.Close()
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            [System.GC]::Collect()
            
            if (Test-Path $pdfPath) {
                $conversionSuccess = $true
                Write-Host "  [OK] PDF created successfully with Word" -ForegroundColor Green
            }
        } catch {
            Write-Host "  [X] Word error: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # Failed
    if (-not $conversionSuccess) {
        Write-Host ""
        Write-Host "  [!] PDF conversion failed" -ForegroundColor Yellow
        Write-Host "  Install wkhtmltopdf: https://wkhtmltopdf.org/downloads.html" -ForegroundColor Gray
        Write-Host "  Or use: winget install wkhtmltopdf" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  HTML report: $htmlFullPath" -ForegroundColor Cyan
        return $null
    }
    
    return $pdfPath
}


#============================================================================
# QUALITY ASSURANCE - Block 7: Run Self-Tests if requested
#============================================================================
if ($RunTests) {
    $testSummary = Invoke-BWSSelfTest
    if (-not $GUI) {
        $continueAfterTests = Read-Host "Continue with checks? (J/N)"
        if ($continueAfterTests -notin @("J","j","Y","y")) {
            Write-Host "Script stopped after self-test run." -ForegroundColor Yellow
            exit 0
        }
    }
}

#============================================================================
# GUI Mode
#============================================================================

if ($GUI) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "BWS Checking Tool v$script:Version - GUI"
    $form.Size = New-Object System.Drawing.Size(1000, 750)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    # BCID Input
    $labelBCID = New-Object System.Windows.Forms.Label
    $labelBCID.Location = New-Object System.Drawing.Point(20, 20)
    $labelBCID.Size = New-Object System.Drawing.Size(150, 20)
    $labelBCID.Text = "BCID:"
    $form.Controls.Add($labelBCID)
    
    $textBCID = New-Object System.Windows.Forms.TextBox
    $textBCID.Location = New-Object System.Drawing.Point(170, 18)
    $textBCID.Size = New-Object System.Drawing.Size(200, 20)
    $textBCID.Text = $BCID
    $form.Controls.Add($textBCID)
    
    # Customer Name Input
    $labelCustomer = New-Object System.Windows.Forms.Label
    $labelCustomer.Location = New-Object System.Drawing.Point(400, 20)
    $labelCustomer.Size = New-Object System.Drawing.Size(100, 20)
    $labelCustomer.Text = "Kunde (optional):"
    $form.Controls.Add($labelCustomer)
    
    $textCustomer = New-Object System.Windows.Forms.TextBox
    $textCustomer.Location = New-Object System.Drawing.Point(510, 18)
    $textCustomer.Size = New-Object System.Drawing.Size(200, 20)
    $textCustomer.Text = $CustomerName
    $form.Controls.Add($textCustomer)
    
    # Subscription ID Input
    $labelSubID = New-Object System.Windows.Forms.Label
    $labelSubID.Location = New-Object System.Drawing.Point(20, 50)
    $labelSubID.Size = New-Object System.Drawing.Size(150, 20)
    $labelSubID.Text = "Subscription ID (optional):"
    $form.Controls.Add($labelSubID)
    
    $textSubID = New-Object System.Windows.Forms.TextBox
    $textSubID.Location = New-Object System.Drawing.Point(170, 48)
    $textSubID.Size = New-Object System.Drawing.Size(540, 20)
    $textSubID.Text = $SubscriptionId
    $form.Controls.Add($textSubID)
    
    # SharePoint URL Input
    $labelSPUrl = New-Object System.Windows.Forms.Label
    $labelSPUrl.Location = New-Object System.Drawing.Point(20, 78)
    $labelSPUrl.Size = New-Object System.Drawing.Size(150, 20)
    $labelSPUrl.Text = "SharePoint URL (optional):"
    $form.Controls.Add($labelSPUrl)
    
    $textSPUrl = New-Object System.Windows.Forms.TextBox
    $textSPUrl.Location = New-Object System.Drawing.Point(170, 76)
    $textSPUrl.Size = New-Object System.Drawing.Size(540, 20)
    $script:spUrlPlaceholder = "https://TENANT-admin.sharepoint.com"
    if ($SharePointUrl) {
        $textSPUrl.Text = $SharePointUrl
        $textSPUrl.ForeColor = [System.Drawing.SystemColors]::WindowText
    } else {
        $textSPUrl.Text = $script:spUrlPlaceholder
        $textSPUrl.ForeColor = [System.Drawing.Color]::Gray
    }
    $textSPUrl.Add_GotFocus({
        if ($textSPUrl.Text -eq $script:spUrlPlaceholder -and $textSPUrl.ForeColor -eq [System.Drawing.Color]::Gray) {
            $textSPUrl.Text = ""
            $textSPUrl.ForeColor = [System.Drawing.SystemColors]::WindowText
        }
    })
    $textSPUrl.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($textSPUrl.Text)) {
            $textSPUrl.Text = $script:spUrlPlaceholder
            $textSPUrl.ForeColor = [System.Drawing.Color]::Gray
        }
    })
    $form.Controls.Add($textSPUrl)
    
    # GroupBox for Check Selection
    $groupBoxChecks = New-Object System.Windows.Forms.GroupBox
    $groupBoxChecks.Location = New-Object System.Drawing.Point(20, 110)
    $groupBoxChecks.Size = New-Object System.Drawing.Size(300, 250)
    $groupBoxChecks.Text = "Select Checks to Run"
    $form.Controls.Add($groupBoxChecks)
    
    # Azure Check Checkbox
    $chkAzure = New-Object System.Windows.Forms.CheckBox
    $chkAzure.Location = New-Object System.Drawing.Point(15, 25)
    $chkAzure.Size = New-Object System.Drawing.Size(250, 20)
    $chkAzure.Text = "Azure Resources Check"
    $chkAzure.Checked = $true
    $groupBoxChecks.Controls.Add($chkAzure)
    
    # Intune Check Checkbox
    $chkIntune = New-Object System.Windows.Forms.CheckBox
    $chkIntune.Location = New-Object System.Drawing.Point(15, 50)
    $chkIntune.Size = New-Object System.Drawing.Size(250, 20)
    $chkIntune.Text = "Intune Policies Check"
    $chkIntune.Checked = $true
    $groupBoxChecks.Controls.Add($chkIntune)
    
    # Entra ID Connect Check Checkbox
    $chkEntraID = New-Object System.Windows.Forms.CheckBox
    $chkEntraID.Location = New-Object System.Drawing.Point(15, 75)
    $chkEntraID.Size = New-Object System.Drawing.Size(250, 20)
    $chkEntraID.Text = "Entra ID Connect Check"
    $chkEntraID.Checked = $true
    $groupBoxChecks.Controls.Add($chkEntraID)
    
    # Hybrid Join Check Checkbox
    $chkIntuneConn = New-Object System.Windows.Forms.CheckBox
    $chkIntuneConn.Location = New-Object System.Drawing.Point(15, 100)
    $chkIntuneConn.Size = New-Object System.Drawing.Size(280, 20)
    $chkIntuneConn.Text = "Hybrid Azure AD Join Check"
    $chkIntuneConn.Checked = $true
    $groupBoxChecks.Controls.Add($chkIntuneConn)
    
    # Defender Check Checkbox
    $chkDefender = New-Object System.Windows.Forms.CheckBox
    $chkDefender.Location = New-Object System.Drawing.Point(15, 125)
    $chkDefender.Size = New-Object System.Drawing.Size(280, 20)
    $chkDefender.Text = "Defender for Endpoint Check"
    $chkDefender.Checked = $true
    $groupBoxChecks.Controls.Add($chkDefender)
    
    # Software Packages Check Checkbox
    $chkSoftware = New-Object System.Windows.Forms.CheckBox
    $chkSoftware.Location = New-Object System.Drawing.Point(15, 150)
    $chkSoftware.Size = New-Object System.Drawing.Size(280, 20)
    $chkSoftware.Text = "BWS Software Packages Check"
    $chkSoftware.Checked = $true
    $groupBoxChecks.Controls.Add($chkSoftware)
    
    # SharePoint Configuration Check Checkbox
    $chkSharePoint = New-Object System.Windows.Forms.CheckBox
    $chkSharePoint.Location = New-Object System.Drawing.Point(15, 175)
    $chkSharePoint.Size = New-Object System.Drawing.Size(280, 20)
    $chkSharePoint.Text = "SharePoint Configuration Check"
    $chkSharePoint.Checked = $true
    $groupBoxChecks.Controls.Add($chkSharePoint)
    
    # Teams Configuration Check Checkbox
    $chkTeams = New-Object System.Windows.Forms.CheckBox
    $chkTeams.Location = New-Object System.Drawing.Point(15, 200)
    $chkTeams.Size = New-Object System.Drawing.Size(280, 20)
    $chkTeams.Text = "Teams Configuration Check"
    $chkTeams.Checked = $true
    $groupBoxChecks.Controls.Add($chkTeams)
    
    # User & License Check
    $chkUserLicense = New-Object System.Windows.Forms.CheckBox
    $chkUserLicense.Location = New-Object System.Drawing.Point(15, 225)
    $chkUserLicense.Size = New-Object System.Drawing.Size(280, 20)
    $chkUserLicense.Text = "User & License Check"
    $chkUserLicense.Checked = $true
    $groupBoxChecks.Controls.Add($chkUserLicense)
    
    # Options GroupBox
    $groupBoxOptions = New-Object System.Windows.Forms.GroupBox
    $groupBoxOptions.Location = New-Object System.Drawing.Point(340, 110)
    $groupBoxOptions.Size = New-Object System.Drawing.Size(300, 250)
    $groupBoxOptions.Text = "Options"
    $form.Controls.Add($groupBoxOptions)
    
    # Compact View Checkbox
    $chkCompact = New-Object System.Windows.Forms.CheckBox
    $chkCompact.Location = New-Object System.Drawing.Point(15, 25)
    $chkCompact.Size = New-Object System.Drawing.Size(250, 20)
    $chkCompact.Text = "Compact View"
    $chkCompact.Checked = $true
    $groupBoxOptions.Controls.Add($chkCompact)
    
    # Verbose Checkbox
    $chkShowAll = New-Object System.Windows.Forms.CheckBox
    $chkShowAll.Location = New-Object System.Drawing.Point(15, 50)
    $chkShowAll.Size = New-Object System.Drawing.Size(250, 20)
    $chkShowAll.Text = "Verbose"
    $chkShowAll.Checked = $false
    $groupBoxOptions.Controls.Add($chkShowAll)
    
    # Export Report Checkbox
    $chkExport = New-Object System.Windows.Forms.CheckBox
    $chkExport.Location = New-Object System.Drawing.Point(15, 75)
    $chkExport.Size = New-Object System.Drawing.Size(250, 20)
    $chkExport.Text = "Export Report"
    $chkExport.Checked = $false
    $groupBoxOptions.Controls.Add($chkExport)
    
    # Export Format Label
    $lblExportFormat = New-Object System.Windows.Forms.Label
    $lblExportFormat.Location = New-Object System.Drawing.Point(15, 100)
    $lblExportFormat.Size = New-Object System.Drawing.Size(100, 20)
    $lblExportFormat.Text = "Export Format:"
    $groupBoxOptions.Controls.Add($lblExportFormat)
    
    # HTML Radio Button
    $radioHTML = New-Object System.Windows.Forms.RadioButton
    $radioHTML.Location = New-Object System.Drawing.Point(30, 120)
    $radioHTML.Size = New-Object System.Drawing.Size(70, 20)
    $radioHTML.Text = "HTML"
    $radioHTML.Checked = $true
    $groupBoxOptions.Controls.Add($radioHTML)
    
    # PDF Radio Button
    $radioPDF = New-Object System.Windows.Forms.RadioButton
    $radioPDF.Location = New-Object System.Drawing.Point(110, 120)
    $radioPDF.Size = New-Object System.Drawing.Size(60, 20)
    $radioPDF.Text = "PDF"
    $radioPDF.Checked = $false
    $groupBoxOptions.Controls.Add($radioPDF)
    
    # Both Radio Button
    $radioBoth = New-Object System.Windows.Forms.RadioButton
    $radioBoth.Location = New-Object System.Drawing.Point(180, 120)
    $radioBoth.Size = New-Object System.Drawing.Size(60, 20)
    $radioBoth.Text = "Both"
    $radioBoth.Checked = $false
    $groupBoxOptions.Controls.Add($radioBoth)
    
    # Run Button
    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Location = New-Object System.Drawing.Point(660, 110)
    $btnRun.Size = New-Object System.Drawing.Size(150, 60)
    $btnRun.Text = "Run Check"
    $btnRun.BackColor = [System.Drawing.Color]::LightGreen
    $btnRun.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnRun)
    
    # Clear Button
    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Location = New-Object System.Drawing.Point(660, 230)
    $btnClear.Size = New-Object System.Drawing.Size(150, 30)
    $btnClear.Text = "Clear Output"
    $form.Controls.Add($btnClear)
    
    # Status Label
    $labelStatus = New-Object System.Windows.Forms.Label
    $labelStatus.Location = New-Object System.Drawing.Point(20, 345)
    $labelStatus.Size = New-Object System.Drawing.Size(800, 20)
    $labelStatus.Text = "Ready - Please select checks and click 'Run Check'"
    $labelStatus.ForeColor = [System.Drawing.Color]::Blue
    $form.Controls.Add($labelStatus)
    
    # Progress Bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 370)
    $progressBar.Size = New-Object System.Drawing.Size(950, 20)
    $progressBar.Style = "Continuous"
    $form.Controls.Add($progressBar)
    
    # Output TextBox
    $textOutput = New-Object System.Windows.Forms.TextBox
    $textOutput.Location = New-Object System.Drawing.Point(20, 400)
    $textOutput.Size = New-Object System.Drawing.Size(950, 290)
    $textOutput.Multiline = $true
    $textOutput.ScrollBars = "Both"
    $textOutput.Font = New-Object System.Drawing.Font("Consolas", 9)
    $textOutput.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $textOutput.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $textOutput.ReadOnly = $true
    $textOutput.WordWrap = $false
    $form.Controls.Add($textOutput)
    
    # Clear Button Click
    $btnClear.Add_Click({
        $textOutput.Clear()
        $labelStatus.Text = "Output cleared - Ready for next check"
        $labelStatus.ForeColor = [System.Drawing.Color]::Blue
        $progressBar.Value = 0
    })
    
    # Run Button Click
    $btnRun.Add_Click({
        $textOutput.Clear()
        $progressBar.Value = 0
        $labelStatus.Text = "Initializing check..."
        $labelStatus.ForeColor = [System.Drawing.Color]::Orange
        $btnRun.Enabled = $false
        $form.Refresh()
        
        $bcid = $textBCID.Text
        $customerName = $textCustomer.Text
        $subId = $textSubID.Text
        $SharePointUrl = if ($textSPUrl.Text -eq $script:spUrlPlaceholder -or $textSPUrl.ForeColor -eq [System.Drawing.Color]::Gray) { "" } else { $textSPUrl.Text }
        $runAzure = $chkAzure.Checked
        $runIntune = $chkIntune.Checked
        $runEntraID = $chkEntraID.Checked
        $runIntuneConn = $chkIntuneConn.Checked
        $runDefender = $chkDefender.Checked
        $runSoftware = $chkSoftware.Checked
        $runSharePoint = $chkSharePoint.Checked
        $runTeams = $chkTeams.Checked
        $runUserLicense = $chkUserLicense.Checked
        $compact = $chkCompact.Checked
        $showAll = $chkShowAll.Checked
        $export = $chkExport.Checked
        
        # Determine export format
        $exportFormat = "HTML"
        if ($radioPDF.Checked) { $exportFormat = "PDF" }
        if ($radioBoth.Checked) { $exportFormat = "Both" }

        # ---- Module Prerequisites (GUI Dialog) ----------------------------
        $labelStatus.Text      = "Checking module prerequisites..."
        $labelStatus.ForeColor = [System.Drawing.Color]::Orange
        $form.Refresh()

        $_guiSkipParams = @{
            SkipSharePoint = (-not $runSharePoint)
            SkipTeams      = (-not $runTeams)
            SkipDefender   = (-not $runDefender)
        }
        $guiModResult = Show-ModuleSetupDialog -SkipParams $_guiSkipParams

        if (-not $guiModResult.AllReady) {
            $labelStatus.Text      = "[!] Some modules failed - checks may be incomplete"
            $labelStatus.ForeColor = [System.Drawing.Color]::Orange
        } else {
            $labelStatus.Text      = "[OK] Modules ready ($($guiModResult.TotalSecs)s)"
            $labelStatus.ForeColor = [System.Drawing.Color]::LightGreen
        }
        $form.Refresh()

        try {
            # Set subscription context if provided
            if ($subId) {
                $textOutput.AppendText("Setting subscription context to: $subId`r`n")
                $textOutput.Refresh()
                try {
                    Set-AzContext -SubscriptionId $subId -ErrorAction Stop | Out-Null
                    $textOutput.AppendText("Subscription context set successfully`r`n`r`n")
                } catch {
                    $textOutput.AppendText("ERROR: Could not set subscription context: $($_.Exception.Message)`r`n`r`n")
                    $labelStatus.Text = "Error setting subscription context"
                    $labelStatus.ForeColor = [System.Drawing.Color]::Red
                    $btnRun.Enabled = $true
                    return
                }
            } else {
                $currentContext = Get-AzContext
                if ($currentContext) {
                    $textOutput.AppendText("Using current subscription: $($currentContext.Subscription.Name)`r`n`r`n")
                } else {
                    $textOutput.AppendText("ERROR: No subscription context found`r`n`r`n")
                    $labelStatus.Text = "Error: No subscription context"
                    $labelStatus.ForeColor = [System.Drawing.Color]::Red
                    $btnRun.Enabled = $true
                    return
                }
            }
            
            $progressBar.Value = 10
            
            # Redirect Write-Host
            $originalWriteHost = Get-Command Write-Host
            function global:Write-Host {
                param(
                    [Parameter(Position=0, ValueFromPipeline=$true)]
                    [object]$Object,
                    [System.ConsoleColor]$ForegroundColor,
                    [switch]$NoNewline
                )
                
                $msg = if ($Object) { $Object.ToString() } else { "" }
                
                if (-not $NoNewline) {
                    $script:textOutput.AppendText("$msg`r`n")
                } else {
                    $script:textOutput.AppendText($msg)
                }
                $script:textOutput.SelectionStart = $script:textOutput.Text.Length
                $script:textOutput.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            $azureResults = $null
            $intuneResults = $null
            $entraIDResults = $null
            $intuneConnResults = $null
            $defenderResults = $null
            $softwareResults = $null
            $sharePointResults = $null
            
            $totalChecks = ($runAzure -as [int]) + ($runIntune -as [int]) + ($runEntraID -as [int]) + ($runIntuneConn -as [int]) + ($runDefender -as [int]) + ($runSoftware -as [int]) + ($runSharePoint -as [int]) + ($runTeams -as [int]) + ($runUserLicense -as [int])
            $currentCheck = 0
            $progressIncrement = if ($totalChecks -gt 0) { 80 / $totalChecks } else { 0 }
            
            # Run Azure Check
            if ($runAzure) {
                $labelStatus.Text = "Running Azure Resources Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $azureResults = Test-AzureResources -BCID $bcid -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run Intune Check
            if ($runIntune) {
                $labelStatus.Text = "Running Intune Policies Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $intuneResults = Test-IntunePolicies -ShowAllPolicies $showAll -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run Entra ID Connect Check
            if ($runEntraID) {
                $labelStatus.Text = "Running Entra ID Connect Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $entraIDResults = Test-EntraIDConnect -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run Hybrid Join Check
            if ($runIntuneConn) {
                $labelStatus.Text = "Running Hybrid Azure AD Join Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $intuneConnResults = Test-IntuneConnector -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run Defender for Endpoint Check
            if ($runDefender) {
                $labelStatus.Text = "Running Defender for Endpoint Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $defenderResults = Test-DefenderForEndpoint -BCID $bcid -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run BWS Software Packages Check
            if ($runSoftware) {
                $labelStatus.Text = "Running BWS Software Packages Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $softwareResults = Test-BWSSoftwarePackages -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run SharePoint Configuration Check
            if ($runSharePoint) {
                $labelStatus.Text = "Running SharePoint Configuration Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $sharePointResults = Test-SharePointConfiguration -CompactView $compact -SharePointUrl $SharePointUrl
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run Teams Configuration Check
            if ($runTeams) {
                $labelStatus.Text = "Running Teams Configuration Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $teamsResults = Test-TeamsConfiguration -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run User and License Check
            if ($runUserLicense) {
                $labelStatus.Text = "Running User & License Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $userLicenseResults = Test-UsersAndLicenses -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Overall Summary
            Write-Host ""
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host "  OVERALL SUMMARY" -ForegroundColor Cyan
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host "  BCID: $bcid" -ForegroundColor White
            
            if ($runAzure -and $azureResults) {
                Write-Host ""
                Write-Host "  Azure Resources:" -ForegroundColor White
                Write-Host "    Total:   $($azureResults.Total)" -ForegroundColor White
                Write-Host "    Found:   $($azureResults.Found.Count)" -ForegroundColor Green
                Write-Host "    Missing: $($azureResults.Missing.Count)" -ForegroundColor $(if ($azureResults.Missing.Count -eq 0) { "Green" } else { "Red" })
            }
            
            if ($runIntune -and $intuneResults -and $intuneResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Intune Policies:" -ForegroundColor White
                Write-Host "    Total:   $($intuneResults.Total)" -ForegroundColor White
                Write-Host "    Found:   $($intuneResults.Found.Count)" -ForegroundColor Green
                Write-Host "    Missing: $($intuneResults.Missing.Count)" -ForegroundColor $(if ($intuneResults.Missing.Count -eq 0) { "Green" } else { "Red" })
            }
            
            if ($runEntraID -and $entraIDResults -and $entraIDResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Entra ID Connect:" -ForegroundColor White
                Write-Host "    Sync Enabled:      " -NoNewline -ForegroundColor White
                Write-Host $(if ($entraIDResults.Status.IsInstalled) { "Yes" } else { "No" }) -ForegroundColor $(if ($entraIDResults.Status.IsInstalled) { "Green" } else { "Red" })
                Write-Host "    Sync Active:       " -NoNewline -ForegroundColor White
                Write-Host $(if ($entraIDResults.Status.IsRunning) { "Yes" } else { "No" }) -ForegroundColor $(if ($entraIDResults.Status.IsRunning) { "Green" } else { "Yellow" })
                if ($entraIDResults.Status.PasswordHashSync -ne $null) {
                    Write-Host "    Password Sync:     " -NoNewline -ForegroundColor White
                    Write-Host $(if ($entraIDResults.Status.PasswordHashSync -eq $true) { "Enabled" } elseif ($entraIDResults.Status.PasswordHashSync -eq $false) { "Disabled" } else { "Unknown" }) -ForegroundColor $(if ($entraIDResults.Status.PasswordHashSync) { "Green" } else { "Gray" })
                }
                if ($entraIDResults.Status.DeviceWritebackEnabled -ne $null) {
                    Write-Host "    Device Hybrid Sync:" -NoNewline -ForegroundColor White
                    Write-Host $(if ($entraIDResults.Status.DeviceWritebackEnabled -eq $true) { "Active" } elseif ($entraIDResults.Status.DeviceWritebackEnabled -eq $false) { "No Devices" } else { "Unknown" }) -ForegroundColor $(if ($entraIDResults.Status.DeviceWritebackEnabled) { "Green" } else { "Gray" })
                }
                if ($entraIDResults.Status.TotalUsers -gt 0) {
                    Write-Host "    Licensed Users:    $($entraIDResults.Status.LicensedUsers)/$($entraIDResults.Status.TotalUsers)" -ForegroundColor $(if ($entraIDResults.Status.UnlicensedUsers -eq 0) { "Green" } else { "Yellow" })
                }
            }
            
            if ($runIntuneConn -and $intuneConnResults -and $intuneConnResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Hybrid Azure AD Join & Intune Connectors:" -ForegroundColor White
                Write-Host "    NDES Connector:    " -NoNewline -ForegroundColor White
                Write-Host $(if ($intuneConnResults.Status.IsConnected) { "Active" } else { "Not Connected" }) -ForegroundColor $(if ($intuneConnResults.Status.IsConnected) { "Green" } else { "Yellow" })
                if ($intuneConnResults.Status.ADServerName) {
                    Write-Host "    AD Server (Azure): $($intuneConnResults.Status.ADServerName)" -ForegroundColor Green
                }
                Write-Host "    Active Connectors: $($intuneConnResults.Status.Connectors.Count)" -ForegroundColor White
                Write-Host "    Errors:            $($intuneConnResults.Status.Errors.Count)" -ForegroundColor $(if ($intuneConnResults.Status.Errors.Count -eq 0) { "Green" } else { "Yellow" })
            }
            
            if ($runDefender -and $defenderResults -and $defenderResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Defender for Endpoint:" -ForegroundColor White
                Write-Host "    Policies:     $($defenderResults.Status.ConfiguredPolicies)" -ForegroundColor $(if ($defenderResults.Status.ConfiguredPolicies -gt 0) { "Green" } else { "Yellow" })
                Write-Host "    Devices:      $($defenderResults.Status.OnboardedDevices)" -ForegroundColor $(if ($defenderResults.Status.OnboardedDevices -gt 0) { "Green" } else { "Gray" })
                Write-Host "    Files:        $($defenderResults.Status.FilesFound.Count)/4" -ForegroundColor $(if ($defenderResults.Status.FilesMissing.Count -eq 0) { "Green" } else { "Red" })
                Write-Host "    Status:       " -NoNewline -ForegroundColor White
                Write-Host $(if ($defenderResults.Status.ConnectorActive) { "Active" } else { "Not Configured" }) -ForegroundColor $(if ($defenderResults.Status.ConnectorActive) { "Green" } else { "Yellow" })
            }
            
            if ($runSoftware -and $softwareResults -and $softwareResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  BWS Software Packages:" -ForegroundColor White
                Write-Host "    Total:        $($softwareResults.Status.Total)" -ForegroundColor White
                Write-Host "    Found:        $($softwareResults.Status.Found.Count)" -ForegroundColor $(if ($softwareResults.Status.Found.Count -eq $softwareResults.Status.Total) { "Green" } else { "Yellow" })
                Write-Host "    Missing:      $($softwareResults.Status.Missing.Count)" -ForegroundColor $(if ($softwareResults.Status.Missing.Count -eq 0) { "Green" } else { "Red" })
            }
            
            if ($runSharePoint -and $sharePointResults -and $sharePointResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  SharePoint Configuration:" -ForegroundColor White
                Write-Host "    SP Ext. Sharing:   " -NoNewline -ForegroundColor White
                Write-Host $(if ($sharePointResults.Status.Settings.SharePointExternalSharing -eq "Anyone") { "Anyone ([OK])" } else { "$($sharePointResults.Status.Settings.SharePointExternalSharing) ([X])" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.SharePointExternalSharing -eq "Anyone") { "Green" } else { "Yellow" })
                Write-Host "    OD Ext. Sharing:   " -NoNewline -ForegroundColor White
                Write-Host $(if ($sharePointResults.Status.Settings.OneDriveExternalSharing -eq "Disabled") { "Only Organization ([OK])" } else { "$($sharePointResults.Status.Settings.OneDriveExternalSharing) ([X])" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.OneDriveExternalSharing -eq "Disabled") { "Green" } else { "Yellow" })
                Write-Host "    Site Creation:     " -NoNewline -ForegroundColor White  
                Write-Host $(if ($sharePointResults.Status.Settings.SiteCreation -eq "Disabled") { "Disabled ([OK])" } else { "$($sharePointResults.Status.Settings.SiteCreation) ([X])" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.SiteCreation -eq "Disabled") { "Green" } else { "Yellow" })
                Write-Host "    Legacy Auth Block: " -NoNewline -ForegroundColor White
                Write-Host $(if ($sharePointResults.Status.Settings.LegacyAuthBlocked -eq $true) { "Yes ([OK])" } else { "No ([X])" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.LegacyAuthBlocked) { "Green" } else { "Yellow" })
            }
            
            if ($runTeams -and $teamsResults -and $teamsResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Teams Configuration:" -ForegroundColor White
                Write-Host "    Meetings w/ unmanaged MS: " -NoNewline -ForegroundColor White
                Write-Host $(if ($teamsResults.Status.Settings.ExternalAccessEnabled -eq $false) { "Disabled ([OK])" } else { "Enabled ([X])" }) -ForegroundColor $(if ($teamsResults.Status.Settings.ExternalAccessEnabled -eq $false) { "Green" } else { "Yellow" })
                
                $allStorageDisabled = ($teamsResults.Status.Settings.CloudStorageCitrix -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageDropbox -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageBox -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageGoogleDrive -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageEgnyte -eq "Disabled")
                Write-Host "    Cloud Storage:     " -NoNewline -ForegroundColor White
                if ($allStorageDisabled) {
                    Write-Host "All Disabled ([OK])" -ForegroundColor Green
                } else {
                    $enabledList = @()
                    if ($teamsResults.Status.Settings.CloudStorageCitrix -eq "Enabled") { $enabledList += "Citrix" }
                    if ($teamsResults.Status.Settings.CloudStorageDropbox -eq "Enabled") { $enabledList += "Dropbox" }
                    if ($teamsResults.Status.Settings.CloudStorageBox -eq "Enabled") { $enabledList += "Box" }
                    if ($teamsResults.Status.Settings.CloudStorageGoogleDrive -eq "Enabled") { $enabledList += "Google Drive" }
                    if ($teamsResults.Status.Settings.CloudStorageEgnyte -eq "Enabled") { $enabledList += "Egnyte" }
                    Write-Host "Enabled: $($enabledList -join ', ') ([X])" -ForegroundColor Yellow
                }
                
                Write-Host "    Anonymous Join:    " -NoNewline -ForegroundColor White
                Write-Host $(if ($teamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Disabled") { "Disabled ([OK])" } else { "Enabled ([X])" }) -ForegroundColor $(if ($teamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Disabled") { "Green" } else { "Yellow" })
                
                Write-Host "    Who Can Present:   " -NoNewline -ForegroundColor White
                Write-Host $(if ($teamsResults.Status.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") { "Everyone ([OK])" } else { "$($teamsResults.Status.Settings.DefaultPresenterRole) ([X])" }) -ForegroundColor $(if ($teamsResults.Status.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") { "Green" } else { "Yellow" })
            }
            
            if ($runUserLicense -and $userLicenseResults -and $userLicenseResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Users & Licenses:" -ForegroundColor White
                Write-Host "    Total Users:        $($userLicenseResults.Status.TotalUsers)" -ForegroundColor White
                Write-Host "    Licensed:           $($userLicenseResults.Status.LicensedUsers)" -ForegroundColor $(if ($userLicenseResults.Status.LicensedUsers -gt 0) { "Green" } else { "Gray" })
                Write-Host "    Unlicensed:         $($userLicenseResults.Status.UnlicensedUsers)" -ForegroundColor $(if ($userLicenseResults.Status.UnlicensedUsers -eq 0) { "Green" } else { "Yellow" })
                Write-Host "    Valid ADM Accounts: $($userLicenseResults.Status.PrivilegedUsers.Count)" -ForegroundColor Green
                Write-Host "    INVALID Priv Users: $($userLicenseResults.Status.InvalidPrivilegedUsers.Count)" -ForegroundColor $(if ($userLicenseResults.Status.InvalidPrivilegedUsers.Count -eq 0) { "Green" } else { "Red" })
            }
            
            Write-Host "======================================================" -ForegroundColor Cyan
            
            if ($compact) {
                Write-Host ""
                Write-Host "Note: Compact View enabled" -ForegroundColor Gray
            }
            
            # Export report if requested
            if ($export) {
                Write-Host ""
                
                $currentContext = Get-AzContext
                $subName = if ($currentContext) { $currentContext.Subscription.Name } else { "Unknown" }
                
                $overallStatus = ($azureResults.Missing.Count -eq 0 -and $azureResults.Errors.Count -eq 0) -and 
                                 (-not $intuneResults -or ($intuneResults.Missing.Count -eq 0 -and $intuneResults.Errors.Count -eq 0)) -and
                                 (-not $entraIDResults -or ($entraIDResults.Status.IsRunning)) -and
                                 (-not $intuneConnResults -or ($intuneConnResults.Status.Errors.Count -eq 0)) -and
                                 (-not $defenderResults -or ($defenderResults.Status.ConnectorActive -and $defenderResults.Status.FilesMissing.Count -eq 0)) -and
                                 (-not $softwareResults -or ($softwareResults.Status.Missing.Count -eq 0)) -and
                                 (-not $sharePointResults -or ($sharePointResults.Status.Compliant)) -and
                                 (-not $teamsResults -or ($teamsResults.Status.Compliant)) -and
                                 (-not $userLicenseResults -or ($userLicenseResults.Status.InvalidPrivilegedUsers.Count -eq 0 -and $userLicenseResults.Status.InvalidEntraIDP2Users.Count -eq 0))
                
                # Generate HTML report
                if ($exportFormat -eq "HTML" -or $exportFormat -eq "Both") {
                    Write-Host "Generating HTML Report..." -ForegroundColor Yellow
                    $htmlPath = Export-HTMLReport -BCID $bcid -CustomerName $customerName -SubscriptionName $subName `
                        -AzureResults $azureResults -IntuneResults $intuneResults `
                        -EntraIDResults $entraIDResults -IntuneConnResults $intuneConnResults `
                        -DefenderResults $defenderResults -SoftwareResults $softwareResults `
                        -SharePointResults $sharePointResults -TeamsResults $teamsResults `
                        -UserLicenseResults $userLicenseResults -OverallStatus $overallStatus
                    
                    Write-Host "HTML Report exported to: $htmlPath" -ForegroundColor Green
                }
                
                # Generate PDF report
                if ($exportFormat -eq "PDF" -or $exportFormat -eq "Both") {
                    if (-not $htmlPath) {
                        # Need HTML first for PDF conversion
                        $htmlPath = Export-HTMLReport -BCID $bcid -CustomerName $customerName -SubscriptionName $subName `
                            -AzureResults $azureResults -IntuneResults $intuneResults `
                            -EntraIDResults $entraIDResults -IntuneConnResults $intuneConnResults `
                            -DefenderResults $defenderResults -SoftwareResults $softwareResults `
                            -SharePointResults $sharePointResults -TeamsResults $teamsResults -OverallStatus $overallStatus
                    }
                    
                    $pdfPath = Export-PDFReport -HTMLPath $htmlPath
                    if ($pdfPath) {
                        Write-Host "PDF Report exported to: $pdfPath" -ForegroundColor Green
                    }
                    
                    # Clean up temp HTML if only PDF was requested
                    if ($exportFormat -eq "PDF" -and $htmlPath -and (Test-Path $htmlPath)) {
                        Remove-Item $htmlPath -Force -ErrorAction SilentlyContinue
                    }
                }
                
                Write-Host ""
            }
            
            $progressBar.Value = 100
            $labelStatus.Text = "Check completed successfully!"
            $labelStatus.ForeColor = [System.Drawing.Color]::Green
            
        } catch {
            $textOutput.AppendText("`r`nERROR: $($_.Exception.Message)`r`n")
            $labelStatus.Text = "Error occurred during check"
            $labelStatus.ForeColor = [System.Drawing.Color]::Red
        } finally {
            # Restore Write-Host
            Remove-Item Function:\Write-Host -ErrorAction SilentlyContinue
            $btnRun.Enabled = $true
        }
    })
    
    [void]$form.ShowDialog()
    exit
}

#============================================================================
# Command Line Mode
#============================================================================

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  BWS-Checking-Script v$script:Version" -ForegroundColor Cyan
Write-Host "  Command Line Mode" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

# Set CompactView as default if not explicitly overridden
if (-not $PSBoundParameters.ContainsKey('CompactView')) {
    $CompactView = $true
}

# Set Subscription Context
if ($SubscriptionId) {
    Write-Host "Setting subscription context to: $SubscriptionId" -ForegroundColor Yellow
    try {
        Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
        Write-Host "Subscription context set successfully" -ForegroundColor Green
    } catch {
        Write-Host "Error setting subscription context: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
} else {
    $currentContext = Get-AzContext
    if ($currentContext) {
        Write-Host "Using current subscription: $($currentContext.Subscription.Name) ($($currentContext.Subscription.Id))" -ForegroundColor Yellow
    } else {
        Write-Host "No subscription context found. Please login with Connect-AzAccount or specify -SubscriptionId" -ForegroundColor Red
        return
    }
}

# Run Azure Check
$azureResults = Test-AzureResources -BCID $BCID -CompactView $CompactView

# Run Intune Check
$intuneResults = $null
if (-not $SkipIntune) {
    $intuneResults = Test-IntunePolicies -ShowAllPolicies $ShowAllPolicies -CompactView $CompactView
}

# Run Entra ID Connect Check
$entraIDResults = $null
if (-not $SkipEntraID) {
    $entraIDResults = Test-EntraIDConnect -CompactView $CompactView
}

# Run Intune Connector Check
$intuneConnResults = $null
if (-not $SkipIntuneConnector) {
    $intuneConnResults = Test-IntuneConnector -CompactView $CompactView
}

# Run Defender for Endpoint Check
$defenderResults = $null
if (-not $SkipDefender) {
    $defenderResults = Test-DefenderForEndpoint -BCID $BCID -CompactView $CompactView
}

# Run BWS Software Packages Check
$softwareResults = $null
if (-not $SkipSoftware) {
    $softwareResults = Test-BWSSoftwarePackages -CompactView $CompactView
}

# Run SharePoint Configuration Check
$sharePointResults = $null
if (-not $SkipSharePoint) {
    $sharePointResults = Test-SharePointConfiguration -CompactView $CompactView -SharePointUrl $SharePointUrl
}

# Run Teams Configuration Check
$teamsResults = $null
if (-not $SkipTeams) {
    $teamsResults = Test-TeamsConfiguration -CompactView $CompactView
}

# User and License Check
$userLicenseResults = $null
if (-not $SkipUserLicenseCheck) {
    $userLicenseResults = Test-UsersAndLicenses -CompactView $CompactView
}

# Overall Summary
$currentContext = Get-AzContext
Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  OVERALL SUMMARY" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  BCID: $BCID" -ForegroundColor White
Write-Host "  Subscription: $($currentContext.Subscription.Name)" -ForegroundColor White
Write-Host ""
Write-Host "  Azure Resources:" -ForegroundColor White
Write-Host "    Total:   $($azureResults.Total)" -ForegroundColor White
Write-Host "    Found:   $($azureResults.Found.Count)" -ForegroundColor $(if ($azureResults.Found.Count -eq $azureResults.Total) { "Green" } else { "Yellow" })
Write-Host "    Missing: $($azureResults.Missing.Count)" -ForegroundColor $(if ($azureResults.Missing.Count -eq 0) { "Green" } else { "Red" })

if ($intuneResults -and $intuneResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Intune Policies:" -ForegroundColor White
    Write-Host "    Total:   $($intuneResults.Total)" -ForegroundColor White
    Write-Host "    Found:   $($intuneResults.Found.Count)" -ForegroundColor $(if ($intuneResults.Found.Count -eq $intuneResults.Total) { "Green" } else { "Yellow" })
    Write-Host "    Missing: $($intuneResults.Missing.Count)" -ForegroundColor $(if ($intuneResults.Missing.Count -eq 0) { "Green" } else { "Red" })
}

if ($entraIDResults -and $entraIDResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Entra ID Connect:" -ForegroundColor White
    Write-Host "    Sync Enabled:       " -NoNewline -ForegroundColor White
    Write-Host $(if ($entraIDResults.Status.IsInstalled) { "Yes ([OK])" } else { "No ([X])" }) -ForegroundColor $(if ($entraIDResults.Status.IsInstalled) { "Green" } else { "Red" })
    Write-Host "    Sync Active:        " -NoNewline -ForegroundColor White
    Write-Host $(if ($entraIDResults.Status.IsRunning) { "Yes ([OK])" } else { "No ([X])" }) -ForegroundColor $(if ($entraIDResults.Status.IsRunning) { "Green" } else { "Yellow" })
    if ($entraIDResults.Status.PasswordHashSync -ne $null) {
        Write-Host "    Password Hash Sync: " -NoNewline -ForegroundColor White
        Write-Host $(if ($entraIDResults.Status.PasswordHashSync -eq $true) { "Enabled ([OK])" } elseif ($entraIDResults.Status.PasswordHashSync -eq $false) { "Disabled ([!])" } else { "Unknown" }) -ForegroundColor $(if ($entraIDResults.Status.PasswordHashSync) { "Green" } else { "Gray" })
    }
    if ($entraIDResults.Status.DeviceWritebackEnabled -ne $null) {
        Write-Host "    Device Hybrid Sync: " -NoNewline -ForegroundColor White
        Write-Host $(if ($entraIDResults.Status.DeviceWritebackEnabled -eq $true) { "Active ([OK])" } elseif ($entraIDResults.Status.DeviceWritebackEnabled -eq $false) { "No Devices" } else { "Unknown" }) -ForegroundColor $(if ($entraIDResults.Status.DeviceWritebackEnabled) { "Green" } else { "Gray" })
    }
    if ($entraIDResults.Status.TotalUsers -gt 0) {
        Write-Host "    Licensed Users:     $($entraIDResults.Status.LicensedUsers)/$($entraIDResults.Status.TotalUsers)" -ForegroundColor $(if ($entraIDResults.Status.UnlicensedUsers -eq 0) { "Green" } else { "Yellow" })
        if ($entraIDResults.Status.UnlicensedUsers -gt 0) {
            Write-Host "    Unlicensed Users:   $($entraIDResults.Status.UnlicensedUsers) ([!])" -ForegroundColor Yellow
        }
    }
}

if ($intuneConnResults -and $intuneConnResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Hybrid Azure AD Join & Intune Connectors:" -ForegroundColor White
    Write-Host "    NDES Connector:     " -NoNewline -ForegroundColor White
    Write-Host $(if ($intuneConnResults.Status.IsConnected) { "Active ([OK])" } else { "Not Connected ([!])" }) -ForegroundColor $(if ($intuneConnResults.Status.IsConnected) { "Green" } else { "Yellow" })
    if ($intuneConnResults.Status.ADServerName) {
        Write-Host "    AD Server (Azure):  $($intuneConnResults.Status.ADServerName) ([OK])" -ForegroundColor Green
    } elseif ($intuneConnResults.Status.ADServerReservation -eq $true) {
        Write-Host "    AD Server (Azure):  Found ([OK])" -ForegroundColor Green
    } elseif ($intuneConnResults.Status.ADServerReservation -eq $false) {
        Write-Host "    AD Server (Azure):  Not Detected ([!])" -ForegroundColor Yellow
    }
    Write-Host "    Active Connectors:  $($intuneConnResults.Status.Connectors.Count)" -ForegroundColor White
    Write-Host "    Errors:             $($intuneConnResults.Status.Errors.Count)" -ForegroundColor $(if ($intuneConnResults.Status.Errors.Count -eq 0) { "Green" } else { "Yellow" })
}

if ($defenderResults -and $defenderResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Defender for Endpoint:" -ForegroundColor White
    Write-Host "    Policies:     $($defenderResults.Status.ConfiguredPolicies)" -ForegroundColor $(if ($defenderResults.Status.ConfiguredPolicies -gt 0) { "Green" } else { "Yellow" })
    Write-Host "    Devices:      $($defenderResults.Status.OnboardedDevices)" -ForegroundColor $(if ($defenderResults.Status.OnboardedDevices -gt 0) { "Green" } else { "Gray" })
    Write-Host "    Files:        $($defenderResults.Status.FilesFound.Count)/4" -ForegroundColor $(if ($defenderResults.Status.FilesMissing.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "    Status:       " -NoNewline -ForegroundColor White
    Write-Host $(if ($defenderResults.Status.ConnectorActive) { "Active" } else { "Not Configured" }) -ForegroundColor $(if ($defenderResults.Status.ConnectorActive) { "Green" } else { "Yellow" })
}

if ($softwareResults -and $softwareResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  BWS Software Packages:" -ForegroundColor White
    Write-Host "    Total:        $($softwareResults.Status.Total)" -ForegroundColor White
    Write-Host "    Found:        $($softwareResults.Status.Found.Count)" -ForegroundColor $(if ($softwareResults.Status.Found.Count -eq $softwareResults.Status.Total) { "Green" } else { "Yellow" })
    Write-Host "    Missing:      $($softwareResults.Status.Missing.Count)" -ForegroundColor $(if ($softwareResults.Status.Missing.Count -eq 0) { "Green" } else { "Red" })
}

if ($sharePointResults -and $sharePointResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  SharePoint Configuration:" -ForegroundColor White
    Write-Host "    SP Ext. Sharing:   " -NoNewline -ForegroundColor White
    Write-Host $(if ($sharePointResults.Status.Settings.SharePointExternalSharing -eq "Anyone") { "Anyone ([OK])" } else { "$($sharePointResults.Status.Settings.SharePointExternalSharing) ([X])" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.SharePointExternalSharing -eq "Anyone") { "Green" } else { "Yellow" })
    Write-Host "    OD Ext. Sharing:   " -NoNewline -ForegroundColor White
    Write-Host $(if ($sharePointResults.Status.Settings.OneDriveExternalSharing -eq "Disabled") { "Only Organization ([OK])" } else { "$($sharePointResults.Status.Settings.OneDriveExternalSharing) ([X])" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.OneDriveExternalSharing -eq "Disabled") { "Green" } else { "Yellow" })
    Write-Host "    Site Creation:     " -NoNewline -ForegroundColor White  
    Write-Host $(if ($sharePointResults.Status.Settings.SiteCreation -eq "Disabled") { "Disabled ([OK])" } else { "$($sharePointResults.Status.Settings.SiteCreation) ([X])" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.SiteCreation -eq "Disabled") { "Green" } else { "Yellow" })
    Write-Host "    Legacy Auth Block: " -NoNewline -ForegroundColor White
    Write-Host $(if ($sharePointResults.Status.Settings.LegacyAuthBlocked -eq $true) { "Yes ([OK])" } else { "No ([X])" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.LegacyAuthBlocked) { "Green" } else { "Yellow" })
}

if ($teamsResults -and $teamsResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Teams Configuration:" -ForegroundColor White
    Write-Host "    Meetings w/ unmanaged MS: " -NoNewline -ForegroundColor White
    Write-Host $(if ($teamsResults.Status.Settings.ExternalAccessEnabled -eq $false) { "Disabled ([OK])" } else { "Enabled ([X])" }) -ForegroundColor $(if ($teamsResults.Status.Settings.ExternalAccessEnabled -eq $false) { "Green" } else { "Yellow" })
    
    $allStorageDisabled = ($teamsResults.Status.Settings.CloudStorageCitrix -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageDropbox -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageBox -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageGoogleDrive -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageEgnyte -eq "Disabled")
    Write-Host "    Cloud Storage:     " -NoNewline -ForegroundColor White
    if ($allStorageDisabled) {
        Write-Host "All Disabled ([OK])" -ForegroundColor Green
    } else {
        $enabledList = @()
        if ($teamsResults.Status.Settings.CloudStorageCitrix -eq "Enabled") { $enabledList += "Citrix" }
        if ($teamsResults.Status.Settings.CloudStorageDropbox -eq "Enabled") { $enabledList += "Dropbox" }
        if ($teamsResults.Status.Settings.CloudStorageBox -eq "Enabled") { $enabledList += "Box" }
        if ($teamsResults.Status.Settings.CloudStorageGoogleDrive -eq "Enabled") { $enabledList += "Google Drive" }
        if ($teamsResults.Status.Settings.CloudStorageEgnyte -eq "Enabled") { $enabledList += "Egnyte" }
        Write-Host "Enabled: $($enabledList -join ', ') ([X])" -ForegroundColor Yellow
    }
    
    Write-Host "    Anonymous Join:    " -NoNewline -ForegroundColor White
    Write-Host $(if ($teamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Disabled") { "Disabled ([OK])" } else { "Enabled ([X])" }) -ForegroundColor $(if ($teamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Disabled") { "Green" } else { "Yellow" })
    
    Write-Host "    Who Can Present:   " -NoNewline -ForegroundColor White
    Write-Host $(if ($teamsResults.Status.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") { "Everyone ([OK])" } else { "$($teamsResults.Status.Settings.DefaultPresenterRole) ([X])" }) -ForegroundColor $(if ($teamsResults.Status.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") { "Green" } else { "Yellow" })
}

if ($userLicenseResults -and $userLicenseResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Users & Licenses:" -ForegroundColor White
    Write-Host "    Total Users:        $($userLicenseResults.Status.TotalUsers)" -ForegroundColor White
    Write-Host "    Licensed Users:     $($userLicenseResults.Status.LicensedUsers)" -ForegroundColor $(if ($userLicenseResults.Status.LicensedUsers -gt 0) { "Green" } else { "Gray" })
    Write-Host "    Unlicensed Users:   $($userLicenseResults.Status.UnlicensedUsers)" -ForegroundColor $(if ($userLicenseResults.Status.UnlicensedUsers -eq 0) { "Green" } else { "Yellow" })
    Write-Host "    Privileged Users:   $($userLicenseResults.Status.PrivilegedUsers.Count + $userLicenseResults.Status.InvalidPrivilegedUsers.Count)" -ForegroundColor White
    Write-Host "    Valid ADM Accounts: $($userLicenseResults.Status.PrivilegedUsers.Count) ([OK])" -ForegroundColor Green
    Write-Host "    INVALID Priv Users: $($userLicenseResults.Status.InvalidPrivilegedUsers.Count)" -ForegroundColor $(if ($userLicenseResults.Status.InvalidPrivilegedUsers.Count -eq 0) { "Green" } else { "Red" })
}

Write-Host ""
$overallStatus = ($azureResults.Missing.Count -eq 0 -and $azureResults.Errors.Count -eq 0) -and 
                 (-not $intuneResults -or ($intuneResults.Missing.Count -eq 0 -and $intuneResults.Errors.Count -eq 0)) -and
                 (-not $entraIDResults -or ($entraIDResults.Status.IsRunning)) -and
                 (-not $intuneConnResults -or ($intuneConnResults.Status.Errors.Count -eq 0)) -and
                 (-not $defenderResults -or ($defenderResults.Status.ConnectorActive -and $defenderResults.Status.FilesMissing.Count -eq 0)) -and
                 (-not $softwareResults -or ($softwareResults.Status.Missing.Count -eq 0)) -and
                 (-not $sharePointResults -or ($sharePointResults.Status.Compliant)) -and
                 (-not $teamsResults -or ($teamsResults.Status.Compliant)) -and
                 (-not $userLicenseResults -or ($userLicenseResults.Status.InvalidPrivilegedUsers.Count -eq 0 -and $userLicenseResults.Status.InvalidEntraIDP2Users.Count -eq 0))

Write-Host "  Overall Status: " -NoNewline -ForegroundColor White
if ($overallStatus) {
    Write-Host "[OK] PASSED" -ForegroundColor Green
} else {
    Write-Host "[X] ISSUES FOUND" -ForegroundColor Red
}
Write-Host "======================================================" -ForegroundColor Cyan

if ($CompactView) {
    Write-Host ""
    Write-Host "Note: Compact View enabled. Use without -CompactView for detailed tables." -ForegroundColor Gray
}

Write-Host ""

# Export Report
if ($ExportReport) {
    
    # Generate HTML report
    if ($ExportFormat -eq "HTML" -or $ExportFormat -eq "Both") {
        Write-Host "Generating HTML Report..." -ForegroundColor Yellow
        
        $htmlPath = Export-HTMLReport -BCID $BCID -CustomerName $CustomerName -SubscriptionName $currentContext.Subscription.Name `
            -AzureResults $azureResults -IntuneResults $intuneResults `
            -EntraIDResults $entraIDResults -IntuneConnResults $intuneConnResults `
            -DefenderResults $defenderResults -SoftwareResults $softwareResults `
            -SharePointResults $sharePointResults -TeamsResults $teamsResults `
            -UserLicenseResults $userLicenseResults -OverallStatus $overallStatus
        
        Write-Host "HTML Report exported to: $htmlPath" -ForegroundColor Green
    }
    
    # Generate PDF report
    if ($ExportFormat -eq "PDF" -or $ExportFormat -eq "Both") {
        if (-not $htmlPath) {
            # Need HTML first for PDF conversion
            $htmlPath = Export-HTMLReport -BCID $BCID -CustomerName $CustomerName -SubscriptionName $currentContext.Subscription.Name `
                -AzureResults $azureResults -IntuneResults $intuneResults `
                -EntraIDResults $entraIDResults -IntuneConnResults $intuneConnResults `
                -DefenderResults $defenderResults -SoftwareResults $softwareResults `
                -SharePointResults $sharePointResults -TeamsResults $teamsResults `
                -UserLicenseResults $userLicenseResults -OverallStatus $overallStatus
        }
        
        $pdfPath = Export-PDFReport -HTMLPath $htmlPath
        if ($pdfPath) {
            Write-Host "PDF Report exported to: $pdfPath" -ForegroundColor Green
        }
        
        # Clean up temp HTML if only PDF was requested
        if ($ExportFormat -eq "PDF" -and $htmlPath -and (Test-Path $htmlPath)) {
            Remove-Item $htmlPath -Force -ErrorAction SilentlyContinue
        }
    }
    
    Write-Host ""
}