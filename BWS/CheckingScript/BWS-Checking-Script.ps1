#Requires -Version 5.1
# PS edition check is done at runtime (see below) - no #Requires directive so the
# script can be launched from pwsh/Core terminals while still targeting PS 5.1 Desktop.
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
.PARAMETER Full
    Run module installation/import and login checks first, then run all selected
    checks. Without -Full (or without clicking "Prerequisites" in the GUI),
    only the checks themselves are executed - no module setup, no login checks.
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
    [switch]$SkipSecurity,
    
    [Parameter(Mandatory=$false)]
    [switch]$SupportBundle,
    
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
               HelpMessage="Check, install and display all required PowerShell modules without running any checks.")]
    [switch]$Full,

    [Parameter(Mandatory=$false,
               HelpMessage="Run built-in Pester-style unit tests before executing checks.")]
    [switch]$RunTests,

    [Parameter(Mandatory=$false,
               HelpMessage="Show diagnostics: logged-in user, tenant, subscription, module versions, PS version.")]
    [switch]$Diagnostics
)

# Runtime PS edition / version guard
if ($PSVersionTable.PSEdition -eq 'Core') {
    Write-Warning "This script is designed for Windows PowerShell 5.1 (Desktop edition)."
    Write-Warning "You are running PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))."
    Write-Warning "The Microsoft.Online.SharePoint.PowerShell module requires Desktop (PS 5.1)."
    Write-Warning "SharePoint checks will be SKIPPED automatically."
    Write-Warning "All other checks will continue normally."
    Write-Host ""
    # Force SharePoint skip when not on Desktop edition
    $script:SkipSharePoint = $true
}
if ($PSVersionTable.PSVersion.Major -lt 5 -or
    ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) {
    if ($GUI) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            "PowerShell 5.1 or higher is required.`nCurrent: $($PSVersionTable.PSVersion)",
            "BWS Checking Script - Version Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } else {
        Write-Host "[X] PowerShell 5.1 or higher is required. Current: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    }
    exit 1
}


# Script Version
$script:Version = "3.0.0"

# -- Account Session State ------------------------------------------------
# The script uses exactly TWO accounts:
#   1. GlobalAdmin  - has all M365 + Azure Subscription rights
#      Used for: Az (Graph/ARM), Teams, SharePoint, Entra/Intune checks
#   2. DomainAdmin  - local Active Directory only (not used in this script
#      for online checks; reserved for future on-prem AD queries)
#
# The GlobalAdmin credential is asked once at startup and reused for all
# services to avoid multiple interactive browser windows.
# Microsoft Learn:
#   Connect-MicrosoftTeams -AccountId -TenantId (reuses AAD session)
#   Connect-SPOService -Credential (PSCredential, PS 5.1 only)
# -------------------------------------------------------------------------
$script:GlobalAdminUPN       = $null   # set by Connect-BWsGlobalAdmin
$script:GlobalAdminTenantId  = $null   # set by Connect-BWsGlobalAdmin
$script:GlobalAdminConnected = $false  # $true once Az + Graph confirmed
$script:TeamsConnected       = $false  # $true after Connect-MicrosoftTeams
$script:SharePointConnected  = $false  # $true after Connect-SPOService

# -- Central Error Log (all check runs accumulate here) --------------------
# Each entry is a PSCustomObject with Code, Severity, Category, Function,
# CheckStep, Message, Detail, HttpStatus, Timestamp.
# MS Learn (ErrorRecord): learn.microsoft.com/powershell/scripting/developer/
#   cmdlet/adding-non-terminating-error-reporting-to-your-cmdlet
# Use Write-BWsError instead of .Errors += "string" in all check functions.
# Export with Get-BWsSupportBundle for support tickets.
$script:BWS_ErrorLog = [System.Collections.Generic.List[object]]::new()

#============================================================================
# QUALITY ASSURANCE - Block 1b: PSScriptAnalyzer (optional, -RunAnalyzer)
# Runs FIRST - before any module import, login, or network call.
#============================================================================
if ($RunAnalyzer) {
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Magenta
    Write-Host "  PSScriptAnalyzer - Static Code Analysis" -ForegroundColor Magenta
    Write-Host "======================================================" -ForegroundColor Magenta
    Write-Host ""

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
            Write-Host "  [OK] No errors or warnings found." -ForegroundColor Green
        }
        Import-Module PSScriptAnalyzer -ErrorAction SilentlyContinue | Out-Null
    }

    Write-Host "======================================================" -ForegroundColor Magenta
    Write-Host ""

    if (-not $GUI) {
        $continueAfterAnalysis = Read-Host "Continue with checks? (J/N)"
        if ($continueAfterAnalysis -notin @("J","j","Y","y")) {
            Write-Host "Script stopped after PSScriptAnalyzer run." -ForegroundColor Yellow
            exit 0
        }
    }
}

#============================================================================
# QUALITY ASSURANCE - Block 1: Strict Mode
#============================================================================
# Set-StrictMode -Version Latest

# Remove any Microsoft.Graph modules that may have been loaded in this PS session
# (e.g. from a previous script run). Microsoft.Graph SDK v2.x cannot load in PS 5.1
# and causes a GetTokenAsync / .NET Framework incompatibility error if left in session.
$null = Get-Module -Name 'Microsoft.Graph*' -ErrorAction SilentlyContinue |
        Remove-Module -Force -ErrorAction SilentlyContinue

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
    
    # In GUI mode: auto-disable SharePoint (no Read-Host possible without console)
    # In console mode: ask user
    if (-not $SkipSharePoint) {
        if ($GUI) {
            $script:SkipSharePoint = $true
            $SkipSharePoint        = $true
            Write-Host "SharePoint-Check automatisch deaktiviert (PS Core + GUI-Modus)." -ForegroundColor Yellow
        } else {
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
}

#============================================================================

#============================================================================
# Required Modules Definition
#============================================================================

$script:RequiredModules = @(
    #
    # MODULE REQUIREMENTS  -  Windows PowerShell 5.1 (Desktop edition)
    # -----------------------------------------------------------------
    # Only modules whose PS cmdlets are ACTUALLY CALLED in this script.
    # All Graph API calls use Invoke-AzRestMethod (Az.Accounts) via REST -
    # no Microsoft.Graph SDK modules needed or installed.
    # All Azure resource queries use Find-BWsAzResource (ARM REST) -
    # Az.Resources is NOT needed.
    #
    # Required=true  -> install + import, abort check on failure
    # Required=false -> install + import only if SkipParam not set
    #
    # Microsoft Learn install guidance:
    #   learn.microsoft.com/powershell/azure/install-azps-windows
    #   Import order: Az.Accounts MUST be imported before Az.Storage
    # -----------------------------------------------------------------

    # -- CORE (always required) -------------------------------------------
    # Az.Accounts: Connect-AzAccount, Get-AzContext, Set-AzContext,
    #              Invoke-AzRestMethod (handles ALL Graph + ARM REST auth)
    # MS Learn: github.com/Azure/azure-powershell - PS 5.1 compatible
    @{ Name="Az.Accounts";
       MinVersion="2.4.0";  MaxVersion="";
       Description="Azure Auth + Invoke-AzRestMethod (Graph/ARM)";
       Scope="CurrentUser"; Required=$true;  SkipParam="" },

    # -- CONDITIONAL (skipped if corresponding -Skip flag is set) ---------
    # Az.Storage: Get-AzStorageAccount, Get-AzStorageContainer,
    #             Get-AzStorageBlob  (Defender file check)
    # MS Learn: learn.microsoft.com/powershell/module/az.storage
    @{ Name="Az.Storage";
       MinVersion="5.0.0";  MaxVersion="";
       Description="Storage Blobs (Defender file check)";
       Scope="CurrentUser"; Required=$false; SkipParam="SkipDefender" },

    # Microsoft.Online.SharePoint.PowerShell:
    #   Connect-SPOService, Get-SPOTenant, Get-SPOSite
    #   NOTE: PS 5.1 (Desktop edition) ONLY - not supported in PS Core / PS 7
    # MS Learn: learn.microsoft.com/powershell/sharepoint/sharepoint-online/connect-sharepoint-online
    @{ Name="Microsoft.Online.SharePoint.PowerShell";
       MinVersion="16.0.0"; MaxVersion="";
       Description="SharePoint Online Admin (PS 5.1 only)";
       Scope="CurrentUser"; Required=$false; SkipParam="SkipSharePoint" },

    # MicrosoftTeams: Connect-MicrosoftTeams, Get-CsTeamsMeetingPolicy,
    #                 Get-CsMeetingConfiguration, Get-CsTeamsClientConfiguration
    # MS Learn: learn.microsoft.com/microsoftteams/teams-powershell-install
    @{ Name="MicrosoftTeams";
       MinVersion="4.0.0";  MaxVersion="";
       Description="Teams Admin (Connect-MicrosoftTeams)";
       Scope="CurrentUser"; Required=$false; SkipParam="SkipTeams" },

    # -- OPTIONAL ---------------------------------------------------------
    # PSScriptAnalyzer: Invoke-ScriptAnalyzer  (only with -RunAnalyzer)
    # MS Learn: learn.microsoft.com/powershell/utility-modules/psscriptanalyzer/overview
    @{ Name="PSScriptAnalyzer";
       MinVersion="1.20.0"; MaxVersion="";
       Description="Static code analysis (-RunAnalyzer only)";
       Scope="CurrentUser"; Required=$false; SkipParam="" }
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

#============================================================================
# ERROR SYSTEM - Write-BWsError / Get-BWsSupportBundle
#============================================================================

function Write-BWsError {
    <#
    .SYNOPSIS
        Central error/warning recorder for all BWS check functions.
        Replaces plain .Errors += "string" with structured error records.
    .NOTES
        Error code scheme:  BWS-{CAT}-{NNN}
          AUTH   Login, Graph token, Az session
          GRAPH  Graph API HTTP errors, throttling
          INTUNE MDM policies, connectors, apps
          ENTRA  Entra ID sync, PIM, roles
          SPO    SharePoint Online connection/config
          TEAMS  Teams configuration
          SEC    Security config (MDM, grp, SICT, PIM)
          SYS    PowerShell environment, modules

        Usage:
          $err = Write-BWsError -Code "BWS-ENTRA-001" -Message "Sync stale"
          $status.Errors += $err

        Or just log without appending to local status:
          Write-BWsError -Code "BWS-AUTH-003" -Message "Graph 403" -Severity Warning
    .PARAMETER Code
        Error code in BWS-CAT-NNN format.
    .PARAMETER Message
        Short human-readable description (shown in console and report).
    .PARAMETER Severity
        Error / Warning / Info  (default: Error)
    .PARAMETER Detail
        Full exception message or technical detail for support.
    .PARAMETER HttpStatus
        HTTP status code from Graph API if applicable (0 = N/A).
    .PARAMETER CheckStep
        Which check step produced this error (e.g. "[1/2] Org sync").
    .PARAMETER SuppressConsole
        Suppress the Write-Host output (log only). Useful in tight loops.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Code,
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("Error","Warning","Info")]
        [string]$Severity      = "Error",
        [string]$Detail        = "",
        [int]   $HttpStatus    = 0,
        [string]$CheckStep     = "",
        [switch]$SuppressConsole
    )

    # Derive category from code (BWS-AUTH-003 -> AUTH)
    $category = if ($Code -match '^BWS-([A-Z]+)-\d+$') { $Matches[1] } else { "GEN" }

    # Capture calling function name (skip Write-BWsError itself = frame 0)
    $callerFn = try {
        $stack = Get-PSCallStack
        if ($stack.Count -ge 2) { $stack[1].Command } else { "Unknown" }
    } catch { "Unknown" }

    $record = [PSCustomObject]@{
        Code       = $Code
        Severity   = $Severity
        Category   = $category
        Function   = $callerFn
        CheckStep  = $CheckStep
        Message    = $Message
        Detail     = if ($Detail) { $Detail } else { "" }
        HttpStatus = $HttpStatus
        Timestamp  = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Resolved   = $false
    }

    # Always add to global log
    $script:BWS_ErrorLog.Add($record)

    # Console output (unless suppressed)
    if (-not $SuppressConsole) {
        $icon  = switch ($Severity) { "Error"{"[X]"} "Warning"{"[!]"} default{"[i]"} }
        $color = switch ($Severity) { "Error"{"Red"} "Warning"{"Yellow"} default{"Cyan"} }
        Write-Host "    $icon [$Code] $Message" -ForegroundColor $color
        if ($Detail -and -not $script:CompactViewGlobal) {
            Write-Host "        Detail : $Detail" -ForegroundColor DarkGray
        }
        if ($CheckStep) {
            Write-Host "        Step   : $CheckStep" -ForegroundColor DarkGray
        }
    }

    # Return the record so callers can do: $status.Errors += Write-BWsError ...
    return $record
}


function Get-BWsSupportBundle {
    <#
    .SYNOPSIS
        Generates a structured support bundle (JSON) from the current script run.
        Contains: error log, environment info, script version, run timestamp.
        Attach this file to any BWS support ticket.
    .PARAMETER BCID
        Business Continuity ID for the filename.
    .PARAMETER OutputPath
        Directory to write the JSON to. Defaults to current directory.
    .PARAMETER PassThru
        Return the bundle hashtable instead of just the file path.
    #>
    param(
        [string]$BCID       = "0000",
        [string]$OutputPath = "",
        [switch]$PassThru
    )

    Write-Host ""
    Write-Host "=====================================================" -ForegroundColor Cyan
    Write-Host "  BWS SUPPORT BUNDLE" -ForegroundColor Cyan
    Write-Host "=====================================================" -ForegroundColor Cyan

    # Environment via existing Get-BWsDiagnostics
    $envInfo = try { Get-BWsDiagnostics -AsObject } catch { @{} }

    $errorsByCategory = @{}
    foreach ($e in $script:BWS_ErrorLog) {
        if (-not $errorsByCategory.ContainsKey($e.Category)) {
            $errorsByCategory[$e.Category] = @()
        }
        $errorsByCategory[$e.Category] += $e
    }

    $topErrors = $script:BWS_ErrorLog |
        Where-Object { $_.Severity -eq "Error" } |
        Group-Object Code |
        Sort-Object Count -Descending |
        Select-Object -First 10 |
        ForEach-Object { @{ Code=$_.Name; Count=$_.Count; Message=($_.Group[0].Message) } }

    $bundle = [ordered]@{
        BundleVersion      = "1.0"
        ScriptVersion      = $script:Version
        GeneratedAt        = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        BCID               = $BCID
        GlobalAdminAccount = if ($script:GlobalAdminUPN) { $script:GlobalAdminUPN } else { "N/A" }
        TenantId           = if ($script:GlobalAdminTenantId) { $script:GlobalAdminTenantId } else { "N/A" }
        Environment        = $envInfo
        Summary            = @{
            TotalEntries   = $script:BWS_ErrorLog.Count
            ErrorCount     = ($script:BWS_ErrorLog | Where-Object { $_.Severity -eq "Error" }).Count
            WarningCount   = ($script:BWS_ErrorLog | Where-Object { $_.Severity -eq "Warning" }).Count
            InfoCount      = ($script:BWS_ErrorLog | Where-Object { $_.Severity -eq "Info" }).Count
            CategoriesHit  = ($script:BWS_ErrorLog | Select-Object -ExpandProperty Category -Unique)
        }
        TopErrors          = $topErrors
        ErrorsByCategory   = $errorsByCategory
        FullLog            = @($script:BWS_ErrorLog)
    }

    # Write JSON
    $ts       = Get-Date -Format "yyyyMMdd_HHmmss"
    $fileName = "BWS_Support_${BCID}_${ts}.json"
    $path     = if ($OutputPath) { Join-Path $OutputPath $fileName } else { Join-Path (Get-Location).Path $fileName }
    $bundle | ConvertTo-Json -Depth 15 | Out-File -FilePath $path -Encoding UTF8

    # Console summary
    $errCount  = $bundle.Summary.ErrorCount
    $warnCount = $bundle.Summary.WarningCount
    Write-Host "  File     : $path" -ForegroundColor Green
    Write-Host "  Errors   : $errCount"   -ForegroundColor $(if ($errCount  -eq 0){"Green"}else{"Red"})
    Write-Host "  Warnings : $warnCount"  -ForegroundColor $(if ($warnCount -eq 0){"Green"}else{"Yellow"})
    if ($topErrors) {
        Write-Host ""
        Write-Host "  Top error codes:" -ForegroundColor Cyan
        foreach ($e in $topErrors) {
            Write-Host "    $($e.Code.PadRight(22)) $($e.Count)x  $($e.Message.Substring(0,[Math]::Min(45,$e.Message.Length)))" -ForegroundColor White
        }
    }
    Write-Host "=====================================================" -ForegroundColor Cyan
    Write-Host "  Attach this file to your BWS support ticket." -ForegroundColor Gray
    Write-Host "=====================================================" -ForegroundColor Cyan
    Write-Host ""

    if ($PassThru) { return $bundle } else { return $path }
}


function Show-BWsErrorSummary {
    <#
    .SYNOPSIS
        Prints a compact error table to the console at the end of a run.
        Called automatically if there are any errors in the log.
    #>
    param([switch]$OnlyErrors)

    $items = if ($OnlyErrors) {
        @($script:BWS_ErrorLog | Where-Object { $_.Severity -eq "Error" })
    } else {
        @($script:BWS_ErrorLog | Where-Object { $_.Severity -in @("Error","Warning") })
    }

    if ($items.Count -eq 0) { return }

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  ERROR LOG SUMMARY  ($($items.Count) issue(s))" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan

    $errCount  = ($items | Where-Object Severity -eq "Error").Count
    $warnCount = ($items | Where-Object Severity -eq "Warning").Count
    Write-Host "  Errors:   $errCount" -ForegroundColor $(if($errCount  -eq 0){"Green"}else{"Red"})
    Write-Host "  Warnings: $warnCount" -ForegroundColor $(if($warnCount -eq 0){"Green"}else{"Yellow"})
    Write-Host ""

    $grouped = $items | Group-Object Category | Sort-Object Count -Descending
    foreach ($grp in $grouped) {
        Write-Host "  [$($grp.Name)] ($($grp.Count) issue(s)):" -ForegroundColor Yellow
        foreach ($e in $grp.Group | Sort-Object Severity) {
            $icon  = if ($e.Severity -eq "Error") { "[X]" } else { "[!]" }
            $col   = if ($e.Severity -eq "Error") { "Red" } else { "Yellow" }
            $short = $e.Message.Substring(0, [Math]::Min(55, $e.Message.Length))
            Write-Host "    $icon [$($e.Code)] $short" -ForegroundColor $col
        }
        Write-Host ""
    }
    Write-Host "  Run Get-BWsSupportBundle for a full support export." -ForegroundColor Gray
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Connect-BWsGlobalAdmin {
    <#
    .SYNOPSIS
        Performs the single interactive GlobalAdmin login for the script.
        ONE login covers Az (Graph/ARM), Teams, and SharePoint.
    .NOTES
        Account model:
          GlobalAdmin  - All M365 + Azure rights. Used for:
                           * Connect-AzAccount  (Az.Accounts)
                           * Connect-MicrosoftTeams -AccountId -TenantId
                           * Connect-SPOService -Credential (PS 5.1 only)
          DomainAdmin  - Local AD only (on-prem queries, not online)

        Microsoft Learn references:
          Connect-AzAccount: learn.microsoft.com/powershell/module/az.accounts/connect-azaccount
          Connect-MicrosoftTeams -AccountId: learn.microsoft.com/powershell/module/microsoftteams/connect-microsoftteams
          Connect-SPOService -Credential: learn.microsoft.com/powershell/module/sharepoint-online/connect-sposervice

        PS 5.1 requirement:
          Microsoft.Online.SharePoint.PowerShell is DESKTOP (PS 5.1) only.
          The script enforces or warns accordingly.

    .PARAMETER SharePointAdminUrl
        SharePoint admin centre URL, e.g. https://contoso-admin.sharepoint.com
        Derived automatically from tenant domain if not provided.
    .PARAMETER ForceRelogin
        Force a new Connect-AzAccount even if a session already exists.
    .OUTPUTS
        Returns $true on success, $false on failure.
    #>
    param(
        [string]$SharePointAdminUrl = "",
        [switch]$ForceRelogin,
        [switch]$SkipTeams,
        [switch]$SkipSharePoint
    )

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  ACCOUNT LOGIN" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Account 1: GlobalAdmin (M365 + Azure)" -ForegroundColor White
    Write-Host "  Account 2: DomainAdmin (on-premises AD - not required" -ForegroundColor Gray
    Write-Host "             for online checks in this script)" -ForegroundColor Gray
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    # -- Step 1: Az / Graph login (GlobalAdmin) -------------------------
    Write-Host "  [1/3] Azure + Graph login (GlobalAdmin)" -ForegroundColor Yellow
    Write-Host "        A browser window will open for interactive sign-in." -ForegroundColor Gray
    Write-Host ""

    try {
        $azCtx = $null

        if (-not $ForceRelogin) {
            $azCtx = Get-AzContext -ErrorAction SilentlyContinue
            if ($azCtx -and $azCtx.Account) {
                Write-Host "    [OK] Already logged in as: $($azCtx.Account.Id)" -ForegroundColor Green
                Write-Host "         Tenant : $($azCtx.Tenant.Id)" -ForegroundColor Gray
                Write-Host "         Sub    : $($azCtx.Subscription.Name)" -ForegroundColor Gray
            } else {
                $azCtx = $null
            }
        }

        if (-not $azCtx) {
            Write-Host "    Launching browser login for GlobalAdmin..." -ForegroundColor Cyan
            # PS 5.1 + MFA: Connect-AzAccount opens a browser window
            # -SkipContextPopulation avoids loading all 25 subscriptions (faster)
            $azCtx = Connect-AzAccount -ErrorAction Stop
            $azCtx = Get-AzContext -ErrorAction Stop
            Write-Host "    [OK] Logged in as: $($azCtx.Account.Id)" -ForegroundColor Green
            Write-Host "         Tenant : $($azCtx.Tenant.Id)" -ForegroundColor Gray
            Write-Host "         Sub    : $($azCtx.Subscription.Name)" -ForegroundColor Gray
        }

        # Store for reuse by Teams + SharePoint
        $script:GlobalAdminUPN      = $azCtx.Account.Id
        $script:GlobalAdminTenantId = $azCtx.Tenant.Id
        $script:GlobalAdminConnected = $true

        # Verify Graph access
        try {
            $probe = Invoke-AzRestMethod -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id' -Method GET -ErrorAction Stop
            if ($probe.StatusCode -ge 200 -and $probe.StatusCode -lt 300) {
                Write-Host "    [OK] Graph API reachable via Az session" -ForegroundColor Green
                $script:BWSGraphConnected = $true
            } else {
                Write-Host "    [!]  Graph probe returned HTTP $($probe.StatusCode)" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "    [!]  Graph probe failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }

    } catch {
        Write-Host "    [X]  Azure login failed: $($_.Exception.Message)" -ForegroundColor Red
        $script:GlobalAdminConnected = $false
        return $false
    }

    # -- Step 2: Teams login (reuse Az session) ------------------------
    if (-not $SkipTeams) {
        Write-Host ""
        Write-Host "  [2/3] Microsoft Teams (reusing GlobalAdmin session)" -ForegroundColor Yellow
        try {
            # Check if already connected
            $teamsAlreadyOk = $false
            try {
                $null = Get-CsTeamsClientConfiguration -ErrorAction Stop
                $teamsAlreadyOk = $true
                Write-Host "    [OK] Teams session already active" -ForegroundColor Green
            } catch {}

            if (-not $teamsAlreadyOk) {
                # Connect-MicrosoftTeams with -AccountId and -TenantId reuses the
                # existing AAD token cache - no new browser window needed.
                # MS Learn: learn.microsoft.com/powershell/module/microsoftteams/connect-microsoftteams
                Write-Host "    Connecting to Microsoft Teams (-AccountId $($script:GlobalAdminUPN))..." -ForegroundColor Cyan
                $null = Connect-MicrosoftTeams `
                    -AccountId $script:GlobalAdminUPN `
                    -TenantId  $script:GlobalAdminTenantId `
                    -ErrorAction Stop 2>$null 3>$null
                Write-Host "    [OK] Teams connected (no additional login required)" -ForegroundColor Green
            }
            $script:TeamsConnected = $true
        } catch {
            Write-Host "    [!]  Teams connection failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "         Teams checks will prompt for login when needed." -ForegroundColor Gray
            # Non-fatal: Teams check will fall through to its own connection
        }
    } else {
        Write-Host "  [2/3] Teams: skipped (-SkipTeams)" -ForegroundColor Gray
    }

    # -- Step 3: SharePoint (PS 5.1 Desktop only) ---------------------
    if (-not $SkipSharePoint) {
        Write-Host ""
        Write-Host "  [3/3] SharePoint Online (PS 5.1 Desktop only)" -ForegroundColor Yellow

        # SharePoint Management Shell REQUIRES Windows PowerShell 5.1
        # It does NOT work in PowerShell 7 / Core.
        # MS Learn: learn.microsoft.com/powershell/sharepoint/sharepoint-online/connect-sharepoint-online
        if ($PSVersionTable.PSEdition -ne "Desktop") {
            Write-Host "    [!]  SharePoint module requires Windows PowerShell 5.1 (Desktop edition)." -ForegroundColor Yellow
            Write-Host "         Current: PS $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))" -ForegroundColor Gray
            Write-Host "         SharePoint checks will be SKIPPED." -ForegroundColor Yellow
            $script:SkipSharePoint = $true
        } else {
            # Derive SharePoint admin URL if not provided
            $spoUrl = $SharePointAdminUrl
            if ([string]::IsNullOrEmpty($spoUrl) -and $script:GlobalAdminUPN) {
                # Derive from UPN: user@contoso.onmicrosoft.com -> https://contoso-admin.sharepoint.com
                $upnDomain = ($script:GlobalAdminUPN -split '@' | Select-Object -Last 1) -replace '\.onmicrosoft\.com$',''
                if ($upnDomain) {
                    $spoUrl = "https://$upnDomain-admin.sharepoint.com"
                    Write-Host "    [i]  Auto-derived SharePoint URL: $spoUrl" -ForegroundColor Gray
                    Write-Host "         Use -SharePointUrl to override if incorrect." -ForegroundColor Gray
                }
            }

            if (-not [string]::IsNullOrEmpty($spoUrl)) {
                try {
                    # Try if already connected
                    $spoAlreadyOk = $false
                    try {
                        $null = Get-SPOTenant -ErrorAction Stop
                        $spoAlreadyOk = $true
                        Write-Host "    [OK] SharePoint session already active" -ForegroundColor Green
                    } catch {}

                    if (-not $spoAlreadyOk) {
                        Write-Host "    Connecting to SharePoint Online..." -ForegroundColor Cyan
                        Write-Host "    URL: $spoUrl" -ForegroundColor Gray
                        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop 2>$null 3>$null

                        # Use ModernAuth for MFA accounts (recommended by MS Learn)
                        # If -Credential is passed it uses legacy auth (no MFA support)
                        # Without -Credential it opens a browser window (MFA-compatible)
                        try {
                            Connect-SPOService -Url $spoUrl `
                                -ModernAuth $true `
                                -AuthenticationUrl "https://login.microsoftonline.com/organizations" `
                                -ErrorAction Stop
                            Write-Host "    [OK] SharePoint connected (ModernAuth)" -ForegroundColor Green
                        } catch {
                            # Fallback: plain connect (triggers browser if needed)
                            Write-Host "    [!]  ModernAuth failed, trying standard connect..." -ForegroundColor Yellow
                            Connect-SPOService -Url $spoUrl -ErrorAction Stop
                            Write-Host "    [OK] SharePoint connected" -ForegroundColor Green
                        }
                    }
                    $script:SharePointConnected = $true
                } catch {
                    Write-Host "    [!]  SharePoint connection failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    Write-Host "         Use -SharePointUrl to specify the correct admin URL." -ForegroundColor Gray
                    Write-Host "         SharePoint checks will prompt for login when needed." -ForegroundColor Gray
                }
            } else {
                Write-Host "    [!]  No SharePoint admin URL available." -ForegroundColor Yellow
                Write-Host "         Use -SharePointUrl https://<tenant>-admin.sharepoint.com" -ForegroundColor Gray
            }
        }
    } else {
        Write-Host "  [3/3] SharePoint: skipped (-SkipSharePoint)" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  LOGIN SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  GlobalAdmin     : $(if ($script:GlobalAdminConnected) { $script:GlobalAdminUPN + ' ([OK])' } else { 'NOT LOGGED IN ([X])' })" `
        -ForegroundColor $(if ($script:GlobalAdminConnected) { "Green" } else { "Red" })
    Write-Host "  Tenant          : $(if ($script:GlobalAdminTenantId) { $script:GlobalAdminTenantId } else { 'Unknown' })" -ForegroundColor Gray
    Write-Host "  Az/Graph        : $(if ($script:GlobalAdminConnected) { '[OK]' } else { '[X]' })" `
        -ForegroundColor $(if ($script:GlobalAdminConnected) { "Green" } else { "Red" })
    Write-Host "  Teams           : $(if ($script:TeamsConnected) { '[OK] Connected' } elseif ($SkipTeams) { 'Skipped' } else { '[!] Not connected (will prompt)' })" `
        -ForegroundColor $(if ($script:TeamsConnected) { "Green" } elseif ($SkipTeams) { "Gray" } else { "Yellow" })
    Write-Host "  SharePoint      : $(if ($script:SharePointConnected) { '[OK] Connected' } elseif ($SkipSharePoint -or $script:SkipSharePoint) { 'Skipped' } else { '[!] Not connected (will prompt)' })" `
        -ForegroundColor $(if ($script:SharePointConnected) { "Green" } elseif ($SkipSharePoint -or $script:SkipSharePoint) { "Gray" } else { "Yellow" })
    Write-Host "  DomainAdmin     : Not required for online checks" -ForegroundColor Gray
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    return $script:GlobalAdminConnected
}

function Connect-BWsGraph {
    <#
    .SYNOPSIS
        Validates the active Az session can reach Graph via Invoke-AzRestMethod.
    .NOTES
        Authentication approach confirmed by Microsoft Learn:
        - Invoke-AzRestMethod (Az.Accounts) handles tokens automatically
          from the active Connect-AzAccount session. No Graph SDK needed.
        - Official example: Invoke-AzRestMethod https://graph.microsoft.com/v1.0/me
        - Source: learn.microsoft.com/powershell/module/az.accounts/invoke-azrestmethod
        - The -Scopes parameter is a no-op here: permissions are determined
          by app registration consent, not by this function.
    #>
    param([string[]]$Scopes = @())   # kept for call-site compatibility; no-op

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
        $probe = Invoke-AzRestMethod -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id' -Method GET -ErrorAction Stop
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
        Returns a flat array of all items. Max 200 pages (safety guard).
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [string]$Method  = "GET",
        [int]   $MaxPages = 200
    )
    if ($Uri -notmatch '^https://') {
        $Uri = "https://graph.microsoft.com/v1.0/$($Uri.TrimStart('/'))"
    }
    $allItems = [System.Collections.Generic.List[object]]::new()
    $nextUri  = $Uri
    $page     = 0
    do {
        $page++
        if ($page -gt $MaxPages) {
            Write-Host "  [!] Paged request exceeded $MaxPages pages, stopping early." -ForegroundColor Yellow
            break
        }
        $resp = Invoke-BWsGraphRequest -Uri $nextUri -Method $Method -ErrorAction Stop
        if ($null -eq $resp) { break }
        if ($resp.PSObject.Properties['value'] -and $resp.value) {
            foreach ($item in $resp.value) { $allItems.Add($item) }
        } elseif (-not $resp.PSObject.Properties['value']) {
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


#============================================================================
# Azure REST Helpers - replaces Az.Resources / Az.Accounts PS 5.1 compat
#============================================================================

function Get-BWsAzContext {
    <#
    .SYNOPSIS
        Thin wrapper around Get-AzContext (Az.Accounts only).
        Keeps Az.Resources unloaded to avoid .NET 5+ dependency errors on PS 5.1.
        Confirmed: Get-AzContext is in Az.Accounts module (not Az.Resources).
        Microsoft Learn: learn.microsoft.com/powershell/module/az.accounts/get-azcontext
    #>
    param([string]$ErrorAction = "SilentlyContinue")
    return (Get-AzContext -ErrorAction $ErrorAction)
}
function Get-BWsDiagnostics {
    <#
    .SYNOPSIS Collects environment diagnostics. Used by -Diagnostics and GUI Diagnostics button.#>
    param([switch]$AsObject)

    $d = [ordered]@{}
    $d['PS_Version']  = $PSVersionTable.PSVersion.ToString()
    $d['PS_Edition']  = $PSVersionTable.PSEdition
    $d['PS_Host']     = $Host.Name
    $d['OS']          = if ($PSVersionTable.OS) { $PSVersionTable.OS } else {
                            "$([System.Environment]::OSVersion.Version)" }

    $azCtx = $null
    try { $azCtx = Get-AzContext -ErrorAction SilentlyContinue } catch {}

    if ($azCtx) {
        $d['AZ_LoggedIn']         = "Yes"
        $d['AZ_Account']          = if ($azCtx.Account)      { $azCtx.Account.Id }      else { "(unknown)" }
        $d['AZ_AccountType']      = if ($azCtx.Account)      { $azCtx.Account.Type }    else { "(unknown)" }
        $d['AZ_TenantId']         = if ($azCtx.Tenant)       { $azCtx.Tenant.Id }       else { "(unknown)" }
        $d['AZ_SubscriptionId']   = if ($azCtx.Subscription) { $azCtx.Subscription.Id }   else { "(none)" }
        $d['AZ_SubscriptionName'] = if ($azCtx.Subscription) { $azCtx.Subscription.Name } else { "(none)" }
        $d['AZ_Environment']      = if ($azCtx.Environment)  { $azCtx.Environment.Name }  else { "(unknown)" }

        try {
            $orgResp = Invoke-AzRestMethod -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName,verifiedDomains' -Method GET -ErrorAction Stop
            if ($orgResp.StatusCode -eq 200) {
                $org = ($orgResp.Content | ConvertFrom-Json).value | Select-Object -First 1
                $d['AZ_TenantName']   = if ($org.displayName) { $org.displayName } else { "(unknown)" }
                $primary = ($org.verifiedDomains | Where-Object { $_.isDefault -eq $true } | Select-Object -First 1).name
                if ($primary) { $d['AZ_PrimaryDomain'] = $primary }
            }
        } catch { $d['AZ_TenantName'] = "(Graph REST not available)" }

        try {
            $meResp = Invoke-AzRestMethod -Uri 'https://graph.microsoft.com/v1.0/me?$select=displayName,userPrincipalName,jobTitle,mail' -Method GET -ErrorAction Stop
            if ($meResp.StatusCode -eq 200) {
                $me = $meResp.Content | ConvertFrom-Json
                $d['USER_DisplayName'] = if ($me.displayName)      { $me.displayName }       else { "(unknown)" }
                $d['USER_UPN']         = if ($me.userPrincipalName) { $me.userPrincipalName } else { "(unknown)" }
                $d['USER_JobTitle']    = if ($me.jobTitle)          { $me.jobTitle }          else { "(not set)" }
                $d['USER_Mail']        = if ($me.mail)              { $me.mail }              else { "(not set)" }
            }
        } catch { $d['USER_DisplayName'] = "(Graph /me not available)" }
    } else {
        $d['AZ_LoggedIn'] = "No - run Connect-AzAccount"
    }

    foreach ($m in @('Az.Accounts','Az.Storage','Microsoft.Online.SharePoint.PowerShell','MicrosoftTeams','PSScriptAnalyzer')) {
        $loaded = Get-Module -Name $m -ErrorAction SilentlyContinue
        $avail  = Get-Module -Name $m -ListAvailable -ErrorAction SilentlyContinue | Select-Object -First 1
        $d["MOD_$m"] = if ($loaded) { "Loaded v$($loaded.Version)" }
                       elseif ($avail) { "Available v$($avail.Version) (not imported)" }
                       else { "Not installed" }
    }

    if ($AsObject) { return $d }

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  DIAGNOSTICS" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan

    $sections = [ordered]@{
        "PowerShell"       = @('PS_Version','PS_Edition','PS_Host','OS')
        "Azure / Entra ID" = @('AZ_LoggedIn','AZ_Account','AZ_AccountType','AZ_TenantId','AZ_TenantName','AZ_PrimaryDomain','AZ_SubscriptionId','AZ_SubscriptionName','AZ_Environment')
        "Signed-in User"   = @('USER_DisplayName','USER_UPN','USER_JobTitle','USER_Mail')
        "Modules"          = ($d.Keys | Where-Object { $_ -like 'MOD_*' })
    }
    foreach ($sec in $sections.Keys) {
        Write-Host ""
        Write-Host "  [$sec]" -ForegroundColor Yellow
        foreach ($key in $sections[$sec]) {
            if ($d.Contains($key)) {
                $lbl = ($key -replace '^(PS_|AZ_|USER_|MOD_)','').PadRight(22)
                $val = $d[$key]
                $col = if ($val -like "*No -*" -or $val -like "*not installed*") { "Yellow" }
                       elseif ($val -like "*Loaded*" -or $val -like "*Yes*") { "Green" }
                       else { "White" }
                Write-Host "    $lbl : $val" -ForegroundColor $col
            }
        }
    }
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
}


function Set-BWsAzSubscription {
    <#
    .SYNOPSIS
        Sets the subscription context using Invoke-AzRestMethod probe
        instead of Set-AzContext (which loads Az.Resources ResourceManagementClient).
    .PARAMETER SubscriptionId
        The Azure subscription GUID to switch to.
    #>
    param([Parameter(Mandatory=$true)][string]$SubscriptionId)

    # Set-AzContext from Az.Accounts alone works fine - it's Az.Resources that
    # triggers get_SerializationSettings. We load Az.Accounts only, so Set-AzContext
    # is safe as long as Az.Resources is NOT loaded.
    # Unload Az.Resources if it somehow ended up in session
    Get-Module -Name 'Az.Resources' -ErrorAction SilentlyContinue |
        Remove-Module -Force -ErrorAction SilentlyContinue 2>$null 3>$null

    try {
        # Az.Accounts' Set-AzContext is safe without Az.Resources loaded
        $null = Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop 2>$null 3>$null
        return $true
    } catch {
        throw "Could not switch to subscription $SubscriptionId`: $($_.Exception.Message)"
    }
}

function Find-BWsAzResource {
    <#
    .SYNOPSIS
        Finds an Azure resource by name and type using the ARM REST API.
        Replaces Get-AzResource (Az.Resources) which is not PS 5.1 compatible.
    .PARAMETER Name
        Resource display name
    .PARAMETER ResourceType
        ARM resource type, e.g. "Microsoft.Storage/storageAccounts"
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][string]$ResourceType
    )

    try {
        $ctx  = Get-BWsAzContext -ErrorAction Stop
        $subId = $ctx.Subscription.Id
        # ARM REST: list all resources of given type, filter by name client-side
        $uri = "https://management.azure.com/subscriptions/$subId/resources" +
               "?`$filter=resourceType eq '$ResourceType'&`$top=1000&api-version=2021-04-01"
        $resp = Invoke-AzRestMethod -Uri $uri -Method GET -ErrorAction Stop
        if ($resp.StatusCode -ne 200) { return $null }
        $data = $resp.Content | ConvertFrom-Json
        $match = $data.value | Where-Object { $_.name -eq $Name } | Select-Object -First 1
        if (-not $match) { return $null }
        # Return object mimicking Get-AzResource output
        return [PSCustomObject]@{
            Name              = $match.name
            ResourceType      = $match.type
            Location          = $match.location
            ResourceGroupName = ($match.id -split '/')[4]
            ResourceId        = $match.id
            Tags              = $match.tags
        }
    } catch {
        return $null
    }
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
        elseif ($s -like "*INFO*")   { $c = "Cyan"     }
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

    # ----------------------------------------------------------------
    # Microsoft Learn prerequisites (install-azps-windows) for PS 5.1:
    #  - TLS 1.2 required (PSGallery, .NET 4.x defaults to TLS 1.0/1.1)
    #  - AzureRM must NOT coexist with Az on PS 5.1 (causes TypeLoadException)
    #  - .NET Framework 4.7.2+ required for Az module
    # ----------------------------------------------------------------

    # TLS 1.2 - mandatory for PSGallery on PS 5.1
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Host "  [OK] TLS 1.2 enforced (required for PSGallery on PS 5.1)" -ForegroundColor Gray
    } catch {
        Write-Host "  [!] TLS 1.2 could not be set: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # AzureRM conflict check - Az + AzureRM incompatible on PS 5.1 (Microsoft Learn)
    $azureRMPresent = Get-Module -Name AzureRM -ListAvailable -ErrorAction SilentlyContinue |
                      Select-Object -First 1
    if ($azureRMPresent) {
        Write-Host ""
        Write-Host "  [!] AzureRM v$($azureRMPresent.Version) CONFLICT - Az + AzureRM incompatible on PS 5.1" -ForegroundColor Red
        Write-Host "      MS Learn: Uninstall-Module AzureRM -AllVersions -Force" -ForegroundColor Yellow
        Get-Module -Name 'AzureRM*' -ErrorAction SilentlyContinue | Remove-Module -Force -ErrorAction SilentlyContinue
        Write-Host "  [i] AzureRM unloaded from session (disk copy unchanged)" -ForegroundColor Gray
        Write-Host ""
    }

    # .NET Framework 4.7.2+ check - required for Az module on PS 5.1
    if ($PSVersionTable.PSEdition -eq "Desktop") {
        try {
            $ndpKey = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
            $ndpRelease = (Get-ItemProperty -Path $ndpKey -ErrorAction Stop).Release
            $ndpVersion = (Get-ItemProperty -Path $ndpKey -ErrorAction Stop).Version
            if ($ndpRelease -ge 461808) {  # 461808 = 4.7.2, 528040 = 4.8
                Write-Host "  [OK] .NET Framework $ndpVersion (>= 4.7.2 required)" -ForegroundColor Gray
            } else {
                Write-Host "  [!] .NET Framework $ndpVersion < 4.7.2 required for Az" -ForegroundColor Yellow
                Write-Host "      https://dotnet.microsoft.com/download/dotnet-framework" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  [i] .NET Framework version check skipped" -ForegroundColor Gray
        }
    }

    # Ensure NuGet provider is available (required in PS 5.1 for Install-Module)
    # Microsoft Learn: learn.microsoft.com/powershell/azure/install-azps-windows
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

    # Update PowerShellGet if running PS 5.1 with old version (< 2.x)
    # Old PS 5.1 ships with PowerShellGet 1.x which has known Install-Module issues
    if ($PSVersionTable.PSEdition -eq "Desktop") {
        $psGet = Get-Module -Name PowerShellGet -ListAvailable -ErrorAction SilentlyContinue |
                 Sort-Object Version -Descending | Select-Object -First 1
        # Microsoft Learn: PowerShellGet >= 2.2.3 required for reliable module installation
        if ($psGet -and $psGet.Version -lt [version]"2.2.3") {
            Write-Host "  Updating PowerShellGet (current: $($psGet.Version), recommended: 2.2.3+)..." -ForegroundColor Yellow
            try {
                Install-Module -Name PowerShellGet -MinimumVersion 2.2.3 -Force `
                    -Scope CurrentUser -AllowClobber -Repository PSGallery -ErrorAction Stop | Out-Null
                Write-Host "  [OK] PowerShellGet updated to 2.2.3+ - new session recommended" -ForegroundColor Green
            } catch {
                Write-Host "  [i] PowerShellGet update skipped: $($_.Exception.Message)" -ForegroundColor Gray
            }
        } elseif ($psGet) {
            Write-Host "  [OK] PowerShellGet $($psGet.Version) (>= 2.2.3)" -ForegroundColor Gray
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
        # "AlwaysSkip" = show installed version info but never try to import
        if ($mod['SkipParam'] -eq 'AlwaysSkip') {
            $doSkip = $true
        } elseif ($mod['SkipParam'] -and $SkipParams.ContainsKey($mod['SkipParam']) -and $SkipParams[$mod['SkipParam']] -eq $true) {
            $doSkip = $true
        }

        Write-Progress -Activity "BWS Module Setup" `
            -Status "[$modCount/$($script:RequiredModules.Count)] $($mod.Name)" `
            -PercentComplete ([int](($modCount-1) / $script:RequiredModules.Count * 100))

        if ($doSkip) {
            # For AlwaysSkip modules: show the installed version number if present,
            # otherwise show "not installed" - so admins can see what is on the system
            if ($mod['SkipParam'] -eq 'AlwaysSkip') {
                $stSkip = Get-ModuleStatus -Name $mod.Name -MinVersion $mod.MinVersion
                $row.Status      = if ($stSkip.InstalledVer) { "INFO v$($stSkip.InstalledVer)" } else { "INFO not installed" }
                $row.InstallTime = "n/a"
                $row.ImportTime  = "n/a (REST)"
            } else {
                $row.Status = "SKIP"; $row.InstallTime = "n/a"; $row.ImportTime = "n/a"
            }
            $row.Skipped = $true
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
        elseif ($s -like "*INFO*")    { $c = [System.Drawing.Color]::FromArgb( 80,160,200) }
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
# PREREQUISITES + STARTUP FLOW
# -Full  : install/import modules, verify login, then run checks
# (none) : skip all setup - run checks only (modules must already be loaded)
# -GUI   : handled inside Show-BWSGUI; "Prerequisites" button = Full flow,
#          "Run Check" button = checks only
#============================================================================

$_skipParams = @{
    SkipSharePoint = [bool]$SkipSharePoint
    SkipTeams      = [bool]$SkipTeams
    SkipDefender   = [bool]$SkipDefender
}

# -Diagnostics: print diagnostics and exit (console) or open dialog (GUI)
if ($Diagnostics -and -not $GUI) {
    Get-BWsDiagnostics
    exit 0
}

if ($Full -and $GUI) {
    # GUI + Full: open Prerequisites dialog, then continue to launch the main form
    Add-Type -AssemblyName System.Windows.Forms
    $null = Show-ModuleSetupDialog -SkipParams $_skipParams
}

if ($Full -and -not $GUI) {
    # Console + Full: module setup, then GlobalAdmin login
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  PREREQUISITES - MODULE SETUP" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    $script:moduleSetupResult = Install-BWSDependencies -SkipParams $_skipParams
    if (-not $script:moduleSetupResult.AllReady) {
        Write-Host "  [!] One or more required modules could not be installed." -ForegroundColor Yellow
        Write-Host "      The script will continue but some checks may fail." -ForegroundColor Yellow
        Write-Host ""
    }

    # GlobalAdmin login - ONE login for Az + Graph + Teams + SharePoint
    $loginOK = Connect-BWsGlobalAdmin `
        -SharePointAdminUrl $SharePointUrl `
        -SkipTeams:$SkipTeams `
        -SkipSharePoint:$SkipSharePoint

    if (-not $loginOK) {
        Write-Host "[X] GlobalAdmin login failed. Cannot continue." -ForegroundColor Red
        exit 1
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
            $azResource = Find-BWsAzResource -Name $resource.Name -ResourceType $resource.Type
            
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
        CheckPerformed = $true
        Found    = $foundResources
        Missing  = $missingResources
        Errors   = $errorResources
        Total    = $azureResourcesToCheck.Count
    }
}

function Test-IntunePolicies {
    <#
    .SYNOPSIS
        Checks Intune for all 26 BWS Standard Policies.
        Queries four API endpoints (confirmed field names from Microsoft Learn):
          deviceConfigurations       -> displayName   (v1.0)
          deviceCompliancePolicies   -> displayName   (v1.0)
          configurationPolicies      -> name          (beta) *** NOT displayName ***
          compliancePolicies         -> name          (beta) *** new v2 compliance endpoint ***
        Returns per-policy boolean: Found=$true / Found=$false
    #>
    param(
        [bool]$ShowAllPolicies = $false,
        [bool]$CompactView     = $false
    )

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  INTUNE POLICY CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    # Result arrays: each entry is [ordered]@{ PolicyName; Found; PolicyId; Endpoint }
    $policyResults   = [System.Collections.Generic.List[object]]::new()
    $retrievalErrors = [System.Collections.Generic.List[string]]::new()

    try {
        # -- Graph connection ------------------------------------------------
        $graphContext = Get-BWsGraphContext -ErrorAction SilentlyContinue
        if (-not $graphContext) {
            try {
                Connect-BWsGraph -Scopes "DeviceManagementConfiguration.Read.All",
                                         "DeviceManagementManagedDevices.Read.All" -ErrorAction Stop
            } catch {
                Write-Host "  [X] Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
                foreach ($p in $script:intuneStandardPolicies) {
                    $policyResults.Add([PSCustomObject]@{
                        PolicyName = $p; Found = $false; PolicyId = $null; Endpoint = "N/A"
                    })
                }
                return @{
                    Found          = @()
                    Missing        = $script:intuneStandardPolicies
                    PolicyResults  = $policyResults.ToArray()
                    Errors         = @("Graph connection failed: $($_.Exception.Message)")
                    Total          = $script:intuneStandardPolicies.Count
                    FoundCount     = 0
                    MissingCount   = $script:intuneStandardPolicies.Count
                    AllFound       = $false
                    CheckPerformed = $false
                }
            }
        }

        # -- Collect all policies from all four endpoints -------------------
        # Unified list: each entry = @{ PolicyName (the name to match); Id; SourceEndpoint }
        $allPolicies = [System.Collections.Generic.List[object]]::new()

        # 1. deviceConfigurations (v1.0)  -> field: displayName
        Write-Host "  [1/4] deviceConfigurations (v1.0, displayName)..." -NoNewline -ForegroundColor Gray
        try {
            $dc = Invoke-BWsGraphPagedRequest -Uri 'deviceManagement/deviceConfigurations?$top=999&$select=id,displayName' -ErrorAction Stop
            if ($dc) {
                foreach ($p in $dc) {
                    if ($p.displayName) {
                        $allPolicies.Add(@{ PolicyName = $p.displayName; Id = $p.id; SourceEndpoint = "deviceConfigurations" })
                    }
                }
            }
            Write-Host " $($dc.Count) policies" -ForegroundColor Gray
        } catch {
            $retrievalErrors.Add("deviceConfigurations: $($_.Exception.Message)")
            Write-Host " [!] $($_.Exception.Message)" -ForegroundColor Yellow
        }

        # 2. deviceCompliancePolicies (v1.0)  -> field: displayName
        Write-Host "  [2/4] deviceCompliancePolicies (v1.0, displayName)..." -NoNewline -ForegroundColor Gray
        try {
            $dcp = Invoke-BWsGraphPagedRequest -Uri 'deviceManagement/deviceCompliancePolicies?$top=999&$select=id,displayName' -ErrorAction Stop
            if ($dcp) {
                foreach ($p in $dcp) {
                    if ($p.displayName) {
                        $allPolicies.Add(@{ PolicyName = $p.displayName; Id = $p.id; SourceEndpoint = "deviceCompliancePolicies" })
                    }
                }
            }
            Write-Host " $($dcp.Count) policies" -ForegroundColor Gray
        } catch {
            $retrievalErrors.Add("deviceCompliancePolicies: $($_.Exception.Message)")
            Write-Host " [!] $($_.Exception.Message)" -ForegroundColor Yellow
        }

        # 3. configurationPolicies / Settings Catalog (beta)  -> field: name (NOT displayName)
        Write-Host "  [3/4] configurationPolicies / Settings Catalog (beta, name)..." -NoNewline -ForegroundColor Gray
        try {
            $cp = Invoke-BWsBetaGraphRequest -Uri 'deviceManagement/configurationPolicies?$top=999&$select=id,name' -ErrorAction Stop
            $cpItems = if ($cp.PSObject.Properties['value']) { @($cp.value) } else { @() }
            foreach ($p in $cpItems) {
                if ($p.name) {
                    $allPolicies.Add(@{ PolicyName = $p.name; Id = $p.id; SourceEndpoint = "configurationPolicies" })
                }
            }
            Write-Host " $($cpItems.Count) policies" -ForegroundColor Gray
        } catch {
            $retrievalErrors.Add("configurationPolicies: $($_.Exception.Message)")
            Write-Host " [!] $($_.Exception.Message)" -ForegroundColor Yellow
        }

        # 4. compliancePolicies / v2 Compliance (beta)  -> field: name (NOT displayName)
        Write-Host "  [4/4] compliancePolicies / v2 Compliance (beta, name)..." -NoNewline -ForegroundColor Gray
        try {
            $v2cp = Invoke-BWsBetaGraphRequest -Uri 'deviceManagement/compliancePolicies?$top=999&$select=id,name' -ErrorAction Stop
            $v2Items = if ($v2cp.PSObject.Properties['value']) { @($v2cp.value) } else { @() }
            foreach ($p in $v2Items) {
                if ($p.name) {
                    $allPolicies.Add(@{ PolicyName = $p.name; Id = $p.id; SourceEndpoint = "compliancePolicies" })
                }
            }
            Write-Host " $($v2Items.Count) policies" -ForegroundColor Gray
        } catch {
            # compliancePolicies endpoint may not exist on all tenants - treat as info only
            Write-Host " [i] not available on this tenant" -ForegroundColor Gray
        }

        Write-Host ""
        Write-Host "  Total policies retrieved: $($allPolicies.Count)" -ForegroundColor Cyan

        if ($ShowAllPolicies) {
            Write-Host ""
            Write-Host "  [DEBUG] All retrieved policies:" -ForegroundColor Magenta
            $allPolicies | Sort-Object { $_['PolicyName'] } | ForEach-Object {
                Write-Host "    [$($_['SourceEndpoint'])] $($_['PolicyName'])" -ForegroundColor Gray
            }
        }

        Write-Host ""

        # -- Per-policy boolean matching ------------------------------------
        foreach ($required in $script:intuneStandardPolicies) {
            $normalizedRequired = Normalize-PolicyName $required
            $match = $allPolicies | Where-Object {
                (Normalize-PolicyName $_['PolicyName']) -eq $normalizedRequired
            } | Select-Object -First 1

            $found    = ($null -ne $match)   # Boolean: $true or $false
            $policyId = if ($found) { $match['Id'] } else { $null }
            $endpoint = if ($found) { $match['SourceEndpoint'] } else { $null }

            Write-Host "  " -NoNewline
            if ($found) {
                Write-Host "[FOUND]   " -NoNewline -ForegroundColor Green
            } else {
                Write-Host "[MISSING] " -NoNewline -ForegroundColor Red
            }
            Write-Host $required -ForegroundColor $(if ($found) { "White" } else { "Yellow" })
            if ($found -and -not $CompactView) {
                Write-Host "            -> $endpoint  (id: $policyId)" -ForegroundColor Gray
            }

            $policyResults.Add([PSCustomObject]@{
                PolicyName = $required
                Found      = [bool]$found
                PolicyId   = $policyId
                Endpoint   = $endpoint
            })
        }

    } catch {
        Write-Host "  [X] Unexpected error: $($_.Exception.Message)" -ForegroundColor Red
        $retrievalErrors.Add("Unexpected error: $($_.Exception.Message)")
    }

    # -- Derive summary from boolean results --------------------------------
    $foundPolicies   = @($policyResults | Where-Object { $_.Found -eq $true  })
    $missingPolicies = @($policyResults | Where-Object { $_.Found -eq $false })
    $allFound        = [bool]($missingPolicies.Count -eq 0)

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  INTUNE POLICIES SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Total required  : $($script:intuneStandardPolicies.Count)" -ForegroundColor White
    Write-Host "  Found           : $($foundPolicies.Count)" -ForegroundColor $(if ($foundPolicies.Count -eq $script:intuneStandardPolicies.Count) { "Green" } else { "Yellow" })
    Write-Host "  Missing         : $($missingPolicies.Count)" -ForegroundColor $(if ($missingPolicies.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Retrieval errors: $($retrievalErrors.Count)" -ForegroundColor $(if ($retrievalErrors.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "  All found       : $(if ($allFound) { 'Yes' } else { 'No' })" -ForegroundColor $(if ($allFound) { "Green" } else { "Red" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    if (-not $CompactView) {
        if ($missingPolicies.Count -gt 0) {
            Write-Host "MISSING POLICIES:" -ForegroundColor Red
            $missingPolicies | ForEach-Object { Write-Host "  - $($_.PolicyName)" -ForegroundColor Yellow }
            Write-Host ""
        }
        if ($retrievalErrors.Count -gt 0) {
            Write-Host "RETRIEVAL ERRORS:" -ForegroundColor Yellow
            $retrievalErrors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
            Write-Host ""
        }
    }

    return @{
        Found          = $foundPolicies
        Missing        = $missingPolicies
        PolicyResults  = $policyResults.ToArray()
        Errors         = $retrievalErrors.ToArray()
        Total          = $script:intuneStandardPolicies.Count
        FoundCount     = $foundPolicies.Count
        MissingCount   = $missingPolicies.Count
        AllFound       = $allFound
        CheckPerformed = $true
    }
}

function Test-EntraIDConnect {
    <#
    .SYNOPSIS
        Checks Entra ID Connect (Azure AD Connect) installation and sync status.
    .NOTES
        Microsoft Learn API references (v1.0):
        1. organization resource
           GET /v1.0/organization?$select=displayName,onPremisesSyncEnabled,
               onPremisesLastSyncDateTime,onPremisesLastPasswordSyncDateTime,verifiedDomains
           Fields: onPremisesSyncEnabled (bool|null), onPremisesLastSyncDateTime (DateTimeOffset),
                   onPremisesLastPasswordSyncDateTime (DateTimeOffset)
           Source: learn.microsoft.com/graph/api/resources/organization

        2. onPremisesDirectorySynchronization resource (v1.0)
           GET /v1.0/directory/onPremisesSynchronization
           features.passwordSyncEnabled         (bool)
           features.deviceWritebackEnabled       (bool)
           features.groupWriteBackEnabled        (bool)
           features.userWritebackEnabled         (bool)
           features.passwordWritebackEnabled     (bool)
           features.directoryExtensionsEnabled   (bool)
           features.synchronizeUpnForManagedUsersEnabled (bool)
           Source: learn.microsoft.com/graph/api/resources/onpremisesdirectorysynchronizationfeature

        Sync age thresholds (Microsoft recommended):
          <= 30 min  : healthy
          <= 3 h     : warning
          >  3 h     : error
    #>
    param([bool]$CompactView = $false)

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  ENTRA ID CONNECT CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    $status = @{
        # Sync state
        SyncEnabled              = $false    # bool: onPremisesSyncEnabled from org
        SyncActive               = $false    # bool: last sync within 3 hours
        LastSyncDateTime         = $null     # DateTimeOffset
        LastPasswordSyncDateTime = $null     # DateTimeOffset
        SyncAgeMinutes           = $null     # int
        SyncAgeStatus            = "Unknown" # "OK" / "Warning" / "Error" / "Unknown"

        # Features (from onPremisesDirectorySynchronization.features)
        PasswordSyncEnabled      = $null     # bool
        PasswordWritebackEnabled = $null     # bool
        DeviceWritebackEnabled   = $null     # bool
        GroupWritebackEnabled    = $null     # bool
        UserWritebackEnabled     = $null     # bool
        DirectoryExtensionsEnabled = $null   # bool

        # Tenant info
        TenantDisplayName        = $null
        VerifiedDomains          = @()
        OnPremDomains            = @()

        # Errors
        Errors                   = @()
        Warnings                 = @()
    }

    try {
        # -- Graph connection ------------------------------------------------
        $graphCtx = Get-BWsGraphContext -ErrorAction SilentlyContinue
        if (-not $graphCtx) {
            try {
                Connect-BWsGraph -Scopes "Directory.Read.All","Organization.Read.All" -ErrorAction Stop
            } catch {
                Write-Host "  [X] Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
                $status.Errors += (Write-BWsError -Code "BWS-AUTH-001" -Message "Graph connection failed" -Detail $_.Exception.Message -CheckStep "Graph connect" -SuppressConsole)
                return @{ Status = $status; CheckPerformed = $false }
            }
        }

        # ===================================================================
        # CHECK 1: Organization sync state
        # GET /v1.0/organization?$select=...
        # Fields: onPremisesSyncEnabled, onPremisesLastSyncDateTime,
        #         onPremisesLastPasswordSyncDateTime, displayName, verifiedDomains
        # ===================================================================
        Write-Host "  [1/2] Organization sync state (Graph v1.0/organization)" -ForegroundColor Yellow
        try {
            $orgUri = 'https://graph.microsoft.com/v1.0/organization?$select=displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesLastPasswordSyncDateTime,verifiedDomains'
            $orgResp = Invoke-BWsGraphRequest -Uri $orgUri -Method GET -ErrorAction Stop
            $org = if ($orgResp.PSObject.Properties['value']) { $orgResp.value | Select-Object -First 1 } else { $orgResp }

            if ($org) {
                $status.TenantDisplayName = $org.displayName

                # Verified / on-prem domains
                if ($org.verifiedDomains) {
                    $status.VerifiedDomains = @($org.verifiedDomains | Select-Object -ExpandProperty name)
                    $status.OnPremDomains   = @($org.verifiedDomains |
                                                Where-Object { $_.isVerified -and $_.name -notlike "*.onmicrosoft.com" } |
                                                Select-Object -ExpandProperty name)
                }

                # Sync enabled?
                # onPremisesSyncEnabled: true=syncing, false=was syncing/stopped, null=never synced
                $status.SyncEnabled = ($org.onPremisesSyncEnabled -eq $true)

                if ($status.SyncEnabled) {
                    Write-Host "    [OK] Sync enabled" -ForegroundColor Green
                    Write-Host "         Tenant : $($status.TenantDisplayName)" -ForegroundColor Gray

                    # Last sync time
                    $status.LastSyncDateTime = $org.onPremisesLastSyncDateTime
                    $status.LastPasswordSyncDateTime = $org.onPremisesLastPasswordSyncDateTime

                    if ($status.LastSyncDateTime) {
                        try {
                            $age = (Get-Date) - [DateTime]$status.LastSyncDateTime
                            $status.SyncAgeMinutes = [int]$age.TotalMinutes

                            if ($age.TotalMinutes -le 30) {
                                $status.SyncAgeStatus = "OK"
                                $status.SyncActive    = $true
                                Write-Host "    [OK] Last sync : $($status.LastSyncDateTime)  ($('{0:0}' -f $age.TotalMinutes) min ago)" -ForegroundColor Green
                            } elseif ($age.TotalHours -le 3) {
                                $status.SyncAgeStatus = "Warning"
                                $status.SyncActive    = $true
                                Write-Host "    [!]  Last sync : $($status.LastSyncDateTime)  ($([int]$age.TotalMinutes) min ago - WARNING)" -ForegroundColor Yellow
                                $status.Warnings += "Last sync $([int]$age.TotalMinutes) min ago (> 30 min threshold)"
                            } else {
                                $status.SyncAgeStatus = "Error"
                                $status.SyncActive    = $false
                                Write-Host "    [X]  Last sync : $($status.LastSyncDateTime)  ($([int]$age.TotalHours) h ago - STALE)" -ForegroundColor Red
                                $status.Errors += (Write-BWsError -Code "BWS-ENTRA-001" -Message "Sync stale: $([int]$age.TotalHours)h ago (threshold: 3h)" -Detail "LastSyncDateTime: $($status.LastSyncDateTime)" -CheckStep "[1/2] Org sync state")
                            }
                        } catch {
                            Write-Host "    [!]  Last sync time could not be parsed: $($status.LastSyncDateTime)" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "    [!]  Last sync time : not available" -ForegroundColor Yellow
                        $status.Warnings += "onPremisesLastSyncDateTime not returned"
                    }

                    if ($status.LastPasswordSyncDateTime) {
                        Write-Host "         Last pwd sync : $($status.LastPasswordSyncDateTime)" -ForegroundColor Gray
                    }

                    # On-prem domains
                    if ($status.OnPremDomains.Count -gt 0) {
                        Write-Host "         On-prem domains : $($status.OnPremDomains -join ', ')" -ForegroundColor Gray
                    }

                } elseif ($org.onPremisesSyncEnabled -eq $false) {
                    Write-Host "    [!]  Sync was previously enabled but is now DISABLED" -ForegroundColor Yellow
                    $status.Errors += (Write-BWsError -Code "BWS-ENTRA-002" -Message "onPremisesSyncEnabled=false (sync disabled)" -Detail "Sync was previously enabled but is now stopped." -Severity "Warning" -CheckStep "[1/2] Org sync state")
                } else {
                    Write-Host "    [i]  Sync not enabled (cloud-only tenant)" -ForegroundColor Gray
                }
            }
        } catch {
            Write-Host "    [X] Organization query failed: $($_.Exception.Message)" -ForegroundColor Red
            $status.Errors += (Write-BWsError -Code "BWS-GRAPH-010" -Message "Organization (onPremises) query failed" -Detail $_.Exception.Message -CheckStep "[1/2] Org sync state")
        }

        # ===================================================================
        # CHECK 2: onPremisesDirectorySynchronization features
        # GET /v1.0/directory/onPremisesSynchronization
        # Returns: features.passwordSyncEnabled, deviceWritebackEnabled, etc.
        # Only meaningful if sync is enabled
        # ===================================================================
        Write-Host ""
        Write-Host "  [2/2] Sync feature flags (Graph v1.0/directory/onPremisesSynchronization)" -ForegroundColor Yellow
        try {
            $syncUri = 'https://graph.microsoft.com/v1.0/directory/onPremisesSynchronization'
            $syncResp = Invoke-BWsGraphRequest -Uri $syncUri -Method GET -ErrorAction Stop

            # The response may be a collection or single object
            $syncObj = if ($syncResp.PSObject.Properties['value']) {
                $syncResp.value | Select-Object -First 1
            } else { $syncResp }

            if ($syncObj -and $syncObj.PSObject.Properties['features']) {
                $f = $syncObj.features

                $status.PasswordSyncEnabled      = [bool]$f.passwordSyncEnabled
                $status.PasswordWritebackEnabled = [bool]$f.passwordWritebackEnabled
                $status.DeviceWritebackEnabled   = [bool]$f.deviceWritebackEnabled
                $status.GroupWritebackEnabled    = [bool]$f.groupWriteBackEnabled
                $status.UserWritebackEnabled     = [bool]$f.userWritebackEnabled
                $status.DirectoryExtensionsEnabled = [bool]$f.directoryExtensionsEnabled

                $rows = @(
                    @{ Label="Password Hash Sync";           Val=$status.PasswordSyncEnabled;         Key="passwordSyncEnabled" },
                    @{ Label="Password Writeback";           Val=$status.PasswordWritebackEnabled;    Key="passwordWritebackEnabled" },
                    @{ Label="Device Writeback";             Val=$status.DeviceWritebackEnabled;      Key="deviceWritebackEnabled" },
                    @{ Label="Group Writeback";              Val=$status.GroupWritebackEnabled;       Key="groupWriteBackEnabled" },
                    @{ Label="User Writeback";               Val=$status.UserWritebackEnabled;        Key="userWritebackEnabled" },
                    @{ Label="Directory Extensions";         Val=$status.DirectoryExtensionsEnabled;  Key="directoryExtensionsEnabled" }
                )

                foreach ($r in $rows) {
                    $icon  = if ($r.Val) { "[OK]" } else { "[ ]" }
                    $color = if ($r.Val) { "Green" } else { "Gray" }
                    Write-Host ("    {0}  {1}" -f $icon, $r.Label.PadRight(28)) -ForegroundColor $color
                }

                if (-not $status.PasswordSyncEnabled -and $status.SyncEnabled) {
                    $status.Warnings += "Password Hash Sync is disabled"
                }
            } else {
                Write-Host "    [i]  No sync feature data returned (tenant may be cloud-only)" -ForegroundColor Gray
            }
        } catch {
            if ($_.Exception.Message -match "403|Forbidden|Unauthorized") {
                Write-Host "    [!]  Access denied to /directory/onPremisesSynchronization" -ForegroundColor Yellow
                Write-Host "         Requires OnPremDirectorySynchronization.Read.All permission" -ForegroundColor Gray
                $status.Warnings += (Write-BWsError -Code "BWS-AUTH-010" -Message "Sync feature flags: access denied" -Detail "Requires OnPremDirectorySynchronization.Read.All" -Severity "Warning" -HttpStatus 403 -CheckStep "[2/2] Sync features")
            } else {
                Write-Host "    [!]  Feature flags query failed: $($_.Exception.Message)" -ForegroundColor Yellow
                $status.Warnings += "Feature flags query: $($_.Exception.Message)"
            }
        }

    } catch {
        Write-Host "  [X] Unexpected error: $($_.Exception.Message)" -ForegroundColor Red
        $status.Errors += "Unexpected error: $($_.Exception.Message)"
    }

    # ===================================================================
    # SUMMARY
    # ===================================================================
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  ENTRA ID CONNECT SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan

    $syncCol  = if ($status.SyncEnabled) { "Green" } else { "Red" }
    $actCol   = if ($status.SyncActive)  { "Green" } elseif ($status.SyncEnabled) { "Yellow" } else { "Gray" }
    $ageLabel = switch ($status.SyncAgeStatus) {
        "OK"      { "[OK]" }
        "Warning" { "[!]" }
        "Error"   { "[X]" }
        default   { "[-]" }
    }

    Write-Host "  Sync enabled           : $(if ($status.SyncEnabled) { 'Yes ([OK])' } else { 'No' })" -ForegroundColor $syncCol
    Write-Host "  Sync active (< 3h)     : $(if ($status.SyncActive) { 'Yes ([OK])' } else { 'No' })" -ForegroundColor $actCol

    if ($status.LastSyncDateTime) {
        $ageStr = if ($status.SyncAgeMinutes -ne $null) {
            if ($status.SyncAgeMinutes -lt 60) { "$($status.SyncAgeMinutes) min ago" }
            else { "$([int]($status.SyncAgeMinutes/60)) h ago" }
        } else { "N/A" }
        Write-Host "  Last sync              : $($status.LastSyncDateTime) ($ageStr)" -ForegroundColor $(if ($status.SyncAgeStatus -eq "OK") { "Green" } elseif ($status.SyncAgeStatus -eq "Warning") { "Yellow" } else { "Red" })
    }
    if ($status.LastPasswordSyncDateTime) {
        Write-Host "  Last password sync     : $($status.LastPasswordSyncDateTime)" -ForegroundColor Gray
    }

    # Feature flags
    $flagsAvail = $status.PasswordSyncEnabled -ne $null
    if ($flagsAvail) {
        Write-Host ""
        Write-Host "  Feature Flags (from onPremisesDirectorySynchronization):" -ForegroundColor White
        foreach ($item in @(
            @{ N="Password Hash Sync";    V=$status.PasswordSyncEnabled },
            @{ N="Password Writeback";    V=$status.PasswordWritebackEnabled },
            @{ N="Device Writeback";      V=$status.DeviceWritebackEnabled },
            @{ N="Group Writeback";       V=$status.GroupWritebackEnabled },
            @{ N="User Writeback";        V=$status.UserWritebackEnabled },
            @{ N="Directory Extensions";  V=$status.DirectoryExtensionsEnabled }
        )) {
            $lbl = "    $($item.N.PadRight(22)) :"
            if ($item.V -eq $true)  { Write-Host "$lbl Enabled" -ForegroundColor Green }
            elseif ($item.V -eq $false) { Write-Host "$lbl Disabled" -ForegroundColor Gray }
            else                    { Write-Host "$lbl Unknown" -ForegroundColor Gray }
        }
    }

    if ($status.OnPremDomains.Count -gt 0) {
        Write-Host ""
        Write-Host "  On-prem domains        : $($status.OnPremDomains -join ', ')" -ForegroundColor White
    }
    Write-Host ""
    Write-Host "  Errors                 : $($status.Errors.Count)" -ForegroundColor $(if ($status.Errors.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Warnings               : $($status.Warnings.Count)" -ForegroundColor $(if ($status.Warnings.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    if (-not $CompactView) {
        if ($status.Errors.Count -gt 0) {
            Write-Host "  ERRORS:" -ForegroundColor Red
            $status.Errors   | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
            Write-Host ""
        }
        if ($status.Warnings.Count -gt 0) {
            Write-Host "  WARNINGS:" -ForegroundColor Yellow
            $status.Warnings | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
            Write-Host ""
        }
    }

    return @{
        Status         = $status
        CheckPerformed = $true
    }
}

function Test-IntuneConnector {
    <#
    .SYNOPSIS
        Checks the Intune Connector for Active Directory (ODJ/Hybrid Autopilot connector).
        API: GET /beta/deviceManagement/ndesConnectors
        Resource: ndesConnector (id, displayName, machineName, state, lastConnectionDateTime,
                  enrolledDateTime, connectorVersion)
        Deprecated threshold: versions < 6.2501.2000.5 (Microsoft, June 2025)
    #>
    param(
        [bool]$CompactView = $false
    )

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  INTUNE CONNECTOR FOR ACTIVE DIRECTORY CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    # Minimum supported connector version (Microsoft deprecation June 2025)
    $minSupportedVersion = [Version]"6.2501.2000.5"

    $connectorStatus = @{
        IsConnected       = $false
        ConnectorCount    = 0
        ActiveCount       = 0
        InactiveCount     = 0
        DeprecatedCount   = 0
        Connectors        = @()
        OnPremSyncEnabled = $false
        LastSyncDateTime  = $null
        Errors            = @()
        Warnings          = @()
    }

    try {
        # -- Graph connection --------------------------------------------
        $graphCtx = Get-BWsGraphContext -ErrorAction SilentlyContinue
        if (-not $graphCtx) {
            try {
                Connect-BWsGraph -Scopes "DeviceManagementServiceConfig.Read.All","DeviceManagementConfiguration.Read.All" -ErrorAction Stop
            } catch {
                Write-Host "  [X] Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
                $connectorStatus.Errors += (Write-BWsError -Code "BWS-AUTH-001" -Message "Graph connection failed" -Detail $_.Exception.Message -CheckStep "Graph connect")
                return @{ Status = $connectorStatus; CheckPerformed = $false }
            }
        }

        # -- 1. NDES / ODJ Connectors  (confirmed API from Microsoft Learn) --
        Write-Host "  [1/3] Intune Connector for Active Directory (ODJ)" -ForegroundColor Yellow
        try {
            $ndesResp = Invoke-BWsBetaGraphRequest -Uri 'deviceManagement/ndesConnectors' -Method GET -ErrorAction Stop

            if ($ndesResp -and $ndesResp.PSObject.Properties['value']) {
                $allConnectors = @($ndesResp.value)
            } elseif ($ndesResp) {
                $allConnectors = @($ndesResp)
            } else {
                $allConnectors = @()
            }

            $connectorStatus.ConnectorCount = $allConnectors.Count

            if ($allConnectors.Count -eq 0) {
                Write-Host "    [i] No connectors configured" -ForegroundColor Gray
            } else {
                foreach ($c in $allConnectors) {
                    $cState   = if ($c.state)         { $c.state }         else { "unknown" }
                    $cName    = if ($c.displayName)   { $c.displayName }   else { "(unnamed)" }
                    $cMachine = if ($c.machineName)   { $c.machineName }   else { "(unknown)" }
                    $cVer     = if ($c.connectorVersion) { $c.connectorVersion } else { "(unknown)" }
                    $cCheckin = $c.lastConnectionDateTime
                    $cEnrolled= $c.enrolledDateTime

                    # Version deprecation check
                    $isDeprecated = $false
                    $verObj = $null
                    try {
                        if ($cVer -ne "(unknown)") {
                            $verObj = [Version]$cVer
                            $isDeprecated = ($verObj -lt $minSupportedVersion)
                        }
                    } catch { $isDeprecated = $false }

                    # Last check-in freshness
                    $freshness = "Unknown"
                    $freshnessColor = "Gray"
                    if ($cCheckin) {
                        try {
                            $age = (Get-Date) - [DateTime]$cCheckin
                            if ($age.TotalHours -le 1) {
                                $freshness = "Recent (< 1h)"
                                $freshnessColor = "Green"
                            } elseif ($age.TotalHours -le 24) {
                                $freshness = "Warning (< 24h)"
                                $freshnessColor = "Yellow"
                                $connectorStatus.Warnings += "$cName : last check-in $([int]$age.TotalHours)h ago"
                            } else {
                                $freshness = "STALE ($([int]$age.TotalDays)d ago)"
                                $freshnessColor = "Red"
                                $connectorStatus.Errors += (Write-BWsError -Code "BWS-INTUNE-010" -Message "NDES connector stale: $cName last check-in $([int]$age.TotalDays) days ago" -Detail "Machine: $cMachine" -CheckStep "[1/3] NDES connectors")
                            }
                        } catch { $freshness = "Parse error" }
                    }

                    # Build connector record
                    $connRecord = @{
                        DisplayName    = $cName
                        MachineName    = $cMachine
                        State          = $cState
                        Version        = $cVer
                        IsDeprecated   = $isDeprecated
                        LastCheckin    = $cCheckin
                        EnrolledDate   = $cEnrolled
                        Freshness      = $freshness
                    }
                    $connectorStatus.Connectors += $connRecord

                    # Count by state
                    if ($cState -eq "active") {
                        $connectorStatus.ActiveCount++
                        $connectorStatus.IsConnected = $true
                    } else {
                        $connectorStatus.InactiveCount++
                    }
                    if ($isDeprecated) { $connectorStatus.DeprecatedCount++ }

                    # Console output
                    $stateColor = if ($cState -eq "active") { "Green" } else { "Red" }
                    Write-Host "    Connector : $cName" -ForegroundColor White
                    Write-Host "      Machine   : $cMachine" -ForegroundColor Gray
                    Write-Host "      State     : " -NoNewline -ForegroundColor Gray
                    Write-Host $cState.ToUpper() -ForegroundColor $stateColor
                    Write-Host "      Version   : " -NoNewline -ForegroundColor Gray
                    if ($isDeprecated) {
                        Write-Host "$cVer  [DEPRECATED - update required!]" -ForegroundColor Red
                        $connectorStatus.Errors += (Write-BWsError -Code "BWS-INTUNE-011" -Message "NDES connector deprecated: $cName version $cVer (min: $minSupportedVersion)" -Detail "Upgrade required: https://learn.microsoft.com/en-us/mem/intune/enrollment/windows-enrollment-prerequisites" -CheckStep "[1/3] NDES connectors")
                    } else {
                        Write-Host $cVer -ForegroundColor $(if ($cVer -eq "(unknown)") { "Gray" } else { "Green" })
                    }
                    if ($cCheckin) {
                        Write-Host "      Last check-in : $cCheckin" -NoNewline -ForegroundColor Gray
                        Write-Host "  [$freshness]" -ForegroundColor $freshnessColor
                    }
                    if ($cEnrolled) {
                        Write-Host "      Enrolled  : $cEnrolled" -ForegroundColor Gray
                    }
                    Write-Host ""
                }
            }
        } catch {
            Write-Host "    [!] Could not query ndesConnectors: $($_.Exception.Message)" -ForegroundColor Yellow
            $connectorStatus.Errors += (Write-BWsError -Code "BWS-GRAPH-020" -Message "ndesConnectors API query failed" -Detail $_.Exception.Message -CheckStep "[1/3] NDES connectors")
        }

        # -- 2. On-Premises Sync (Entra ID Connect / AAD Sync) ----------
        Write-Host "  [2/3] On-Premises Directory Sync (Entra ID Connect)" -ForegroundColor Yellow
        try {
            $orgResp = Invoke-BWsGraphRequest -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesLastPasswordSyncDateTime' -Method GET -ErrorAction Stop
            $org = if ($orgResp.PSObject.Properties['value']) { $orgResp.value | Select-Object -First 1 } else { $orgResp }

            if ($org) {
                $connectorStatus.OnPremSyncEnabled = [bool]$org.onPremisesSyncEnabled
                $connectorStatus.LastSyncDateTime  = $org.onPremisesLastSyncDateTime

                if ($org.onPremisesSyncEnabled) {
                    Write-Host "    [OK] Sync enabled" -ForegroundColor Green
                    if ($org.onPremisesLastSyncDateTime) {
                        Write-Host "      Last sync           : $($org.onPremisesLastSyncDateTime)" -ForegroundColor Gray
                    }
                    if ($org.onPremisesLastPasswordSyncDateTime) {
                        Write-Host "      Last password sync  : $($org.onPremisesLastPasswordSyncDateTime)" -ForegroundColor Gray
                    }
                } else {
                    Write-Host "    [i] On-premises sync not enabled" -ForegroundColor Gray
                }
            }
        } catch {
            Write-Host "    [!] Could not query org sync status: $($_.Exception.Message)" -ForegroundColor Yellow
            $connectorStatus.Warnings += "Org sync query failed: $($_.Exception.Message)"
        }

        # -- 3. AD Server in Azure (ARM REST, no Az.Resources) ----------
        Write-Host ""
        Write-Host "  [3/3] AD / Sync Server presence in Azure" -ForegroundColor Yellow
        try {
            $azCtx = Get-BWsAzContext -ErrorAction SilentlyContinue
            if ($azCtx) {
                # Find VMs matching AD/DC/Sync naming patterns via ARM REST (PS 5.1 safe)
                $adVMs = Find-BWsAzResource -ResourceType "Microsoft.Compute/virtualMachines" -NamePattern "^[0-9]{4,5}-S[0-9]{2}$|DC|ADDS|Sync|AAD" -ErrorAction SilentlyContinue

                if ($adVMs -and $adVMs.Count -gt 0) {
                    Write-Host "    [OK] Found $($adVMs.Count) potential AD/Sync server(s)" -ForegroundColor Green
                    foreach ($vm in $adVMs) {
                        Write-Host "      $($vm.name)  ($($vm.location))" -ForegroundColor Gray
                        $connectorStatus.Connectors += @{
                            DisplayName  = $vm.name
                            MachineName  = $vm.name
                            State        = "azure-vm"
                            Version      = "N/A"
                            IsDeprecated = $false
                            LastCheckin  = $null
                            EnrolledDate = $null
                            Freshness    = "N/A"
                        }
                    }
                } else {
                    Write-Host "    [i] No matching AD/Sync VMs found" -ForegroundColor Gray
                }
            } else {
                Write-Host "    [i] No Azure connection available" -ForegroundColor Gray
            }
        } catch {
            Write-Host "    [!] Azure VM check failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }

    } catch {
        Write-Host "  [X] General error: $($_.Exception.Message)" -ForegroundColor Red
        $connectorStatus.Errors += "General error: $($_.Exception.Message)"
    }

    # -- Summary ---------------------------------------------------------
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  INTUNE CONNECTOR SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan

    $overallOK = ($connectorStatus.IsConnected -and
                  $connectorStatus.DeprecatedCount -eq 0 -and
                  $connectorStatus.Errors.Count -eq 0)

    Write-Host "  Connectors configured  : $($connectorStatus.ConnectorCount)" -ForegroundColor White
    Write-Host "  Active                 : $($connectorStatus.ActiveCount)" -ForegroundColor $(if ($connectorStatus.ActiveCount -gt 0) { "Green" } else { "Red" })
    Write-Host "  Inactive               : $($connectorStatus.InactiveCount)" -ForegroundColor $(if ($connectorStatus.InactiveCount -gt 0) { "Yellow" } else { "Green" })

    if ($connectorStatus.DeprecatedCount -gt 0) {
        Write-Host "  DEPRECATED versions    : $($connectorStatus.DeprecatedCount)  [UPDATE REQUIRED]" -ForegroundColor Red
    } else {
        Write-Host "  Deprecated versions    : 0" -ForegroundColor Green
    }

    Write-Host "  On-prem sync enabled   : $(if ($connectorStatus.OnPremSyncEnabled) {'Yes'} else {'No/Unknown'})" -ForegroundColor $(if ($connectorStatus.OnPremSyncEnabled) { "Green" } else { "Gray" })
    if ($connectorStatus.LastSyncDateTime) {
        Write-Host "  Last sync              : $($connectorStatus.LastSyncDateTime)" -ForegroundColor Gray
    }
    Write-Host "  Errors                 : $($connectorStatus.Errors.Count)" -ForegroundColor $(if ($connectorStatus.Errors.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Warnings               : $($connectorStatus.Warnings.Count)" -ForegroundColor $(if ($connectorStatus.Warnings.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "  Overall status         : " -NoNewline -ForegroundColor White
    Write-Host $(if ($overallOK) { "[OK] HEALTHY" } elseif ($connectorStatus.IsConnected) { "[!] ACTIVE WITH ISSUES" } else { "[X] NOT CONNECTED" }) `
        -ForegroundColor $(if ($overallOK) { "Green" } elseif ($connectorStatus.IsConnected) { "Yellow" } else { "Red" })

    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    if (-not $CompactView -and ($connectorStatus.Errors.Count -gt 0 -or $connectorStatus.Warnings.Count -gt 0)) {
        if ($connectorStatus.Errors.Count -gt 0) {
            Write-Host "  ERRORS:" -ForegroundColor Red
            $connectorStatus.Errors | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
            Write-Host ""
        }
        if ($connectorStatus.Warnings.Count -gt 0) {
            Write-Host "  WARNINGS:" -ForegroundColor Yellow
            $connectorStatus.Warnings | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
            Write-Host ""
        }
    }

    return @{
        Status         = $connectorStatus
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
    <#
    .SYNOPSIS
        Checks BWS standard software packages in Intune.
        Detects Mac clients and conditionally checks BeyondTrust + Printix apps.
    .NOTES
        Microsoft Learn API references:
        - Managed devices: GET /beta/deviceManagement/managedDevices?$filter=operatingSystem eq 'macOS'
          Source: learn.microsoft.com/graph/api/intune-devices-manageddevice-list
        - Mobile apps:    GET /beta/deviceAppManagement/mobileApps?$filter=isof('microsoft.graph.macOSLobApp')
          Source: learn.microsoft.com/graph/api/intune-apps-mobileapp-list
    #>
    param([bool]$CompactView = $false)

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  BWS STANDARD SOFTWARE PACKAGES CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    $softwareStatus = @{
        Total          = 7
        Found          = [System.Collections.Generic.List[hashtable]]::new()
        Missing        = [System.Collections.Generic.List[hashtable]]::new()
        Errors         = @()
        HasMacClients  = $false
        MacDeviceCount = 0
        BeyondTrustOk  = $null   # $true/$false/$null=not checked
        PrintixOk      = $null
    }

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
        # -- Graph connection -------------------------------------------------
        if (-not (Get-BWsGraphContext)) {
            try { Connect-BWsGraph -ErrorAction Stop } catch {
                $softwareStatus.Errors += "Graph connection failed: $($_.Exception.Message)"
                return @{ Status = $softwareStatus; CheckPerformed = $false }
            }
        }

        # ===================================================================
        # CHECK 1: Detect Mac clients
        # GET /beta/deviceManagement/managedDevices?$filter=operatingSystem eq 'macOS'
        # Microsoft Learn: learn.microsoft.com/graph/api/intune-devices-manageddevice-list
        # ===================================================================
        Write-Host "  [1/3] Mac client detection..." -ForegroundColor Yellow
        try {
            $macUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices" +
                      "?`$filter=operatingSystem eq 'macOS'&`$select=id,deviceName,operatingSystem&`$top=1"
            $macResp = Invoke-BWsBetaGraphRequest -Uri $macUri -Method GET -ErrorAction Stop
            $macCount = if ($macResp.PSObject.Properties['value']) { $macResp.value.Count } else { 0 }
            # Use @odata.count if present, else count returned items
            if ($macResp.PSObject.Properties['@odata.count']) { $macCount = [int]$macResp.'@odata.count' }
            $softwareStatus.MacDeviceCount = $macCount
            $softwareStatus.HasMacClients  = ($macCount -gt 0)
            if ($softwareStatus.HasMacClients) {
                Write-Host "    [i]  Mac clients detected: $macCount device(s)" -ForegroundColor Cyan
            } else {
                Write-Host "    [OK] No Mac clients - Mac-specific app checks skipped" -ForegroundColor Gray
            }
        } catch {
            Write-Host "    [!]  Mac detection failed: $($_.Exception.Message)" -ForegroundColor Yellow
            $softwareStatus.Errors += "Mac detection: $($_.Exception.Message)"
        }

        # ===================================================================
        # CHECK 2: Load all Intune apps (Win32 + Store + M365 + macOS)
        # ===================================================================
        Write-Host "  [2/3] Retrieving Intune app inventory..." -ForegroundColor Yellow
        $allApps = [System.Collections.Generic.List[object]]::new()
        foreach ($filter in @(
            "isof('microsoft.graph.win32LobApp')",
            "isof('microsoft.graph.winGetApp')",
            "isof('microsoft.graph.officeSuiteApp')",
            "isof('microsoft.graph.macOSLobApp')",
            "isof('microsoft.graph.macOSMicrosoftEdgeApp')"
        )) {
            try {
                $apps = Invoke-BWsGraphPagedRequest -Uri "deviceAppManagement/mobileApps?`$filter=$filter&`$top=999" -ErrorAction SilentlyContinue
                if ($apps) { foreach ($a in $apps) { $allApps.Add($a) } }
            } catch {}
        }
        # Fallback: all apps if specific filters returned nothing
        if ($allApps.Count -eq 0) {
            try {
                $allAppsRaw = Invoke-BWsGraphPagedRequest -Uri "deviceAppManagement/mobileApps?`$top=999"
                foreach ($a in $allAppsRaw) { $allApps.Add($a) }
            } catch {
                $softwareStatus.Errors += "App retrieval failed: $($_.Exception.Message)"
            }
        }
        Write-Host "    [i]  Total apps in Intune: $($allApps.Count)" -ForegroundColor Gray

        # ===================================================================
        # CHECK 3a: Standard software packages (7 required apps)
        # ===================================================================
        Write-Host "  [3/3] Checking standard software packages..." -ForegroundColor Yellow
        Write-Host ""
        foreach ($sw in $requiredSoftware) {
            $found = $allApps | Where-Object { $_.displayName -like "*$sw*" } | Select-Object -First 1
            if (-not $found) {
                # Fuzzy: match key words
                $words = $sw -split '\s+' | Where-Object { $_.Length -gt 3 }
                foreach ($w in $words) {
                    $found = $allApps | Where-Object { $_.displayName -like "*$w*" } | Select-Object -First 1
                    if ($found) { break }
                }
            }
            if ($found) {
                Write-Host "    [OK] $sw" -ForegroundColor Green
                Write-Host "         Found: $($found.displayName)" -ForegroundColor Gray
                $softwareStatus.Found.Add(@{ SoftwareName=$sw; ActualName=$found.displayName; AppId=$found.id })
            } else {
                Write-Host "    [X]  $sw  -  NOT FOUND IN INTUNE" -ForegroundColor Red
                $softwareStatus.Missing.Add(@{ SoftwareName=$sw })
            }
        }

        # ===================================================================
        # CHECK 3b: BeyondTrust  (only if Mac clients present)
        # Requirement: "Check if customer has Mac clients and if so check if
        #               BeyondTrust is deployed on Intune"
        # App name in Intune: varies - search for 'BeyondTrust' or 'beyond Trust'
        # ===================================================================
        Write-Host ""
        if ($softwareStatus.HasMacClients) {
            Write-Host "  [Mac] BeyondTrust Remote Support check (Mac clients detected)" -ForegroundColor Cyan
            $btApp = $allApps | Where-Object { $_.displayName -like "*BeyondTrust*" -or $_.displayName -like "*beyond Trust*" } | Select-Object -First 1
            if ($btApp) {
                Write-Host "    [OK] BeyondTrust found: $($btApp.displayName)" -ForegroundColor Green
                $softwareStatus.BeyondTrustOk = $true
                $softwareStatus.Found.Add(@{ SoftwareName="BeyondTrust Remote Support (Mac)"; ActualName=$btApp.displayName; AppId=$btApp.id })
            } else {
                Write-Host "    [X]  BeyondTrust NOT FOUND in Intune - required for Mac clients" -ForegroundColor Red
                $softwareStatus.BeyondTrustOk = $false
                $softwareStatus.Missing.Add(@{ SoftwareName="BeyondTrust Remote Support (Mac)" })
                $softwareStatus.Errors += (Write-BWsError -Code "BWS-INTUNE-020" -Message "BeyondTrust Remote Support not found in Intune (Mac clients: $($softwareStatus.MacDeviceCount))" -Detail "Deploy BeyondTrust via Intune for macOS. Mac devices detected: $($softwareStatus.MacDeviceCount)." -CheckStep "[3b] BeyondTrust")
            }
        } else {
            Write-Host "  [Mac] BeyondTrust: skipped (no Mac clients)" -ForegroundColor Gray
            $softwareStatus.BeyondTrustOk = $null
        }

        # ===================================================================
        # CHECK 3c: Printix App  (only if Mac clients present)
        # Requirement: "Check if customer has Mac clients, if so check if
        #               the Printix App is deployed on Intune"
        # ===================================================================
        if ($softwareStatus.HasMacClients) {
            Write-Host "  [Mac] Printix App check (Mac clients detected)" -ForegroundColor Cyan
            $printixApp = $allApps | Where-Object { $_.displayName -like "*Printix*" } | Select-Object -First 1
            if ($printixApp) {
                Write-Host "    [OK] Printix found: $($printixApp.displayName)" -ForegroundColor Green
                $softwareStatus.PrintixOk = $true
                $softwareStatus.Found.Add(@{ SoftwareName="Printix (Mac)"; ActualName=$printixApp.displayName; AppId=$printixApp.id })
            } else {
                Write-Host "    [X]  Printix NOT FOUND in Intune - required for Mac clients" -ForegroundColor Red
                $softwareStatus.PrintixOk = $false
                $softwareStatus.Missing.Add(@{ SoftwareName="Printix (Mac)" })
                $softwareStatus.Errors += (Write-BWsError -Code "BWS-INTUNE-021" -Message "Printix not found in Intune (Mac clients: $($softwareStatus.MacDeviceCount))" -Detail "Deploy Printix app via Intune for macOS. Mac devices detected: $($softwareStatus.MacDeviceCount)." -CheckStep "[3c] Printix")
            }
        } else {
            Write-Host "  [Mac] Printix: skipped (no Mac clients)" -ForegroundColor Gray
            $softwareStatus.PrintixOk = $null
        }

    } catch {
        $softwareStatus.Errors += "Unexpected error: $($_.Exception.Message)"
    }

    $softwareStatus.Total = $softwareStatus.Found.Count + $softwareStatus.Missing.Count

    # -- Summary ------------------------------------------------------------
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  BWS SOFTWARE PACKAGES SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Mac Clients:    $(if ($softwareStatus.HasMacClients) { "$($softwareStatus.MacDeviceCount) device(s)" } else { 'None' })" -ForegroundColor $(if ($softwareStatus.HasMacClients) { "Cyan" } else { "Gray" })
    Write-Host "  Found:          $($softwareStatus.Found.Count)" -ForegroundColor $(if ($softwareStatus.Missing.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "  Missing:        $($softwareStatus.Missing.Count)" -ForegroundColor $(if ($softwareStatus.Missing.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Errors:         $($softwareStatus.Errors.Count)" -ForegroundColor $(if ($softwareStatus.Errors.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    if (-not $CompactView) {
        if ($softwareStatus.Missing.Count -gt 0) {
            Write-Host "  MISSING:" -ForegroundColor Red
            $softwareStatus.Missing | ForEach-Object { Write-Host "    [X] $($_.SoftwareName)" -ForegroundColor Red }
            Write-Host ""
        }
        if ($softwareStatus.Errors.Count -gt 0) {
            Write-Host "  ERRORS:" -ForegroundColor Yellow
            $softwareStatus.Errors | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
            Write-Host ""
        }
    }

    return @{ Status = $softwareStatus; CheckPerformed = $true }
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
        
        # -- PS 5.1 / Desktop enforcement ------------------------------
        # Microsoft.Online.SharePoint.PowerShell requires Windows PowerShell 5.1
        # It does NOT support PowerShell 7 / Core.
        # MS Learn: learn.microsoft.com/powershell/sharepoint/sharepoint-online/connect-sharepoint-online
        if ($PSVersionTable.PSEdition -ne "Desktop") {
            Write-Host "  [!] SharePoint module requires Windows PowerShell 5.1 (Desktop edition)." -ForegroundColor Yellow
            Write-Host "      Current: PS $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))" -ForegroundColor Gray
            Write-Host "      SharePoint check SKIPPED - please run in powershell.exe (PS 5.1)" -ForegroundColor Yellow
            $spConfig.Errors += (Write-BWsError -Code "BWS-SYS-010" -Message "SharePoint requires PS 5.1 Desktop (current: $($PSVersionTable.PSEdition))" -Detail "Microsoft.Online.SharePoint.PowerShell only runs on Windows PowerShell 5.1" -Severity "Warning" -CheckStep "PS edition check")
            return @{ Status = $spConfig; CheckPerformed = $false }
        }

        # -- Module detection ----------------------------------------------
        $spoModuleAvailable = $false
        $moduleType = $null
        if (Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell" -ErrorAction SilentlyContinue) {
            $spoModuleAvailable = $true
            $moduleType = "SPO"
        }

        if ($spoModuleAvailable) {
            Write-Host "  [SharePoint] Using $moduleType module (PS 5.1 Desktop)" -ForegroundColor Gray

            # -- Auto-connect using stored GlobalAdmin session -------------
            # If Connect-BWsGlobalAdmin ran, $script:SharePointConnected may already be $true.
            $needsConnection = $false
            $tenant = $null
            try {
                $tenant = Get-SPOTenant -ErrorAction Stop
                Write-Host "  [SharePoint] Session active" -ForegroundColor Gray
            } catch {
                $needsConnection = $true
            }

            if ($needsConnection) {
                # Determine SPO admin URL
                $spoConnUrl = $SharePointUrl
                if ([string]::IsNullOrEmpty($spoConnUrl) -and $script:GlobalAdminUPN) {
                    $upnDomain = ($script:GlobalAdminUPN -split '@' | Select-Object -Last 1) -replace '\.onmicrosoft\.com$',''
                    if ($upnDomain) { $spoConnUrl = "https://$upnDomain-admin.sharepoint.com" }
                }

                if (-not [string]::IsNullOrEmpty($spoConnUrl)) {
                    Write-Host "  [SharePoint] Connecting to: $spoConnUrl" -ForegroundColor Yellow
                    try {
                        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop 2>$null 3>$null
                        # ModernAuth = MFA compatible (MS Learn recommended)
                        try {
                            Connect-SPOService -Url $spoConnUrl `
                                -ModernAuth $true `
                                -AuthenticationUrl "https://login.microsoftonline.com/organizations" `
                                -ErrorAction Stop
                        } catch {
                            # Fallback for environments without ModernAuth parameter (older module)
                            Connect-SPOService -Url $spoConnUrl -ErrorAction Stop
                        }
                        Write-Host "  [SharePoint] Connected successfully" -ForegroundColor Green
                        $tenant = Get-SPOTenant -ErrorAction Stop
                        $script:SharePointConnected = $true
                        $needsConnection = $false
                    } catch {
                        Write-Host "  [SharePoint] Connection failed: $($_.Exception.Message)" -ForegroundColor Red
                        $spConfig.Errors += (Write-BWsError -Code "BWS-SPO-001" -Message "SharePoint connection failed" -Detail $_.Exception.Message -CheckStep "SPO connect")
                    }
                } else {
                    Write-Host "  [SharePoint] No admin URL available. Use -SharePointUrl parameter." -ForegroundColor Yellow
                    $spConfig.Errors += (Write-BWsError -Code "BWS-SPO-002" -Message "SharePoint admin URL not provided" -Detail "Use -SharePointUrl https://<tenant>-admin.sharepoint.com" -Severity "Warning" -CheckStep "SPO connect")
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
                        $spConfig.Errors += (Write-BWsError -Code "BWS-SPO-010" -Message "SharePoint External Sharing is not 'Anyone' (current: $spSharingCapability)" -Detail "Required: ExternalUserAndGuestSharing. Entra admin centre > SharePoint > Sharing." -CheckStep "[1] External sharing")
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
                        $spConfig.Errors += (Write-BWsError -Code "BWS-SPO-020" -Message "Legacy auth not blocked in SharePoint (LegacyAuthProtocolsEnabled=true)" -Detail "Set via SPO admin centre > Access control > Apps using legacy auth." -CheckStep "[3] Legacy auth")
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
            Write-Host "  [!] Microsoft.Online.SharePoint.PowerShell module not found" -ForegroundColor Yellow
            Write-Host "      Install with: Install-Module -Name Microsoft.Online.SharePoint.PowerShell" -ForegroundColor Gray
            Write-Host "      Note: Requires Windows PowerShell 5.1 (Desktop edition)" -ForegroundColor Gray
            $spConfig.Errors += "Microsoft.Online.SharePoint.PowerShell not installed (requires PS 5.1 Desktop)"
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

            # -- Auto-connect using stored GlobalAdmin session ------------
            # If Connect-BWsGlobalAdmin ran successfully, $script:TeamsConnected=$true
            # and the Teams session is already established. If not, try to connect now.
            $teamsConnected = $false
            try {
                $csConfig = Get-CsTeamsClientConfiguration -ErrorAction Stop
                $teamsConnected = $true
                Write-Host "  [Teams] Session active" -ForegroundColor Gray
            } catch {
                # Not connected - try using stored GlobalAdmin credentials
                if ($script:GlobalAdminConnected -and $script:GlobalAdminUPN) {
                    Write-Host "  [Teams] Connecting with GlobalAdmin credentials (-AccountId)..." -ForegroundColor Yellow
                    try {
                        # MS Learn: Connect-MicrosoftTeams -AccountId reuses AAD token
                        $null = Connect-MicrosoftTeams `
                            -AccountId $script:GlobalAdminUPN `
                            -TenantId  $script:GlobalAdminTenantId `
                            -ErrorAction Stop 2>$null 3>$null
                        $script:TeamsConnected = $true
                        $teamsConnected = $true
                        Write-Host "  [Teams] Connected (no additional login required)" -ForegroundColor Green
                    } catch {
                        Write-Host "  [Teams] Auto-connect failed: $($_.Exception.Message)" -ForegroundColor Yellow
                        Write-Host "  [Teams] Please connect manually: Connect-MicrosoftTeams" -ForegroundColor Gray
                    }
                } else {
                    Write-Host "  [Teams] Not connected. Run Prerequisites first (-Full or GUI Prerequisites button)." -ForegroundColor Yellow
                    Write-Host "  [Teams] Or connect manually: Connect-MicrosoftTeams" -ForegroundColor Gray
                }
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
                        $teamsConfig.Errors += (Write-BWsError -Code "BWS-TEAMS-010" -Message "External access to unmanaged Teams accounts is enabled (AllowTeamsConsumer=true)" -Detail "Disable in Teams admin centre > External access." -CheckStep "[1] External access")
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
                        $teamsConfig.Errors += (Write-BWsError -Code "BWS-TEAMS-020" -Message "Cloud storage enabled in Teams: $($enabledProviders -join ', ')" -Detail "Disable in Teams admin centre > Teams apps > Permission policies." -CheckStep "[2] Cloud storage")
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
                        $teamsConfig.Errors += (Write-BWsError -Code "BWS-TEAMS-030" -Message "Anonymous meeting join/start is enabled ($($meetingIssues -join ', '))" -Detail "Disable in Teams admin centre > Meetings > Meeting policies." -CheckStep "[3] Meeting settings")
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
        
        if (-not $allUsers) { $allUsers = @() }
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

function Test-EntraSecurityConfig {
    <#
    .SYNOPSIS
        Checks Entra ID security configuration:
        1. MDM enrollment policies (Intune automatic enrollment scope)
        2. grp_UsersWithPriviledge assigned to MS Graph Command Line Tools enterprise app
        3. PIM eligible role assignments for partner accounts
        4. Old direct (non-PIM) privileged role assignments + SICT automation app removal
    .NOTES
        Microsoft Learn API references (v1.0 / beta):

        MDM policies (beta only):
          GET /beta/policies/mobileDeviceManagementPolicies
          Fields: id, displayName, appliesTo (none|all|selected), discoveryUrl
          Intune app ID: 0000000a-0000-0000-c000-000000000000
          Source: learn.microsoft.com/graph/api/mobiledevicemanagementpolicies-list

        Enterprise app group assignment (v1.0):
          GET /v1.0/servicePrincipals?$filter=displayName eq 'Microsoft Graph Command Line Tools'
          GET /v1.0/servicePrincipals/{id}/appRoleAssignedTo
          Fields: principalId, principalType, principalDisplayName
          Best practice: read via appRoleAssignedTo on the resource SP
          Source: learn.microsoft.com/graph/api/serviceprincipal-list-approleassignedto

        PIM eligible roles (v1.0 - iteration 3, current):
          GET /v1.0/roleManagement/directory/roleEligibilitySchedules?$expand=roleDefinition,principal
          Fields: principalId, roleDefinitionId, status, scheduleInfo
          Source: learn.microsoft.com/graph/api/resources/privilegedidentitymanagementv3-overview

        Direct role assignments (non-PIM, v1.0):
          GET /v1.0/roleManagement/directory/roleAssignments?$expand=principal,roleDefinition
          Source: learn.microsoft.com/graph/api/rbacapplication-list-roleassignments

        Service principal lookup (SICT app):
          GET /v1.0/servicePrincipals?$filter=displayName eq 'SICT'
          Source: learn.microsoft.com/graph/api/serviceprincipal-list
    #>
    param([bool]$CompactView = $false)

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  ENTRA ID SECURITY CONFIGURATION CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    $status = @{
        # MDM
        MdmAppliesTo        = $null    # "none"/"all"/"selected"
        MdmDisplayName      = $null
        MdmDiscoveryUrl     = $null
        MdmConfigOk         = $false
        # Enterprise app group
        GraphCliSpFound     = $false
        GraphCliSpId        = $null
        GrpPriviledgeFound  = $false
        GrpPriviledgeAssigned = $false
        # PIM
        PartnerPimUsers     = @()     # users with PIM eligible roles
        DirectAdminCount    = 0       # direct permanent role assignments (non-PIM)
        DirectAdmins        = @()
        # SICT
        SictAppFound        = $false
        SictAppId           = $null
        SictAppName         = $null
        # Errors / Warnings
        Errors              = @()
        Warnings            = @()
    }

    try {
        # Graph connection
        if (-not (Get-BWsGraphContext)) {
            try { Connect-BWsGraph -ErrorAction Stop } catch {
                $status.Errors += (Write-BWsError -Code "BWS-AUTH-001" -Message "Graph connection failed" -Detail $_.Exception.Message -CheckStep "Graph connect")
                return @{ Status = $status; CheckPerformed = $false }
            }
        }

        # =================================================================
        # CHECK 1: MDM Enrollment Policies
        # GET /beta/policies/mobileDeviceManagementPolicies
        # Intune MDM app ID: 0000000a-0000-0000-c000-000000000000
        # appliesTo must be "all" for automatic enrollment of all users
        # Source: learn.microsoft.com/graph/api/mobiledevicemanagementpolicies-list
        # =================================================================
        Write-Host "  [1/4] MDM enrollment policies (Entra ID -> Mobility)" -ForegroundColor Yellow
        try {
            $mdmUri = "https://graph.microsoft.com/beta/policies/mobileDeviceManagementPolicies"
            $mdmResp = Invoke-BWsBetaGraphRequest -Uri $mdmUri -Method GET -ErrorAction Stop
            $mdmPolicies = if ($mdmResp.PSObject.Properties['value']) { $mdmResp.value } else { @($mdmResp) }
            # Microsoft Intune has a well-known app ID
            $intunePolicy = $mdmPolicies | Where-Object {
                $_.id -eq "0000000a-0000-0000-c000-000000000000" -or
                $_.displayName -like "*Intune*" -or
                $_.displayName -like "*Microsoft Intune*"
            } | Select-Object -First 1
            if (-not $intunePolicy) { $intunePolicy = $mdmPolicies | Select-Object -First 1 }

            if ($intunePolicy) {
                $status.MdmAppliesTo   = $intunePolicy.appliesTo
                $status.MdmDisplayName = $intunePolicy.displayName
                $status.MdmDiscoveryUrl = $intunePolicy.discoveryUrl
                $status.MdmConfigOk    = ($intunePolicy.appliesTo -eq "all")

                $col = if ($status.MdmConfigOk) { "Green" } else { "Yellow" }
                $ico = if ($status.MdmConfigOk) { "[OK]" } else { "[!]" }
                Write-Host "    $ico MDM Policy  : $($intunePolicy.displayName)" -ForegroundColor $col
                Write-Host "         Applies To : $($intunePolicy.appliesTo)  (expected: all)" -ForegroundColor $col
                if ($intunePolicy.discoveryUrl) {
                    Write-Host "         Discovery  : $($intunePolicy.discoveryUrl)" -ForegroundColor Gray
                }
                if (-not $status.MdmConfigOk) {
                    $status.Warnings += "MDM enrollment scope is '$($intunePolicy.appliesTo)' - should be 'all' for automatic enrollment"
                }
            } else {
                Write-Host "    [!]  No MDM policy found - Intune auto-enrollment may not be configured" -ForegroundColor Yellow
                $status.Warnings += "No MDM enrollment policy found"
            }

            # Also check MAM policies
            try {
                $mamUri  = "https://graph.microsoft.com/beta/policies/mobileAppManagementPolicies"
                $mamResp = Invoke-BWsBetaGraphRequest -Uri $mamUri -Method GET -ErrorAction SilentlyContinue
                $mamPolicies = if ($mamResp -and $mamResp.PSObject.Properties['value']) { $mamResp.value } else { @() }
                $intuneMAM = $mamPolicies | Where-Object { $_.displayName -like "*Intune*" } | Select-Object -First 1
                if ($intuneMAM) {
                    Write-Host "    [i]  MAM Policy    : $($intuneMAM.displayName) (Applies: $($intuneMAM.appliesTo))" -ForegroundColor Gray
                }
            } catch {}
        } catch {
            $msg = "MDM policy check failed: $($_.Exception.Message)"
            if ($_.Exception.Message -match "403|Forbidden") {
                Write-Host "    [!]  Access denied to /beta/policies/mobileDeviceManagementPolicies" -ForegroundColor Yellow
                Write-Host "         Requires Policy.Read.All permission" -ForegroundColor Gray
                $status.Warnings += (Write-BWsError -Code "BWS-AUTH-011" -Message "MDM policy check: access denied (Policy.Read.All required)" -Detail $msg -Severity "Warning" -HttpStatus 403 -CheckStep "[1/4] MDM policies")
            } else {
                Write-Host "    [X]  $msg" -ForegroundColor Red
                $status.Errors += $msg
            }
        }

        # =================================================================
        # CHECK 2: grp_UsersWithPriviledge assigned to MS Graph Command Line Tools
        # Step 1: GET /v1.0/servicePrincipals?$filter=appId eq '14d82eec-204b-4c2f-b7e8-296a70dab67e'
        #         (well-known appId of Microsoft Graph Command Line Tools)
        # Step 2: GET /v1.0/servicePrincipals/{id}/appRoleAssignedTo
        #         -> check if principalDisplayName contains 'grp_UsersWithPriviledge'
        # Source: learn.microsoft.com/graph/api/serviceprincipal-list-approleassignedto
        # =================================================================
        Write-Host ""
        Write-Host "  [2/4] grp_UsersWithPriviledge -> MS Graph Command Line Tools" -ForegroundColor Yellow
        # Microsoft Graph Command Line Tools well-known appId
        $graphCliAppId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
        try {
            $spUri   = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$graphCliAppId'&`$select=id,displayName,appId"
            $spResp  = Invoke-BWsGraphRequest -Uri $spUri -Method GET -ErrorAction Stop
            $graphCliSP = if ($spResp.PSObject.Properties['value']) { $spResp.value | Select-Object -First 1 } else { $null }
            if ($graphCliSP) {
                $status.GraphCliSpFound = $true
                $status.GraphCliSpId   = $graphCliSP.id
                Write-Host "    [i]  Found SP: $($graphCliSP.displayName) (id: $($graphCliSP.id.Substring(0,8))...)" -ForegroundColor Gray

                # Get all appRoleAssignedTo - best practice per MS Learn
                $assignUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($graphCliSP.id)/appRoleAssignedTo"
                $assignResp = Invoke-BWsGraphPagedRequest -Uri $assignUri
                $grpAssignment = $assignResp | Where-Object {
                    $_.principalDisplayName -like "*grp_UsersWithPriviledge*" -or
                    $_.principalDisplayName -like "*UsersWithPriviledge*" -or
                    $_.principalDisplayName -like "*grp_Users*Priviledge*"
                } | Select-Object -First 1

                if ($grpAssignment) {
                    $status.GrpPriviledgeFound    = $true
                    $status.GrpPriviledgeAssigned = $true
                    Write-Host "    [OK] grp_UsersWithPriviledge is assigned to MS Graph Command Line Tools" -ForegroundColor Green
                    Write-Host "         Group: $($grpAssignment.principalDisplayName)" -ForegroundColor Gray
                } else {
                    $status.GrpPriviledgeFound    = $false
                    $status.GrpPriviledgeAssigned = $false
                    Write-Host "    [X]  grp_UsersWithPriviledge NOT assigned to MS Graph Command Line Tools" -ForegroundColor Red
                    Write-Host "         Total assignments found: $($assignResp.Count)" -ForegroundColor Gray
                    $status.Errors += (Write-BWsError -Code "BWS-SEC-010" -Message "grp_UsersWithPriviledge not assigned to MS Graph Command Line Tools" -Detail "Assign in Entra admin centre > Enterprise apps > Microsoft Graph Command Line Tools > Users and groups." -CheckStep "[2/4] Enterprise app")
                }
            } else {
                Write-Host "    [!]  Microsoft Graph Command Line Tools service principal not found" -ForegroundColor Yellow
                $status.Warnings += "MS Graph Command Line Tools SP not found in tenant"
            }
        } catch {
            Write-Host "    [!]  Enterprise app check failed: $($_.Exception.Message)" -ForegroundColor Yellow
            $status.Warnings += "Enterprise app check: $($_.Exception.Message)"
        }

        # =================================================================
        # CHECK 3: PIM eligible roles for Partner accounts
        # GET /v1.0/roleManagement/directory/roleEligibilitySchedules
        #   ?$expand=roleDefinition,principal&$top=100
        # Source: learn.microsoft.com/graph/api/resources/privilegedidentitymanagementv3-overview
        #
        # Partner accounts are identified by UPN pattern: *@*.onmicrosoft.com
        # or displayName containing "partner" / "BWS" / "ext"
        # =================================================================
        Write-Host ""
        Write-Host "  [3/4] PIM eligible role assignments for partner accounts" -ForegroundColor Yellow
        try {
            $pimUri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedules" +
                      "?`$expand=roleDefinition,principal&`$top=100&`$select=id,principalId,roleDefinitionId,status,scheduleInfo"
            $pimResp = Invoke-BWsGraphRequest -Uri $pimUri -Method GET -ErrorAction Stop
            $pimSchedules = if ($pimResp.PSObject.Properties['value']) { $pimResp.value } else { @() }

            if ($pimSchedules.Count -gt 0) {
                Write-Host "    [i]  Total PIM eligible assignments: $($pimSchedules.Count)" -ForegroundColor Gray

                # -------------------------------------------------------
                # Partner account identification
                # -------------------------------------------------------
                # TODO (TA Integration): When a Technical Assessment (TA) is
                # loaded, replace the generic pattern matching below with the
                # exact partner account UPNs defined in the TA document.
                # The TA contains the authoritative list of partner accounts
                # (e.g. BCID-ADM-Partner@<tenant>.onmicrosoft.com).
                #
                # Current fallback: heuristic matching until TA is available.
                # Matches accounts by UPN/displayName patterns typical for
                # GDAP / external partner accounts:
                #   - UPN ending in @*.onmicrosoft.com
                #   - DisplayName containing "partner", "bws" or "ext"
                # -------------------------------------------------------
                $taPartnerUpns = @()  # TODO: populate from TA when available

                $partnerPim = @($pimSchedules | Where-Object {
                    $upn = if ($_.principal -and $_.principal.PSObject.Properties['userPrincipalName']) { $_.principal.userPrincipalName } else { "" }
                    $dn  = if ($_.principal -and $_.principal.PSObject.Properties['displayName']) { $_.principal.displayName } else { "" }

                    # Priority 1: exact TA UPN match (used once TA integration is done)
                    $taMatch = ($taPartnerUpns.Count -gt 0) -and ($taPartnerUpns -contains $upn)

                    # Priority 2: heuristic fallback (until TA is integrated)
                    $heuristic = $upn -like "*@*.onmicrosoft.com" -or
                                 $dn  -like "*partner*" -or
                                 $dn  -like "*bws*" -or
                                 $dn  -like "*ext*"

                    $taMatch -or $heuristic
                })

                if ($taPartnerUpns.Count -gt 0) {
                    Write-Host "    [i]  Partner accounts from TA: $($taPartnerUpns.Count) defined" -ForegroundColor Cyan
                } else {
                    Write-Host "    [i]  Partner detection: heuristic (no TA loaded - exact names from TA not yet available)" -ForegroundColor Gray
                }

                if ($partnerPim.Count -gt 0) {
                    Write-Host "    [OK] Partner accounts with PIM eligible roles: $($partnerPim.Count)" -ForegroundColor Green
                    $status.PartnerPimUsers = @($partnerPim | ForEach-Object {
                        $upn  = if ($_.principal) { $_.principal.userPrincipalName } else { "(unknown)" }
                        $role = if ($_.roleDefinition) { $_.roleDefinition.displayName } else { $_.roleDefinitionId }
                        @{ UPN=$upn; Role=$role; Status=$_.status }
                    })
                    if (-not $CompactView) {
                        foreach ($p in ($partnerPim | Select-Object -First 10)) {
                            $upn  = if ($p.principal) { $p.principal.userPrincipalName } else { "(unknown)" }
                            $role = if ($p.roleDefinition) { $p.roleDefinition.displayName } else { $p.roleDefinitionId }
                            Write-Host "         $($upn.PadRight(40)) -> $role" -ForegroundColor Cyan
                        }
                    }
                } else {
                    Write-Host "    [!]  No partner accounts found in PIM eligible schedules" -ForegroundColor Yellow
                    Write-Host "         (All PIM schedules are for internal accounts, or PIM is not configured)" -ForegroundColor Gray
                    $status.Warnings += (Write-BWsError -Code "BWS-SEC-020" -Message "No partner accounts found in PIM eligible roles (heuristic search)" -Detail "Check Entra admin centre > Roles and admins > PIM eligible assignments." -Severity "Warning" -CheckStep "[3/4] PIM partner")
                }

                # Also report all PIM-eligible admin accounts
                Write-Host "    [i]  All PIM eligible admin accounts: $($pimSchedules.Count)" -ForegroundColor Gray

            } else {
                Write-Host "    [!]  No PIM eligible role assignments found" -ForegroundColor Yellow
                $status.Warnings += "PIM has no eligible role assignments configured"
            }

            # Check for direct permanent role assignments (non-PIM) - these should be minimal
            Write-Host ""
            Write-Host "  [3b] Checking direct (non-PIM) permanent privileged role assignments..." -ForegroundColor Yellow
            try {
                # Get privileged role definitions first
                $privRoles = @("62e90394-69f5-4237-9190-012177145e10", # Global Administrator
                               "e8611ab8-c189-46e8-94e1-60213ab1f814", # Privileged Role Administrator  
                               "29232cdf-9323-42fd-ade2-1d097af3e4de", # Exchange Administrator
                               "f28a1f50-f6e7-4571-818b-6a12f2af6b6c")  # SharePoint Administrator
                $directUri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments" +
                             "?`$expand=principal,roleDefinition&`$top=200"
                $directResp = Invoke-BWsGraphRequest -Uri $directUri -Method GET -ErrorAction Stop
                $directAssigns = if ($directResp.PSObject.Properties['value']) { $directResp.value } else { @() }

                # Filter to privileged roles with non-PIM (direct) assignments
                $directAdmin = @($directAssigns | Where-Object {
                    $rId = $_.roleDefinitionId
                    $privRoles -contains $rId -and
                    ($_.directoryScopeId -eq "/" -or -not $_.directoryScopeId)
                })
                $status.DirectAdminCount = $directAdmin.Count
                if ($directAdmin.Count -gt 0) {
                    Write-Host "    [!]  Direct privileged role assignments found: $($directAdmin.Count)" -ForegroundColor Yellow
                    Write-Host "         (Best practice: use PIM eligible assignments instead of direct)" -ForegroundColor Gray
                    $status.DirectAdmins = @($directAdmin | ForEach-Object {
                        $dn = if ($_.principal) { $_.principal.displayName } else { $_.principalId }
                        $rn = if ($_.roleDefinition) { $_.roleDefinition.displayName } else { $_.roleDefinitionId }
                        @{ DisplayName=$dn; Role=$rn }
                    })
                    if (-not $CompactView) {
                        foreach ($d in ($directAdmin | Select-Object -First 5)) {
                            $dn = if ($d.principal) { $d.principal.displayName } else { $d.principalId }
                            $rn = if ($d.roleDefinition) { $d.roleDefinition.displayName } else { $d.roleDefinitionId }
                            Write-Host "         $($dn.PadRight(35)) -> $rn" -ForegroundColor Yellow
                        }
                    }
                    $status.Warnings += (Write-BWsError -Code "BWS-SEC-021" -Message "$($directAdmin.Count) direct permanent admin role(s) found (best practice: use PIM eligible)" -Detail "Accounts: $($status.DirectAdmins | ForEach-Object {$_.DisplayName} | Select-Object -First 5 | Join-String -Separator ', ')" -Severity "Warning" -CheckStep "[3/4] PIM direct roles")
                } else {
                    Write-Host "    [OK] No direct permanent privileged role assignments ([OK] all via PIM or no privileged users)" -ForegroundColor Green
                }
            } catch {
                if ($_.Exception.Message -match "403|Forbidden") {
                    Write-Host "    [!]  Direct role assignment check: access denied (requires RoleManagement.Read.All)" -ForegroundColor Yellow
                } else {
                    Write-Host "    [!]  Direct role assignment check: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

        } catch {
            if ($_.Exception.Message -match "403|Forbidden") {
                Write-Host "    [!]  PIM check requires RoleManagement.Read.Directory or Privileged Role Administrator" -ForegroundColor Yellow
                $status.Warnings += "PIM check: access denied (requires RoleEligibilitySchedule.Read.Directory)"
            } else {
                Write-Host "    [X]  PIM check failed: $($_.Exception.Message)" -ForegroundColor Red
                $status.Errors += (Write-BWsError -Code "BWS-GRAPH-030" -Message "PIM roleEligibilitySchedules query failed" -Detail $_.Exception.Message -CheckStep "[3/4] PIM")
            }
        }

        # =================================================================
        # CHECK 4: SICT Automation App removal
        # Requirement: "Check if all old privileged roles are removed and
        #               if SICT the automation app is removed"
        # GET /v1.0/servicePrincipals?$filter=displayName eq 'SICT'
        # Source: learn.microsoft.com/graph/api/serviceprincipal-list
        # =================================================================
        Write-Host ""
        Write-Host "  [4/4] SICT automation app removal check" -ForegroundColor Yellow
        try {
            # Search for SICT app (multiple possible names)
            $sictNames = @('SICT', 'SICT Automation', 'SICT App')
            $sictFound = $null
            foreach ($name in $sictNames) {
                $sictUri = "https://graph.microsoft.com/v1.0/servicePrincipals" +
                           "?`$filter=displayName eq '$name'&`$select=id,displayName,appId,accountEnabled"
                try {
                    $sictResp = Invoke-BWsGraphRequest -Uri $sictUri -Method GET -ErrorAction Stop
                    $sictSP   = if ($sictResp.PSObject.Properties['value']) { $sictResp.value | Select-Object -First 1 } else { $null }
                    if ($sictSP) { $sictFound = $sictSP; break }
                } catch {}
            }

            if ($sictFound) {
                $status.SictAppFound = $true
                $status.SictAppId    = $sictFound.id
                $status.SictAppName  = $sictFound.displayName
                $isEnabled = $sictFound.accountEnabled -eq $true
                if ($isEnabled) {
                    Write-Host "    [X]  SICT automation app FOUND and ENABLED - should be removed" -ForegroundColor Red
                    Write-Host "         Name: $($sictFound.displayName)  ID: $($sictFound.id.Substring(0,8))..." -ForegroundColor Gray
                    $status.Errors += (Write-BWsError -Code "BWS-SEC-030" -Message "SICT automation app is still present and enabled" -Detail "Remove in Entra admin centre > Enterprise apps. App: $($sictFound.displayName)" -CheckStep "[4/4] SICT app")
                } else {
                    Write-Host "    [!]  SICT automation app found but DISABLED" -ForegroundColor Yellow
                    Write-Host "         Name: $($sictFound.displayName) - consider removing it completely" -ForegroundColor Gray
                    $status.Warnings += "SICT automation app is present (disabled) - consider removing"
                }
            } else {
                Write-Host "    [OK] SICT automation app not found (removed or never existed)" -ForegroundColor Green
            }
        } catch {
            Write-Host "    [!]  SICT app check failed: $($_.Exception.Message)" -ForegroundColor Yellow
            $status.Warnings += "SICT app check: $($_.Exception.Message)"
        }

    } catch {
        $status.Errors += "Unexpected error: $($_.Exception.Message)"
    }

    # -- Summary -----------------------------------------------------------
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  SECURITY CONFIGURATION SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  MDM Enrollment Scope  : $(if ($status.MdmAppliesTo) { $status.MdmAppliesTo.ToUpper() } else { 'Unknown' })" -ForegroundColor $(if ($status.MdmConfigOk) { "Green" } elseif ($status.MdmAppliesTo) { "Yellow" } else { "Gray" })
    Write-Host "  Graph CLI Group Assign: $(if ($status.GrpPriviledgeAssigned) { 'OK - grp_UsersWithPriviledge assigned' } elseif ($status.GraphCliSpFound) { 'MISSING - not assigned' } else { 'SP not found' })" -ForegroundColor $(if ($status.GrpPriviledgeAssigned) { "Green" } elseif ($status.GraphCliSpFound) { "Red" } else { "Yellow" })
    Write-Host "  PIM Partner Accounts  : $($status.PartnerPimUsers.Count) with eligible roles" -ForegroundColor $(if ($status.PartnerPimUsers.Count -gt 0) { "Green" } else { "Yellow" })
    Write-Host "  Direct Perm. Admin    : $($status.DirectAdminCount)" -ForegroundColor $(if ($status.DirectAdminCount -eq 0) { "Green" } else { "Yellow" })
    Write-Host "  SICT App Removed      : $(if (-not $status.SictAppFound) { 'Yes ([OK])' } else { 'NO - still present ([X])' })" -ForegroundColor $(if (-not $status.SictAppFound) { "Green" } else { "Red" })
    Write-Host "  Errors                : $($status.Errors.Count)" -ForegroundColor $(if ($status.Errors.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Warnings              : $($status.Warnings.Count)" -ForegroundColor $(if ($status.Warnings.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""

    if (-not $CompactView) {
        if ($status.Errors.Count -gt 0) {
            Write-Host "  ERRORS:" -ForegroundColor Red
            $status.Errors   | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
            Write-Host ""
        }
        if ($status.Warnings.Count -gt 0) {
            Write-Host "  WARNINGS:" -ForegroundColor Yellow
            $status.Warnings | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
            Write-Host ""
        }
    }

    return @{ Status = $status; CheckPerformed = $true }
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
        [object]$SecurityResults,
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
                <li><a href="#security">&rarr; Security Configuration</a></li>
                <li><a href="#errorlog">&rarr; Error Log</a></li>
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
        $intuneClass = if ($IntuneResults.AllFound -eq $true) { "success" } elseif ($IntuneResults.MissingCount -gt 0) { "error" } else { "warning" }
        $html += @"
                    <div class="summary-card $intuneClass">
                        <h3>Intune Policies</h3>
                        <div class="value">$($IntuneResults.FoundCount)/$($IntuneResults.Total)</div>
                        <p>Found</p>
                    </div>
"@
    }

    # 3. Entra ID Connect
    if ($EntraIDResults -and $EntraIDResults.CheckPerformed) {
        $entraClass = if ($EntraIDResults.Status.SyncEnabled -and $EntraIDResults.Status.SyncActive -and $EntraIDResults.Status.Errors.Count -eq 0) { "success" } `
                      elseif ($EntraIDResults.Status.SyncEnabled) { "warning" } else { "error" }
        $entraDetails = if ($EntraIDResults.Status.SyncActive -and $EntraIDResults.Status.PasswordSyncEnabled) { "Sync OK | PHS On" } `
                        elseif ($EntraIDResults.Status.SyncActive) { "Sync Active" } `
                        elseif ($EntraIDResults.Status.SyncEnabled) { "Sync Stale" } `
                        else { "Sync Not Enabled" }
        $html += @"
                    <div class="summary-card $entraClass">
                        <h3>Entra ID Sync</h3>
                        <div class="value">$(if ($EntraIDResults.Status.SyncActive) { '&#10003;' } else { '&#10007;' })</div>
                        <p>$entraDetails</p>
                    </div>
"@
    }

    # Security Config card
    if ($SecurityResults -and $SecurityResults.CheckPerformed) {
        $secCardClass = if ($SecurityResults.Status.Errors.Count -eq 0 -and $SecurityResults.Status.GrpPriviledgeAssigned -and -not $SecurityResults.Status.SictAppFound) { "success" } `
                        elseif ($SecurityResults.Status.Errors.Count -gt 0) { "error" } else { "warning" }
        $secCardDetail = if (-not $SecurityResults.Status.SictAppFound -and $SecurityResults.Status.GrpPriviledgeAssigned) { "MDM OK | Group OK" } `
                         elseif ($SecurityResults.Status.Errors.Count -gt 0) { "$($SecurityResults.Status.Errors.Count) Error(s)" } `
                         else { "Warnings present" }
        $html += @"
                    <div class="summary-card $secCardClass">
                        <h3>Security Config</h3>
                        <div class="value">$(if ($SecurityResults.Status.Errors.Count -eq 0) { '&#10003;' } else { '&#10007;' })</div>
                        <p>$secCardDetail</p>
                    </div>
"@
    }

    # 4. Hybrid Azure AD Join & Intune Connectors
    if ($IntuneConnResults -and $IntuneConnResults.CheckPerformed) {
        $connectorClass = if ($IntuneConnResults.Status.ActiveCount -gt 0 -and $IntuneConnResults.Status.DeprecatedCount -eq 0 -and $IntuneConnResults.Status.Errors.Count -eq 0) { "success" } `
                          elseif ($IntuneConnResults.Status.ActiveCount -gt 0) { "warning" } else { "error" }
        $connectorDetails = if ($IntuneConnResults.Status.DeprecatedCount -gt 0) {
            "Active - $($IntuneConnResults.Status.DeprecatedCount) deprecated!"
        } elseif ($IntuneConnResults.Status.ActiveCount -gt 0) {
            "Active ($($IntuneConnResults.Status.ActiveCount) connector(s))"
        } else {
            "Not Connected"
        }
        $html += @"
                    <div class="summary-card $connectorClass">
                        <h3>Intune AD Connector</h3>
                        <div class="value">$(if ($IntuneConnResults.Status.ActiveCount -gt 0) { '&#10003;' } else { '&#10007;' })</div>
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
        $azTotal   = $AzureResults.Total
        $azFound   = $AzureResults.Found.Count
        $azMissing = $AzureResults.Missing.Count
        $azErrors  = if ($AzureResults.Errors) { $AzureResults.Errors.Count } else { 0 }
        $html += @"
            <div class="section" id="azure">
                <h2><span class="section-icon">&#9729;</span>Azure Resources</h2>
                <ul class="info-list">
                    <li><strong>Total expected:</strong> $azTotal</li>
                    <li><strong>Found:</strong> <span class="$(if ($azFound -eq $azTotal) {'status-found'} else {'status-warning'})">$azFound</span></li>
                    <li><strong>Missing:</strong> <span class="$(if ($azMissing -eq 0) {'status-found'} else {'status-error'})">$azMissing</span></li>
                    $(if ($azErrors -gt 0) { "<li><strong>Query errors:</strong> <span class='status-warning'>$azErrors</span></li>" })
                </ul>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Category</th>
                            <th>Sub-Category</th>
                            <th>Resource Name</th>
                            <th>Resource Group</th>
                            <th>Location</th>
                        </tr>
                    </thead>
                    <tbody>
"@
        foreach ($resource in ($AzureResults.Found | Sort-Object Category, SubCategory)) {
            $html += "                        <tr><td><span class='status-icon status-found'>&#10003;</span></td><td>$($resource.Category)</td><td style='color:#aaa;'>$($resource.SubCategory)</td><td><strong>$($resource.Name)</strong></td><td style='color:#aaa;'>$($resource.ResourceGroupName)</td><td style='color:#aaa;'>$($resource.Location)</td></tr>`n"
        }
        foreach ($resource in ($AzureResults.Missing | Sort-Object Category, SubCategory)) {
            $html += "                        <tr style='background:rgba(220,60,60,0.1);'><td><span class='status-icon status-missing'>&#10007;</span></td><td>$($resource.Category)</td><td style='color:#aaa;'>$($resource.SubCategory)</td><td><strong>$($resource.Name)</strong></td><td colspan='2' style='color:#888;'><em>Not found in subscription</em></td></tr>`n"
        }
        if ($AzureResults.Errors) {
            foreach ($resource in $AzureResults.Errors) {
                $html += "                        <tr style='background:rgba(220,160,0,0.08);'><td><span class='status-icon status-warning'>&#9888;</span></td><td>$($resource.Category)</td><td style='color:#aaa;'>$($resource.SubCategory)</td><td><strong>$($resource.Name)</strong></td><td colspan='2' style='color:#f0a050;font-size:11px;'>$($resource.ErrorMessage)</td></tr>`n"
            }
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
                <ul class="info-list">
                    <li><strong>Required:</strong> $($IntuneResults.Total)</li>
                    <li><strong>Found:</strong> <span class="$(if ($IntuneResults.FoundCount -eq $IntuneResults.Total) { 'status-found' } else { 'status-warning' })">$($IntuneResults.FoundCount)</span></li>
                    <li><strong>Missing:</strong> <span class="$(if ($IntuneResults.MissingCount -eq 0) { 'status-found' } else { 'status-error' })">$($IntuneResults.MissingCount)</span></li>
                    <li><strong>All policies found:</strong> $(if ($IntuneResults.AllFound) { '<span class="status-found">&#10003; Yes</span>' } else { '<span class="status-error">&#10007; No</span>' })</li>
                    $(if ($IntuneResults.Errors.Count -gt 0) { "<li><strong>Retrieval errors:</strong> <span class='status-warning'>$($IntuneResults.Errors.Count)</span></li>" })
                </ul>
                <table>
                    <thead>
                        <tr>
                            <th>Found</th>
                            <th>Policy Name</th>
                            <th>Endpoint</th>
                            <th>Policy ID</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($policy in $IntuneResults.PolicyResults) {
            $statusIcon  = if ($policy.Found) { '<span class="status-icon status-found">&#10003;</span>' } else { '<span class="status-icon status-missing">&#10007;</span>' }
            $rowStyle    = if (-not $policy.Found) { ' style="background:#1e0a0a;"' } else { '' }
            $endpointTxt = if ($policy.Endpoint) { $policy.Endpoint } else { '-' }
            $idTxt       = if ($policy.PolicyId)  { "<span style='font-size:11px;color:#777;'>$($policy.PolicyId)</span>" } else { '-' }
            $html += @"
                        <tr$rowStyle>
                            <td>$statusIcon</td>
                            <td>$($policy.PolicyName)</td>
                            <td style="font-size:12px; color:#8ab4f8;">$endpointTxt</td>
                            <td>$idTxt</td>
                        </tr>
"@
        }

        $html += @"
                    </tbody>
                </table>
"@
        if ($IntuneResults.Errors.Count -gt 0) {
            $html += @"
                <h3>Retrieval Errors:</h3>
                <ul class="info-list">
"@
            foreach ($e in $IntuneResults.Errors) {
                $html += "                    <li><span class='status-warning'>&#9888;</span> $e</li>`n"
            }
            $html += "                </ul>`n"
        }

        $html += "            </div>`n"
    }

    # BWS Software Packages Section
    if ($SoftwareResults -and $SoftwareResults.CheckPerformed) {
        $sw = $SoftwareResults.Status
        $swFound   = $sw.Found.Count
        $swMissing = $sw.Missing.Count
        $swTotal   = $swFound + $swMissing
        $html += @"
            <div class="section" id="software">
                <h2><span class="section-icon">&#128230;</span>BWS Standard Software Packages</h2>
                <ul class="info-list">
                    <li><strong>Total Required:</strong> $swTotal</li>
                    <li><strong>Found:</strong> <span class="$(if ($swFound -eq $swTotal) {'status-found'} else {'status-warning'})">$swFound</span></li>
                    <li><strong>Missing:</strong> <span class="$(if ($swMissing -eq 0) {'status-found'} else {'status-error'})">$swMissing</span></li>
                    <li><strong>Mac Clients Detected:</strong> $(if ($sw.HasMacClients) { "<span class='status-found'>&#10003; Yes ($($sw.MacDeviceCount) devices)</span>" } else { '<span style="color:#888;">No macOS devices in Intune</span>' })</li>
                    $(if ($sw.HasMacClients) {
                        "<li><strong>BeyondTrust (Mac):</strong> $(if ($sw.BeyondTrustOk) { '<span class=''status-found''>&#10003; Deployed</span>' } else { '<span class=''status-error''>&#10007; NOT deployed</span>' })</li>"
                    })
                    $(if ($sw.HasMacClients) {
                        "<li><strong>Printix (Mac):</strong> $(if ($sw.PrintixOk) { '<span class=''status-found''>&#10003; Deployed</span>' } else { '<span class=''status-error''>&#10007; NOT deployed</span>' })</li>"
                    })
                </ul>
                <h3 style="margin-top:14px;">Package Details:</h3>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Required Package</th>
                            <th>Found As</th>
                            <th>Match Type</th>
                        </tr>
                    </thead>
                    <tbody>
"@
        foreach ($app in $SoftwareResults.Status.Found) {
            $html += "                        <tr><td><span class='status-icon status-found'>&#10003;</span></td><td>$($app.SoftwareName)</td><td>$($app.ActualName)</td><td>$($app.MatchType)</td></tr>`n"
        }
        foreach ($app in $SoftwareResults.Status.Missing) {
            $html += "                        <tr style='background:rgba(220,60,60,0.08);'><td><span class='status-icon status-missing'>&#10007;</span></td><td><strong>$($app.SoftwareName)</strong></td><td><em style='color:#888;'>Not found in Intune</em></td><td>-</td></tr>`n"
        }
        $html += "                    </tbody></table>`n"
        if ($SoftwareResults.Status.Errors.Count -gt 0) {
            $html += "                <h3>Errors:</h3><ul class='info-list'>`n"
            foreach ($e in $SoftwareResults.Status.Errors) { $html += "                    <li><span class='status-error'>&#9888;</span> $e</li>`n" }
            $html += "                </ul>`n"
        }
        $html += "            </div>`n"
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
"@
        if ($TeamsResults.Status.Errors -and $TeamsResults.Status.Errors.Count -gt 0) {
            $html += "                <h3>Errors:</h3><ul class='info-list'>`n"
            foreach ($e in $TeamsResults.Status.Errors) {
                $html += "                    <li><span class='status-error'>&#9888;</span> $e</li>`n"
            }
            $html += "                </ul>`n"
        }
        $html += "            </div>`n"
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
        $entraClass = if ($EntraIDResults.Status.SyncEnabled -and $EntraIDResults.Status.SyncActive -and $EntraIDResults.Status.Errors.Count -eq 0) { "success" } `
                      elseif ($EntraIDResults.Status.SyncEnabled) { "warning" } else { "error" }
        $html += @"
            <div class="section" id="entra">
                <h2><span class="section-icon">&#128279;</span>Entra ID Connect</h2>
                <ul class="info-list">
                    $(if ($EntraIDResults.Status.TenantDisplayName) { "<li><strong>Tenant Name:</strong> $($EntraIDResults.Status.TenantDisplayName)</li>" })
                    <li><strong>Sync Enabled:</strong>
                        $(if ($EntraIDResults.Status.SyncEnabled) { '<span class="status-found">&#10003; Yes</span>' } else { '<span class="status-missing">&#10007; No</span>' })</li>
                    <li><strong>Sync Active (&lt; 3h):</strong>
                        $(if ($EntraIDResults.Status.SyncActive) { '<span class="status-found">&#10003; Yes</span>' } else { '<span class="status-error">&#10007; No</span>' })</li>
                    <li><strong>Sync Age Status:</strong>
                        $(if ($EntraIDResults.Status.SyncAgeStatus -eq "OK") { '<span class="status-found">OK</span>' } elseif ($EntraIDResults.Status.SyncAgeStatus -eq "Warning") { '<span class="status-warning">Warning</span>' } elseif ($EntraIDResults.Status.SyncAgeStatus -eq "Error") { '<span class="status-error">Error</span>' } else { '<span class="status-missing">Unknown</span>' })</li>
                    $(if ($EntraIDResults.Status.LastSyncDateTime) { "<li><strong>Last Sync:</strong> $($EntraIDResults.Status.LastSyncDateTime)$(if ($EntraIDResults.Status.SyncAgeMinutes) { ' (' + $EntraIDResults.Status.SyncAgeMinutes + ' min ago)' })</li>" })
                    $(if ($EntraIDResults.Status.LastPasswordSyncDateTime) { "<li><strong>Last Password Sync:</strong> $($EntraIDResults.Status.LastPasswordSyncDateTime)</li>" })
                    $(if ($EntraIDResults.Status.OnPremDomains -and $EntraIDResults.Status.OnPremDomains.Count -gt 0) { "<li><strong>On-prem Domains:</strong> $($EntraIDResults.Status.OnPremDomains -join ', ')</li>" })
                    $(if ($EntraIDResults.Status.VerifiedDomains -and $EntraIDResults.Status.VerifiedDomains.Count -gt 0) { "<li><strong>Verified Domains:</strong> $($EntraIDResults.Status.VerifiedDomains -join ', ')</li>" })
                </ul>
"@
        # Feature flags table
        if ($EntraIDResults.Status.PasswordSyncEnabled -ne $null) {
            $html += @"
                <h3>Sync Feature Flags <small style='color:#888;font-size:12px;'>(onPremisesDirectorySynchronization.features)</small></h3>
                <table style='width:100%;border-collapse:collapse;font-size:13px;'>
                    <thead><tr style='background:#1a2540;color:#8ab4f8;'>
                        <th style='padding:6px 10px;text-align:left;'>Feature</th>
                        <th style='padding:6px 10px;text-align:left;'>Status</th>
                    </tr></thead>
                    <tbody>
"@
            $featureRows = @(
                @{ Label="Password Hash Sync";     Val=$EntraIDResults.Status.PasswordSyncEnabled      },
                @{ Label="Password Writeback";     Val=$EntraIDResults.Status.PasswordWritebackEnabled },
                @{ Label="Device Writeback";       Val=$EntraIDResults.Status.DeviceWritebackEnabled   },
                @{ Label="Group Writeback";        Val=$EntraIDResults.Status.GroupWritebackEnabled    },
                @{ Label="User Writeback";         Val=$EntraIDResults.Status.UserWritebackEnabled     },
                @{ Label="Directory Extensions";   Val=$EntraIDResults.Status.DirectoryExtensionsEnabled }
            )
            foreach ($fr in $featureRows) {
                $fIcon  = if ($fr.Val -eq $true) { '<span class="status-found">&#10003; Enabled</span>' } `
                          elseif ($fr.Val -eq $false) { '<span style="color:#666;">&#9675; Disabled</span>' } `
                          else { '<span class="status-missing">Unknown</span>' }
                $html += "                    <tr><td style='padding:5px 10px;'>$($fr.Label)</td><td style='padding:5px 10px;'>$fIcon</td></tr>`n"
            }
            $html += "                    </tbody></table>`n"
        }
        # Errors and warnings
        if ($EntraIDResults.Status.Errors.Count -gt 0 -or $EntraIDResults.Status.Warnings.Count -gt 0) {
            $html += "                <h3>Errors &amp; Warnings:</h3><ul class='info-list'>`n"
            foreach ($e in $EntraIDResults.Status.Errors)   { $html += "                    <li><span class='status-error'>&#9888; Error:</span> $e</li>`n" }
            foreach ($w in $EntraIDResults.Status.Warnings) { $html += "                    <li><span class='status-warning'>&#9888; Warning:</span> $w</li>`n" }
            $html += "                </ul>`n"
        }
        $html += "            </div>`n"
    }

    # Hybrid Join Section
    if ($IntuneConnResults -and $IntuneConnResults.CheckPerformed) {
        $connSectionClass = if ($IntuneConnResults.Status.ActiveCount -gt 0 -and $IntuneConnResults.Status.DeprecatedCount -eq 0 -and $IntuneConnResults.Status.Errors.Count -eq 0) { "section-ok" } `
                            elseif ($IntuneConnResults.Status.ActiveCount -gt 0) { "section-warning" } else { "section-error" }
        $html += @"
            <div class="section" id="hybrid">
                <h2><span class="section-icon">&#128272;</span>Intune Connector for Active Directory</h2>
                <ul class="info-list">
                    <li><strong>Connector state:</strong>
                        $(if ($IntuneConnResults.Status.ActiveCount -gt 0) { '<span class="status-found">&#10003; Active</span>' } else { '<span class="status-error">&#10007; Not Connected</span>' })</li>
                    <li><strong>Total configured:</strong> $($IntuneConnResults.Status.ConnectorCount)</li>
                    <li><strong>Active:</strong> <span class="$(if ($IntuneConnResults.Status.ActiveCount -gt 0) { 'status-found' } else { 'status-error' })">$($IntuneConnResults.Status.ActiveCount)</span></li>
                    <li><strong>Inactive:</strong> <span class="$(if ($IntuneConnResults.Status.InactiveCount -gt 0) { 'status-warning' } else { 'status-found' })">$($IntuneConnResults.Status.InactiveCount)</span></li>
                    <li><strong>Deprecated versions:</strong>
                        $(if ($IntuneConnResults.Status.DeprecatedCount -gt 0) { "<span class='status-error'>&#9888; $($IntuneConnResults.Status.DeprecatedCount) - UPDATE REQUIRED (min: 6.2501.2000.5)</span>" } else { '<span class="status-found">&#10003; None</span>' })</li>
                    <li><strong>On-premises sync:</strong>
                        $(if ($IntuneConnResults.Status.OnPremSyncEnabled) { '<span class="status-found">&#10003; Enabled</span>' } else { '<span class="status-missing">Not enabled / Unknown</span>' })</li>
                    $(if ($IntuneConnResults.Status.LastSyncDateTime) { "<li><strong>Last sync:</strong> $($IntuneConnResults.Status.LastSyncDateTime)</li>" })
                    <li><strong>Errors:</strong> $(if ($IntuneConnResults.Status.Errors.Count -eq 0) { '<span class="status-found">0</span>' } else { "<span class='status-error'>$($IntuneConnResults.Status.Errors.Count)</span>" })</li>
                    <li><strong>Warnings:</strong> $(if ($IntuneConnResults.Status.Warnings.Count -eq 0) { '<span class="status-found">0</span>' } else { "<span class='status-warning'>$($IntuneConnResults.Status.Warnings.Count)</span>" })</li>
                </ul>
"@

        # Connector details table
        $ndesConnectors = $IntuneConnResults.Status.Connectors | Where-Object { $_.State -ne "azure-vm" }
        if ($ndesConnectors -and @($ndesConnectors).Count -gt 0) {
            $html += @"
                <h3>Connector Details:</h3>
                <table style="width:100%; border-collapse:collapse; margin-top:8px; font-size:13px;">
                    <tr style="background:#1a2540; color:#8ab4f8;">
                        <th style="padding:6px 10px; text-align:left;">Name</th>
                        <th style="padding:6px 10px; text-align:left;">Machine</th>
                        <th style="padding:6px 10px; text-align:left;">State</th>
                        <th style="padding:6px 10px; text-align:left;">Version</th>
                        <th style="padding:6px 10px; text-align:left;">Last Check-in</th>
                        <th style="padding:6px 10px; text-align:left;">Enrolled</th>
                    </tr>
"@
            foreach ($c in $ndesConnectors) {
                $rowBg    = if ($c.State -eq "active") { "#0d1b12" } else { "#1e0a0a" }
                $stCls    = if ($c.State -eq "active") { "status-found" } else { "status-error" }
                $verCls   = if ($c.IsDeprecated) { "status-error" } else { "status-found" }
                $verText  = if ($c.IsDeprecated) { "$($c.Version) &#9888;" } else { $c.Version }
                $html += @"
                    <tr style="background:$rowBg;">
                        <td style="padding:5px 10px;">$($c.DisplayName)</td>
                        <td style="padding:5px 10px; color:#aaa;">$($c.MachineName)</td>
                        <td style="padding:5px 10px;"><span class="$stCls">$($c.State.ToUpper())</span></td>
                        <td style="padding:5px 10px;"><span class="$verCls">$verText</span></td>
                        <td style="padding:5px 10px; color:#aaa;">$(if ($c.LastCheckin) { $c.LastCheckin } else { '-' }) <span class="$( if ($c.Freshness -like '*STALE*') {'status-error'} elseif ($c.Freshness -like '*Warning*') {'status-warning'} else {'status-found'} )">$($c.Freshness)</span></td>
                        <td style="padding:5px 10px; color:#aaa;">$(if ($c.EnrolledDate) { $c.EnrolledDate } else { '-' })</td>
                    </tr>
"@
            }
            $html += "</table>`n"
        }

        # Azure VM servers (AD/Sync)
        $azVMs = $IntuneConnResults.Status.Connectors | Where-Object { $_.State -eq "azure-vm" }
        if ($azVMs -and @($azVMs).Count -gt 0) {
            $html += @"
                <h3>AD / Sync Server(s) in Azure:</h3>
                <ul class="info-list">
"@
            foreach ($vm in $azVMs) {
                $html += @"
                    <li><span class="status-found">&#10003;</span> <strong>$($vm.DisplayName)</strong></li>
"@
            }
            $html += "</ul>`n"
        }

        # Errors and Warnings
        if ($IntuneConnResults.Status.Errors.Count -gt 0 -or $IntuneConnResults.Status.Warnings.Count -gt 0) {
            $html += @"
                <h3>Errors &amp; Warnings:</h3>
                <ul class="info-list">
"@
            foreach ($e in $IntuneConnResults.Status.Errors) {
                $html += "                    <li><span class='status-error'>&#9888; Error:</span> $e</li>`n"
            }
            foreach ($w in $IntuneConnResults.Status.Warnings) {
                $html += "                    <li><span class='status-warning'>&#9888; Warning:</span> $w</li>`n"
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
        
        if ($DefenderResults.Status.Errors -and $DefenderResults.Status.Errors.Count -gt 0) {
            $html += "                <h3>Errors:</h3><ul class='info-list'>`n"
            foreach ($e in $DefenderResults.Status.Errors) {
                $html += "                    <li><span class='status-error'>&#9888;</span> $e</li>`n"
            }
            $html += "                </ul>`n"
        }
        $html += "            </div>`n"
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

    # ---------------------------------------------------------------
    # Security Configuration Section
    # ---------------------------------------------------------------
    if ($SecurityResults -and $SecurityResults.CheckPerformed) {
        $secClass = if ($SecurityResults.Status.GrpPriviledgeAssigned -and
                        -not $SecurityResults.Status.SictAppFound -and
                        $SecurityResults.Status.MdmConfigOk -and
                        $SecurityResults.Status.Errors.Count -eq 0) { "success" } `
                    elseif ($SecurityResults.Status.Errors.Count -gt 0) { "error" } `
                    else { "warning" }
        $html += @"
            <div class="section" id="security">
                <h2><span class="section-icon">&#128274;</span>Security Configuration</h2>
                <ul class="info-list">
                    <li><strong>MDM Enrollment Scope:</strong>
                        $(if ($SecurityResults.Status.MdmConfigOk) { '<span class="status-found">&#10003; All Users</span>' } elseif ($SecurityResults.Status.MdmAppliesTo) { "<span class='status-warning'>&#9888; $($SecurityResults.Status.MdmAppliesTo)</span>" } else { '<span class="status-missing">Unknown</span>' })</li>
                    $(if ($SecurityResults.Status.MdmDiscoveryUrl) { "<li><strong>MDM Discovery URL:</strong> $($SecurityResults.Status.MdmDiscoveryUrl)</li>" })
                    <li><strong>grp_UsersWithPriviledge (MS Graph CLI):</strong>
                        $(if ($SecurityResults.Status.GrpPriviledgeAssigned) { '<span class="status-found">&#10003; Assigned</span>' } else { '<span class="status-error">&#10007; NOT assigned</span>' })</li>
                    <li><strong>PIM Partner Eligible Roles:</strong>
                        $(if ($SecurityResults.Status.PartnerPimUsers.Count -gt 0) { "<span class='status-found'>&#10003; $($SecurityResults.Status.PartnerPimUsers.Count) account(s)</span>" } else { '<span class="status-warning">&#9888; None found</span>' })</li>
                    <li><strong>Direct Permanent Admin Assignments:</strong>
                        $(if ($SecurityResults.Status.DirectAdminCount -eq 0) { '<span class="status-found">&#10003; None (OK)</span>' } else { "<span class='status-warning'>&#9888; $($SecurityResults.Status.DirectAdminCount) direct assignment(s)</span>" })</li>
                    <li><strong>MS Graph CLI Service Principal:</strong>
                        $(if ($SecurityResults.Status.GraphCliSpFound) { '<span class="status-found">&#10003; Found in Entra</span>' } else { '<span class="status-warning">&#9888; Not found</span>' })</li>
                    <li><strong>SICT Automation App:</strong>
                        $(if (-not $SecurityResults.Status.SictAppFound) { '<span class="status-found">&#10003; Removed / Not present</span>' } else { "<span class='status-error'>&#10007; STILL PRESENT$(if ($SecurityResults.Status.SictAppName) { ': ' + $SecurityResults.Status.SictAppName })</span>" })</li>
                </ul>
"@
        # Partner PIM Eligible Roles table
        if ($SecurityResults.Status.PartnerPimUsers -and $SecurityResults.Status.PartnerPimUsers.Count -gt 0) {
            $html += @"
                <h3>PIM Eligible Partner Accounts:</h3>
                <table style='width:100%;border-collapse:collapse;font-size:13px;'>
                    <thead><tr style='background:#1a2540;color:#8ab4f8;'>
                        <th style='padding:6px 10px;text-align:left;'>Account</th>
                        <th style='padding:6px 10px;text-align:left;'>UPN</th>
                        <th style='padding:6px 10px;text-align:left;'>Role</th>
                    </tr></thead><tbody>
"@
            foreach ($p in $SecurityResults.Status.PartnerPimUsers) {
                $html += "                    <tr><td style='padding:5px 10px;'>$($p.DisplayName)</td><td style='padding:5px 10px;color:#aaa;'>$($p.UPN)</td><td style='padding:5px 10px;'>$($p.Role)</td></tr>`n"
            }
            $html += "                    </tbody></table>`n"
        }
        if ($SecurityResults.Status.DirectAdmins.Count -gt 0) {
            $html += @"
                <h3>Direct Permanent Privileged Role Assignments:</h3>
                <table style='width:100%;border-collapse:collapse;font-size:13px;'>
                    <thead><tr style='background:#1a2540;color:#8ab4f8;'>
                        <th style='padding:6px 10px;text-align:left;'>Account</th>
                        <th style='padding:6px 10px;text-align:left;'>Role</th>
                    </tr></thead><tbody>
"@
            foreach ($d in $SecurityResults.Status.DirectAdmins) {
                $html += "                    <tr><td style='padding:5px 10px;'>$($d.DisplayName)</td><td style='padding:5px 10px;color:#f0a050;'>$($d.Role)</td></tr>`n"
            }
            $html += "                    </tbody></table>`n"
        }
        if ($SecurityResults.Status.Errors.Count -gt 0 -or $SecurityResults.Status.Warnings.Count -gt 0) {
            $html += "                <h3>Issues:</h3><ul class='info-list'>`n"
            foreach ($e in $SecurityResults.Status.Errors)   { $html += "                    <li><span class='status-error'>Error: $e</span></li>`n" }
            foreach ($w in $SecurityResults.Status.Warnings) { $html += "                    <li><span class='status-warning'>Warning: $w</span></li>`n" }
            $html += "                </ul>`n"
        }
        $html += "            </div>`n"
    }

    # -- Error Log section in HTML report -----------------------------------
    if ($script:BWS_ErrorLog -and $script:BWS_ErrorLog.Count -gt 0) {
        $errItems   = @($script:BWS_ErrorLog | Where-Object { $_.Severity -eq "Error" })
        $warnItems  = @($script:BWS_ErrorLog | Where-Object { $_.Severity -eq "Warning" })
        $html += @"
            <div class="section" id="errorlog">
                <h2><span class="section-icon">&#128203;</span>Error Log  ($($script:BWS_ErrorLog.Count) issue(s))</h2>
                <p style='color:var(--muted);font-size:13px;margin:0 0 12px;'>
                    Errors: $($errItems.Count)&nbsp;&nbsp;&bull;&nbsp;&nbsp;Warnings: $($warnItems.Count)
                    &nbsp;&nbsp;&bull;&nbsp;&nbsp;Use <code>-SupportBundle</code> to export as JSON.
                </p>
                <table style='width:100%;border-collapse:collapse;font-size:13px;'>
                    <thead><tr style='background:#1a2540;color:#8ab4f8;'>
                        <th style='padding:6px 10px;text-align:left;width:160px;'>Code</th>
                        <th style='padding:6px 10px;text-align:left;width:80px;'>Severity</th>
                        <th style='padding:6px 10px;text-align:left;width:180px;'>Function</th>
                        <th style='padding:6px 10px;text-align:left;'>Message</th>
                        <th style='padding:6px 10px;text-align:left;width:140px;'>Time</th>
                    </tr></thead><tbody>
"@
        foreach ($e in ($script:BWS_ErrorLog | Sort-Object Severity, Timestamp)) {
            $bgCol  = switch ($e.Severity) {
                "Error"   { "rgba(220,60,60,0.12)" }
                "Warning" { "rgba(220,160,30,0.10)" }
                default   { "transparent" }
            }
            $sevCls = switch ($e.Severity) {
                "Error"   { "status-error" }
                "Warning" { "status-warning" }
                default   { "" }
            }
            $html += "                    <tr style='background:$bgCol;border-bottom:1px solid rgba(255,255,255,0.05);'>"
            $html += "<td style='padding:5px 10px;font-family:monospace;'><strong>$($e.Code)</strong></td>"
            $html += "<td style='padding:5px 10px;'><span class='$sevCls'>$($e.Severity)</span></td>"
            $html += "<td style='padding:5px 10px;color:#aaa;'>$($e.Function)</td>"
            $html += "<td style='padding:5px 10px;'>$($e.Message)"
            if ($e.Detail) {
                $html += "<br><span style='font-size:11px;color:#888;'>$($e.Detail.Substring(0,[Math]::Min(100,$e.Detail.Length)))</span>"
            }
            $html += "</td><td style='padding:5px 10px;color:#888;'>$($e.Timestamp)</td></tr>`n"
        }
        $html += "                    </tbody></table></div>`n"
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
    $groupBoxChecks.Size = New-Object System.Drawing.Size(300, 280)
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
    
    $chkSecurity = New-Object System.Windows.Forms.CheckBox
    $chkSecurity.Location = New-Object System.Drawing.Point(15, 250)
    $chkSecurity.Size = New-Object System.Drawing.Size(280, 20)
    $chkSecurity.Text = "Security Config Check (MDM/PIM/SICT)"
    $chkSecurity.Checked = $true
    $groupBoxChecks.Controls.Add($chkSecurity)

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
    
    # Prerequisites Button
    $btnPrereq = New-Object System.Windows.Forms.Button
    $btnPrereq.Location = New-Object System.Drawing.Point(660, 180)
    $btnPrereq.Size     = New-Object System.Drawing.Size(150, 44)
    $btnPrereq.Text     = "Prerequisites"
    $btnPrereq.BackColor = [System.Drawing.Color]::FromArgb(60, 120, 200)
    $btnPrereq.ForeColor = [System.Drawing.Color]::White
    $btnPrereq.Font     = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $btnPrereq.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $form.Controls.Add($btnPrereq)

    # Clear Button
    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Location = New-Object System.Drawing.Point(660, 232)
    $btnClear.Size = New-Object System.Drawing.Size(150, 30)
    $btnClear.Text = "Clear Output"
    $form.Controls.Add($btnClear)

    $btnDiag = New-Object System.Windows.Forms.Button
    $btnDiag.Location  = New-Object System.Drawing.Point(660, 271)
    $btnDiag.Size      = New-Object System.Drawing.Size(150, 44)
    $btnDiag.Text      = "Diagnostics"
    $btnDiag.BackColor = [System.Drawing.Color]::FromArgb(80, 60, 140)
    $btnDiag.ForeColor = [System.Drawing.Color]::White
    $btnDiag.Font      = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $btnDiag.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $form.Controls.Add($btnDiag)

    $btnStop = New-Object System.Windows.Forms.Button
    $btnStop.Location  = New-Object System.Drawing.Point(660, 323)
    $btnStop.Size      = New-Object System.Drawing.Size(150, 44)
    $btnStop.Text      = "Stop"
    $btnStop.BackColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
    $btnStop.ForeColor = [System.Drawing.Color]::White
    $btnStop.Font      = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $btnStop.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnStop.Enabled   = $false
    $btnStop.Visible   = $false
    $form.Controls.Add($btnStop)

    # Stop Button Click
    $btnStop.Add_Click({
        $script:stopRequested = $true
        $btnStop.Enabled  = $false
        $btnStop.Text     = "Stopping..."
        $labelStatus.Text = "[!] Stop requested - finishing current check..."
        $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,160,60)
        $form.Refresh()
    })

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
    $textOutput.ReadOnly  = $true
    $textOutput.WordWrap  = $false
    $textOutput.MaxLength = 0
    $form.Controls.Add($textOutput)
    
    # Prerequisites Button Click  (Full flow: modules + login verification)
    $btnPrereq.Add_Click({
        $btnPrereq.Enabled = $false
        $btnRun.Enabled    = $false
        $labelStatus.Text  = "Running prerequisites..."
        $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(60,120,200)
        $form.Refresh()
        try {
            # Step 1: Module setup dialog
            $labelStatus.Text = "Step 1/2 - Checking modules..."
            $form.Refresh()
            $_preSkipParams = @{
                SkipSharePoint = (-not $chkSharePoint.Checked)
                SkipTeams      = (-not $chkTeams.Checked)
                SkipDefender   = (-not $chkDefender.Checked)
            }
            $preModResult = Show-ModuleSetupDialog -SkipParams $_preSkipParams

            # Step 2: GlobalAdmin login (Az + Graph + Teams + SharePoint)
            # One browser login covers all services.
            $labelStatus.Text = "Step 2/2 - GlobalAdmin login (Az + Teams + SharePoint)..."
            $form.Refresh()
            $preSubId  = $textSubID.Text.Trim()
            $loginOK   = $false
            $loginMsg  = ""

            # Set subscription if provided before login attempt
            if ($preSubId) {
                try { Set-BWsAzSubscription -SubscriptionId $preSubId } catch {}
            }

            try {
                # Connect-BWsGlobalAdmin handles Az + Teams + SharePoint in one call
                $loginOK = Connect-BWsGlobalAdmin `
                    -SharePointAdminUrl $SharePointUrl `
                    -SkipTeams:(-not $chkTeams.Checked) `
                    -SkipSharePoint:(-not $chkSharePoint.Checked)

                if ($loginOK) {
                    $preCtx = Get-BWsAzContext
                    $subLabel = if ($preCtx) { $preCtx.Subscription.Name } else { "Unknown" }
                    $loginMsg = "[OK] GlobalAdmin login complete - Sub: $subLabel - Teams: $(if ($script:TeamsConnected) {'OK'} else {'pending'}) - SPO: $(if ($script:SharePointConnected) {'OK'} else {'pending'})"
                } else {
                    $loginMsg = "[X] GlobalAdmin login failed - check browser window"
                }
            } catch {
                $loginMsg = "[X] Login error: $($_.Exception.Message)"
            }

            # Show result in status label
            $allOK = $preModResult.AllReady -and $loginOK -and $script:GlobalAdminConnected
            $statusText  = if ($allOK) { "[OK] Prerequisites complete - Ready to run checks" } `
                           else { "[!] Prerequisites complete with warnings - see details above" }
            $statusColor = if ($allOK) { [System.Drawing.Color]::FromArgb(60,200,80) } `
                           else { [System.Drawing.Color]::FromArgb(240,160,40) }
            $labelStatus.Text      = $statusText
            $labelStatus.ForeColor = $statusColor

            # Show login result in output box
            $textOutput.AppendText("`r`n--- Prerequisites ---`r`n")
            $textOutput.AppendText("Modules  : $($preModResult.OK) OK / $($preModResult.Failed) failed / $($preModResult.Warnings) warnings`r`n")
            $textOutput.AppendText("Login    : $loginMsg`r`n")
            $textOutput.AppendText("---------------------`r`n")
            $textOutput.ScrollToCaret()
            $form.Refresh()
        } catch {
            $labelStatus.Text      = "[X] Prerequisites failed: $($_.Exception.Message)"
            $labelStatus.ForeColor = [System.Drawing.Color]::Red
        } finally {
            $btnPrereq.Enabled = $true
            $btnRun.Enabled    = $true
        }
    })

    # Clear Button Click
    $btnClear.Add_Click({
        $textOutput.Clear()
        $labelStatus.Text = "Output cleared - Ready for next check"
        $labelStatus.ForeColor = [System.Drawing.Color]::Blue
        $progressBar.Value = 0
    })

    # Diagnostics Button Click
    # Writes directly to $textOutput - no separate dialog, no overlap possible
    $btnDiag.Add_Click({
        $btnDiag.Enabled   = $false
        $btnRun.Enabled    = $false
        $btnPrereq.Enabled = $false
        $labelStatus.Text      = "Collecting diagnostics..."
        $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(120, 80, 200)
        $form.Refresh()
        try {
            $diagData = Get-BWsDiagnostics -AsObject

            $textOutput.Clear()
            $progressBar.Value = 0

            # Helper: append a line with colour simulation via prefix markers
            # (plain TextBox has no per-character colour, so we use layout instead)
            function script:Write-DiagLine {
                param([string]$Text = "", [string]$Prefix = "")
                $script:textOutput.AppendText("$Prefix$Text`r`n")
                $script:textOutput.SelectionStart = $script:textOutput.Text.Length
                $script:textOutput.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
            }

            Write-DiagLine "============================================================"
            Write-DiagLine "  BWS DIAGNOSTICS  -  v$script:Version"
            Write-DiagLine "============================================================"

            $diagSections = [ordered]@{
                "POWERSHELL" = @(
                    @{ K='PS_Version';  L='Version'       },
                    @{ K='PS_Edition';  L='Edition'       },
                    @{ K='PS_Host';     L='Host'          },
                    @{ K='OS';          L='OS'            }
                )
                "AZURE / ENTRA ID" = @(
                    @{ K='AZ_LoggedIn';         L='Logged In'         },
                    @{ K='AZ_Account';          L='Account'           },
                    @{ K='AZ_AccountType';      L='Account Type'      },
                    @{ K='AZ_TenantId';         L='Tenant ID'         },
                    @{ K='AZ_TenantName';       L='Tenant Name'       },
                    @{ K='AZ_TenantDomain';     L='Tenant Domain'     },
                    @{ K='AZ_PrimaryDomain';    L='Primary Domain'    },
                    @{ K='AZ_SubscriptionId';   L='Subscription ID'   },
                    @{ K='AZ_SubscriptionName'; L='Subscription Name' },
                    @{ K='AZ_Environment';      L='Environment'       }
                )
                "SIGNED-IN USER" = @(
                    @{ K='USER_DisplayName'; L='Display Name' },
                    @{ K='USER_UPN';         L='UPN'          },
                    @{ K='USER_JobTitle';    L='Job Title'    },
                    @{ K='USER_Mail';        L='Mail'         }
                )
                "MODULES" = ($diagData.Keys |
                    Where-Object { $_ -like 'MOD_*' } |
                    Sort-Object |
                    ForEach-Object {
                        @{ K=$_; L=($_ -replace '^MOD_','') }
                    })
            }

            foreach ($secName in $diagSections.Keys) {
                Write-DiagLine ""
                Write-DiagLine "  [ $secName ]"
                Write-DiagLine "  ----------------------------------------------------"
                foreach ($row in $diagSections[$secName]) {
                    if ($diagData.Contains($row.K)) {
                        $lbl = $row.L.PadRight(24)
                        $val = $diagData[$row.K]
                        # Status prefix so user can scan quickly
                        $prefix = "    "
                        if ($val -like "*No -*" -or $val -like "*not installed*" -or
                            $val -like "*Not installed*" -or $val -like "*not available*") {
                            $prefix = " !  "
                        } elseif ($val -like "*Yes*" -or $val -like "*Loaded*" -or
                                  $val -like "*loaded*") {
                            $prefix = " OK "
                        }
                        Write-DiagLine "$lbl : $val" -Prefix $prefix
                    }
                }
            }

            Write-DiagLine ""
            Write-DiagLine "============================================================"
            Write-DiagLine "  END OF DIAGNOSTICS  -  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            Write-DiagLine "============================================================"

            $textOutput.SelectionStart = 0
            $textOutput.ScrollToCaret()
            $progressBar.Value = 100

            $labelStatus.Text      = "[OK] Diagnostics ready - scroll to read"
            $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(90, 200, 120)

        } catch {
            $textOutput.AppendText("`r`n[X] Diagnostics error: $($_.Exception.Message)`r`n")
            $labelStatus.Text      = "[X] Diagnostics failed"
            $labelStatus.ForeColor = [System.Drawing.Color]::Red
        } finally {
            Remove-Item Function:Write-DiagLine -ErrorAction SilentlyContinue
            $btnDiag.Enabled   = $true
            $btnRun.Enabled    = $true
            $btnPrereq.Enabled = $true
        }
    })
    # Run Button Click
    $btnRun.Add_Click({
        # Clean up any leftover Write-Host override from a previous failed run
        Remove-Item Function:\Write-Host -ErrorAction SilentlyContinue
        $textOutput.Clear()
        $progressBar.Value = 0
        $labelStatus.Text = "Initializing check..."
        $labelStatus.ForeColor = [System.Drawing.Color]::Orange
        $script:stopRequested = $false
        $btnRun.Enabled    = $false
        $btnPrereq.Enabled = $false
        $btnClear.Enabled  = $false
        $btnDiag.Enabled   = $false
        $btnStop.Enabled   = $true
        $btnStop.Visible   = $true
        $btnStop.Text      = "Stop"
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
        $runSecurity     = $chkSecurity.Checked
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

        try {
            $progressBar.Value = 5
            
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
                # DoEvents allows UI to repaint but is re-entrant safe here because
                # $btnRun, $btnPrereq and $btnClear are all disabled during this block
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            $azureResults        = $null
            $intuneResults       = $null
            $entraIDResults      = $null
            $intuneConnResults   = $null
            $defenderResults     = $null
            $softwareResults     = $null
            $sharePointResults   = $null
            $teamsResults        = $null
            $userLicenseResults  = $null
            $securityResults     = $null
            
            $totalChecks = ($runAzure -as [int]) + ($runIntune -as [int]) + ($runEntraID -as [int]) + ($runIntuneConn -as [int]) + ($runDefender -as [int]) + ($runSoftware -as [int]) + ($runSharePoint -as [int]) + ($runTeams -as [int]) + ($runUserLicense -as [int]) + ($runSecurity -as [int])
            $currentCheck = 0
            $progressIncrement = if ($totalChecks -gt 0) { [Math]::Floor(80 / $totalChecks) } else { 0 }
            
            # Run Azure Check
            if ($runAzure) {
                $labelStatus.Text = "Running Azure Resources Check..."
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
                $form.Refresh()
                
                $azureResults = Test-AzureResources -BCID $bcid -CompactView $compact
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
                $currentCheck++
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
            }
            
            # Run Intune Check
            if ($runIntune) {
                $labelStatus.Text = "Running Intune Policies Check..."
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
                $form.Refresh()
                
                $intuneResults = Test-IntunePolicies -ShowAllPolicies $showAll -CompactView $compact
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
                $currentCheck++
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
            }
            
            # Run Entra ID Connect Check
            if ($runEntraID) {
                $labelStatus.Text = "Running Entra ID Connect Check..."
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
                $form.Refresh()
                
                $entraIDResults = Test-EntraIDConnect -CompactView $compact
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
                $currentCheck++
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
            }
            
            # Run Hybrid Join Check
            if ($runIntuneConn) {
                $labelStatus.Text = "Running Hybrid Azure AD Join Check..."
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
                $form.Refresh()
                
                $intuneConnResults = Test-IntuneConnector -CompactView $compact
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
                $currentCheck++
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
            }
            
            # Run Defender for Endpoint Check
            if ($runDefender) {
                $labelStatus.Text = "Running Defender for Endpoint Check..."
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
                $form.Refresh()
                
                $defenderResults = Test-DefenderForEndpoint -BCID $bcid -CompactView $compact
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
                $currentCheck++
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
            }
            
            # Run BWS Software Packages Check
            if ($runSoftware) {
                $labelStatus.Text = "Running BWS Software Packages Check..."
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
                $form.Refresh()
                
                $softwareResults = Test-BWSSoftwarePackages -CompactView $compact
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
                $currentCheck++
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
            }
            
            # Run SharePoint Configuration Check
            if ($runSharePoint) {
                $labelStatus.Text = "Running SharePoint Configuration Check..."
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
                $form.Refresh()
                
                $sharePointResults = Test-SharePointConfiguration -CompactView $compact -SharePointUrl $SharePointUrl
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
                $currentCheck++
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
            }
            
            # Run Teams Configuration Check
            if ($runTeams) {
                $labelStatus.Text = "Running Teams Configuration Check..."
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
                $form.Refresh()
                
                $teamsResults = Test-TeamsConfiguration -CompactView $compact
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
                $currentCheck++
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
            }
            
            # Run User and License Check
            if ($runUserLicense) {
                $labelStatus.Text = "Running User & License Check..."
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
                $form.Refresh()
                
                $userLicenseResults = Test-UsersAndLicenses -CompactView $compact
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
                $currentCheck++
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
            }

            # Security Configuration Check (MDM, PIM, grp_UsersWithPriviledge, SICT)
            if ($runSecurity) {
                $labelStatus.Text = "Running Security Configuration Check..."
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
                $form.Refresh()
                
                $securityResults = Test-EntraSecurityConfig -CompactView $compact
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
                $currentCheck++
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, [int](10 + ($currentCheck * $progressIncrement))))
            }
            
            # Overall Summary
            Write-Host ""
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host "  OVERALL SUMMARY" -ForegroundColor Cyan
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host "  BCID: $bcid" -ForegroundColor White
            
            if ($runAzure -and $azureResults -and $azureResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Azure Resources:" -ForegroundColor White
                Write-Host "    Total:   $($azureResults.Total)" -ForegroundColor White
                Write-Host "    Found:   $($azureResults.Found.Count)" -ForegroundColor Green
                Write-Host "    Missing: $($azureResults.Missing.Count)" -ForegroundColor $(if ($azureResults.Missing.Count -eq 0) { "Green" } else { "Red" })
            }
            
            if ($runIntune -and $intuneResults -and $intuneResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Intune Policies:" -ForegroundColor White
                Write-Host "    Required        : $($intuneResults.Total)" -ForegroundColor White
                Write-Host "    Found           : $($intuneResults.FoundCount)" -ForegroundColor $(if ($intuneResults.FoundCount -eq $intuneResults.Total) { "Green" } else { "Yellow" })
                Write-Host "    Missing         : $($intuneResults.MissingCount)" -ForegroundColor $(if ($intuneResults.MissingCount -eq 0) { "Green" } else { "Red" })
                Write-Host "    All found       : $(if ($intuneResults.AllFound) { 'Yes ([OK])' } else { 'No ([!])' })" -ForegroundColor $(if ($intuneResults.AllFound) { "Green" } else { "Red" })
                Write-Host "    Retrieval errors: $($intuneResults.Errors.Count)" -ForegroundColor $(if ($intuneResults.Errors.Count -eq 0) { "Green" } else { "Yellow" })
            }
            
            if ($runEntraID -and $entraIDResults -and $entraIDResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Entra ID Connect:" -ForegroundColor White
                Write-Host "    Sync enabled       : $(if ($entraIDResults.Status.SyncEnabled) { 'Yes ([OK])' } else { 'No' })" -ForegroundColor $(if ($entraIDResults.Status.SyncEnabled) { "Green" } else { "Red" })
                Write-Host "    Sync active (< 3h) : $(if ($entraIDResults.Status.SyncActive) { 'Yes ([OK])' } else { 'No' })" -ForegroundColor $(if ($entraIDResults.Status.SyncActive) { "Green" } else { "Yellow" })
                if ($entraIDResults.Status.LastSyncDateTime) {
                    Write-Host "    Last sync          : $($entraIDResults.Status.LastSyncDateTime)" -ForegroundColor Gray
                }
                if ($entraIDResults.Status.PasswordSyncEnabled -ne $null) {
                    Write-Host "    Password Hash Sync : $(if ($entraIDResults.Status.PasswordSyncEnabled) { 'Enabled ([OK])' } else { 'Disabled' })" -ForegroundColor $(if ($entraIDResults.Status.PasswordSyncEnabled) { "Green" } else { "Yellow" })
                }
                if ($entraIDResults.Status.DeviceWritebackEnabled -ne $null) {
                    Write-Host "    Device Writeback   : $(if ($entraIDResults.Status.DeviceWritebackEnabled) { 'Enabled ([OK])' } else { 'Disabled' })" -ForegroundColor $(if ($entraIDResults.Status.DeviceWritebackEnabled) { "Green" } else { "Gray" })
                }
                Write-Host "    Errors             : $($entraIDResults.Status.Errors.Count)" -ForegroundColor $(if ($entraIDResults.Status.Errors.Count -eq 0) { "Green" } else { "Red" })
                Write-Host "    Warnings           : $($entraIDResults.Status.Warnings.Count)" -ForegroundColor $(if ($entraIDResults.Status.Warnings.Count -eq 0) { "Green" } else { "Yellow" })
            }
            
            if ($runIntuneConn -and $intuneConnResults -and $intuneConnResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Intune Connector for Active Directory:" -ForegroundColor White
                $connState = if ($intuneConnResults.Status.ActiveCount -gt 0) { "Active ($($intuneConnResults.Status.ActiveCount))" } else { "Not Connected" }
                $connColor = if ($intuneConnResults.Status.ActiveCount -gt 0) { "Green" } else { "Red" }
                Write-Host "    Connector state    : $connState" -ForegroundColor $connColor
                Write-Host "    Total connectors   : $($intuneConnResults.Status.ConnectorCount)" -ForegroundColor White
                Write-Host "    Active             : $($intuneConnResults.Status.ActiveCount)" -ForegroundColor $(if ($intuneConnResults.Status.ActiveCount -gt 0) { "Green" } else { "Red" })
                Write-Host "    Inactive           : $($intuneConnResults.Status.InactiveCount)" -ForegroundColor $(if ($intuneConnResults.Status.InactiveCount -gt 0) { "Yellow" } else { "Green" })
                if ($intuneConnResults.Status.DeprecatedCount -gt 0) {
                    Write-Host "    DEPRECATED vers.   : $($intuneConnResults.Status.DeprecatedCount)  [UPDATE REQUIRED]" -ForegroundColor Red
                } else {
                    Write-Host "    Deprecated versions: 0" -ForegroundColor Green
                }
                Write-Host "    On-prem sync       : $(if ($intuneConnResults.Status.OnPremSyncEnabled) {'Enabled'} else {'Disabled/Unknown'})" -ForegroundColor $(if ($intuneConnResults.Status.OnPremSyncEnabled) { "Green" } else { "Gray" })
                Write-Host "    Errors             : $($intuneConnResults.Status.Errors.Count)" -ForegroundColor $(if ($intuneConnResults.Status.Errors.Count -eq 0) { "Green" } else { "Red" })
                Write-Host "    Warnings           : $($intuneConnResults.Status.Warnings.Count)" -ForegroundColor $(if ($intuneConnResults.Status.Warnings.Count -eq 0) { "Green" } else { "Yellow" })
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
            if ($runSecurity -and $securityResults -and $securityResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Security Config:" -ForegroundColor White
                Write-Host "    MDM Scope    : $(if ($securityResults.Status.MdmConfigOk) { 'All users ([OK])' } elseif ($securityResults.Status.MdmAppliesTo) { $securityResults.Status.MdmAppliesTo + ' ([!])' } else { 'Unknown' })" -ForegroundColor $(if ($securityResults.Status.MdmConfigOk) { "Green" } else { "Yellow" })
                Write-Host "    Graph CLI Grp: $(if ($securityResults.Status.GrpPriviledgeAssigned) { 'Assigned ([OK])' } else { 'MISSING ([X])' })" -ForegroundColor $(if ($securityResults.Status.GrpPriviledgeAssigned) { "Green" } else { "Red" })
                Write-Host "    PIM Partner  : $($securityResults.Status.PartnerPimUsers.Count) eligible" -ForegroundColor $(if ($securityResults.Status.PartnerPimUsers.Count -gt 0) { "Green" } else { "Yellow" })
                Write-Host "    SICT App     : $(if (-not $securityResults.Status.SictAppFound) { 'Removed ([OK])' } else { 'PRESENT ([X])' })" -ForegroundColor $(if (-not $securityResults.Status.SictAppFound) { "Green" } else { "Red" })
                Write-Host "    Errors       : $($securityResults.Status.Errors.Count)" -ForegroundColor $(if ($securityResults.Status.Errors.Count -eq 0) { "Green" } else { "Red" })
            }
            
            Write-Host "======================================================" -ForegroundColor Cyan
            
            # Show error summary if there are issues
            Show-BWsErrorSummary
            
            if ($compact) {
                Write-Host ""
                Write-Host "Note: Compact View enabled" -ForegroundColor Gray
            }
            
            # Export report if requested
            if ($export) {
                Write-Host ""
                
                $currentContext = Get-BWsAzContext
                $subName = if ($currentContext) { $currentContext.Subscription.Name } else { "Unknown" }
                
                $overallStatus = (-not $azureResults -or ($azureResults.Missing.Count -eq 0 -and $azureResults.Errors.Count -eq 0)) -and 
                                 (-not $intuneResults -or ($intuneResults.AllFound -eq $true)) -and
                                 (-not $entraIDResults -or ($entraIDResults.Status.SyncActive -eq $true)) -and
                                 (-not $intuneConnResults -or ($intuneConnResults.Status.Errors.Count -eq 0 -and $intuneConnResults.Status.DeprecatedCount -eq 0)) -and
                                 (-not $defenderResults -or ($defenderResults.Status.ConnectorActive -and $defenderResults.Status.FilesMissing.Count -eq 0)) -and
                                 (-not $softwareResults -or ($softwareResults.Status.Missing.Count -eq 0)) -and
                                 (-not $sharePointResults -or ($sharePointResults.Status.Compliant)) -and
                                 (-not $teamsResults -or ($teamsResults.Status.Compliant)) -and
                                 (-not $userLicenseResults -or ($userLicenseResults.Status.InvalidPrivilegedUsers.Count -eq 0 -and $userLicenseResults.Status.InvalidEntraIDP2Users.Count -eq 0)) -and
                                 (-not $securityResults -or ($securityResults.Status.GrpPriviledgeAssigned -and -not $securityResults.Status.SictAppFound -and $securityResults.Status.Errors.Count -eq 0))
                
                # Generate HTML report
                if ($exportFormat -eq "HTML" -or $exportFormat -eq "Both") {
                    Write-Host "Generating HTML Report..." -ForegroundColor Yellow
                    $htmlPath = Export-HTMLReport -BCID $bcid -CustomerName $customerName -SubscriptionName $subName `
                        -AzureResults $azureResults -IntuneResults $intuneResults `
                        -EntraIDResults $entraIDResults -IntuneConnResults $intuneConnResults `
                        -DefenderResults $defenderResults -SoftwareResults $softwareResults `
                        -SharePointResults $sharePointResults -TeamsResults $teamsResults `
                        -UserLicenseResults $userLicenseResults -SecurityResults $securityResults -OverallStatus $overallStatus
                    
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
                            -SharePointResults $sharePointResults -TeamsResults $teamsResults `
                            -UserLicenseResults $userLicenseResults -SecurityResults $securityResults -OverallStatus $overallStatus
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
            # Restore Write-Host, hide Stop, re-enable all buttons
            Remove-Item Function:\Write-Host -ErrorAction SilentlyContinue
            $script:stopRequested = $false
            $btnStop.Enabled   = $false
            $btnStop.Visible   = $false
            $btnStop.Text      = "Stop"
            $btnRun.Enabled    = $true
            $btnPrereq.Enabled = $true
            $btnClear.Enabled  = $true
            $btnDiag.Enabled   = $true
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

# Subscription context: set if -SubscriptionId provided, otherwise use current
# (If running with -Full this was already handled above; this is a lightweight
# fallback so that -Full is not mandatory for quick re-runs after initial setup)
if ($SubscriptionId) {
    try {
        Set-BWsAzSubscription -SubscriptionId $SubscriptionId
        Write-Host "Subscription: $SubscriptionId" -ForegroundColor Gray
    } catch {
        Write-Host "[!] Could not set subscription: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} elseif ($Full) {
    $currentContext = Get-BWsAzContext
    if (-not $currentContext) {
        Write-Host "[X] Not logged in. Run Connect-AzAccount or use -SubscriptionId" -ForegroundColor Red
        return
    }
    Write-Host "Subscription: $($currentContext.Subscription.Name)" -ForegroundColor Gray
}

# Run Azure Check
$azureResults = Test-AzureResources -BCID $BCID -CompactView $CompactView
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }

# Run Intune Check
$intuneResults = $null
if (-not $SkipIntune) {
    $intuneResults = Test-IntunePolicies -ShowAllPolicies $ShowAllPolicies -CompactView $CompactView
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
}

# Run Entra ID Connect Check
$entraIDResults = $null
if (-not $SkipEntraID) {
    $entraIDResults = Test-EntraIDConnect -CompactView $CompactView
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
}

# Run Intune Connector Check
$intuneConnResults = $null
if (-not $SkipIntuneConnector) {
    $intuneConnResults = Test-IntuneConnector -CompactView $CompactView
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
}

# Run Defender for Endpoint Check
$defenderResults = $null
if (-not $SkipDefender) {
    $defenderResults = Test-DefenderForEndpoint -BCID $BCID -CompactView $CompactView
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
}

# Run BWS Software Packages Check
$softwareResults = $null
if (-not $SkipSoftware) {
    $softwareResults = Test-BWSSoftwarePackages -CompactView $CompactView
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
}

# Run SharePoint Configuration Check
$sharePointResults = $null
if (-not $SkipSharePoint) {
    $sharePointResults = Test-SharePointConfiguration -CompactView $CompactView -SharePointUrl $SharePointUrl
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
}

# Run Teams Configuration Check
$teamsResults = $null
if (-not $SkipTeams) {
    $teamsResults = Test-TeamsConfiguration -CompactView $CompactView
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
}

# User and License Check
$userLicenseResults = $null
if (-not $SkipUserLicenseCheck) {
    $userLicenseResults = Test-UsersAndLicenses -CompactView $CompactView
                # Check stop request
                if ($script:stopRequested) {
                    $textOutput.AppendText("`r`n[STOP] Check interrupted by user.`r`n")
                    $labelStatus.Text      = "Stopped by user"
                    $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,120,60)
                    return
                }
}

# Security Configuration Check
$securityResults = $null
if (-not $SkipSecurity) {
    $securityResults = Test-EntraSecurityConfig -CompactView $CompactView
}

# Overall Summary
$currentContext = Get-BWsAzContext
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
    Write-Host "    Required        : $($intuneResults.Total)" -ForegroundColor White
    Write-Host "    Found           : $($intuneResults.FoundCount)" -ForegroundColor $(if ($intuneResults.FoundCount -eq $intuneResults.Total) { "Green" } else { "Yellow" })
    Write-Host "    Missing         : $($intuneResults.MissingCount)" -ForegroundColor $(if ($intuneResults.MissingCount -eq 0) { "Green" } else { "Red" })
    Write-Host "    All found       : $(if ($intuneResults.AllFound) { 'Yes ([OK])' } else { 'No ([!])' })" -ForegroundColor $(if ($intuneResults.AllFound) { "Green" } else { "Red" })
    Write-Host "    Retrieval errors: $($intuneResults.Errors.Count)" -ForegroundColor $(if ($intuneResults.Errors.Count -eq 0) { "Green" } else { "Yellow" })
}

if ($entraIDResults -and $entraIDResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Entra ID Connect:" -ForegroundColor White
    Write-Host "    Sync enabled       : $(if ($entraIDResults.Status.SyncEnabled) { 'Yes ([OK])' } else { 'No ([X])' })" -ForegroundColor $(if ($entraIDResults.Status.SyncEnabled) { "Green" } else { "Red" })
    Write-Host "    Sync active (< 3h) : $(if ($entraIDResults.Status.SyncActive) { 'Yes ([OK])' } else { 'No ([X])' })" -ForegroundColor $(if ($entraIDResults.Status.SyncActive) { "Green" } else { "Yellow" })
    if ($entraIDResults.Status.LastSyncDateTime) {
        Write-Host "    Last sync          : $($entraIDResults.Status.LastSyncDateTime)" -ForegroundColor Gray
    }
    if ($entraIDResults.Status.PasswordSyncEnabled -ne $null) {
        Write-Host "    Password Hash Sync : $(if ($entraIDResults.Status.PasswordSyncEnabled) { 'Enabled ([OK])' } else { 'Disabled' })" -ForegroundColor $(if ($entraIDResults.Status.PasswordSyncEnabled) { "Green" } else { "Yellow" })
    }
    if ($entraIDResults.Status.DeviceWritebackEnabled -ne $null) {
        Write-Host "    Device Writeback   : $(if ($entraIDResults.Status.DeviceWritebackEnabled) { 'Enabled' } else { 'Disabled' })" -ForegroundColor $(if ($entraIDResults.Status.DeviceWritebackEnabled) { "Green" } else { "Gray" })
    }
    Write-Host "    Errors             : $($entraIDResults.Status.Errors.Count)" -ForegroundColor $(if ($entraIDResults.Status.Errors.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "    Warnings           : $($entraIDResults.Status.Warnings.Count)" -ForegroundColor $(if ($entraIDResults.Status.Warnings.Count -eq 0) { "Green" } else { "Yellow" })
}

if ($intuneConnResults -and $intuneConnResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Intune Connector for Active Directory:" -ForegroundColor White
    $connState = if ($intuneConnResults.Status.ActiveCount -gt 0) { "Active ($($intuneConnResults.Status.ActiveCount)) ([OK])" } else { "Not Connected ([!])" }
    $connColor = if ($intuneConnResults.Status.ActiveCount -gt 0) { "Green" } else { "Red" }
    Write-Host "    Connector state    : $connState" -ForegroundColor $connColor
    Write-Host "    Total / Active     : $($intuneConnResults.Status.ConnectorCount) / $($intuneConnResults.Status.ActiveCount)" -ForegroundColor White
    if ($intuneConnResults.Status.DeprecatedCount -gt 0) {
        Write-Host "    DEPRECATED vers.   : $($intuneConnResults.Status.DeprecatedCount)  [UPDATE REQUIRED]" -ForegroundColor Red
    } else {
        Write-Host "    Deprecated versions: 0 ([OK])" -ForegroundColor Green
    }
    Write-Host "    On-prem sync       : $(if ($intuneConnResults.Status.OnPremSyncEnabled) {'Enabled ([OK])'} else {'Disabled/Unknown'})" -ForegroundColor $(if ($intuneConnResults.Status.OnPremSyncEnabled) { "Green" } else { "Gray" })
    if ($intuneConnResults.Status.LastSyncDateTime) {
        Write-Host "    Last sync          : $($intuneConnResults.Status.LastSyncDateTime)" -ForegroundColor Gray
    }
    Write-Host "    Errors             : $($intuneConnResults.Status.Errors.Count)" -ForegroundColor $(if ($intuneConnResults.Status.Errors.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "    Warnings           : $($intuneConnResults.Status.Warnings.Count)" -ForegroundColor $(if ($intuneConnResults.Status.Warnings.Count -eq 0) { "Green" } else { "Yellow" })
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

if ($securityResults -and $securityResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Security Config:" -ForegroundColor White
    Write-Host "    MDM Enrollment     : $(if ($securityResults.Status.MdmConfigOk) { 'All users ([OK])' } elseif ($securityResults.Status.MdmAppliesTo) { $securityResults.Status.MdmAppliesTo + ' ([!])' } else { 'Unknown' })" -ForegroundColor $(if ($securityResults.Status.MdmConfigOk) { "Green" } else { "Yellow" })
    Write-Host "    Graph CLI Group    : $(if ($securityResults.Status.GrpPriviledgeAssigned) { 'grp_UsersWithPriviledge assigned ([OK])' } else { 'NOT assigned ([X])' })" -ForegroundColor $(if ($securityResults.Status.GrpPriviledgeAssigned) { "Green" } else { "Red" })
    Write-Host "    PIM Partner Accts  : $($securityResults.Status.PartnerPimUsers.Count) with eligible roles" -ForegroundColor $(if ($securityResults.Status.PartnerPimUsers.Count -gt 0) { "Green" } else { "Yellow" })
    Write-Host "    Direct Perm. Admin : $($securityResults.Status.DirectAdminCount)" -ForegroundColor $(if ($securityResults.Status.DirectAdminCount -eq 0) { "Green" } else { "Yellow" })
    Write-Host "    SICT App           : $(if (-not $securityResults.Status.SictAppFound) { 'Removed ([OK])' } else { 'STILL PRESENT ([X])' })" -ForegroundColor $(if (-not $securityResults.Status.SictAppFound) { "Green" } else { "Red" })
    Write-Host "    Errors             : $($securityResults.Status.Errors.Count)" -ForegroundColor $(if ($securityResults.Status.Errors.Count -eq 0) { "Green" } else { "Red" })
}

Write-Host ""
$overallStatus = ($azureResults.Missing.Count -eq 0 -and $azureResults.Errors.Count -eq 0) -and 
                 (-not $intuneResults -or ($intuneResults.AllFound -eq $true)) -and
                 (-not $entraIDResults -or ($entraIDResults.Status.SyncActive -eq $true)) -and
                 (-not $intuneConnResults -or ($intuneConnResults.Status.Errors.Count -eq 0 -and $intuneConnResults.Status.DeprecatedCount -eq 0)) -and
                 (-not $defenderResults -or ($defenderResults.Status.ConnectorActive -and $defenderResults.Status.FilesMissing.Count -eq 0)) -and
                 (-not $softwareResults -or ($softwareResults.Status.Missing.Count -eq 0)) -and
                 (-not $sharePointResults -or ($sharePointResults.Status.Compliant)) -and
                 (-not $teamsResults -or ($teamsResults.Status.Compliant)) -and
                 (-not $userLicenseResults -or ($userLicenseResults.Status.InvalidPrivilegedUsers.Count -eq 0 -and $userLicenseResults.Status.InvalidEntraIDP2Users.Count -eq 0)) -and
                 (-not $securityResults -or ($securityResults.Status.GrpPriviledgeAssigned -and -not $securityResults.Status.SictAppFound -and $securityResults.Status.Errors.Count -eq 0))

Write-Host "  Overall Status: " -NoNewline -ForegroundColor White
if ($overallStatus) {
    Write-Host "[OK] PASSED" -ForegroundColor Green
} else {
    Write-Host "[X] ISSUES FOUND" -ForegroundColor Red
}
Write-Host "======================================================" -ForegroundColor Cyan

# Show error summary if there are issues
Show-BWsErrorSummary

# Generate support bundle if requested
if ($SupportBundle) {
    $null = Get-BWsSupportBundle -BCID $BCID
}

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
            -UserLicenseResults $userLicenseResults -SecurityResults $securityResults -OverallStatus $overallStatus
        
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
                -UserLicenseResults $userLicenseResults -SecurityResults $securityResults -OverallStatus $overallStatus
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