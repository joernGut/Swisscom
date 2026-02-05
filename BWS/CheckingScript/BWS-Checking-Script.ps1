<#
.SYNOPSIS
  BWS Checking Script (extensible rule-based health checks for Microsoft 365 / Azure services).

.DESCRIPTION
  This script runs a set of "rules" (checks) against a tenant and produces:
   - Console output (default)
   - Optional WPF GUI (Windows only) enabled via -ShowGui
   - Optional HTML report enabled via -GenerateHtmlReport

  IMPORTANT:
   - Authentication / service connections are handled by a separate login script (as per requirement).
   - This script intentionally does NOT call Connect-MgGraph / Connect-ExchangeOnline / etc.

  Initial rules (v1):
   1) Entra ID Connect is enabled and "recently synced" (based on Microsoft Graph /organization properties)
   2) Intune Connector is configured and "recently connected" (based on Microsoft Graph beta /deviceManagement/ndesConnectors)

.EXTENSIBILITY MODEL
  - Built-in rules live in Get-BwsBuiltInRules()
  - Additional rules can be added as plugin scripts in .\Checks\*.ps1 (default).
    Each plugin file should output (return) an array of rule objects with the schema described
    in Get-BwsBuiltInRules().

  Rule schema (pscustomobject):
    Id           : string (unique)
    Name         : string
    Category     : string (e.g., 'Entra ID', 'Intune', 'Exchange Online')
    Description  : string (what is being checked and why)
    Test         : scriptblock param($Context) -> returns pscustomobject with:
                   Status   : Pass | Fail | Warning | Error | NotApplicable
                   Message  : human readable result
                   Evidence : hashtable / object with additional info

.PARAMETER ShowGui
  Shows the WPF GUI. Windows only. In PowerShell 7+, the script relaunches itself in STA mode.

.PARAMETER ShowDebugConsole
  If set together with -ShowGui, shows the debug console panel (errors per rule).
  If not set, the debug console panel is collapsed.

.PARAMETER GenerateHtmlReport
  Generates an HTML report after checks complete.

.PARAMETER ReportPath
  Output path for the HTML report.

.PARAMETER ChecksPath
  Folder containing optional additional rule scripts (*.ps1). Default: .\Checks

.PARAMETER EntraLastSyncMaxHours
  Maximum allowed age (hours) of Entra ID Connect last sync time.

.PARAMETER IntuneConnectorLastSeenMaxHours
  Maximum allowed age (hours) of Intune NDES connector lastConnectionDateTime.

.EXAMPLE
  # Headless run (console output only)
  .\BWS-Checking-Script.ps1

.EXAMPLE
  # GUI run (includes debug console)
  .\BWS-Checking-Script.ps1 -ShowGui -ShowDebugConsole

.EXAMPLE
  # Generate HTML report
  .\BWS-Checking-Script.ps1 -GenerateHtmlReport -ReportPath .\BWS-Report.html

.NOTES
  - WPF GUI requires Windows.
  - Intune NDES connector endpoint is currently in Microsoft Graph /beta and may change.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $false)]
  [switch]$ShowGui,

  [Parameter(Mandatory = $false)]
  [switch]$ShowDebugConsole,

  [Parameter(Mandatory = $false)]
  [switch]$GenerateHtmlReport,

  [Parameter(Mandatory = $false)]
  [ValidateNotNullOrEmpty()]
  [string]$ReportPath = (Join-Path -Path $PSScriptRoot -ChildPath ("BWS-Checking-Report_{0}.html" -f (Get-Date -Format "yyyyMMdd_HHmmss"))),

  [Parameter(Mandatory = $false)]
  [ValidateNotNullOrEmpty()]
  [string]$ChecksPath = (Join-Path -Path $PSScriptRoot -ChildPath "Checks"),

  [Parameter(Mandatory = $false)]
  [ValidateRange(1, 720)]
  [int]$EntraLastSyncMaxHours = 24,

  [Parameter(Mandatory = $false)]
  [ValidateRange(1, 720)]
  [int]$IntuneConnectorLastSeenMaxHours = 24
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------
# Global runtime state (script scope)
# ---------------------------
$script:LogEntries = New-Object System.Collections.Generic.List[object]

# ---------------------------
# Logging helper
# ---------------------------
function Write-BwsLog {
  <#
    .SYNOPSIS
      Adds a structured entry to the in-memory debug log.
    .DESCRIPTION
      Used by rule execution wrapper and by individual rules.
      The GUI debug console and HTML report can display these entries.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [ValidateSet('INFO','WARN','ERROR','DEBUG')]
    [string]$Level,

    [Parameter(Mandatory)]
    [string]$Message,

    [Parameter(Mandatory = $false)]
    [string]$RuleId,

    [Parameter(Mandatory = $false)]
    [System.Exception]$Exception
  )

  $entry = [pscustomobject]@{
    Timestamp = (Get-Date)
    Level     = $Level
    RuleId    = $RuleId
    Message   = $Message
    Exception = if ($Exception) { $Exception.ToString() } else { $null }
  }

  [void]$script:LogEntries.Add($entry)
}

# ---------------------------
# Utility: Convert bound parameters to arguments (for STA relaunch)
# ---------------------------
function ConvertTo-BwsArgumentList {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [hashtable]$BoundParameters
  )

  $args = New-Object System.Collections.Generic.List[string]

  foreach ($kvp in $BoundParameters.GetEnumerator()) {
    $k = $kvp.Key
    $v = $kvp.Value

    if ($v -is [switch]) {
      if ($v.IsPresent) {
        [void]$args.Add("-$k")
      }
      continue
    }

    # Quote values safely for a new PowerShell process.
    $escaped = $v.ToString().Replace('"','`"')
    [void]$args.Add("-$k")
    [void]$args.Add('"' + $escaped + '"')
  }

  return ,$args.ToArray()
}

# ---------------------------
# GUI: Ensure STA on Windows when showing WPF
# ---------------------------
if ($ShowGui) {
  if (-not $IsWindows) {
    throw "WPF GUI is only supported on Windows. Run without -ShowGui on non-Windows platforms."
  }

  $isSta = ([System.Threading.Thread]::CurrentThread.ApartmentState -eq 'STA')
  if (-not $isSta -and -not $env:BWS_STA_RELAUNCH) {
    # PowerShell 7+ is typically MTA in console sessions. WPF requires STA.
    # We relaunch the script in STA mode and exit the current process.
    $env:BWS_STA_RELAUNCH = '1'

    $exe = if ($PSVersionTable.PSEdition -eq 'Core') {
      Join-Path $PSHOME 'pwsh.exe'
    } else {
      Join-Path $PSHOME 'powershell.exe'
    }

    $argList = New-Object System.Collections.Generic.List[string]
    [void]$argList.Add('-NoProfile')
    [void]$argList.Add('-ExecutionPolicy'); [void]$argList.Add('Bypass')
    [void]$argList.Add('-Sta')
    [void]$argList.Add('-File'); [void]$argList.Add('"' + $PSCommandPath.Replace('"','`"') + '"')

    foreach ($a in (ConvertTo-BwsArgumentList -BoundParameters $PSBoundParameters)) {
      [void]$argList.Add($a)
    }

    Start-Process -FilePath $exe -ArgumentList $argList.ToArray() -Wait | Out-Null
    exit $LASTEXITCODE
  }
}

# ---------------------------
# Graph prerequisites
# ---------------------------
function Assert-BwsGraphConnected {
  <#
    .SYNOPSIS
      Ensures Microsoft Graph PowerShell SDK context is available (authentication handled elsewhere).
    .DESCRIPTION
      This script relies on Invoke-MgGraphRequest (Microsoft.Graph.Authentication).
      Your separate login script should have run Connect-MgGraph already.
  #>
  [CmdletBinding()]
  param()

  if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
    throw "Missing Microsoft Graph PowerShell SDK. Install/Import Microsoft.Graph.Authentication (or Microsoft.Graph) so Invoke-MgGraphRequest is available."
  }

  if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) {
    throw "Microsoft Graph PowerShell SDK cmdlets not found. Ensure Microsoft.Graph.Authentication is imported."
  }

  $ctx = Get-MgContext
  if (-not $ctx -or -not $ctx.Account) {
    throw "No Microsoft Graph context found. Please run your separate login script first (Connect-MgGraph)."
  }

  Write-BwsLog -Level INFO -Message ("Graph context found. Account: {0}, TenantId: {1}" -f $ctx.Account, $ctx.TenantId)
}

function Invoke-BwsGraphGet {
  <#
    .SYNOPSIS
      Thin wrapper around Invoke-MgGraphRequest for GET calls with consistent error handling.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Uri
  )

  try {
    return Invoke-MgGraphRequest -Method GET -Uri $Uri
  }
  catch {
    # Re-throw but keep a consistent, readable message for logs.
    throw $_.Exception
  }
}

# ---------------------------
# Rules: Built-in definitions
# ---------------------------
function Get-BwsBuiltInRules {
  <#
    .SYNOPSIS
      Returns the built-in rule definitions.
    .DESCRIPTION
      To extend, either:
       - Add new rule objects here, OR
       - Drop additional rule scripts into the plugin folder (default: .\Checks)
  #>

  $rules = @()

  # Rule 1: Entra ID Connect configured and syncing
  $rules += [pscustomobject]@{
    Id          = 'ENTRA-001'
    Name        = 'Entra ID Connect is enabled and syncing recently'
    Category    = 'Entra ID'
    Description = @'
Checks whether directory synchronization from on-premises is enabled for the tenant
and whether the last sync timestamp is "fresh" (within a configurable threshold).

Data source:
 - Microsoft Graph v1.0 /organization
Key properties:
 - onPremisesSyncEnabled
 - onPremisesLastSyncDateTime

Interpretation:
 - Pass  : onPremisesSyncEnabled = true AND last sync <= threshold hours
 - Fail  : onPremisesSyncEnabled = false OR last sync older than threshold
 - Warn  : enabled but last sync timestamp missing
 - Error : Graph call failed (missing permissions, connectivity, etc.)

Permissions (typical delegated):
 - Organization.Read.All (or higher)
'@
    Test        = {
      param($Context)

      $uri = "https://graph.microsoft.com/v1.0/organization?`$select=displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesLastPasswordSyncDateTime"

      $orgResponse = Invoke-BwsGraphGet -Uri $uri
      $org = $orgResponse.value | Select-Object -First 1

      if (-not $org) {
        return [pscustomobject]@{
          Status   = 'Error'
          Message  = 'Graph returned no organization object.'
          Evidence = @{ Uri = $uri }
        }
      }

      $enabled = [bool]$org.onPremisesSyncEnabled
      $lastSync = $null
      if ($org.onPremisesLastSyncDateTime) {
        $lastSync = [datetime]$org.onPremisesLastSyncDateTime
      }

      if (-not $enabled) {
        return [pscustomobject]@{
          Status   = 'Fail'
          Message  = 'Entra ID Connect (directory sync) is not enabled for this tenant.'
          Evidence = @{
            TenantDisplayName = $org.displayName
            onPremisesSyncEnabled = $org.onPremisesSyncEnabled
            onPremisesLastSyncDateTime = $org.onPremisesLastSyncDateTime
          }
        }
      }

      if (-not $lastSync) {
        return [pscustomobject]@{
          Status   = 'Warning'
          Message  = 'Directory sync is enabled, but the last sync timestamp is missing.'
          Evidence = @{
            TenantDisplayName = $org.displayName
            onPremisesSyncEnabled = $org.onPremisesSyncEnabled
            onPremisesLastSyncDateTime = $org.onPremisesLastSyncDateTime
          }
        }
      }

      $age = New-TimeSpan -Start $lastSync.ToUniversalTime() -End $Context.NowUtc
      $ageHours = [math]::Round($age.TotalHours, 2)

      if ($ageHours -le $Context.EntraLastSyncMaxHours) {
        return [pscustomobject]@{
          Status   = 'Pass'
          Message  = ("Directory sync is enabled and last sync is {0} hours old (<= {1}h)." -f $ageHours, $Context.EntraLastSyncMaxHours)
          Evidence = @{
            TenantDisplayName = $org.displayName
            onPremisesSyncEnabled = $org.onPremisesSyncEnabled
            onPremisesLastSyncDateTimeUtc = $lastSync.ToUniversalTime().ToString("o")
            AgeHours = $ageHours
            ThresholdHours = $Context.EntraLastSyncMaxHours
          }
        }
      }

      return [pscustomobject]@{
        Status   = 'Fail'
        Message  = ("Directory sync is enabled but last sync is stale: {0} hours old (> {1}h)." -f $ageHours, $Context.EntraLastSyncMaxHours)
        Evidence = @{
          TenantDisplayName = $org.displayName
          onPremisesSyncEnabled = $org.onPremisesSyncEnabled
          onPremisesLastSyncDateTimeUtc = $lastSync.ToUniversalTime().ToString("o")
          AgeHours = $ageHours
          ThresholdHours = $Context.EntraLastSyncMaxHours
        }
      }
    }
  }

  # Rule 2: Intune connector configured and "syncing" (interpreted as NDES connector active + recent heartbeat)
  $rules += [pscustomobject]@{
    Id          = 'INTUNE-001'
    Name        = 'Intune NDES Connector is active and connected recently'
    Category    = 'Intune'
    Description = @'
Checks whether an Intune on-prem NDES connector exists, is active, and has connected recently.

Data source:
 - Microsoft Graph /beta/deviceManagement/ndesConnectors
Key properties per connector:
 - state (active/inactive)
 - lastConnectionDateTime

Interpretation:
 - Pass  : >= 1 connector with state=active AND lastConnectionDateTime <= threshold hours
 - Fail  : no connectors OR no active connectors OR last connection older than threshold
 - Error : Graph call failed (missing permissions, licensing, etc.)

Notes:
 - This endpoint is currently in Microsoft Graph /beta, so it may change.
 - Required permissions for this API (per docs):
    Delegated/App: DeviceManagementConfiguration.Read.All (or higher)

If your environment uses a different Intune "connector" concept, add another rule later
(or adjust this rule) using the appropriate endpoint(s).
'@
    Test        = {
      param($Context)

      $uri = "https://graph.microsoft.com/beta/deviceManagement/ndesConnectors"

      $resp = Invoke-BwsGraphGet -Uri $uri
      $connectors = @($resp.value)

      if (-not $connectors -or $connectors.Count -eq 0) {
        return [pscustomobject]@{
          Status   = 'Fail'
          Message  = 'No Intune NDES connectors found (not configured).'
          Evidence = @{ Uri = $uri }
        }
      }

      $active = @($connectors | Where-Object { $_.state -eq 'active' })

      if (-not $active -or $active.Count -eq 0) {
        return [pscustomobject]@{
          Status   = 'Fail'
          Message  = 'NDES connectors exist, but none are in state=active.'
          Evidence = @{
            TotalConnectors = $connectors.Count
            States = ($connectors | Select-Object -ExpandProperty state | Sort-Object | Get-Unique)
          }
        }
      }

      # Find newest lastConnectionDateTime among active connectors
      $latest = $active |
        Where-Object { $_.lastConnectionDateTime } |
        Sort-Object { [datetime]$_.lastConnectionDateTime } -Descending |
        Select-Object -First 1

      if (-not $latest -or -not $latest.lastConnectionDateTime) {
        return [pscustomobject]@{
          Status   = 'Warning'
          Message  = 'Active NDES connectors found, but lastConnectionDateTime is missing.'
          Evidence = @{
            ActiveConnectors = $active.Count
            ConnectorIds = ($active | Select-Object -ExpandProperty id)
          }
        }
      }

      $last = [datetime]$latest.lastConnectionDateTime
      $age = New-TimeSpan -Start $last.ToUniversalTime() -End $Context.NowUtc
      $ageHours = [math]::Round($age.TotalHours, 2)

      if ($ageHours -le $Context.IntuneConnectorLastSeenMaxHours) {
        return [pscustomobject]@{
          Status   = 'Pass'
          Message  = ("Active NDES connector(s) found; last connection is {0} hours old (<= {1}h)." -f $ageHours, $Context.IntuneConnectorLastSeenMaxHours)
          Evidence = @{
            TotalConnectors = $connectors.Count
            ActiveConnectors = $active.Count
            LatestConnectorId = $latest.id
            LatestConnectorName = $latest.displayName
            LatestLastConnectionUtc = $last.ToUniversalTime().ToString("o")
            AgeHours = $ageHours
            ThresholdHours = $Context.IntuneConnectorLastSeenMaxHours
          }
        }
      }

      return [pscustomobject]@{
        Status   = 'Fail'
        Message  = ("Active NDES connector(s) found, but last connection is stale: {0} hours old (> {1}h)." -f $ageHours, $Context.IntuneConnectorLastSeenMaxHours)
        Evidence = @{
          TotalConnectors = $connectors.Count
          ActiveConnectors = $active.Count
          LatestConnectorId = $latest.id
          LatestConnectorName = $latest.displayName
          LatestLastConnectionUtc = $last.ToUniversalTime().ToString("o")
          AgeHours = $ageHours
          ThresholdHours = $Context.IntuneConnectorLastSeenMaxHours
        }
      }
    }
  }

  return $rules
}

function Get-BwsPluginRules {
  <#
    .SYNOPSIS
      Loads additional rules from plugin folder.
    .DESCRIPTION
      Each .ps1 in $ChecksPath should output/return an array of rule objects.
      Example plugin file content:

        @(
          [pscustomobject]@{
            Id='EXO-001'; Name='...'; Category='Exchange Online'; Description='...';
            Test={ param($Context) return [pscustomobject]@{ Status='Pass'; Message='...'; Evidence=@{} } }
          }
        )

      The plugin runs in the current scope, so it can call Write-BwsLog if needed.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [string]$Path
  )

  if (-not (Test-Path -LiteralPath $Path)) {
    Write-BwsLog -Level INFO -Message "ChecksPath not found; no plugin rules loaded: $Path"
    return @()
  }

  $files = Get-ChildItem -LiteralPath $Path -Filter '*.ps1' -File -ErrorAction Stop
  if (-not $files) { return @() }

  $loaded = @()

  foreach ($f in $files) {
    try {
      Write-BwsLog -Level INFO -Message "Loading plugin rule file: $($f.FullName)"
      $r = . $f.FullName
      if ($r) { $loaded += $r }
    }
    catch {
      Write-BwsLog -Level ERROR -Message "Failed to load plugin rule file: $($f.FullName)" -Exception $_.Exception
    }
  }

  return $loaded
}

# ---------------------------
# Rule execution
# ---------------------------
function Invoke-BwsRule {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [pscustomobject]$Rule,

    [Parameter(Mandatory)]
    [pscustomobject]$Context
  )

  $start = Get-Date
  try {
    Write-BwsLog -Level INFO -RuleId $Rule.Id -Message ("Starting rule: {0} - {1}" -f $Rule.Id, $Rule.Name)

    $out = & $Rule.Test $Context

    # Normalize output (defensive programming)
    if (-not $out -or -not $out.Status) {
      $out = [pscustomobject]@{
        Status   = 'Error'
        Message  = 'Rule returned no result or missing Status.'
        Evidence = @{}
      }
    }

    $end = Get-Date
    $dur = New-TimeSpan -Start $start -End $end

    return [pscustomobject]@{
      RuleId      = $Rule.Id
      Name        = $Rule.Name
      Category    = $Rule.Category
      Description = $Rule.Description
      Status      = $out.Status
      Message     = $out.Message
      Evidence    = $out.Evidence
      StartedAt   = $start
      FinishedAt  = $end
      DurationMs  = [int][math]::Round($dur.TotalMilliseconds, 0)
    }
  }
  catch {
    $end = Get-Date
    $dur = New-TimeSpan -Start $start -End $end

    Write-BwsLog -Level ERROR -RuleId $Rule.Id -Message ("Rule crashed: {0}" -f $Rule.Name) -Exception $_.Exception

    return [pscustomobject]@{
      RuleId      = $Rule.Id
      Name        = $Rule.Name
      Category    = $Rule.Category
      Description = $Rule.Description
      Status      = 'Error'
      Message     = $_.Exception.Message
      Evidence    = @{ Exception = $_.Exception.ToString() }
      StartedAt   = $start
      FinishedAt  = $end
      DurationMs  = [int][math]::Round($dur.TotalMilliseconds, 0)
    }
  }
}

function Invoke-BwsAllRules {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [pscustomobject[]]$Rules,

    [Parameter(Mandatory)]
    [pscustomobject]$Context
  )

  $results = foreach ($r in $Rules) {
    Invoke-BwsRule -Rule $r -Context $Context
  }

  return $results
}

# ---------------------------
# HTML report
# ---------------------------
function ConvertTo-BwsHtmlEncoded {
  param([string]$Text)
  if ($null -eq $Text) { return '' }
  return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Export-BwsHtmlReport {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)]
    [pscustomobject[]]$Results,

    [Parameter(Mandatory)]
    [object[]]$LogEntries,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  $now = Get-Date
  $pass = ($Results | Where-Object Status -eq 'Pass').Count
  $fail = ($Results | Where-Object Status -eq 'Fail').Count
  $warn = ($Results | Where-Object Status -eq 'Warning').Count
  $err  = ($Results | Where-Object Status -eq 'Error').Count
  $na   = ($Results | Where-Object Status -eq 'NotApplicable').Count

  $rows = foreach ($r in $Results) {
    $statusClass = switch ($r.Status) {
      'Pass' { 'pass' }
      'Fail' { 'fail' }
      'Warning' { 'warn' }
      'Error' { 'err' }
      default { 'na' }
    }

    $evidenceJson = ''
    try { $evidenceJson = ($r.Evidence | ConvertTo-Json -Depth 8) } catch { $evidenceJson = '<failed to serialize evidence>' }

    @"
<tr>
  <td>$(ConvertTo-BwsHtmlEncoded $r.Category)</td>
  <td><strong>$(ConvertTo-BwsHtmlEncoded $r.RuleId)</strong><br/>$(ConvertTo-BwsHtmlEncoded $r.Name)</td>
  <td class="$statusClass">$(ConvertTo-BwsHtmlEncoded $r.Status)</td>
  <td>$(ConvertTo-BwsHtmlEncoded $r.Message)</td>
  <td><details><summary>Show</summary><pre>$(ConvertTo-BwsHtmlEncoded $evidenceJson)</pre></details></td>
</tr>
"@
  }

  $logRows = foreach ($l in $LogEntries) {
    $ex = if ($l.Exception) { "`n$l.Exception" } else { "" }
    @"
<tr>
  <td>$(ConvertTo-BwsHtmlEncoded ($l.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")))</td>
  <td>$(ConvertTo-BwsHtmlEncoded $l.Level)</td>
  <td>$(ConvertTo-BwsHtmlEncoded $l.RuleId)</td>
  <td><pre>$(ConvertTo-BwsHtmlEncoded ($l.Message + $ex))</pre></td>
</tr>
"@
  }

  $html = @"
<!doctype html>
<html>
<head>
  <meta charset="utf-8"/>
  <title>BWS Checking Script Report</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; }
    .meta { margin-bottom: 16px; }
    .kpi { display: inline-block; margin-right: 12px; padding: 8px 10px; border: 1px solid #ddd; border-radius: 8px; }
    table { border-collapse: collapse; width: 100%; margin-top: 10px; }
    th, td { border: 1px solid #ddd; padding: 8px; vertical-align: top; }
    th { background: #f6f6f6; }
    .pass { background: #eaffea; font-weight: 600; }
    .fail { background: #ffecec; font-weight: 600; }
    .warn { background: #fff7e6; font-weight: 600; }
    .err  { background: #ffd6d6; font-weight: 600; }
    .na   { background: #f2f2f2; font-weight: 600; }
    pre { white-space: pre-wrap; word-wrap: break-word; margin: 0; }
    details summary { cursor: pointer; }
  </style>
</head>
<body>
  <h1>BWS Checking Script Report</h1>
  <div class="meta">
    <div>Generated: <strong>$($now.ToString("yyyy-MM-dd HH:mm:ss"))</strong> (local time)</div>
  </div>

  <div class="kpi">Pass: <strong>$pass</strong></div>
  <div class="kpi">Fail: <strong>$fail</strong></div>
  <div class="kpi">Warning: <strong>$warn</strong></div>
  <div class="kpi">Error: <strong>$err</strong></div>
  <div class="kpi">N/A: <strong>$na</strong></div>

  <h2>Rules</h2>
  <table>
    <thead>
      <tr>
        <th>Category</th>
        <th>Rule</th>
        <th>Status</th>
        <th>Message</th>
        <th>Evidence</th>
      </tr>
    </thead>
    <tbody>
      $($rows -join "`n")
    </tbody>
  </table>

  <h2>Debug Console (errors, warnings, traces)</h2>
  <table>
    <thead>
      <tr>
        <th>Time</th>
        <th>Level</th>
        <th>RuleId</th>
        <th>Message / Exception</th>
      </tr>
    </thead>
    <tbody>
      $($logRows -join "`n")
    </tbody>
  </table>
</body>
</html>
"@

  if ($PSCmdlet.ShouldProcess($Path, "Write HTML report")) {
    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
      New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $html | Out-File -LiteralPath $Path -Encoding UTF8 -Force
    Write-BwsLog -Level INFO -Message "HTML report written to: $Path"
  }
}

# ---------------------------
# WPF GUI
# ---------------------------
function Show-BwsGui {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [pscustomobject[]]$Results,

    [Parameter(Mandatory)]
    [object[]]$LogEntries,

    [Parameter(Mandatory)]
    [bool]$ShowDebugConsolePanel,

    [Parameter(Mandatory)]
    [string]$DefaultReportPath
  )

  Add-Type -AssemblyName PresentationFramework
  Add-Type -AssemblyName PresentationCore
  Add-Type -AssemblyName WindowsBase

  $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="BWS Checking Script" Height="720" Width="1100"
        WindowStartupLocation="CenterScreen">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,10">
      <TextBlock Text="BWS Checking Script" FontSize="18" FontWeight="Bold" Margin="0,0,18,0"/>
      <TextBlock Text="(Results + optional Debug Console)" VerticalAlignment="Center" Foreground="Gray"/>
    </StackPanel>

    <Grid Grid.Row="1">
      <Grid.RowDefinitions>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="220"/>
      </Grid.RowDefinitions>

      <DataGrid x:Name="dgResults" Grid.Row="0" AutoGenerateColumns="False" IsReadOnly="True"
                CanUserAddRows="False" Margin="0,0,0,10">
        <DataGrid.Columns>
          <DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="120"/>
          <DataGridTextColumn Header="RuleId" Binding="{Binding RuleId}" Width="90"/>
          <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="300"/>
          <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="90"/>
          <DataGridTextColumn Header="Message" Binding="{Binding Message}" Width="*"/>
        </DataGrid.Columns>
      </DataGrid>

      <TextBlock Grid.Row="1" Text="Debug Console (rule errors/exceptions):" FontWeight="Bold" Margin="0,0,0,6"/>

      <Border Grid.Row="2" BorderBrush="#DDDDDD" BorderThickness="1" CornerRadius="6">
        <TextBox x:Name="tbDebug" FontFamily="Consolas" FontSize="12"
                 TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                 IsReadOnly="True" Padding="8"/>
      </Border>
    </Grid>

    <StackPanel Orientation="Horizontal" Grid.Row="2" HorizontalAlignment="Right" Margin="0,10,0,0">
      <Button x:Name="btnExport" Content="Export HTML Report" Width="150" Height="30" Margin="0,0,10,0"/>
      <Button x:Name="btnClose" Content="Close" Width="90" Height="30"/>
    </StackPanel>
  </Grid>
</Window>
"@

  [xml]$xml = $xaml
  $reader = New-Object System.Xml.XmlNodeReader $xml
  $window = [System.Windows.Markup.XamlReader]::Load($reader)

  $dgResults = $window.FindName('dgResults')
  $tbDebug   = $window.FindName('tbDebug')
  $btnExport = $window.FindName('btnExport')
  $btnClose  = $window.FindName('btnClose')

  # Bind results
  $dgResults.ItemsSource = $Results

  # Fill debug console text (all log entries)
  if ($ShowDebugConsolePanel) {
    $lines = foreach ($l in $LogEntries) {
      $ts = $l.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")
      $rid = if ($l.RuleId) { $l.RuleId } else { '-' }
      $ex = if ($l.Exception) { "`n$l.Exception" } else { "" }
      "[{0}] {1} [{2}] {3}{4}" -f $ts, $l.Level, $rid, $l.Message, $ex
    }
    $tbDebug.Text = ($lines -join "`r`n")
  }
  else {
    # Collapse debug console if not requested
    $tbDebug.Text = "Debug console hidden (use -ShowDebugConsole to display)."
  }

  $btnExport.Add_Click({
    try {
      Export-BwsHtmlReport -Results $Results -LogEntries $LogEntries -Path $DefaultReportPath -Confirm:$false
      [System.Windows.MessageBox]::Show("HTML report exported to:`n$DefaultReportPath","BWS Checking Script") | Out-Null
    }
    catch {
      [System.Windows.MessageBox]::Show("Failed to export HTML report:`n$($_.Exception.Message)","BWS Checking Script") | Out-Null
    }
  })

  $btnClose.Add_Click({ $window.Close() })

  [void]$window.ShowDialog()
}

# ---------------------------
# Main
# ---------------------------
try {
  Write-BwsLog -Level INFO -Message "BWS Checking Script started."

  Assert-BwsGraphConnected

  $context = [pscustomobject]@{
    NowUtc                        = [datetime]::UtcNow
    EntraLastSyncMaxHours         = $EntraLastSyncMaxHours
    IntuneConnectorLastSeenMaxHours = $IntuneConnectorLastSeenMaxHours
  }

  $rules = @()
  $rules += Get-BwsBuiltInRules
  $rules += Get-BwsPluginRules -Path $ChecksPath

  # Basic sanity check: rule IDs should be unique
  $dup = $rules | Group-Object Id | Where-Object Count -gt 1
  if ($dup) {
    $ids = ($dup | Select-Object -ExpandProperty Name) -join ', '
    throw "Duplicate rule Id(s) detected: $ids"
  }

  $results = Invoke-BwsAllRules -Rules $rules -Context $context

  # Console summary (headless-friendly)
  $results |
    Select-Object Category, RuleId, Name, Status, Message |
    Sort-Object Category, RuleId |
    Format-Table -AutoSize | Out-String | Write-Host

  if ($GenerateHtmlReport) {
    Export-BwsHtmlReport -Results $results -LogEntries $script:LogEntries.ToArray() -Path $ReportPath -Confirm:$false
    Write-Host "HTML report: $ReportPath"
  }

  if ($ShowGui) {
    Show-BwsGui -Results $results -LogEntries $script:LogEntries.ToArray() -ShowDebugConsolePanel:$ShowDebugConsole.IsPresent -DefaultReportPath $ReportPath
  }

  Write-BwsLog -Level INFO -Message "BWS Checking Script finished."

  # Output objects to pipeline (automation / CI usage)
  $results
}
catch {
  Write-BwsLog -Level ERROR -Message "Fatal error." -Exception $_.Exception
  Write-Error $_
  exit 1
}
