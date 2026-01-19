<# 
.SYNOPSIS
  BWSCheckScript - Configuration/Compliance checks for Entra ID, Azure, Azure Virtual Desktop, Intune and Active Directory.
.DESCRIPTION
  - Loads conditions from .\BWSConditions.ps1 (same folder)
  - Optional GUI (-Gui)
  - Preflights + installs/imports required modules (including auth modules)
  - Executes checks and generates an HTML report
.NOTES
  Recommended: PowerShell 7.4+ on Windows
  GUI requires Windows (WinForms).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$Gui,

    [Parameter(Mandatory = $false)]
    [ValidateSet('DeviceCode','Interactive','ClientCertificate','ManagedIdentity')]
    [string]$AuthMode = 'DeviceCode',

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    # App-only (ClientCertificate)
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Join-Path -Path $PWD -ChildPath "BWSReport"),

    [Parameter(Mandatory = $false)]
    [switch]$AutoInstallModules,

    [Parameter(Mandatory = $false)]
    [string[]]$IncludeProducts, # e.g. 'EntraID','Azure','AVD','Intune','AD'

    [Parameter(Mandatory = $false)]
    [string[]]$IncludeTags,     # e.g. 'MFA','Baseline'

    [Parameter(Mandatory = $false)]
    [switch]$NoAuth             # If you connect yourself before running this script
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------- Utilities ----------
function New-BwsRunId {
    (Get-Date).ToString('yyyyMMdd-HHmmss')
}

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Write-BwsLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')]
        [string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Write-Host "[$ts][$Level] $Message"
}

function Test-IsWindows { return $IsWindows }

function Set-TlsForWindowsPowerShell {
    # Windows PowerShell 5.1 often needs TLS 1.2 for PSGallery
    if ($PSVersionTable.PSEdition -eq 'Desktop') {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        } catch {
            Write-BwsLog "Could not set TLS 1.2: $($_.Exception.Message)" "WARN"
        }
    }
}

function Ensure-NuGetProvider {
    # Required for Install-Module on some systems
    $prov = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $prov) {
        if (-not $AutoInstallModules) {
            Write-BwsLog "NuGet package provider is missing. Install it or run with -AutoInstallModules." "WARN"
            return
        }
        Write-BwsLog "Installing NuGet package provider..." "INFO"
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }
}

function Ensure-Module {
    param(
        [Parameter(Mandatory)][string]$Name,
        [switch]$AutoInstall
    )

    if (Get-Module -ListAvailable -Name $Name) { return $true }

    if (-not $AutoInstall) {
        Write-BwsLog "Module '$Name' is missing. Install manually: Install-Module $Name -Scope CurrentUser" "WARN"
        return $false
    }

    Set-TlsForWindowsPowerShell
    Ensure-NuGetProvider

    Write-BwsLog "Installing missing module '$Name' (CurrentUser)..." "INFO"
    Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop

    return [bool](Get-Module -ListAvailable -Name $Name)
}

function Import-RequiredModule {
    param(
        [Parameter(Mandatory)][string]$Name
    )
    try {
        Import-Module $Name -ErrorAction Stop | Out-Null
        Write-BwsLog "Imported module: $Name" "DEBUG"
    } catch {
        throw "Failed to import module '$Name': $($_.Exception.Message)"
    }
}

function Get-Union {
    param([string[]]$A, [string[]]$B)
    @($A + $B | Where-Object { $_ -and $_.Trim() } | Select-Object -Unique)
}

function Get-FilteredConditions {
    param(
        [Parameter(Mandatory)][object[]]$Conditions,
        [string[]]$IncludeProducts,
        [string[]]$IncludeTags
    )
    $filtered = $Conditions

    if ($IncludeProducts -and $IncludeProducts.Count -gt 0) {
        $set = $IncludeProducts | ForEach-Object { $_.Trim() }
        $filtered = $filtered | Where-Object { $set -contains $_.Product }
    }
    if ($IncludeTags -and $IncludeTags.Count -gt 0) {
        $tags = $IncludeTags | ForEach-Object { $_.Trim() }
        $filtered = $filtered | Where-Object {
            $_.Tags -and (@($_.Tags) | Where-Object { $tags -contains $_ }).Count -gt 0
        }
    }
    return @($filtered)
}

# ---------- Condition Loading ----------
function Import-BwsConditions {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "BWSConditions not found: $Path"
    }
    $conds = & $Path
    if (-not $conds) { throw "BWSConditions returned no conditions." }
    if ($conds -isnot [System.Collections.IEnumerable]) { throw "BWSConditions must return an array/enumerable." }
    return @($conds)
}

# ---------- Module Preflight ----------
function Get-ModulesRequiredForConditions {
    param([Parameter(Mandatory)][object[]]$Conditions)

    $needGraph = ($Conditions | Where-Object { $_.RequiresGraph -eq $true }).Count -gt 0
    $needAzure = ($Conditions | Where-Object { $_.RequiresAzure -eq $true }).Count -gt 0
    $needAD    = ($Conditions | Where-Object { $_.RequiresAD -eq $true }).Count -gt 0

    $modules = New-Object System.Collections.Generic.List[string]

    # Graph stack
    if ($needGraph) {
        [void]$modules.Add("Microsoft.Graph") # includes Connect-MgGraph / Graph cmdlets
    }

    # Az stack
    if ($needAzure) {
        [void]$modules.Add("Az.Accounts") # Connect-AzAccount
        [void]$modules.Add("Az.Resources") # common for subscription/resource queries

        # If any AVD product checks exist -> ensure Az.DesktopVirtualization
        $needAvdModule = ($Conditions | Where-Object { $_.Product -eq 'AVD' }).Count -gt 0
        if ($needAvdModule) { [void]$modules.Add("Az.DesktopVirtualization") }
    }

    # Active Directory (RSAT)
    if ($needAD) {
        [void]$modules.Add("ActiveDirectory")
    }

    # Optional per-condition module list (future-proof)
    $Conditions | ForEach-Object {
        if ($_.PSObject.Properties.Name -contains 'RequiredModules' -and $_.RequiredModules) {
            foreach ($m in @($_.RequiredModules)) { if ($m) { [void]$modules.Add([string]$m) } }
        }
    }

    return @($modules | Select-Object -Unique)
}

function Ensure-AndImportModulesForRun {
    param([Parameter(Mandatory)][object[]]$Conditions)

    $modules = Get-ModulesRequiredForConditions -Conditions $Conditions
    if (-not $modules -or $modules.Count -eq 0) {
        Write-BwsLog "No external modules required based on current filters." "INFO"
        return
    }

    Write-BwsLog "Required modules: $($modules -join ', ')" "INFO"

    foreach ($m in $modules) {
        $ok = Ensure-Module -Name $m -AutoInstall:$AutoInstallModules
        if (-not $ok) {
            throw "Missing module '$m'. Install it or run with -AutoInstallModules."
        }
        Import-RequiredModule -Name $m
    }

    # Sanity checks for auth cmdlets
    if (($Conditions | Where-Object { $_.RequiresGraph }).Count -gt 0) {
        if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
            throw "Connect-MgGraph not found even after importing Microsoft.Graph."
        }
    }
    if (($Conditions | Where-Object { $_.RequiresAzure }).Count -gt 0) {
        if (-not (Get-Command Connect-AzAccount -ErrorAction SilentlyContinue)) {
            throw "Connect-AzAccount not found even after importing Az.Accounts."
        }
    }
}

# ---------- Auth / Connections ----------
function Connect-BwsGraph {
    param(
        [Parameter(Mandatory)][string[]]$Scopes,
        [string]$TenantId,
        [ValidateSet('DeviceCode','Interactive','ClientCertificate','ManagedIdentity')]
        [string]$AuthMode,
        [string]$ClientId,
        [string]$CertificateThumbprint
    )

    Write-BwsLog "Connecting to Microsoft Graph ($AuthMode) with scopes: $($Scopes -join ', ')" "INFO"

    $params = @{}
    if ($TenantId) { $params.TenantId = $TenantId }

    switch ($AuthMode) {
        'DeviceCode' {
            Connect-MgGraph @params -Scopes $Scopes -UseDeviceAuthentication | Out-Null
        }
        'Interactive' {
            Connect-MgGraph @params -Scopes $Scopes | Out-Null
        }
        'ClientCertificate' {
            if (-not $ClientId -or -not $CertificateThumbprint) {
                throw "ClientCertificate requires -ClientId and -CertificateThumbprint."
            }
            Connect-MgGraph @params -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint | Out-Null
        }
        'ManagedIdentity' {
            Connect-MgGraph @params -Identity | Out-Null
        }
    }

    $ctx = Get-MgContext
    if (-not $ctx) { throw "Graph context is empty. Authentication failed?" }
    return $ctx
}

function Connect-BwsAzure {
    param(
        [ValidateSet('DeviceCode','Interactive','ClientCertificate','ManagedIdentity')]
        [string]$AuthMode,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint
    )

    Write-BwsLog "Connecting to Azure (Az.Accounts) via $AuthMode" "INFO"

    $params = @{}
    if ($TenantId) { $params.Tenant = $TenantId }

    switch ($AuthMode) {
        'DeviceCode'      { Connect-AzAccount @params -UseDeviceAuthentication | Out-Null }
        'Interactive'     { Connect-AzAccount @params | Out-Null }
        'ClientCertificate' {
            if (-not $ClientId -or -not $CertificateThumbprint) {
                throw "ClientCertificate requires -ClientId and -CertificateThumbprint."
            }
            Connect-AzAccount @params -ServicePrincipal -ApplicationId $ClientId -CertificateThumbprint $CertificateThumbprint | Out-Null
        }
        'ManagedIdentity' { Connect-AzAccount @params -Identity | Out-Null }
    }

    return (Get-AzContext)
}

# ---------- Execution ----------
function Invoke-BwsChecks {
    param(
        [Parameter(Mandatory)][object[]]$Conditions,
        [string[]]$IncludeProducts,
        [string[]]$IncludeTags,
        [switch]$NoAuth
    )

    $filtered = Get-FilteredConditions -Conditions $Conditions -IncludeProducts $IncludeProducts -IncludeTags $IncludeTags
    if (-not $filtered -or $filtered.Count -eq 0) { throw "No conditions left after filtering." }

    # Preflight modules for filtered run
    Ensure-AndImportModulesForRun -Conditions $filtered

    # Determine required connections
    $needGraph = ($filtered | Where-Object { $_.RequiresGraph -eq $true }).Count -gt 0
    $needAzure = ($filtered | Where-Object { $_.RequiresAzure -eq $true }).Count -gt 0
    $needAD    = ($filtered | Where-Object { $_.RequiresAD -eq $true }).Count -gt 0

    $graphScopes = @()
    if ($needGraph) {
        $filtered | Where-Object { $_.GraphScopes } | ForEach-Object {
            $graphScopes = Get-Union -A $graphScopes -B @($_.GraphScopes)
        }
        if (-not $graphScopes -or $graphScopes.Count -eq 0) {
            $graphScopes = @("Directory.Read.All")
        }
    }

    $context = [ordered]@{
        RunId          = New-BwsRunId
        GraphContext   = $null
        AzContext      = $null
        NeedGraph      = $needGraph
        NeedAzure      = $needAzure
        NeedAD         = $needAD
        StartTime      = Get-Date
        Errors         = @()
    }

    if (-not $NoAuth) {
        if ($needGraph) {
            $context.GraphContext = Connect-BwsGraph -Scopes $graphScopes -TenantId $TenantId -AuthMode $AuthMode -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        if ($needAzure) {
            $context.AzContext = Connect-BwsAzure -AuthMode $AuthMode -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        if ($needAD) {
            # ActiveDirectory already imported by preflight; still verify cmdlet availability
            if (-not (Get-Command Get-ADDomain -ErrorAction SilentlyContinue)) {
                Write-BwsLog "ActiveDirectory cmdlets not available. AD checks may fail (RSAT missing?)." "WARN"
            }
        }
    } else {
        Write-BwsLog "NoAuth set: expecting you already connected (Connect-MgGraph / Connect-AzAccount)." "WARN"
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($c in $filtered) {
        $started = Get-Date
        $status = 'Error'
        $isCompliant = $false
        $actual = $null
        $expected = $c.Expected
        $evidence = $null
        $message = $null

        try {
            if (-not $c.Test -or $c.Test -isnot [scriptblock]) {
                throw "Condition '$($c.Id)' has no valid Test scriptblock."
            }

            $r = & $c.Test -Context $context -Condition $c
            if ($null -eq $r) { throw "Test returned no result object." }

            $isCompliant = [bool]$r.IsCompliant
            $actual      = $r.Actual
            if ($r.PSObject.Properties.Name -contains 'Expected' -and $r.Expected) { $expected = $r.Expected }
            $evidence    = $r.Evidence
            $message     = $r.Message

            $status = if ($isCompliant) { 'Pass' } else { 'Fail' }
        }
        catch {
            $status = 'Error'
            $message = $_.Exception.Message
            $context.Errors += $_.Exception
        }

        $ended = Get-Date
        $results.Add([pscustomobject]@{
            RunId       = $context.RunId
            Product     = $c.Product
            Id          = $c.Id
            Title       = $c.Title
            Severity    = $c.Severity
            Tags        = ($c.Tags -join ', ')
            Status      = $status
            IsCompliant = $isCompliant
            Expected    = ($expected | Out-String).Trim()
            Actual      = ($actual   | Out-String).Trim()
            Evidence    = ($evidence | Out-String).Trim()
            Message     = $message
            Remediation = $c.Remediation
            Started     = $started
            DurationMs  = [int]((New-TimeSpan -Start $started -End $ended).TotalMilliseconds)
        }) | Out-Null
    }

    return [pscustomobject]@{
        Context = [pscustomobject]$context
        Results = @($results)
    }
}

# ---------- Reporting ----------
function Convert-BwsResultsToHtml {
    param(
        [Parameter(Mandatory)][object]$Run,
        [Parameter(Mandatory)][string]$OutFile
    )

    $ctx  = $Run.Context
    $rows = $Run.Results

    $summary = $rows | Group-Object Status | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{ Status = $_.Name; Count = $_.Count }
    }

    $css = @"
    body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
    h1, h2 { margin-bottom: 6px; }
    .meta { color: #555; margin-bottom: 18px; }
    table { border-collapse: collapse; width: 100%; margin: 12px 0 20px 0; }
    th, td { border: 1px solid #ddd; padding: 8px; vertical-align: top; }
    th { background: #f3f3f3; text-align: left; }
    .Pass  { background: #e9f7ef; }
    .Fail  { background: #fff3cd; }
    .Error { background: #f8d7da; }
    .badge { display:inline-block; padding:2px 8px; border-radius: 10px; font-size: 12px; background:#eee; margin-right:6px; }
    .small { font-size: 12px; color:#666; }
    details summary { cursor: pointer; }
    pre { white-space: pre-wrap; word-break: break-word; }
"@

    $metaHtml = @"
    <div class='meta'>
      <div><span class='badge'>RunId</span> $($ctx.RunId)</div>
      <div><span class='badge'>Start</span> $($ctx.StartTime)</div>
      <div>
        <span class='badge'>NeedGraph</span> $($ctx.NeedGraph)
        <span class='badge'>NeedAzure</span> $($ctx.NeedAzure)
        <span class='badge'>NeedAD</span> $($ctx.NeedAD)
      </div>
    </div>
"@

    $summaryHtml = ($summary | ConvertTo-Html -Fragment -PreContent "<h2>Summary</h2>")

    $byProduct = $rows | Sort-Object Product, Severity, Id | Group-Object Product

    $sections = foreach ($g in $byProduct) {
        $prod = $g.Name

        $tblRows = foreach ($r in $g.Group) {
            $cls = $r.Status

            $exp = [System.Net.WebUtility]::HtmlEncode($r.Expected)
            $act = [System.Net.WebUtility]::HtmlEncode($r.Actual)
            $evi = [System.Net.WebUtility]::HtmlEncode($r.Evidence)
            $msg = [System.Net.WebUtility]::HtmlEncode($r.Message)
            $rem = [System.Net.WebUtility]::HtmlEncode($r.Remediation)

@"
<tr class='$cls'>
  <td>$($r.Severity)</td>
  <td><b>$($r.Id)</b><br/><span class='small'>$($r.Tags)</span></td>
  <td>$($r.Title)</td>
  <td><b>$($r.Status)</b><br/><span class='small'>$($r.DurationMs) ms</span></td>
  <td>
    <details><summary>Details</summary>
      <div><b>Expected:</b><pre>$exp</pre></div>
      <div><b>Actual:</b><pre>$act</pre></div>
      <div><b>Evidence:</b><pre>$evi</pre></div>
      <div><b>Message:</b><pre>$msg</pre></div>
    </details>
  </td>
  <td><pre>$rem</pre></td>
</tr>
"@
        }

@"
<h2>$prod</h2>
<table>
  <thead>
    <tr>
      <th>Severity</th>
      <th>Id / Tags</th>
      <th>Title</th>
      <th>Status</th>
      <th>Details</th>
      <th>Remediation</th>
    </tr>
  </thead>
  <tbody>
    $($tblRows -join "`n")
  </tbody>
</table>
"@
    }

    $html = @"
<!doctype html>
<html>
<head>
  <meta charset="utf-8"/>
  <title>BWS Check Report - $($ctx.RunId)</title>
  <style>$css</style>
</head>
<body>
  <h1>BWS Check Report</h1>
  $metaHtml
  $summaryHtml
  $($sections -join "`n")
</body>
</html>
"@

    Set-Content -LiteralPath $OutFile -Value $html -Encoding UTF8
}

# ---------- Optional GUI ----------
function Show-BwsGuiAndRun {
    param(
        [Parameter(Mandatory)][object[]]$Conditions
    )

    if (-not (Test-IsWindows)) {
        throw "GUI is only available on Windows."
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "BWSCheckScript"
    $form.Width = 980
    $form.Height = 640
    $form.StartPosition = "CenterScreen"

    $lblOut = New-Object System.Windows.Forms.Label
    $lblOut.Text = "Output path:"
    $lblOut.Left = 12
    $lblOut.Top = 14
    $lblOut.Width = 90

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Left = 110
    $txtOut.Top = 10
    $txtOut.Width = 700
    $txtOut.Text = $OutputPath

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "..."
    $btnBrowse.Left = 820
    $btnBrowse.Top = 9
    $btnBrowse.Width = 40
    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.SelectedPath = $txtOut.Text
        if ($dlg.ShowDialog() -eq "OK") { $txtOut.Text = $dlg.SelectedPath }
    })

    $lblProd = New-Object System.Windows.Forms.Label
    $lblProd.Text = "Products (filter):"
    $lblProd.Left = 12
    $lblProd.Top = 50
    $lblProd.Width = 140

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Left = 12
    $clb.Top = 72
    $clb.Width = 200
    $clb.Height = 200

    $products = $Conditions | Select-Object -ExpandProperty Product -Unique | Sort-Object
    foreach ($p in $products) { [void]$clb.Items.Add($p, $true) }

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Run checks"
    $btnRun.Left = 12
    $btnRun.Top = 285
    $btnRun.Width = 200

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 230
    $grid.Top = 72
    $grid.Width = 720
    $grid.Height = 500
    $grid.ReadOnly = $true
    $grid.AutoSizeColumnsMode = "Fill"

    $status = New-Object System.Windows.Forms.Label
    $status.Left = 230
    $status.Top = 50
    $status.Width = 720
    $status.Text = "Ready."

    $btnRun.Add_Click({
        try {
            $sel = @()
            for ($i=0; $i -lt $clb.Items.Count; $i++) {
                if ($clb.GetItemChecked($i)) { $sel += [string]$clb.Items[$i] }
            }
            if (-not $sel -or $sel.Count -eq 0) { throw "No product selected." }

            $status.Text = "Running..."
            $form.Refresh()

            $script:OutputPath = $txtOut.Text
            Ensure-Folder -Path $script:OutputPath

            $run = Invoke-BwsChecks -Conditions $Conditions -IncludeProducts $sel -IncludeTags $IncludeTags -NoAuth:$NoAuth

            $reportFile = Join-Path $script:OutputPath ("BWSReport-{0}.html" -f $run.Context.RunId)
            Convert-BwsResultsToHtml -Run $run -OutFile $reportFile

            $grid.DataSource = $run.Results
            $status.Text = "Done. Report: $reportFile"
        } catch {
            $status.Text = "ERROR: " + $_.Exception.Message
        }
    })

    $form.Controls.AddRange(@($lblOut,$txtOut,$btnBrowse,$lblProd,$clb,$btnRun,$status,$grid))
    [void]$form.ShowDialog()
}

# ---------- Main ----------
try {
    $runId = New-BwsRunId
    Ensure-Folder -Path $OutputPath
    $logFile = Join-Path $OutputPath "BWSCheck-$runId.log"
    Start-Transcript -LiteralPath $logFile -Append | Out-Null

    $condPath = Join-Path $PSScriptRoot "BWSConditions.ps1"
    $conditions = Import-BwsConditions -Path $condPath

    if ($Gui) {
        Show-BwsGuiAndRun -Conditions $conditions
        return
    }

    # CLI mode
    $run = Invoke-BwsChecks -Conditions $conditions -IncludeProducts $IncludeProducts -IncludeTags $IncludeTags -NoAuth:$NoAuth

    $reportFile = Join-Path $OutputPath ("BWSReport-{0}.html" -f $run.Context.RunId)
    Convert-BwsResultsToHtml -Run $run -OutFile $reportFile

    Write-BwsLog "Report created: $reportFile" "INFO"
    Write-BwsLog "Log file: $logFile" "INFO"
}
finally {
    try { Stop-Transcript | Out-Null } catch {}
}
