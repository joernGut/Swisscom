<#
.BWSCheckMain.ps1
Zweck:
- Lädt Regeln aus BWSCheckRules.ps1
- Optional GUI zur Auswahl (Out-GridView)
- Führt Checks aus
- Erstellt HTML Report

Wichtige Parameter:
- -NoAuth: keine Authentifizierung starten (z.B. wenn schon via Enterprise App/Pipeline authentifiziert)
- -Gui: Rule-Auswahl via Out-GridView (wenn verfügbar)
- -Rules: explizite Rule-Ids
- -Products: Filter (z.B. 'EntraID','Intune','Teams','SharePoint Online','OneDrive','Exchange Online')
- -OutputPath: Pfad für HTML Report
#>

[CmdletBinding()]
param(
    [switch]$Gui,

    [string[]]$Rules,
    [string[]]$Products,

    [switch]$NoAuth,

    # Falls Main selbst Auth anstoßen soll (über BWSCheckAuth.ps1)
    [ValidateSet('Interactive','DeviceCode','AppOnlyCertificate')]
    [string]$AuthMode = 'Interactive',
    [string]$TenantId,
    [string]$AccountId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$Organization,
    [string]$SharePointAdminUrl,

    [switch]$AllowInstall,

    [string]$OutputPath,

    [switch]$OpenReport,
    [switch]$DisconnectWhenDone
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------
# Load companion scripts
# ---------------------------
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$authPath   = Join-Path $scriptRoot 'BWSCheckAuth.ps1'
$rulesPath  = Join-Path $scriptRoot 'BWSCheckRules.ps1'

if (-not (Test-Path $rulesPath)) { throw "Rules file nicht gefunden: $rulesPath" }
. $rulesPath

if (-not $NoAuth -and (Test-Path $authPath)) {
    . $authPath
}

# ---------------------------
# Helper: Graph wrapper
# ---------------------------
function Invoke-BwsGraphRequest {
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST','PUT','PATCH','DELETE')]
        [string]$Method,
        [Parameter(Mandatory)]
        [string]$PathOrUri
    )

    if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        throw "Invoke-MgGraphRequest nicht verfügbar. Installiere Microsoft.Graph.Authentication und verbinde dich (Connect-MgGraph)."
    }

    $uri = if ($PathOrUri -match '^https?://') { $PathOrUri } else { "https://graph.microsoft.com$PathOrUri" }
    return Invoke-MgGraphRequest -Method $Method -Uri $uri
}

function Test-OutGridViewAvailable {
    return [bool](Get-Command Out-GridView -ErrorAction SilentlyContinue)
}

function Get-DefaultOutputPath {
    $ts = (Get-Date).ToString("yyyyMMdd_HHmmss")
    return Join-Path $scriptRoot ("BWS_Report_{0}.html" -f $ts)
}

# ---------------------------
# Select rules
# ---------------------------
$allRules = @($script:BwsCheckRules)

if ($Products -and $Products.Count -gt 0) {
    $allRules = $allRules | Where-Object { $Products -contains $_.Product }
}

if ($Rules -and $Rules.Count -gt 0) {
    $selected = $allRules | Where-Object { $Rules -contains $_.Id }
}
elseif ($Gui) {
    if (Test-OutGridViewAvailable) {
        $selected = $allRules |
            Select-Object Id,Product,Severity,Name,Description |
            Out-GridView -Title 'BWS Checks auswählen (Mehrfachauswahl möglich)' -PassThru |
            ForEach-Object {
                $allRules | Where-Object Id -eq $_.Id
            }
    }
    else {
        Write-Warning "Out-GridView nicht verfügbar. Starte ohne GUI (alle Regeln)."
        $selected = $allRules
    }
}
else {
    $selected = $allRules
}

if (-not $selected -or $selected.Count -eq 0) {
    throw "Keine Regeln ausgewählt."
}

# ---------------------------
# Aggregate requirements (Scopes/Services)
# ---------------------------
$needGraph = $false
$needExo   = $false
$needTeams = $false
$needSpo   = $false

$scopes = New-Object System.Collections.Generic.HashSet[string]
foreach ($r in $selected) {
    if ($r.Requires.Graph) { $needGraph = $true }
    if ($r.Requires.Exchange) { $needExo = $true }
    if ($r.Requires.Teams) { $needTeams = $true }
    if ($r.Requires.SharePoint) { $needSpo = $true }

    foreach ($s in @($r.MinimumScopes)) { [void]$scopes.Add($s) }
}

# Für interaktive Graph-Auth: wenn Regeln Scopes verlangen, sonst fallback:
$finalScopes = @()
if ($scopes.Count -gt 0) { $finalScopes = $scopes.ToArray() } else { $finalScopes = @('User.Read') }

# ---------------------------
# Auth (optional)
# ---------------------------
$authContext = $null
if (-not $NoAuth) {
    if (Get-Command Connect-BwsServices -ErrorAction SilentlyContinue) {
        $authContext = Connect-BwsServices `
            -Mode $AuthMode `
            -TenantId $TenantId `
            -AccountId $AccountId `
            -ClientId $ClientId `
            -CertificateThumbprint $CertificateThumbprint `
            -Organization $Organization `
            -Scopes $finalScopes `
            -ConnectGraph:($needGraph) `
            -ConnectExchange:($needExo) `
            -ConnectTeams:($needTeams) `
            -ConnectSharePoint:($needSpo) `
            -SharePointAdminUrl $SharePointAdminUrl `
            -AllowInstall:$AllowInstall `
            -Quiet:$false
    }
    else {
        Write-Warning "BWSCheckAuth.ps1 nicht geladen/gefunden. Versuche minimale Graph-Auth (falls nötig)."
        if ($needGraph -and (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
            Connect-MgGraph -Scopes $finalScopes | Out-Null
        }
    }
}
else {
    Write-Host "NoAuth aktiv: Es wird keine Authentifizierung durchgeführt. Es wird von bestehenden Sessions/Token ausgegangen." -ForegroundColor Yellow
}

# ---------------------------
# Execution context for rules
# ---------------------------
$ctx = [pscustomobject]@{
    RunId   = [guid]::NewGuid().ToString()
    Started = (Get-Date)
    Auth    = $authContext
    Helper  = [pscustomobject]@{
        InvokeGraph = { param($m,$p) Invoke-BwsGraphRequest -Method $m -PathOrUri $p }.GetNewClosure()
    }
    Rule    = $null
}

# ---------------------------
# Run rules
# ---------------------------
$results = New-Object System.Collections.Generic.List[object]

foreach ($rule in $selected) {
    Write-Host ("[{0}] {1} - {2}" -f $rule.Product, $rule.Id, $rule.Name) -ForegroundColor Cyan

    $ctx.Rule = $rule

    try {
        $res = & $rule.ScriptBlock $ctx
        if (-not $res) {
            $res = [pscustomobject]@{
                Timestamp=(Get-Date).ToString("s"); RuleId=$rule.Id; RuleName=$rule.Name; Product=$rule.Product
                Severity=$rule.Severity; Status='Error'; Summary='Regel lieferte kein Ergebnisobjekt.'
                Details=$null; Remediation='Prüfe den Rule ScriptBlock.'; Evidence=$null
            }
        }
        $results.Add($res)
    }
    catch {
        $results.Add([pscustomobject]@{
            Timestamp=(Get-Date).ToString("s"); RuleId=$rule.Id; RuleName=$rule.Name; Product=$rule.Product
            Severity=$rule.Severity; Status='Error'; Summary=$_.Exception.Message
            Details=$_.ScriptStackTrace; Remediation='StackTrace prüfen / Berechtigungen / API/Module prüfen.'; Evidence=$null
        })
    }
}

# ---------------------------
# HTML Report
# ---------------------------
if (-not $OutputPath) { $OutputPath = Get-DefaultOutputPath }

$summary = $results |
    Group-Object Status |
    Select-Object Name,Count |
    Sort-Object Name

$counts = @{
    Pass    = ($results | Where-Object Status -eq 'Pass').Count
    Fail    = ($results | Where-Object Status -eq 'Fail').Count
    Warn    = ($results | Where-Object Status -eq 'Warn').Count
    Info    = ($results | Where-Object Status -eq 'Info').Count
    Error   = ($results | Where-Object Status -eq 'Error').Count
    Skipped = ($results | Where-Object Status -eq 'Skipped').Count
}

$css = @"
:root { font-family: Segoe UI, Arial, sans-serif; }
body { margin: 20px; }
h1 { margin-bottom: 4px; }
.meta { color: #555; margin-bottom: 18px; }
.badge { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 12px; margin-right: 6px; border: 1px solid #ddd; }
.Pass { background: #eaffea; }
.Fail { background: #ffecec; }
.Warn { background: #fff6e0; }
.Info { background: #eef5ff; }
.Error { background: #ffecec; border-color: #ffb3b3; }
.Skipped { background: #f2f2f2; }
table { border-collapse: collapse; width: 100%; margin-top: 10px; }
th, td { border: 1px solid #ddd; padding: 8px; vertical-align: top; }
th { background: #f7f7f7; text-align: left; }
.small { font-size: 12px; color: #666; }
details { margin: 10px 0; }
pre { white-space: pre-wrap; word-break: break-word; background: #f7f7f7; padding: 10px; border-radius: 8px; }
"@

function Convert-ObjToPre {
    param($obj)
    if ($null -eq $obj) { return '' }
    try { return ($obj | ConvertTo-Json -Depth 8) } catch { return ($obj | Out-String) }
}

$now = Get-Date
$report = New-Object System.Text.StringBuilder
[void]$report.AppendLine("<html><head><meta charset='utf-8'><title>BWS Check Report</title><style>$css</style></head><body>")
[void]$report.AppendLine("<h1>BWS Check Report</h1>")
[void]$report.AppendLine(("<div class='meta'>RunId: <b>{0}</b> &nbsp;|&nbsp; Start: {1} &nbsp;|&nbsp; Generated: {2}</div>" -f $ctx.RunId, $ctx.Started, $now))

foreach ($k in @('Pass','Fail','Warn','Info','Error','Skipped')) {
    [void]$report.AppendLine(("<span class='badge {0}'>{0}: {1}</span>" -f $k, $counts[$k]))
}

[void]$report.AppendLine("<h2>Ergebnisübersicht</h2>")
[void]$report.AppendLine("<table><tr><th>Status</th><th>Anzahl</th></tr>")
foreach ($s in $summary) {
    [void]$report.AppendLine(("<tr><td>{0}</td><td>{1}</td></tr>" -f $s.Name, $s.Count))
}
[void]$report.AppendLine("</table>")

[void]$report.AppendLine("<h2>Details</h2>")
foreach ($r in ($results | Sort-Object Status, Product, RuleId)) {
    $detailsJson = Convert-ObjToPre $r.Details
    $evidence = if ($r.Evidence) { ($r.Evidence -join "`n") } else { '' }

    [void]$report.AppendLine("<details>")
    [void]$report.AppendLine(("<summary><span class='badge {0}'>{0}</span> <b>[{1}] {2}</b> <span class='small'>({3} / {4})</span></summary>" -f $r.Status, $r.Product, $r.RuleName, $r.RuleId, $r.Severity))
    [void]$report.AppendLine(("<p><b>Summary:</b> {0}</p>" -f [System.Web.HttpUtility]::HtmlEncode($r.Summary)))
    if ($r.Remediation) {
        [void]$report.AppendLine(("<p><b>Remediation:</b> {0}</p>" -f [System.Web.HttpUtility]::HtmlEncode($r.Remediation)))
    }
    if ($evidence) {
        [void]$report.AppendLine("<p><b>Evidence:</b></p>")
        [void]$report.AppendLine("<pre>$([System.Web.HttpUtility]::HtmlEncode($evidence))</pre>")
    }
    if ($detailsJson) {
        [void]$report.AppendLine("<p><b>Details (JSON):</b></p>")
        [void]$report.AppendLine("<pre>$([System.Web.HttpUtility]::HtmlEncode($detailsJson))</pre>")
    }
    [void]$report.AppendLine("</details>")
}

[void]$report.AppendLine("</body></html>")

$report.ToString() | Set-Content -Path $OutputPath -Encoding UTF8
Write-Host "HTML Report erstellt: $OutputPath" -ForegroundColor Green

if ($OpenReport) {
    try { Start-Process $OutputPath | Out-Null } catch { Write-Warning "Konnte Report nicht öffnen: $($_.Exception.Message)" }
}

# ---------------------------
# Disconnect (optional)
# ---------------------------
if ($DisconnectWhenDone -and (Get-Command Disconnect-BwsServices -ErrorAction SilentlyContinue)) {
    Disconnect-BwsServices -Graph:$needGraph -Exchange:$needExo -Teams:$needTeams | Out-Null
    Write-Host "Verbindungen getrennt." -ForegroundColor DarkGray
}
