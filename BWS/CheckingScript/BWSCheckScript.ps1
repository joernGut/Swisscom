<# 
.SYNOPSIS
  BWSCheckScript - Configuration/Compliance checks for Entra ID, Azure, Azure Virtual Desktop, Intune and Active Directory.

.DESCRIPTION
  - Loads conditions from .\BWSConditions.ps1 (same folder)
  - Optional GUI (-Gui)
  - Preflights + installs/imports required modules (including auth modules)
  - Graph: imports ONLY Microsoft.Graph.Authentication (Connect-MgGraph) and ensures Microsoft.Graph is installed
  - Executes checks and generates an HTML report

  Fixes included:
  - Robust GUI error display (even if error occurs before GUI is shown)
  - GUI shows full error details in a dedicated textbox + MessageBox
  - Error output includes file + line + column when possible (fallback parsing from ScriptStackTrace)
  - "Count cannot be found": use @(...).Count for scalar/pipeline normalization
  - "Argument types do not match": converts Context/Condition to the declared parameter types of each Test scriptblock,
    including Hashtable/IDictionary/OrderedDictionary/Generic dictionaries and custom classes via property mapping

.NOTES
  Recommended: PowerShell 7+ on Windows
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

# -----------------------------
# Utilities / Logging / Errors
# -----------------------------
function New-BwsRunId { (Get-Date).ToString('yyyyMMdd-HHmmss') }

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

function New-BwsFolder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Initialize-BwsWinForms {
    if (-not (Test-IsWindows)) { return }
    try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop } catch {}
    try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop } catch {}
}

function Get-BwsSourceLineFromFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$LineNumber
    )
    try {
        if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $null }
        $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
        $idx = $LineNumber - 1
        if ($idx -ge 0 -and $idx -lt $lines.Count) { return $lines[$idx] }
    } catch {}
    return $null
}

function Get-BwsLocationFromError {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $scriptName = $null
    $lineNo = $null
    $col = $null
    $codeLine = $null

    $inv = $ErrorRecord.InvocationInfo
    if ($inv) {
        if ($inv.ScriptName) { $scriptName = $inv.ScriptName }
        if ($inv.ScriptLineNumber -gt 0) { $lineNo = $inv.ScriptLineNumber }
        if ($inv.OffsetInLine -gt 0) { $col = $inv.OffsetInLine }
        if ($inv.Line) { $codeLine = $inv.Line }
    }

    # Fallback: parse ScriptStackTrace if InvocationInfo is missing (common for .NET exceptions)
    if (-not $lineNo -or -not $scriptName) {
        $sst = $ErrorRecord.ScriptStackTrace
        if ($sst) {
            $m = [regex]::Match($sst, '(?<path>[A-Za-z]:\\[^:]+):\s*line\s*(?<line>\d+)', 'IgnoreCase')
            if ($m.Success) {
                if (-not $scriptName) { $scriptName = $m.Groups['path'].Value }
                if (-not $lineNo) { $lineNo = [int]$m.Groups['line'].Value }
            }
        }
    }

    # If we got path + line but no code line, try reading it from file
    if ($scriptName -and $lineNo -and -not $codeLine) {
        $codeLine = Get-BwsSourceLineFromFile -Path $scriptName -LineNumber $lineNo
    }

    return [pscustomobject]@{
        ScriptName = $scriptName
        Line       = $lineNo
        Column     = $col
        Code       = $codeLine
    }
}

function Format-BwsError {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $loc = Get-BwsLocationFromError -ErrorRecord $ErrorRecord

    $inner = $ErrorRecord.Exception.InnerException
    $innerMsg = if ($inner) { $inner.Message } else { $null }

    $stack = $ErrorRecord.ScriptStackTrace

    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add("ERROR: $($ErrorRecord.Exception.Message)") | Out-Null

    if ($loc.ScriptName -or $loc.Line -or $loc.Column) {
        $colText  = if ($loc.Column) { $loc.Column } else { "n/a" }
        $lineText = if ($loc.Line)   { $loc.Line }   else { "n/a" }
        $parts.Add("Location: $($loc.ScriptName) (Line $lineText, Column $colText)") | Out-Null
    }
    if ($loc.Code) { $parts.Add("Code: $($loc.Code)") | Out-Null }
    if ($innerMsg) { $parts.Add("InnerException: $innerMsg") | Out-Null }
    if ($stack)    { $parts.Add("ScriptStackTrace: $stack") | Out-Null }

    return ($parts -join [Environment]::NewLine)
}

function Show-BwsErrorDialog {
    param(
        [Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [string]$Title = "BWSCheckScript - Error"
    )

    $details = Format-BwsError -ErrorRecord $ErrorRecord

    if (Test-IsWindows) {
        Initialize-BwsWinForms
        try {
            [System.Windows.Forms.MessageBox]::Show(
                $details,
                $Title,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            return
        } catch {}
    }

    Write-BwsLog $details "ERROR"
}

# -----------------------------
# PowerShellGet / Modules
# -----------------------------
function Set-TlsForWindowsPowerShell {
    # Windows PowerShell 5.1 often needs TLS 1.2 for PSGallery
    if ($PSVersionTable.PSEdition -eq 'Desktop') {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 }
        catch { Write-BwsLog "Could not set TLS 1.2: $($_.Exception.Message)" "WARN" }
    }
}

function Install-BwsNuGetProviderIfMissing {
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

function Set-BwsPSGalleryTrustedIfAutoInstall {
    if (-not $AutoInstallModules) { return }
    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction Stop
        if ($repo.InstallationPolicy -ne 'Trusted') {
            Write-BwsLog "Setting PSGallery InstallationPolicy to Trusted..." "WARN"
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }
    } catch {
        Write-BwsLog "Could not validate PSGallery repository: $($_.Exception.Message)" "WARN"
    }
}

function Install-BwsModuleIfNeeded {
    param(
        [Parameter(Mandatory)][string]$Name,
        [switch]$AutoInstall,
        [version]$MinimumVersion,
        [version]$RequiredVersion
    )

    $available = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending
    if ($available) {
        if ($RequiredVersion) {
            if ($available | Where-Object { $_.Version -eq $RequiredVersion }) { return $true }
        } elseif ($MinimumVersion) {
            if ($available[0].Version -ge $MinimumVersion) { return $true }
        } else {
            return $true
        }
    }

    if (-not $AutoInstall) {
        $vMsg = if ($RequiredVersion) { " (RequiredVersion: $RequiredVersion)" }
        elseif ($MinimumVersion) { " (MinimumVersion: $MinimumVersion)" }
        else { "" }
        Write-BwsLog "Module '$Name' is missing or does not meet the required version$vMsg. Install manually or run with -AutoInstallModules." "WARN"
        return $false
    }

    Set-TlsForWindowsPowerShell
    Install-BwsNuGetProviderIfMissing
    Set-BwsPSGalleryTrustedIfAutoInstall

    $installParams = @{
        Name         = $Name
        Scope        = 'CurrentUser'
        Force        = $true
        AllowClobber = $true
        ErrorAction  = 'Stop'
    }
    if ($RequiredVersion) { $installParams.RequiredVersion = $RequiredVersion }
    elseif ($MinimumVersion) { $installParams.MinimumVersion = $MinimumVersion }

    Write-BwsLog "Installing module '$Name' (CurrentUser)..." "INFO"
    Install-Module @installParams

    $available2 = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending
    if (-not $available2) { return $false }

    if ($RequiredVersion) { return [bool]($available2 | Where-Object { $_.Version -eq $RequiredVersion }) }
    if ($MinimumVersion) { return [bool]($available2[0].Version -ge $MinimumVersion) }
    return $true
}

function Import-ModuleSafe {
    param(
        [Parameter(Mandatory)][string]$Name,
        [version]$RequiredVersion
    )
    $params = @{ Name = $Name; ErrorAction = 'Stop'; Force = $true }
    if ($RequiredVersion) { $params.RequiredVersion = $RequiredVersion }
    Import-Module @params | Out-Null
}

# -----------------------------
# Conditions: Load / Filter
# -----------------------------
function Import-BwsConditions {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "BWSConditions not found: $Path" }

    $conds = & $Path
    if (-not $conds) { throw "BWSConditions returned no conditions." }

    $arr = @($conds)
    if ($arr.Count -eq 0) { throw "BWSConditions returned an empty set." }

    return $arr
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

    $filtered = @($Conditions)

    if ($IncludeProducts -and @($IncludeProducts).Count -gt 0) {
        $set = @($IncludeProducts) | ForEach-Object { $_.Trim() }
        $filtered = @($filtered | Where-Object { $set -contains $_.Product })
    }

    if ($IncludeTags -and @($IncludeTags).Count -gt 0) {
        $tags = @($IncludeTags) | ForEach-Object { $_.Trim() }
        $filtered = @($filtered | Where-Object {
            $_.Tags -and (@(@($_.Tags) | Where-Object { $tags -contains $_ })).Count -gt 0
        })
    }

    return @($filtered)
}

# -----------------------------
# Graph stack (robust)
# -----------------------------
function Import-BwsGraphModules {
    # Install + import ONLY authentication. Ensure Microsoft.Graph is installed (submodules can autoload).
    $okAuth = Install-BwsModuleIfNeeded -Name "Microsoft.Graph.Authentication" -AutoInstall:$AutoInstallModules
    if (-not $okAuth) { throw "Microsoft.Graph.Authentication is missing. Install it or run with -AutoInstallModules." }

    Import-ModuleSafe -Name "Microsoft.Graph.Authentication"

    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw "Connect-MgGraph not found even after importing Microsoft.Graph.Authentication."
    }

    $okMg = Install-BwsModuleIfNeeded -Name "Microsoft.Graph" -AutoInstall:$AutoInstallModules
    if (-not $okMg) { throw "Microsoft.Graph is missing. Install it or run with -AutoInstallModules." }
}

# -----------------------------
# Module preflight for conditions
# -----------------------------
function Get-ModulesRequiredForConditions {
    param([Parameter(Mandatory)][object[]]$Conditions)

    $needAzure = @($Conditions | Where-Object { $_.RequiresAzure -eq $true }).Count -gt 0
    $needAD    = @($Conditions | Where-Object { $_.RequiresAD -eq $true }).Count -gt 0

    $modules = New-Object System.Collections.Generic.List[string]

    if ($needAzure) {
        [void]$modules.Add("Az.Accounts")
        [void]$modules.Add("Az.Resources")
        $needAvd = @($Conditions | Where-Object { $_.Product -eq 'AVD' }).Count -gt 0
        if ($needAvd) { [void]$modules.Add("Az.DesktopVirtualization") }
    }

    if ($needAD) { [void]$modules.Add("ActiveDirectory") }

    foreach ($c in @($Conditions)) {
        if ($c.PSObject.Properties.Name -contains 'RequiredModules' -and $c.RequiredModules) {
            foreach ($m in @($c.RequiredModules)) {
                if ($m) { [void]$modules.Add([string]$m) }
            }
        }
    }

    return @($modules)
}

function Import-BwsModulesForRun {
    param([Parameter(Mandatory)][object[]]$Conditions)

    $needGraph = @($Conditions | Where-Object { $_.RequiresGraph -eq $true }).Count -gt 0
    if ($needGraph) { Import-BwsGraphModules }

    $modules = @(Get-ModulesRequiredForConditions -Conditions $Conditions)
    if (@($modules).Count -eq 0) { return }

    # Keep order but dedupe
    $seen = @{}
    $orderedUnique = foreach ($m in $modules) {
        if (-not $seen.ContainsKey($m)) { $seen[$m] = $true; $m }
    }

    foreach ($m in $orderedUnique) {
        $ok = Install-BwsModuleIfNeeded -Name $m -AutoInstall:$AutoInstallModules
        if (-not $ok) { throw "Missing module '$m'. Install it or run with -AutoInstallModules." }
        Import-ModuleSafe -Name $m
    }
}

# -----------------------------
# Auth / Connections
# -----------------------------
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
        'DeviceCode'      { Connect-MgGraph @params -Scopes $Scopes -UseDeviceAuthentication | Out-Null }
        'Interactive'     { Connect-MgGraph @params -Scopes $Scopes | Out-Null }
        'ClientCertificate' {
            if (-not $ClientId -or -not $CertificateThumbprint) {
                throw "ClientCertificate requires -ClientId and -CertificateThumbprint."
            }
            Connect-MgGraph @params -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint | Out-Null
        }
        'ManagedIdentity' { Connect-MgGraph @params -Identity | Out-Null }
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

# -----------------------------
# Type conversion (fix for Argument types do not match)
# -----------------------------
function ConvertTo-BwsHashtable {
    param([Parameter(Mandatory)][object]$InputObject)

    if ($InputObject -is [hashtable]) { return $InputObject }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $InputObject.Keys) { $h[$k] = $InputObject[$k] }
        return $h
    }

    $h2 = @{}
    foreach ($p in $InputObject.PSObject.Properties) { $h2[$p.Name] = $p.Value }
    return $h2
}

function ConvertTo-BwsOrderedDictionary {
    param([Parameter(Mandatory)][object]$InputObject)

    $src = ConvertTo-BwsHashtable -InputObject $InputObject
    $od = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($k in $src.Keys) { [void]$od.Add($k, $src[$k]) }
    return $od
}

function ConvertTo-BwsGenericDictionary {
    param(
        [Parameter(Mandatory)][object]$InputObject,
        [Parameter(Mandatory)][Type]$TargetType
    )

    $src = ConvertTo-BwsHashtable -InputObject $InputObject
    $genArgs = $TargetType.GetGenericArguments()
    $keyType = $genArgs[0]
    $valType = $genArgs[1]

    $dictType = [System.Collections.Generic.Dictionary`2].MakeGenericType($keyType, $valType)
    $dict = [Activator]::CreateInstance($dictType)

    foreach ($k in $src.Keys) {
        $kk = [System.Management.Automation.LanguagePrimitives]::ConvertTo($k, $keyType)
        $vv = $src[$k]
        if ($null -ne $vv) { $vv = [System.Management.Automation.LanguagePrimitives]::ConvertTo($vv, $valType) }
        $dict.Add($kk, $vv)
    }
    return $dict
}

function ConvertTo-BwsCustomObjectType {
    param(
        [Parameter(Mandatory)][object]$InputObject,
        [Parameter(Mandatory)][Type]$TargetType
    )

    if ($TargetType.IsAbstract -or $TargetType.IsInterface) {
        throw "TargetType '$($TargetType.FullName)' is abstract/interface - cannot instantiate."
    }

    $ctor = $TargetType.GetConstructor(@())
    if (-not $ctor) {
        throw "TargetType '$($TargetType.FullName)' has no parameterless constructor."
    }

    $src = ConvertTo-BwsHashtable -InputObject $InputObject

    # case-insensitive lookup
    $lookup = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($k in $src.Keys) { $lookup[[string]$k] = $src[$k] }

    $obj = [Activator]::CreateInstance($TargetType)

    foreach ($prop in $TargetType.GetProperties([Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::Instance)) {
        if (-not $prop.CanWrite) { continue }
        if (-not $lookup.ContainsKey($prop.Name)) { continue }

        $val = $lookup[$prop.Name]
        try {
            if ($null -ne $val -and $prop.PropertyType -ne [object] -and -not $prop.PropertyType.IsInstanceOfType($val)) {
                $val = [System.Management.Automation.LanguagePrimitives]::ConvertTo($val, $prop.PropertyType)
            }
            $prop.SetValue($obj, $val)
        } catch {
            # ignore property conversion errors
        }
    }

    return $obj
}

function ConvertTo-BwsTargetType {
    param(
        [Parameter(Mandatory)][object]$Value,
        [Parameter(Mandatory)][Type]$TargetType
    )

    if (-not $TargetType -or $TargetType -eq [object]) {
        return [pscustomobject](ConvertTo-BwsHashtable -InputObject $Value)
    }

    if ($TargetType.IsInstanceOfType($Value)) { return $Value }

    if ($TargetType -eq [hashtable] -or $TargetType.FullName -eq 'System.Collections.Hashtable') {
        return (ConvertTo-BwsHashtable -InputObject $Value)
    }

    if ($TargetType.FullName -eq 'System.Collections.Specialized.OrderedDictionary') {
        return (ConvertTo-BwsOrderedDictionary -InputObject $Value)
    }

    if ([System.Collections.IDictionary].IsAssignableFrom($TargetType)) {
        return (ConvertTo-BwsHashtable -InputObject $Value)
    }

    if ($TargetType.IsGenericType) {
        $def = $TargetType.GetGenericTypeDefinition().FullName
        if ($def -like 'System.Collections.Generic.Dictionary`2' -or
            $def -like 'System.Collections.Generic.IDictionary`2' -or
            $def -like 'System.Collections.Generic.IReadOnlyDictionary`2') {
            return (ConvertTo-BwsGenericDictionary -InputObject $Value -TargetType $TargetType)
        }
    }

    if ($TargetType.FullName -in @('System.Management.Automation.PSCustomObject','System.Management.Automation.PSObject')) {
        return [pscustomobject](ConvertTo-BwsHashtable -InputObject $Value)
    }

    if ($TargetType.IsClass) {
        try { return (ConvertTo-BwsCustomObjectType -InputObject $Value -TargetType $TargetType) } catch {}
    }

    return [System.Management.Automation.LanguagePrimitives]::ConvertTo($Value, $TargetType)
}

function Get-BwsScriptBlockParamInfo {
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock)

    $list = New-Object System.Collections.Generic.List[object]
    try {
        $pb = $ScriptBlock.Ast.ParamBlock
        if (-not $pb) { return @() }

        foreach ($p in @($pb.Parameters)) {
            $name = $p.Name.VariablePath.UserPath
            $type = $p.StaticType
            $list.Add([pscustomobject]@{ Name = $name; Type = $type }) | Out-Null
        }
    } catch {
        return @()
    }
    return @($list)
}

function Invoke-BwsConditionTest {
    param(
        [Parameter(Mandatory)][scriptblock]$Test,
        [Parameter(Mandatory)][hashtable]$ContextHash,
        [Parameter(Mandatory)][object]$Condition
    )

    $paramInfo = @(Get-BwsScriptBlockParamInfo -ScriptBlock $Test)

    # Prefer mapping by param name. If not possible, default: [0]=Context, [1]=Condition
    $ctxIdx  = -1
    $condIdx = -1
    if ($paramInfo.Count -gt 0) {
        for ($i=0; $i -lt $paramInfo.Count; $i++) {
            $n = $paramInfo[$i].Name
            if ($ctxIdx  -lt 0 -and $n -match '^(Context|Ctx)$')    { $ctxIdx  = $i }
            if ($condIdx -lt 0 -and $n -match '^(Condition|Cond)$') { $condIdx = $i }
        }
    }
    if ($ctxIdx -lt 0 -and $condIdx -lt 0) { $ctxIdx = 0; $condIdx = 1 }
    elseif ($ctxIdx -ge 0 -and $condIdx -lt 0) { $condIdx = if ($ctxIdx -eq 0) { 1 } else { 0 } }
    elseif ($condIdx -ge 0 -and $ctxIdx -lt 0) { $ctxIdx  = if ($condIdx -eq 0) { 1 } else { 0 } }

    $argCount = if ($paramInfo.Count -gt 0) { $paramInfo.Count } else { 2 }
    $maxIdx = [Math]::Max($ctxIdx, $condIdx)
    if ($argCount -lt ($maxIdx + 1)) { $argCount = $maxIdx + 1 }

    $args = @()
    for ($i=0; $i -lt $argCount; $i++) { $args += $null }

    if ($paramInfo.Count -gt 0 -and $ctxIdx -lt $paramInfo.Count) {
        $args[$ctxIdx] = ConvertTo-BwsTargetType -Value $ContextHash -TargetType $paramInfo[$ctxIdx].Type
    } else {
        $args[$ctxIdx] = [pscustomobject]$ContextHash
    }

    if ($paramInfo.Count -gt 0 -and $condIdx -lt $paramInfo.Count) {
        $args[$condIdx] = ConvertTo-BwsTargetType -Value $Condition -TargetType $paramInfo[$condIdx].Type
    } else {
        $args[$condIdx] = $Condition
    }

    try {
        return & $Test @args
    } catch {
        $expected = if ($paramInfo.Count -gt 0) {
            ($paramInfo | ForEach-Object { "$($_.Name): $($_.Type.FullName)" }) -join '; '
        } else { '(no param block detected)' }
        throw "Argument types do not match when invoking condition test. Expected parameters: $expected. Inner error: $($_.Exception.Message)"
    }
}

# -----------------------------
# Execution
# -----------------------------
function Invoke-BwsChecks {
    param(
        [Parameter(Mandatory)][object[]]$Conditions,
        [string[]]$IncludeProducts,
        [string[]]$IncludeTags,
        [switch]$NoAuth
    )

    $filtered = @(Get-FilteredConditions -Conditions $Conditions -IncludeProducts $IncludeProducts -IncludeTags $IncludeTags)
    if ($filtered.Count -eq 0) { throw "No conditions left after filtering." }

    Import-BwsModulesForRun -Conditions $filtered

    $needGraph = @($filtered | Where-Object { $_.RequiresGraph -eq $true }).Count -gt 0
    $needAzure = @($filtered | Where-Object { $_.RequiresAzure -eq $true }).Count -gt 0
    $needAD    = @($filtered | Where-Object { $_.RequiresAD -eq $true }).Count -gt 0

    $graphScopes = @()
    if ($needGraph) {
        foreach ($c in @($filtered | Where-Object { $_.GraphScopes })) {
            $graphScopes = Get-Union -A $graphScopes -B @($c.GraphScopes)
        }
        if (@($graphScopes).Count -eq 0) { $graphScopes = @("Directory.Read.All") }
    }

    $context = @{}
    $context.RunId        = New-BwsRunId
    $context.GraphContext = $null
    $context.AzContext    = $null
    $context.NeedGraph    = $needGraph
    $context.NeedAzure    = $needAzure
    $context.NeedAD       = $needAD
    $context.StartTime    = Get-Date
    $context.Errors       = @()

    if (-not $NoAuth) {
        if ($needGraph) { $context.GraphContext = Connect-BwsGraph -Scopes $graphScopes -TenantId $TenantId -AuthMode $AuthMode -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint }
        if ($needAzure) { $context.AzContext    = Connect-BwsAzure -AuthMode $AuthMode -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint }
        if ($needAD) {
            if (-not (Get-Command Get-ADDomain -ErrorAction SilentlyContinue)) {
                Write-BwsLog "ActiveDirectory cmdlets not available. AD checks may fail (RSAT missing?)." "WARN"
            }
        }
    } else {
        Write-BwsLog "NoAuth set: expecting you already connected (Connect-MgGraph / Connect-AzAccount)." "WARN"
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($c in @($filtered)) {
        $started = Get-Date
        $status = 'Error'
        $isCompliant = $false
        $actual = $null
        $expected = $c.Expected
        $evidence = $null
        $message = $null
        $location = $null

        try {
            if (-not $c.Test -or $c.Test -isnot [scriptblock]) {
                throw "Condition '$($c.Id)' has no valid Test scriptblock."
            }

            $r = Invoke-BwsConditionTest -Test $c.Test -ContextHash $context -Condition $c
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

            $loc = Get-BwsLocationFromError -ErrorRecord $_
            $colText  = if ($loc.Column) { $loc.Column } else { "n/a" }
            $lineText = if ($loc.Line)   { $loc.Line }   else { "n/a" }
            if ($loc.ScriptName -or $loc.Line -or $loc.Column) {
                $location = "$($loc.ScriptName) (Line $lineText, Column $colText)"
            }
            if ($loc.Code) { $evidence = $loc.Code }

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
            Location    = $location
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

# -----------------------------
# Reporting (HTML)
# -----------------------------
function Convert-BwsResultsToHtml {
    param(
        [Parameter(Mandatory)][object]$Run,
        [Parameter(Mandatory)][string]$OutFile
    )

    $ctx  = $Run.Context
    $rows = @($Run.Results)

    $summary = @($rows | Group-Object Status | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{ Status = $_.Name; Count = $_.Count }
    })

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

    $byProduct = @($rows | Sort-Object Product, Severity, Id | Group-Object Product)

    $sections = foreach ($g in $byProduct) {
        $prod = $g.Name

        $tblRows = foreach ($r in @($g.Group)) {
            $cls = $r.Status

            $exp = [System.Net.WebUtility]::HtmlEncode([string]$r.Expected)
            $act = [System.Net.WebUtility]::HtmlEncode([string]$r.Actual)
            $evi = [System.Net.WebUtility]::HtmlEncode([string]$r.Evidence)
            $msg = [System.Net.WebUtility]::HtmlEncode([string]$r.Message)
            $loc = [System.Net.WebUtility]::HtmlEncode([string]$r.Location)
            $rem = [System.Net.WebUtility]::HtmlEncode([string]$r.Remediation)

@"
<tr class='$cls'>
  <td>$($r.Severity)</td>
  <td><b>$($r.Id)</b><br/><span class='small'>$($r.Tags)</span></td>
  <td>$($r.Title)</td>
  <td><b>$($r.Status)</b><br/><span class='small'>$($r.DurationMs) ms</span></td>
  <td>
    <details><summary>Details</summary>
      <div><b>Location:</b><pre>$loc</pre></div>
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

# -----------------------------
# Optional GUI
# -----------------------------
function Show-BwsGuiAndRun {
    param([Parameter(Mandatory)][object[]]$Conditions)

    if (-not (Test-IsWindows)) { throw "GUI is only available on Windows." }

    Initialize-BwsWinForms

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "BWSCheckScript"
    $form.Width = 1100
    $form.Height = 720
    $form.StartPosition = "CenterScreen"

    $lblOut = New-Object System.Windows.Forms.Label
    $lblOut.Text = "Output path:"
    $lblOut.Left = 12
    $lblOut.Top = 14
    $lblOut.Width = 90

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Left = 110
    $txtOut.Top = 10
    $txtOut.Width = 760
    $txtOut.Text = $OutputPath

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "..."
    $btnBrowse.Left = 880
    $btnBrowse.Top = 9
    $btnBrowse.Width = 40
    $btnBrowse.Add_Click({
        try {
            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
            $dlg.SelectedPath = $txtOut.Text
            if ($dlg.ShowDialog() -eq "OK") { $txtOut.Text = $dlg.SelectedPath }
        } catch {
            Show-BwsErrorDialog -ErrorRecord $_ -Title "Browse Error"
        }
    })

    $lblProd = New-Object System.Windows.Forms.Label
    $lblProd.Text = "Products (filter):"
    $lblProd.Left = 12
    $lblProd.Top = 50
    $lblProd.Width = 140

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Left = 12
    $clb.Top = 72
    $clb.Width = 220
    $clb.Height = 220

    $products = @($Conditions | Select-Object -ExpandProperty Product -Unique | Sort-Object)
    foreach ($p in $products) { [void]$clb.Items.Add($p, $true) }

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Run checks"
    $btnRun.Left = 12
    $btnRun.Top = 305
    $btnRun.Width = 220

    $status = New-Object System.Windows.Forms.Label
    $status.Left = 250
    $status.Top = 50
    $status.Width = 820
    $status.Text = "Ready."

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 250
    $grid.Top = 72
    $grid.Width = 820
    $grid.Height = 420
    $grid.ReadOnly = $true
    $grid.AutoSizeColumnsMode = "Fill"

    $lblErr = New-Object System.Windows.Forms.Label
    $lblErr.Left = 12
    $lblErr.Top = 350
    $lblErr.Width = 220
    $lblErr.Text = "Last error (details):"

    $txtErr = New-Object System.Windows.Forms.TextBox
    $txtErr.Left = 12
    $txtErr.Top = 372
    $txtErr.Width = 1058
    $txtErr.Height = 300
    $txtErr.Multiline = $true
    $txtErr.ScrollBars = "Vertical"
    $txtErr.ReadOnly = $true

    $btnRun.Add_Click({
        try {
            $txtErr.Text = ""

            $sel = @()
            for ($i=0; $i -lt $clb.Items.Count; $i++) {
                if ($clb.GetItemChecked($i)) { $sel += [string]$clb.Items[$i] }
            }
            if (@($sel).Count -eq 0) { throw "No product selected." }

            $status.Text = "Running..."
            $form.Refresh()

            $script:OutputPath = $txtOut.Text
            New-BwsFolder -Path $script:OutputPath

            $run = Invoke-BwsChecks -Conditions $Conditions -IncludeProducts $sel -IncludeTags $IncludeTags -NoAuth:$NoAuth

            $reportFile = Join-Path $script:OutputPath ("BWSReport-{0}.html" -f $run.Context.RunId)
            Convert-BwsResultsToHtml -Run $run -OutFile $reportFile

            $grid.DataSource = $run.Results
            $status.Text = "Done. Report: $reportFile"
        } catch {
            $details = Format-BwsError -ErrorRecord $_
            $txtErr.Text = $details

            $loc = Get-BwsLocationFromError -ErrorRecord $_
            $colText  = if ($loc.Column) { $loc.Column } else { "n/a" }
            $lineText = if ($loc.Line)   { $loc.Line }   else { "n/a" }
            $status.Text = ("ERROR: {0} (Line {1}, Col {2})" -f $_.Exception.Message, $lineText, $colText)

            Show-BwsErrorDialog -ErrorRecord $_ -Title "BWSCheckScript - Run Error"
        }
    })

    $form.Controls.AddRange(@(
        $lblOut,$txtOut,$btnBrowse,
        $lblProd,$clb,$btnRun,
        $status,$grid,
        $lblErr,$txtErr
    ))

    [void]$form.ShowDialog()
}

# -----------------------------
# Main
# -----------------------------
try {
    $runId = New-BwsRunId
    New-BwsFolder -Path $OutputPath

    $logFile = Join-Path $OutputPath "BWSCheck-$runId.log"
    Start-Transcript -LiteralPath $logFile -Append | Out-Null

    $condPath = Join-Path $PSScriptRoot "BWSConditions.ps1"
    $conditions = @(Import-BwsConditions -Path $condPath)

    if ($Gui) {
        Show-BwsGuiAndRun -Conditions $conditions
        return
    }

    $run = Invoke-BwsChecks -Conditions $conditions -IncludeProducts $IncludeProducts -IncludeTags $IncludeTags -NoAuth:$NoAuth

    $reportFile = Join-Path $OutputPath ("BWSReport-{0}.html" -f $run.Context.RunId)
    Convert-BwsResultsToHtml -Run $run -OutFile $reportFile

    Write-BwsLog "Report created: $reportFile" "INFO"
    Write-BwsLog "Log file: $logFile" "INFO"
}
catch {
    if ($Gui) {
        Show-BwsErrorDialog -ErrorRecord $_ -Title "BWSCheckScript - Startup Error"
    }
    Write-BwsLog (Format-BwsError -ErrorRecord $_) "ERROR"
    throw
}
finally {
    try { Stop-Transcript | Out-Null } catch {}
}
