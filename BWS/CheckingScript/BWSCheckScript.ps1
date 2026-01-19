<# 
.SYNOPSIS
  BWSCheckScript - Configuration/Compliance checks for Entra ID, Azure, Azure Virtual Desktop, Intune and Active Directory.

.DESCRIPTION
  - Loads conditions from .\BWSConditions.ps1 (same folder)
  - Optional GUI (-Gui)
  - Preflights + installs/imports required modules (including auth modules)
  - IMPORTANT: To avoid Microsoft.Graph <-> Microsoft.Graph.Authentication exact-version dependency issues,
               this script imports ONLY Microsoft.Graph.Authentication (for Connect-MgGraph) and ensures
               Microsoft.Graph is installed (for submodules). It does NOT import the Microsoft.Graph meta module.
  - Executes checks and generates an HTML report

  Fixes included:
  - "Count cannot be found on this object": all Count usage on pipeline/scalar results is wrapped with @(...)
  - "Argument types do not match": condition Test scriptblocks are invoked POSITIONALLY with arguments converted
    to the exact declared parameter types (hashtable, OrderedDictionary, generic Dictionary<TKey,TValue>, etc.).

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

function Set-TlsForWindowsPowerShell {
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
            Write-BwsLog "Setting PSGallery InstallationPolicy to Trusted (may prompt in some environments)..." "WARN"
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
    try {
        $params = @{ Name = $Name; ErrorAction = 'Stop'; Force = $true }
        if ($RequiredVersion) { $params.RequiredVersion = $RequiredVersion }
        Import-Module @params | Out-Null
        $loaded = Get-Module -Name $Name -ErrorAction SilentlyContinue
        $verMsg = if ($loaded) { $loaded.Version.ToString() } else { "?" }
        Write-BwsLog "Imported module: $Name ($verMsg)" "DEBUG"
    } catch {
        throw "Failed to import module '$Name': $($_.Exception.Message)"
    }
}

function Get-Union {
    param([string[]]$A, [string[]]$B)
    @($A + $B | Where-Object { $_ -and $_.Trim() } | Select-Object -Unique)
}

# ---------- Safe conversions (for type-strict condition parameters) ----------
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

    # Supports Dictionary<TKey,TValue> and IDictionary<TKey,TValue> declared types by creating Dictionary<TKey,TValue>
    $src = ConvertTo-BwsHashtable -InputObject $InputObject

    if (-not $TargetType.IsGenericType) { throw "TargetType is not generic." }
    $genArgs = $TargetType.GetGenericArguments()
    if ($genArgs.Count -ne 2) { throw "Generic dictionary must have 2 generic arguments." }

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

function ConvertTo-BwsTargetType {
    param(
        [Parameter(Mandatory)][object]$Value,
        [Parameter(Mandatory)][Type]$TargetType,
        [ValidateSet('Context','Condition')]
        [string]$Role = 'Context'
    )

    # If no type / object, keep as is (but prefer PSCustomObject for convenience)
    if (-not $TargetType -or $TargetType -eq [object]) {
        if ($Role -eq 'Context') { return [pscustomobject](ConvertTo-BwsHashtable -InputObject $Value) }
        return $Value
    }

    # Hashtable / IDictionary
    if ($TargetType -eq [hashtable] -or $TargetType.FullName -eq 'System.Collections.Hashtable') {
        return (ConvertTo-BwsHashtable -InputObject $Value)
    }
    if ([System.Collections.IDictionary].IsAssignableFrom($TargetType)) {
        # OrderedDictionary is IDictionary but not interchangeable; handle explicitly below
        if ($TargetType.FullName -eq 'System.Collections.Specialized.OrderedDictionary') {
            return (ConvertTo-BwsOrderedDictionary -InputObject $Value)
        }
        # Hashtable implements IDictionary
        return (ConvertTo-BwsHashtable -InputObject $Value)
    }

    # Generic dictionary requested (Dictionary<TKey,TValue>, IDictionary<TKey,TValue>, IReadOnlyDictionary<TKey,TValue>)
    if ($TargetType.IsGenericType) {
        $def = $TargetType.GetGenericTypeDefinition().FullName
        if ($def -like 'System.Collections.Generic.Dictionary`2' -or
            $def -like 'System.Collections.Generic.IDictionary`2' -or
            $def -like 'System.Collections.Generic.IReadOnlyDictionary`2') {
            return (ConvertTo-BwsGenericDictionary -InputObject $Value -TargetType $TargetType)
        }
    }

    # PSCustomObject / PSObject types
    if ($TargetType.FullName -eq 'System.Management.Automation.PSCustomObject' -or
        $TargetType.FullName -eq 'System.Management.Automation.PSObject') {
        return [pscustomobject](ConvertTo-BwsHashtable -InputObject $Value)
    }

    # Last resort: try PowerShell conversion
    try {
        return [System.Management.Automation.LanguagePrimitives]::ConvertTo($Value, $TargetType)
    } catch {
        # Fallback to PSCustomObject (least surprising for most checks)
        if ($Role -eq 'Context') { return [pscustomobject](ConvertTo-BwsHashtable -InputObject $Value) }
        return [pscustomobject](ConvertTo-BwsHashtable -InputObject $Value)
    }
}

function Get-BwsScriptBlockParamInfo {
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock)

    $info = New-Object System.Collections.Generic.List[object]
    try {
        $pb = $ScriptBlock.Ast.ParamBlock
        if (-not $pb) { return @() }

        foreach ($p in @($pb.Parameters)) {
            $name = $p.Name.VariablePath.UserPath
            $type = $p.StaticType
            $info.Add([pscustomobject]@{ Name = $name; Type = $type }) | Out-Null
        }
    } catch {
        return @()
    }
    return @($info)
}

function Invoke-BwsConditionTest {
    <#
      This is the hard fix for: "Argument types do not match"
      - We inspect the scriptblock param types and CONVERT context/condition to the declared types.
      - We then invoke POSITIONALLY to avoid named-binding issues.
    #>
    param(
        [Parameter(Mandatory)][scriptblock]$Test,
        [Parameter(Mandatory)][hashtable]$ContextHash,
        [Parameter(Mandatory)][object]$Condition
    )

    $paramInfo = @(Get-BwsScriptBlockParamInfo -ScriptBlock $Test)

    # Decide mapping: by name if possible; else assume [0]=Context, [1]=Condition
    $ctxIdx  = -1
    $condIdx = -1

    if ($paramInfo.Count -gt 0) {
        for ($i=0; $i -lt $paramInfo.Count; $i++) {
            $n = $paramInfo[$i].Name
            if ($ctxIdx -lt 0 -and $n -match '^(Context|Ctx)$') { $ctxIdx = $i }
            if ($condIdx -lt 0 -and $n -match '^(Condition|Cond)$') { $condIdx = $i }
        }
    }

    if ($ctxIdx -lt 0 -and $condIdx -lt 0) {
        $ctxIdx  = 0
        $condIdx = 1
    } elseif ($ctxIdx -ge 0 -and $condIdx -lt 0) {
        $condIdx = if ($ctxIdx -eq 0) { 1 } else { 0 }
    } elseif ($condIdx -ge 0 -and $ctxIdx -lt 0) {
        $ctxIdx = if ($condIdx -eq 0) { 1 } else { 0 }
    }

    # Determine how many args to pass (at least up to the max idx we touch)
    $argCount = 0
    if ($paramInfo.Count -gt 0) {
        $argCount = $paramInfo.Count
    } else {
        $argCount = 2
    }
    $maxIdx = [Math]::Max($ctxIdx, $condIdx)
    if ($argCount -lt ($maxIdx + 1)) { $argCount = $maxIdx + 1 }

    $args = @()
    for ($i=0; $i -lt $argCount; $i++) { $args += $null }

    # Prepare base objects
    $ctxBase  = $ContextHash
    $condBase = $Condition

    # Convert to declared types if we have them
    if ($paramInfo.Count -gt 0 -and $ctxIdx -ge 0 -and $ctxIdx -lt $paramInfo.Count) {
        $t = $paramInfo[$ctxIdx].Type
        $args[$ctxIdx] = ConvertTo-BwsTargetType -Value $ctxBase -TargetType $t -Role 'Context'
    } else {
        $args[$ctxIdx] = [pscustomobject]$ContextHash
    }

    if ($paramInfo.Count -gt 0 -and $condIdx -ge 0 -and $condIdx -lt $paramInfo.Count) {
        $t2 = $paramInfo[$condIdx].Type
        $args[$condIdx] = ConvertTo-BwsTargetType -Value $condBase -TargetType $t2 -Role 'Condition'
    } else {
        $args[$condIdx] = $Condition
    }

    try {
        return & $Test @args
    } catch {
        # Add diagnostics so you see the real expected types in the log/HTML message
        $expected = if ($paramInfo.Count -gt 0) {
            ($paramInfo | ForEach-Object { "$($_.Name): $($_.Type.FullName)" }) -join '; '
        } else {
            '(no param block detected)'
        }

        throw "Argument types do not match when invoking condition test. Expected parameters: $expected. Inner error: $($_.Exception.Message)"
    }
}

# ---------- Filtering / Conditions ----------
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

function Import-BwsConditions {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "BWSConditions not found: $Path" }

    $conds = & $Path
    if (-not $conds) { throw "BWSConditions returned no conditions." }

    $arr = @($conds)
    if ($arr.Count -eq 0) { throw "BWSConditions returned an empty set." }

    return $arr
}

# ---------- Graph Stack ----------
function Import-BwsGraphModules {
    $okAuth = Install-BwsModuleIfNeeded -Name "Microsoft.Graph.Authentication" -AutoInstall:$AutoInstallModules
    if (-not $okAuth) { throw "Microsoft.Graph.Authentication is missing. Install it or run with -AutoInstallModules." }

    Import-ModuleSafe -Name "Microsoft.Graph.Authentication"
    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw "Connect-MgGraph not found even after importing Microsoft.Graph.Authentication."
    }

    $okMg = Install-BwsModuleIfNeeded -Name "Microsoft.Graph" -AutoInstall:$AutoInstallModules
    if (-not $okMg) { throw "Microsoft.Graph is missing. Install it or run with -AutoInstallModules." }

    Write-BwsLog "Microsoft.Graph is installed. Graph submodules will auto-load on demand (meta module import skipped)." "INFO"
}

# ---------- Module Preflight ----------
function Get-ModulesRequiredForConditions {
    param([Parameter(Mandatory)][object[]]$Conditions)

    $needAzure = @($Conditions | Where-Object { $_.RequiresAzure -eq $true }).Count -gt 0
    $needAD    = @($Conditions | Where-Object { $_.RequiresAD -eq $true }).Count -gt 0

    $modules = New-Object System.Collections.Generic.List[string]

    if ($needAzure) {
        [void]$modules.Add("Az.Accounts")
        [void]$modules.Add("Az.Resources")
        $needAvdModule = @($Conditions | Where-Object { $_.Product -eq 'AVD' }).Count -gt 0
        if ($needAvdModule) { [void]$modules.Add("Az.DesktopVirtualization") }
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
    if ($needGraph) {
        Write-BwsLog "Preparing Microsoft Graph module stack..." "INFO"
        Import-BwsGraphModules
    }

    $modules = @(Get-ModulesRequiredForConditions -Conditions $Conditions)
    if (@($modules).Count -eq 0) { return }

    $seen = @{}
    $orderedUnique = foreach ($m in $modules) {
        if (-not $seen.ContainsKey($m)) { $seen[$m] = $true; $m }
    }

    Write-BwsLog "Required modules: $($orderedUnique -join ', ')" "INFO"

    foreach ($m in $orderedUnique) {
        $ok = Install-BwsModuleIfNeeded -Name $m -AutoInstall:$AutoInstallModules
        if (-not $ok) { throw "Missing module '$m'. Install it or run with -AutoInstallModules." }
        Import-ModuleSafe -Name $m
    }

    if (@($Conditions | Where-Object { $_.RequiresAzure -eq $true }).Count -gt 0) {
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
        'DeviceCode' { Connect-MgGraph @params -Scopes $Scopes -UseDeviceAuthentication | Out-Null }
        'Interactive' { Connect-MgGraph @params -Scopes $Scopes | Out-Null }
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

# ---------- Execution ----------
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

    # Context in canonical form: hashtable
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
        if ($needGraph) {
            $context.GraphContext = Connect-BwsGraph -Scopes $graphScopes -TenantId $TenantId -AuthMode $AuthMode -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        if ($needAzure) {
            $context.AzContext = Connect-BwsAzure -AuthMode $AuthMode -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
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

        try {
            if (-not $c.Test -or $c.Test -isnot [scriptblock]) {
                throw "Condition '$($c.Id)' has no valid Test scriptblock."
            }

            # Critical: this is where Argument types do not match is fixed
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
            $rem = [System.Net.WebUtility]::HtmlEncode([string]$r.Remediation)

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
    param([Parameter(Mandatory)][object[]]$Conditions)

    if (-not (Test-IsWindows)) { throw "GUI is only available on Windows." }

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

    $products = @($Conditions | Select-Object -ExpandProperty Product -Unique | Sort-Object)
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
            $status.Text = "ERROR: " + $_.Exception.Message
        }
    })

    $form.Controls.AddRange(@($lblOut,$txtOut,$btnBrowse,$lblProd,$clb,$btnRun,$status,$grid))
    [void]$form.ShowDialog()
}

# ---------- Main ----------
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
finally {
    try { Stop-Transcript | Out-Null } catch {}
}
