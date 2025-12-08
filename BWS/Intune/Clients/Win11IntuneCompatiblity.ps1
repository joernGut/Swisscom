<#
.SYNOPSIS
    Checks basic Windows 11 hardware requirements and
    client compatibility for Microsoft Intune, with optional Excel export.

.DESCRIPTION
    Windows 11 checks:
    - TPM: presence, status, SpecVersion (Win32_Tpm), Windows 11 requirement (>= 2.0)
    - Firmware: BIOS vs. UEFI
    - Secure Boot: enabled / disabled / not supported
    - RAM: >= 4 GB
    - CPU: 64-bit, >= 2 cores, >= 1 GHz (no full Microsoft CPU support list check)
    - System drive: >= 64 GB total capacity

    Intune checks (client compatibility):
    - Windows client OS (no server), version:
        * Windows 10, build >= 14393 (1607) or
        * Windows 11 (build >= 22000)
    - Intune Management Extension prerequisites:
        * Supported OS version (see above)
        * .NET Framework >= 4.7.2
        * PowerShell 5.1 available

    Optional:
    - Export results to a formatted Excel file using the ImportExcel module
      (table, filters, frozen header, autosize).

    Tested on Windows 10/11 with Windows PowerShell 5.1 and PowerShell 7+.
#>

[CmdletBinding()]
param(
    [switch]$ExportExcel,
    [string]$ExcelPath = "$env:USERPROFILE\Desktop\Win11-Intune-Compatibility.xlsx",
    [switch]$ShowExcel
)

#region Helper functions (generic)

function Test-IsAdmin {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function New-CheckResult {
    param(
        [Parameter(Mandatory)][string]$Check,
        [bool]$Ok,
        [string]$Value,
        [string]$Required,
        [string]$Notes
    )

    [PSCustomObject]@{
        Check    = $Check
        Ok       = $Ok
        Value    = $Value
        Required = $Required
        Notes    = $Notes
    }
}

#endregion

#region Windows 11 hardware checks

function Test-Tpm {
    [CmdletBinding()]
    param()

    $ok       = $false
    $value    = ''
    $required = 'TPM 2.0, enabled'
    $notes    = ''

    try {
        # WMI/CIM class Win32_Tpm provides SpecVersion and vendor info
        $tpmCim = Get-CimInstance -Namespace 'root/cimv2/Security/MicrosoftTpm' -ClassName 'Win32_Tpm' -ErrorAction Stop
    }
    catch {
        return New-CheckResult -Check 'TPM' -Ok $false -Value 'Error querying TPM' -Required $required -Notes $_.Exception.Message
    }

    if (-not $tpmCim) {
        return New-CheckResult -Check 'TPM' -Ok $false -Value 'No TPM found' -Required $required -Notes 'Win32_Tpm returned no data.'
    }

    $specRaw = $tpmCim.SpecVersion
    $value   = "SpecVersion: $specRaw; Manufacturer: $($tpmCim.ManufacturerIdTxt) $($tpmCim.ManufacturerVersion)"

    if ([string]::IsNullOrWhiteSpace($specRaw)) {
        $notes = 'TPM present, but SpecVersion could not be determined.'
        $ok    = $false
    }
    else {
        # Use the first entry as the "major" TPM version (e.g. "2.0" from "2.0, 0, 1.16")
        $majorString = ($specRaw -split ',')[0].Trim()
        try {
            $version = [version]$majorString
            if ($version -ge [version]'2.0') {
                $ok    = $true
                $notes = "TPM version >= 2.0 (found: $majorString)"
            }
            else {
                $ok    = $false
                $notes = "TPM version < 2.0 (found: $majorString)"
            }
        }
        catch {
            $ok    = $false
            $notes = "TPM SpecVersion could not be parsed: $majorString"
        }
    }

    return New-CheckResult -Check 'TPM' -Ok $ok -Value $value -Required $required -Notes $notes
}

function Test-Firmware {
    [CmdletBinding()]
    param()

    $required = 'UEFI firmware'
    $notes    = ''
    $ok       = $false
    $value    = ''

    try {
        $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Control'
        $peType  = (Get-ItemProperty -Path $regPath -Name 'PEFirmwareType' -ErrorAction Stop).PEFirmwareType

        switch ($peType) {
            1 { $value = 'BIOS (Legacy)'; $ok = $false; $notes = 'System boots in legacy BIOS mode.' }
            2 { $value = 'UEFI';          $ok = $true;  $notes = 'System uses UEFI firmware.' }
            default {
                $value = "Unknown firmware type ($peType)"
                $ok    = $false
                $notes = 'PEFirmwareType registry value is set but has an unknown value.'
            }
        }
    }
    catch {
        $value = 'Error detecting firmware type'
        $ok    = $false
        $notes = $_.Exception.Message
    }

    New-CheckResult -Check 'Firmware (UEFI)' -Ok $ok -Value $value -Required $required -Notes $notes
}

function Test-SecureBoot {
    [CmdletBinding()]
    param()

    $required = 'Secure Boot enabled'
    $ok       = $false
    $value    = ''
    $notes    = ''

    try {
        # Cmdlet is only available on UEFI systems
        $sb = Confirm-SecureBootUEFI -ErrorAction Stop
        if ($sb -eq $true) {
            $ok    = $true
            $value = 'Secure Boot: enabled'
            $notes = 'Secure Boot is enabled.'
        }
        elseif ($sb -eq $false) {
            $ok    = $false
            $value = 'Secure Boot: disabled'
            $notes = 'Secure Boot is available, but disabled.'
        }
        else {
            $ok    = $false
            $value = "Secure Boot: unknown state ($sb)"
            $notes = 'Cmdlet returned an unexpected result.'
        }
    }
    catch {
        $ok    = $false
        $value = 'Secure Boot not supported or error'
        $notes = $_.Exception.Message
    }

    New-CheckResult -Check 'Secure Boot' -Ok $ok -Value $value -Required $required -Notes $notes
}

function Test-Ram {
    [CmdletBinding()]
    param()

    $required = '>= 4 GB RAM'
    $ok       = $false
    $value    = ''
    $notes    = ''

    try {
        $cs   = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $ramB = [int64]$cs.TotalPhysicalMemory
        $ramGB = [math]::Round($ramB / 1GB, 2)
        $value = "$ramGB GB"

        if ($ramB -ge 4GB) {
            $ok    = $true
            $notes = 'Sufficient physical memory.'
        }
        else {
            $ok    = $false
            $notes = 'Less than 4 GB RAM.'
        }
    }
    catch {
        $ok    = $false
        $value = 'Error querying RAM'
        $notes = $_.Exception.Message
    }

    New-CheckResult -Check 'RAM' -Ok $ok -Value $value -Required $required -Notes $notes
}

function Test-Cpu {
    [CmdletBinding()]
    param()

    $required = '64-bit, >= 2 cores, >= 1 GHz (no full CPU support list check)'
    $ok       = $true
    $notes    = @()
    $value    = ''

    try {
        $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop | Select-Object -First 1

        $name         = $cpu.Name
        $cores        = $cpu.NumberOfCores
        $clockMHz     = $cpu.MaxClockSpeed
        $addressWidth = $cpu.AddressWidth  # 64 = 64-bit

        $clockGHz = [math]::Round($clockMHz / 1000, 2)
        $value    = "$name | Cores: $cores | Clock: $clockGHz GHz | Architecture: $addressWidth-bit"

        if ($addressWidth -lt 64) {
            $ok = $false
            $notes += 'CPU is not 64-bit.'
        }
        if ($cores -lt 2) {
            $ok = $false
            $notes += 'Less than 2 physical cores.'
        }
        if ($clockMHz -lt 1000) {
            $ok = $false
            $notes += 'Clock frequency < 1 GHz.'
        }

        if (-not $notes) {
            $notes = 'CPU meets the basic minimum requirements. Note: not checked against Microsoft''s official CPU support list.'
        }
        else {
            $notes = $notes -join ' '
        }
    }
    catch {
        $ok    = $false
        $value = 'Error querying CPU'
        $notes = $_.Exception.Message
    }

    New-CheckResult -Check 'CPU (basic requirements)' -Ok $ok -Value $value -Required $required -Notes $notes
}

function Test-SystemDrive {
    [CmdletBinding()]
    param()

    $required = '>= 64 GB total capacity on system drive'
    $ok       = $false
    $value    = ''
    $notes    = ''

    try {
        $systemDrive = $env:SystemDrive
        $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter ("DeviceID='$systemDrive'") -ErrorAction Stop

        $sizeGB = [math]::Round($disk.Size / 1GB, 2)
        $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
        $value  = "Size: $sizeGB GB, free: $freeGB GB"

        if ($disk.Size -ge 64GB) {
            $ok    = $true
            $notes = 'System drive has at least 64 GB total capacity.'
        }
        else {
            $ok    = $false
            $notes = 'System drive is smaller than 64 GB.'
        }
    }
    catch {
        $ok    = $false
        $value = 'Error querying system drive'
        $notes = $_.Exception.Message
    }

    New-CheckResult -Check 'System drive' -Ok $ok -Value $value -Required $required -Notes $notes
}

function Test-OsInfo {
    [CmdletBinding()]
    param()

    $required = 'Windows 10/11, 64-bit (for upgrade path; actual HW requirements checked separately)'
    $ok       = $true
    $notes    = ''
    $value    = ''

    try {
        $os   = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $ver  = [version]$os.Version
        $caption = $os.Caption
        $arch = $os.OSArchitecture

        $value = "$caption ($($os.Version)), Architecture: $arch"

        if ($ver.Major -lt 10) {
            $ok    = $false
            $notes = 'Windows version < 10. Upgrade to Windows 11 is not directly supported.'
        }
        elseif ($arch -notmatch '64') {
            $ok    = $false
            $notes = 'Operating system is not 64-bit.'
        }
        else {
            $notes = 'Operating system is generally upgradable; see other lines for full HW checks.'
        }
    }
    catch {
        $ok    = $false
        $value = 'Error querying OS'
        $notes = $_.Exception.Message
    }

    New-CheckResult -Check 'Operating system' -Ok $ok -Value $value -Required $required -Notes $notes
}

#endregion

#region Intune checks

function Test-IntuneOsSupport {
    [CmdletBinding()]
    param()

    $required = 'Windows client: Windows 10 (1607+) or Windows 11 (no server OS)'
    $ok       = $false
    $value    = ''
    $notes    = ''

    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop

        $caption     = $os.Caption
        $versionStr  = $os.Version
        $buildStr    = $os.BuildNumber
        $productType = $os.ProductType   # 1=Workstation, 2=Domain Controller, 3=Server

        $ver   = [version]$versionStr
        [int]$build = $buildStr

        $value = "$caption (Version: $versionStr, Build: $buildStr, ProductType: $productType)"

        if ($productType -ne 1) {
            $ok    = $false
            $notes = 'Intune does not support Windows Server OS (client OS only).'
        }
        elseif ($ver.Major -lt 10) {
            $ok    = $false
            $notes = 'OS is older than Windows 10 – not intended for modern Intune management.'
        }
        else {
            # Windows 10 and 11 both use version 10.x; distinguish by build number
            if ($build -ge 22000) {
                # Windows 11
                $ok    = $true
                $notes = 'Windows 11 client OS, supported by Intune.'
            }
            elseif ($build -ge 14393) {
                # Windows 10 1607 or later
                $ok    = $true
                $notes = 'Windows 10 client (1607 or later). Intune support is available; some features may require newer builds.'
            }
            else {
                $ok    = $false
                $notes = 'Windows 10 build < 1607 – not suitable for Intune/Intune Management Extension.'
            }
        }
    }
    catch {
        $ok    = $false
        $value = 'Error querying OS (Intune check)'
        $notes = $_.Exception.Message
    }

    New-CheckResult -Check 'Intune: OS support' -Ok $ok -Value $value -Required $required -Notes $notes
}

function Get-DotNetReleaseVersion {
    try {
        $key = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction Stop
        return [int]$key.Release
    }
    catch {
        return $null
    }
}

function Get-DotNetFriendlyVersion {
    param(
        [int]$Release
    )

    # Rough mapping for display only
    if ($Release -ge 528040) { return '4.8 or higher' }
    elseif ($Release -ge 461808) { return '4.7.2' }
    elseif ($Release -ge 461308) { return '4.7.1' }
    elseif ($Release -ge 460798) { return '4.7' }
    else { return "Release key: $Release (older than 4.7.2)" }
}

function Test-IntuneImePrereqs {
    [CmdletBinding()]
    param()

    $required = '.NET >= 4.7.2, PowerShell 5.1, Windows 10 (1607+) or Windows 11'
    $ok       = $true
    $notes    = @()
    $value    = ''

    # OS check again (can differ slightly from generic OS check)
    try {
        $os   = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $ver  = [version]$os.Version
        [int]$build = $os.BuildNumber
    }
    catch {
        $ok    = $false
        $notes += 'OS version could not be determined.'
        $build = 0
    }

    if ($build -lt 14393) {
        $ok    = $false
        $notes += 'OS build < 14393 (Windows 10 1607) – Intune Management Extension is not supported.'
    }

    # Check .NET version
    $release = Get-DotNetReleaseVersion
    if ($null -eq $release) {
        $ok    = $false
        $notes += '.NET Framework v4 (Full) not found.'
        $dotnetInfo = 'unknown'
    }
    else {
        $dotnetInfo = Get-DotNetFriendlyVersion -Release $release
        if ($release -lt 461808) {
            $ok    = $false
            $notes += '.NET Framework < 4.7.2 – Intune Management Extension requires >= 4.7.2.'
        }
    }

    # Check PowerShell 5.1 (for script execution via IME)
    $psCurrent = $PSVersionTable.PSVersion
    $ps51Installed = Test-Path "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"

    if (-not $ps51Installed) {
        $ok    = $false
        $notes += 'Windows PowerShell 5.1 (powershell.exe) was not found.'
    }

    $value = "OS build: $build; .NET: $dotnetInfo; PS session: $psCurrent; PS 5.1 installed: $ps51Installed"

    if (-not $notes) {
        $notes = 'All basic prerequisites for the Intune Management Extension are met (from local machine perspective).'
    }
    else {
        $notes = $notes -join ' '
    }

    New-CheckResult -Check 'Intune: Management Extension prerequisites' -Ok $ok -Value $value -Required $required -Notes $notes
}

#endregion Intune checks

#region Excel export (optional, ImportExcel module)

function Export-CompatReportToExcel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Results,

        [Parameter(Mandatory)]
        [string]$Path,

        [switch]$Show
    )

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Warning 'The ImportExcel module is required for formatted Excel output. Install it with: Install-Module -Name ImportExcel -Scope CurrentUser'
        return
    }

    Import-Module ImportExcel -ErrorAction Stop

    # Prepare formatted Excel output: table, filters, frozen header, autosize
    $excelParams = @{
        Path          = $Path
        WorksheetName = 'CompatReport'
        AutoSize      = $true
        AutoFilter    = $true
        FreezeTopRow  = $true
        BoldTopRow    = $true
        TableName     = 'CompatReport'
        TableStyle    = 'Medium9'    # can be changed as desired
        ClearSheet    = $true
    }

    if ($Show) {
        $excelParams.Show = $true
    }

    $Results |
        Sort-Object Check |
        Export-Excel @excelParams

    Write-Host "Excel report written to $Path" -ForegroundColor Cyan
}

#endregion Excel export

#region Main

Write-Host '=== Windows 11 & Intune Compatibility Check (local PC) ===' -ForegroundColor Cyan
Write-Host ''

if (-not (Test-IsAdmin)) {
    Write-Warning 'It is recommended to run this script as Administrator – otherwise TPM/Secure Boot/registry queries may fail.'
    Write-Host ''
}

$results = @()

# OS / Intune-specific checks
$results += Test-OsInfo
$results += Test-IntuneOsSupport
$results += Test-IntuneImePrereqs

# Hardware / Windows 11
$results += Test-Cpu
$results += Test-Ram
$results += Test-SystemDrive
$results += Test-Firmware
$results += Test-SecureBoot
$results += Test-Tpm

$results | Format-Table -AutoSize

# Optional Excel export
if ($ExportExcel) {
    Export-CompatReportToExcel -Results $results -Path $ExcelPath -Show:$ShowExcel
}

Write-Host ''
$failed = $results | Where-Object { $_.Ok -eq $false }

if (-not $failed) {
    Write-Host '=> Result: This PC meets all Windows 11 and Intune requirements checked by this script.' -ForegroundColor Green
}
else {
    Write-Host '=> Result: This PC does NOT meet all checked requirements.' -ForegroundColor Yellow
    Write-Host 'Details for failed checks:' -ForegroundColor Yellow
    $failed | Select-Object Check, Value, Required, Notes | Format-List
}

Write-Host ''
Write-Host 'Note: CPU compatibility is only checked at a basic level. Intune tenant/licensing/MDM settings cannot be validated from the local client.' -ForegroundColor DarkGray

#endregion Main
