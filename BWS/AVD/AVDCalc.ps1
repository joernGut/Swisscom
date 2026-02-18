<#
================================================================================
  Author: Jörn Gutting
  Date:   2025-02-17
  Script: Azure Virtual Desktop (AVD) Sizing Calculator (WPF GUI) - PowerShell 7
  Version: 2.2.1.1

================================================================================
#>

#requires -Version 7.0

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptVersion  = '2.2.0'
$ScriptBuildUtc = '2025-02-17 16:00:00Z'

#region Ensure STA for WPF
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  $proc = Start-Process -FilePath 'pwsh' -ArgumentList @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', "`"$PSCommandPath`""
  ) -PassThru -Wait
  exit $proc.ExitCode
}
#endregion

#region Assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
#endregion

#region UI message helpers
function Write-UiInfo {
  param([Parameter(Mandatory)][string]$Message, [string]$Title = 'AVD Sizing Calculator')
  [System.Windows.MessageBox]::Show($Message, $Title, 'OK', 'Information') | Out-Null
}
function Write-UiWarning {
  param([Parameter(Mandatory)][string]$Message, [string]$Title = 'AVD Sizing Calculator')
  [System.Windows.MessageBox]::Show($Message, $Title, 'OK', 'Warning') | Out-Null
}
function Write-UiError {
  param([Parameter(Mandatory)][string]$Message, [string]$Title = 'AVD Sizing Calculator')
  [System.Windows.MessageBox]::Show($Message, $Title, 'OK', 'Error') | Out-Null
}
#endregion

#region Safe parsing
function ConvertTo-DoubleSafe {
  param([AllowNull()][string]$Text, [double]$Default = 0)
  $t = ($Text ?? '').Trim()
  if ([string]::IsNullOrWhiteSpace($t)) { return $Default }
  $t = $t -replace [char]0x00A0, '' -replace [char]0x202F, '' -replace "[\s']", '' -replace ',', '.'
  $v = 0.0
  if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v)) { return $v }
  return $Default
}
function ConvertTo-IntSafe {
  param([AllowNull()][string]$Text, [int]$Default = 0)
  $t = ($Text ?? '').Trim()
  if ([string]::IsNullOrWhiteSpace($t)) { return $Default }
  $t = $t -replace [char]0x00A0, '' -replace [char]0x202F, '' -replace "[\s']", ''
  if ($t -match '^\d{1,3}([.,]\d{3})+$') { $t2 = $t -replace '[\.,]', ''; $iv = 0; if ([int]::TryParse($t2, [ref]$iv)) { return $iv } }
  $t = $t -replace ',', '.'
  $dv = 0.0
  if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture, [ref]$dv)) { return [int][Math]::Floor($dv) }
  $digitsOnly = ($t -replace '\D', '')
  if (-not [string]::IsNullOrWhiteSpace($digitsOnly)) { $iv2 = 0; if ([int]::TryParse($digitsOnly, [ref]$iv2)) { return $iv2 } }
  return $Default
}
#endregion

#region Small helpers
function Get-CeilingInt { param([Parameter(Mandatory)][double]$Value); [int][Math]::Ceiling($Value) }
function Get-ComboText {
  param([Parameter(Mandatory)][System.Windows.Controls.ComboBox]$Combo)
  if (-not $Combo.SelectedItem) { return '' }
  $item = $Combo.SelectedItem
  try { if ($item.Content) { return [string]$item.Content } } catch {}
  return [string]$item
}
function ConvertTo-ArmRegionName {
  param([Parameter(Mandatory)][string]$LocationText)
  $norm = $LocationText.Trim().ToLowerInvariant()
  if ([string]::IsNullOrWhiteSpace($norm)) { return $norm }
  $normNoSpace = ($norm -replace '\s+', '') -replace '-', ''
  $map = @{
    'switzerlandnorth'='switzerlandnorth'; 'switzerland north'='switzerlandnorth'
    'switzerlandwest'='switzerlandwest';   'switzerland west'='switzerlandwest'
    'westeurope'='westeurope';             'west europe'='westeurope'
    'northeurope'='northeurope';           'north europe'='northeurope'
    'germanywestcentral'='germanywestcentral'; 'germany west central'='germanywestcentral'
    'francecentral'='francecentral';       'france central'='francecentral'
    'uksouth'='uksouth'; 'uk south'='uksouth'; 'eastus'='eastus'; 'east us'='eastus'
    'eastus2'='eastus2'; 'east us 2'='eastus2'; 'westus2'='westus2'; 'west us 2'='westus2'
    'westus3'='westus3'; 'west us 3'='westus3'; 'centralus'='centralus'; 'central us'='centralus'
  }
  if ($map.ContainsKey($norm)) { return $map[$norm] }
  if ($map.ContainsKey($normNoSpace)) { return $map[$normNoSpace] }
  return $normNoSpace
}
function Get-CurrencyCode {
  param([Parameter(Mandatory)][System.Windows.Controls.ComboBox]$Combo)
  $t = (Get-ComboText -Combo $Combo).Trim().ToUpperInvariant()
  switch ($t) { 'USD' { return 'USD' } 'EUR' { return 'EUR' } 'CHF' { return 'CHF' } default { return 'USD' } }
}
function Get-TextDebugInfo {
  param([AllowNull()][string]$Text)
  $t = if ($null -eq $Text) { '<null>' } else { $Text }
  $chars = @(); foreach ($ch in ($t.ToCharArray())) { $code = [int][char]$ch; $chars += "$(if($ch -match '\s'){'<ws>'}else{$ch})=0x$($code.ToString('X4'))" }
  [pscustomobject]@{ Value=$t; Length=$t.Length; Codes=($chars -join ', ') }
}
#endregion

#region Application Profiles
# Each app profile defines additional resource overhead per user/instance
$script:ApplicationCatalog = [ordered]@{
  # --- Client Applications ---
  'Microsoft 365 (Word/Excel/PPT)' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=0.3; RamMBPerUser=512; DiskIOPSPerUser=2
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Standard Office suite. Included in baseline Medium/Heavy workload profiles.'
  }
  'Microsoft Teams' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=0.5; RamMBPerUser=768; DiskIOPSPerUser=3
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Teams with AV optimisation. Use media optimisation redirects where possible.'
  }
  'Web Browser (Edge/Chrome)' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=0.4; RamMBPerUser=1024; DiskIOPSPerUser=2
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Modern browsers are memory-hungry. Limit tab count via GPO if possible.'
  }
  'Microsoft Outlook' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=0.2; RamMBPerUser=384; DiskIOPSPerUser=3
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Cached mode increases FSLogix profile size significantly.'
  }
  'PDF Editor (Acrobat/Foxit)' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=0.1; RamMBPerUser=256; DiskIOPSPerUser=1
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Minimal overhead for typical usage.'
  }
  'ERP Client (SAP GUI / Dynamics)' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=0.3; RamMBPerUser=512; DiskIOPSPerUser=2
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Thin-client ERP access. Backend processing runs on separate servers.'
  }
  'Power BI Desktop' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=1.0; RamMBPerUser=2048; DiskIOPSPerUser=5
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Memory-intensive for large datasets. Consider Power BI Service instead.'
  }

  # --- Database Engines (on session host) ---
  'SQL Server Express/Developer' = [pscustomobject]@{
    Category='Database'; CpuOverheadPerUser=2.0; RamMBPerUser=4096; DiskIOPSPerUser=50
    DiskGBPerUser=50; RequiresGPU=$false; RequiresDataDisk=$true
    Notes='SQL Server best practice: E-series VM, 8:1 memory:vCore, Premium SSD v2 data disks, separate data/log/tempdb. Not recommended for multi-session pooled.'
    PreferredVmSeries='E'; MinMemoryToVcoreRatio=8; DataDiskType='Premium SSD v2'
    DataDiskMinIOPS=5000; DataDiskMinMBps=200
  }
  'SQL Server Standard/Enterprise' = [pscustomobject]@{
    Category='Database'; CpuOverheadPerUser=4.0; RamMBPerUser=8192; DiskIOPSPerUser=100
    DiskGBPerUser=100; RequiresGPU=$false; RequiresDataDisk=$true
    Notes='Production SQL: E-series or M-series, 8:1+ memory:vCore, separate data/log/tempdb on Premium SSD v2 or Ultra Disk. Personal host pool strongly recommended.'
    PreferredVmSeries='E'; MinMemoryToVcoreRatio=8; DataDiskType='Premium SSD v2 or Ultra Disk'
    DataDiskMinIOPS=10000; DataDiskMinMBps=400
  }
  'PostgreSQL' = [pscustomobject]@{
    Category='Database'; CpuOverheadPerUser=1.5; RamMBPerUser=4096; DiskIOPSPerUser=40
    DiskGBPerUser=40; RequiresGPU=$false; RequiresDataDisk=$true
    Notes='PostgreSQL on session host: use shared_buffers=25% RAM. Premium SSD data disk.'
    PreferredVmSeries='E'; MinMemoryToVcoreRatio=4; DataDiskType='Premium SSD v2'
    DataDiskMinIOPS=3000; DataDiskMinMBps=125
  }
  'MySQL / MariaDB' = [pscustomobject]@{
    Category='Database'; CpuOverheadPerUser=1.5; RamMBPerUser=3072; DiskIOPSPerUser=30
    DiskGBPerUser=30; RequiresGPU=$false; RequiresDataDisk=$true
    Notes='InnoDB buffer pool = 70-80% RAM. Use Premium SSD data disks.'
    PreferredVmSeries='E'; MinMemoryToVcoreRatio=4; DataDiskType='Premium SSD v2'
    DataDiskMinIOPS=3000; DataDiskMinMBps=125
  }
  'SQLite / MS Access (local DB)' = [pscustomobject]@{
    Category='Database'; CpuOverheadPerUser=0.5; RamMBPerUser=512; DiskIOPSPerUser=15
    DiskGBPerUser=5; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='File-based DB. Profile/local storage. No separate data disk needed.'
  }

  # --- Development Tools ---
  'Visual Studio' = [pscustomobject]@{
    Category='Development'; CpuOverheadPerUser=2.0; RamMBPerUser=4096; DiskIOPSPerUser=10
    DiskGBPerUser=20; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Heavy IDE. Consider Personal pool. IntelliSense/compilation are CPU-intensive.'
  }
  'VS Code' = [pscustomobject]@{
    Category='Development'; CpuOverheadPerUser=0.5; RamMBPerUser=1024; DiskIOPSPerUser=5
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Lightweight IDE. Extensions can increase memory usage.'
  }
  'Docker Desktop' = [pscustomobject]@{
    Category='Development'; CpuOverheadPerUser=2.0; RamMBPerUser=4096; DiskIOPSPerUser=20
    DiskGBPerUser=40; RequiresGPU=$false; RequiresDataDisk=$true
    Notes='Requires nested virtualisation. Personal pool only. Data disk for images.'
    DataDiskType='Premium SSD v2'; DataDiskMinIOPS=3000; DataDiskMinMBps=125
  }
  'Git / Build Tools' = [pscustomobject]@{
    Category='Development'; CpuOverheadPerUser=0.5; RamMBPerUser=512; DiskIOPSPerUser=5
    DiskGBPerUser=10; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Compilation spikes CPU. Use SSD-backed storage for repos.'
  }

  # --- CAD / GPU Applications ---
  'AutoCAD / AutoCAD LT' = [pscustomobject]@{
    Category='CAD/GPU'; CpuOverheadPerUser=2.0; RamMBPerUser=4096; DiskIOPSPerUser=10
    DiskGBPerUser=0; RequiresGPU=$true; RequiresDataDisk=$false
    Notes='2D: NV-series sufficient. 3D: NVads A10 v5 recommended.'
    PreferredVmSeries='NV'
  }
  'Revit / 3ds Max' = [pscustomobject]@{
    Category='CAD/GPU'; CpuOverheadPerUser=4.0; RamMBPerUser=8192; DiskIOPSPerUser=15
    DiskGBPerUser=0; RequiresGPU=$true; RequiresDataDisk=$false
    Notes='Heavy 3D. NV-series with dedicated GPU. Personal pool recommended.'
    PreferredVmSeries='NV'
  }
  'SolidWorks / CATIA' = [pscustomobject]@{
    Category='CAD/GPU'; CpuOverheadPerUser=4.0; RamMBPerUser=8192; DiskIOPSPerUser=15
    DiskGBPerUser=0; RequiresGPU=$true; RequiresDataDisk=$false
    Notes='Engineering CAD. NV-series A10/T4. Personal pool. ISV GPU certification required.'
    PreferredVmSeries='NV'
  }
  'Video Editing (Premiere/DaVinci)' = [pscustomobject]@{
    Category='CAD/GPU'; CpuOverheadPerUser=4.0; RamMBPerUser=16384; DiskIOPSPerUser=30
    DiskGBPerUser=100; RequiresGPU=$true; RequiresDataDisk=$true
    Notes='Heavy GPU + storage. NV-series + Premium SSD v2 data disks. Personal pool only.'
    PreferredVmSeries='NV'; DataDiskType='Premium SSD v2'; DataDiskMinIOPS=10000; DataDiskMinMBps=400
  }
}

function Get-ApplicationOverhead {
  <#
  .SYNOPSIS
    Calculates aggregate resource overhead from selected applications.
  #>
  param(
    [Parameter(Mandatory)][string[]]$SelectedApps,
    [int]$UsersPerHost = 1
  )

  $totalCpu = 0.0; $totalRamMB = 0; $totalDiskIOPS = 0; $totalDiskGB = 0
  $requiresGPU = $false; $requiresDataDisk = $false
  $preferredSeries = $null; $minMemVcoreRatio = 0
  $dataDiskType = $null; $dataDiskMinIOPS = 0; $dataDiskMinMBps = 0
  $notes = [System.Collections.Generic.List[string]]::new()
  $dbApps = [System.Collections.Generic.List[string]]::new()

  foreach ($appName in $SelectedApps) {
    if (-not $script:ApplicationCatalog.Contains($appName)) { continue }
    $app = $script:ApplicationCatalog[$appName]

    $totalCpu      += $app.CpuOverheadPerUser * $UsersPerHost
    $totalRamMB    += $app.RamMBPerUser * $UsersPerHost
    $totalDiskIOPS += $app.DiskIOPSPerUser * $UsersPerHost
    $totalDiskGB   += $app.DiskGBPerUser

    if ($app.RequiresGPU) { $requiresGPU = $true }
    if ($app.RequiresDataDisk) { $requiresDataDisk = $true }

    if ($app.PSObject.Properties.Name -contains 'PreferredVmSeries' -and $app.PreferredVmSeries) {
      $preferredSeries = $app.PreferredVmSeries
    }
    if ($app.PSObject.Properties.Name -contains 'MinMemoryToVcoreRatio' -and $app.MinMemoryToVcoreRatio -gt $minMemVcoreRatio) {
      $minMemVcoreRatio = $app.MinMemoryToVcoreRatio
    }
    if ($app.PSObject.Properties.Name -contains 'DataDiskType' -and $app.DataDiskType) {
      $dataDiskType = $app.DataDiskType
    }
    if ($app.PSObject.Properties.Name -contains 'DataDiskMinIOPS' -and $app.DataDiskMinIOPS -gt $dataDiskMinIOPS) {
      $dataDiskMinIOPS = $app.DataDiskMinIOPS; $dataDiskMinMBps = $app.DataDiskMinMBps
    }
    if ($app.Category -eq 'Database') { $dbApps.Add($appName) }
    $notes.Add("$appName : +$($app.CpuOverheadPerUser) vCPU, +$($app.RamMBPerUser) MB RAM/user ($($app.Category))")
  }

  if ($dbApps.Count -gt 0) {
    $notes.Add("DATABASE WARNING: Running DB engines ($($dbApps -join ', ')) on AVD session hosts is atypical.")
    $notes.Add("  Best practice: Use Azure SQL, Azure DB for PostgreSQL/MySQL as managed PaaS.")
    $notes.Add("  If local DB required: Personal host pool, E-series VM, separate data disks.")
  }
  if ($requiresGPU) {
    $notes.Add("GPU REQUIRED: Selected apps need NV-series VMs (GPU-enabled).")
  }

  [pscustomobject]@{
    TotalCpuOverhead     = [Math]::Round($totalCpu, 2)
    TotalRamOverheadMB   = $totalRamMB
    TotalRamOverheadGB   = [Math]::Round($totalRamMB / 1024.0, 2)
    TotalDiskIOPS        = $totalDiskIOPS
    TotalDiskGB          = $totalDiskGB
    RequiresGPU          = $requiresGPU
    RequiresDataDisk     = $requiresDataDisk
    PreferredVmSeries    = $preferredSeries
    MinMemoryToVcoreRatio = $minMemVcoreRatio
    DataDiskType         = $dataDiskType
    DataDiskMinIOPS      = $dataDiskMinIOPS
    DataDiskMinMBps      = $dataDiskMinMBps
    HasDatabaseEngine    = ($dbApps.Count -gt 0)
    DatabaseApps         = @([string[]]$dbApps.ToArray())
    Notes                = @([string[]]$notes.ToArray())
    SelectedApps         = @([string[]]$SelectedApps)
  }
}
#endregion

#region Azure Managed Disk catalog
$script:AzureDiskCatalog = @(
  [pscustomobject]@{ Sku='P1';  SizeGiB=4;     IOPS=120;   MBps=25;   BurstIOPS=3500;  BurstMBps=170  }
  [pscustomobject]@{ Sku='P2';  SizeGiB=8;     IOPS=120;   MBps=25;   BurstIOPS=3500;  BurstMBps=170  }
  [pscustomobject]@{ Sku='P3';  SizeGiB=16;    IOPS=120;   MBps=25;   BurstIOPS=3500;  BurstMBps=170  }
  [pscustomobject]@{ Sku='P4';  SizeGiB=32;    IOPS=120;   MBps=25;   BurstIOPS=3500;  BurstMBps=170  }
  [pscustomobject]@{ Sku='P6';  SizeGiB=64;    IOPS=240;   MBps=50;   BurstIOPS=3500;  BurstMBps=170  }
  [pscustomobject]@{ Sku='P10'; SizeGiB=128;   IOPS=500;   MBps=100;  BurstIOPS=3500;  BurstMBps=170  }
  [pscustomobject]@{ Sku='P15'; SizeGiB=256;   IOPS=1100;  MBps=125;  BurstIOPS=3500;  BurstMBps=170  }
  [pscustomobject]@{ Sku='P20'; SizeGiB=512;   IOPS=2300;  MBps=150;  BurstIOPS=3500;  BurstMBps=170  }
  [pscustomobject]@{ Sku='P30'; SizeGiB=1024;  IOPS=5000;  MBps=200;  BurstIOPS=30000; BurstMBps=1000 }
  [pscustomobject]@{ Sku='P40'; SizeGiB=2048;  IOPS=7500;  MBps=250;  BurstIOPS=30000; BurstMBps=1000 }
  [pscustomobject]@{ Sku='P50'; SizeGiB=4096;  IOPS=7500;  MBps=250;  BurstIOPS=30000; BurstMBps=1000 }
  [pscustomobject]@{ Sku='P60'; SizeGiB=8192;  IOPS=16000; MBps=500;  BurstIOPS=30000; BurstMBps=1000 }
  [pscustomobject]@{ Sku='P70'; SizeGiB=16384; IOPS=18000; MBps=750;  BurstIOPS=30000; BurstMBps=1000 }
  [pscustomobject]@{ Sku='P80'; SizeGiB=32768; IOPS=20000; MBps=900;  BurstIOPS=30000; BurstMBps=1000 }
)
function Get-OptimalOsDiskSku {
  param([Parameter(Mandatory)][int]$MinSizeGiB, [Parameter(Mandatory)][int]$TargetIOPS, [int]$TargetMBps = 100)
  # Prefer disk where provisioned (sustained) IOPS meets the target
  foreach ($disk in $script:AzureDiskCatalog) {
    if ($disk.SizeGiB -ge $MinSizeGiB -and $disk.IOPS -ge $TargetIOPS -and $disk.MBps -ge $TargetMBps) { return $disk }
  }
  # Fallback: provisioned IOPS meets target (relax MBps)
  foreach ($disk in $script:AzureDiskCatalog) {
    if ($disk.SizeGiB -ge $MinSizeGiB -and $disk.IOPS -ge $TargetIOPS) { return $disk }
  }
  # Last resort: burst IOPS (credit-based, max 30 min for P20 and smaller)
  foreach ($disk in $script:AzureDiskCatalog) {
    if ($disk.SizeGiB -ge $MinSizeGiB -and $disk.BurstIOPS -ge $TargetIOPS) { return $disk }
  }
  return $script:AzureDiskCatalog[-1]
}
#endregion

#region VM naming + AVD ranking
function Get-VmNameMetadata {
  param([Parameter(Mandatory)][string]$Name)
  $gen = 0; if ($Name -match '_v(?<g>\d+)$') { if ($Matches.ContainsKey('g') -and $Matches['g']) { $gen = [int]$Matches['g'] } }
  $suffix = ''; if ($Name -match '^Standard_[A-Za-z]+(?<n>\d+)(?<s>[A-Za-z]+)?_v\d+$') { if ($Matches.ContainsKey('s') -and $Matches['s']) { $suffix = [string]$Matches['s'] } }
  $hasS = (-not [string]::IsNullOrWhiteSpace($suffix)) -and ($suffix -match 's')
  $hasD = (-not [string]::IsNullOrWhiteSpace($suffix)) -and ($suffix -match 'd')
  $family = ''; if ($Name -match '^Standard_(?<fam>[A-Za-z]+)\d') { if ($Matches.ContainsKey('fam')) { $family = [string]$Matches['fam'] } }
  [pscustomobject]@{ Generation=$gen; Suffix=$suffix; HasS=[bool]$hasS; HasD=[bool]$hasD; Family=$family }
}
function Get-AvdVmFamilyRank {
  param([Parameter(Mandatory)][string]$Family, [string]$Workload='Medium', [bool]$HasDatabase=$false, [bool]$RequiresGPU=$false)
  $fam = $Family.ToUpperInvariant()

  # GPU ALWAYS takes priority when GPU apps are selected (CAD, Video, 3D)
  # Even when a database is also needed, GPU is the harder constraint
  if ($RequiresGPU) {
    switch -Regex ($fam) {
      '^NV' { return 0 }   # NV-series: designed for VDI/CAD/visualisation
      '^NC' { return 1 }   # NC-series: heavier GPU compute
      '^E'  { return 6 }   # E-series: no GPU but good RAM
      '^D'  { return 7 }   # D-series: no GPU
      '^F'  { return 8 }   # F-series: no GPU, low RAM
      '^B'  { return 9 }   # B-series: burstable, never for GPU
      default { return 5 }
    }
  }

  # Database (no GPU): E-series preferred for 8:1 memory:vCore ratio
  if ($HasDatabase) {
    switch -Regex ($fam) {
      '^E'  { return 0 }   # E-series: 8 GB/vCPU, ideal for SQL Server
      '^D'  { return 1 }   # D-series: 4 GB/vCPU, acceptable
      '^M'  { return 2 }   # M-series: ultra-high memory, expensive
      '^F'  { return 5 }   # F-series: 2 GB/vCPU, too little for DB
      '^B'  { return 9 }   # B-series: never for production DB
      default { return 4 }
    }
  }

  # Power workload WITHOUT GPU: use D-series (standard ranking)
  # Power workload WITH GPU: handled by RequiresGPU block above

  # Standard AVD workloads
  switch -Regex ($fam) {
    '^D'  { return 0 }   # D-series: general purpose, 4 GB/vCPU
    '^F'  { return 1 }   # F-series: compute optimised, 2 GB/vCPU
    '^E'  { return 2 }   # E-series: memory optimised, 8 GB/vCPU
    '^NV' { return 3 }   # NV-series: GPU
    '^NC' { return 4 }   # NC-series: GPU compute
    '^B'  { return 8 }   # B-series: burstable, penalised
    default { return 5 }
  }
}

# RAM per vCPU by VM family (based on real Azure VM specifications)
# Used to determine how much RAM a VM of a given series/size actually has
$script:VmSeriesRamRatio = @{
  # D-series: General purpose, 4 GB/vCPU
  # D2s_v5=8, D4s_v5=16, D8s_v5=32, D16s_v5=64, D32s_v5=128
  'D' = @{ GBperVcpu=4; Sizes=@{ 2=8; 4=16; 8=32; 12=48; 16=64; 24=96; 32=128; 48=192; 64=256 } }

  # E-series: Memory optimised, 8 GB/vCPU
  # E2s_v5=16, E4s_v5=32, E8s_v5=64, E16s_v5=128, E32s_v5=256
  'E' = @{ GBperVcpu=8; Sizes=@{ 2=16; 4=32; 8=64; 16=128; 20=160; 24=192; 32=256; 48=384; 64=512 } }

  # F-series: Compute optimised, 2 GB/vCPU
  # F2s_v2=4, F4s_v2=8, F8s_v2=16, F16s_v2=32, F32s_v2=64
  'F' = @{ GBperVcpu=2; Sizes=@{ 2=4; 4=8; 8=16; 16=32; 32=64; 48=96; 72=144 } }

  # NV-series: GPU visualisation (NVadsA10_v5 = AMD + NVIDIA A10)
  # NV6ads_A10_v5=55, NV12ads_A10_v5=110, NV18ads_A10_v5=220, NV36ads_A10_v5=440
  # NV-series v3: NV12s_v3=112, NV24s_v3=224, NV48s_v3=448
  'NV' = @{ GBperVcpu=7; Sizes=@{ 6=55; 12=110; 18=220; 24=220; 36=440 } }

  # NC-series: GPU compute (NC-series A100/H100)
  # NC6s_v3=112, NC12s_v3=224, NC24s_v3=448
  'NC' = @{ GBperVcpu=8; Sizes=@{ 6=112; 12=224; 24=448 } }

  # B-series: Burstable, variable RAM
  # B2s=4, B4ms=16, B8ms=32, B12ms=48, B16ms=64, B20ms=80
  'B' = @{ GBperVcpu=4; Sizes=@{ 2=4; 4=16; 8=32; 12=48; 16=64; 20=80 } }

  # M-series: Ultra memory (28 GB/vCPU)
  'M' = @{ GBperVcpu=28; Sizes=@{ 8=218; 16=432; 32=875; 64=1750 } }
}

function Get-VmSeriesRamGB {
  # Returns the expected RAM in GB for a given VM series and vCPU count
  param([string]$Series='D', [int]$Vcpu=8)
  $s = $Series.ToUpperInvariant()
  # Try exact lookup first
  foreach ($key in $script:VmSeriesRamRatio.Keys) {
    if ($s -match "^$key") {
      $info = $script:VmSeriesRamRatio[$key]
      # Exact size match
      if ($info.Sizes.ContainsKey($Vcpu)) { return [double]$info.Sizes[$Vcpu] }
      # Closest smaller size
      $closest = ($info.Sizes.Keys | ForEach-Object { [int]$_ } | Sort-Object | Where-Object { $_ -le $Vcpu } | Select-Object -Last 1)
      if ($closest) { return [double]$info.Sizes[[int]$closest] }
      # Fallback: ratio-based
      return [double]($Vcpu * $info.GBperVcpu)
    }
  }
  # Unknown series: assume D-series ratio (4 GB/vCPU)
  return [double]($Vcpu * 4)
}
function Get-GenerationRank { param([Parameter(Mandatory)][int]$Generation)
  if ($Generation -ge 6) { return 0 }; if ($Generation -eq 5) { return 0 }; if ($Generation -eq 4) { return 1 }
  if ($Generation -eq 3) { return 2 }; if ($Generation -eq 2) { return 3 }; return 9
}
function Get-SuffixRank { param([Parameter(Mandatory)][bool]$HasS, [bool]$HasD=$false)
  if ($HasS -and $HasD) { return 0 }; if ($HasS) { return 1 }; return 3
}
#endregion

#region Microsoft-verified Guidelines + CPU Scaling Factor
# Source: https://learn.microsoft.com/en-us/windows-server/remote/remote-desktop-services/session-host-virtual-machine-sizing-guidelines
# Verified against September 2025 revision

$Guidelines = [ordered]@{
  MultiSession = [ordered]@{
    # RAM calculation now uses $VmSeriesRamRatio catalog (real Azure VM specs)
    # Per-user RAM from GO-EUC research: Light=2GB, Medium=3-4GB, Heavy=5-6GB
    # Users/host = min(CPU-based, RAM-based) - dual-limit system
    # MinOsDiskGB: Windows 11 multi-session image ~25 GB + apps + updates + paging = 128 GB minimum
    Light  = [ordered]@{ UsersPerVcpu=6; MinVcpu=8; MinRamGB=16; MinOsDiskGB=128; MinProfileGB=30; RamPerUserGB=2
      Examples='D8s_v5, D8s_v4, F8s_v2, D8as_v4, D16s_v5' }
    Medium = [ordered]@{ UsersPerVcpu=4; MinVcpu=8; MinRamGB=16; MinOsDiskGB=128; MinProfileGB=30; RamPerUserGB=4
      Examples='D8s_v5, D8s_v4, F8s_v2, D8as_v4, D16s_v5' }
    Heavy  = [ordered]@{ UsersPerVcpu=2; MinVcpu=8; MinRamGB=16; MinOsDiskGB=128; MinProfileGB=30; RamPerUserGB=6
      Examples='D8s_v5, D8s_v4, F8s_v2, D16s_v5, D16s_v4' }
    Power  = [ordered]@{ UsersPerVcpu=1; MinVcpu=6; MinRamGB=56; MinOsDiskGB=256; MinProfileGB=30; RamPerUserGB=8
      Examples='D16ds_v5, D16s_v4, NV6, NV16as_v4' }
  }
  SingleSession = [ordered]@{
    Light  = [ordered]@{ Vcpu=2; RamGB=8;  MinOsDiskGB=128; MinProfileGB=30; Examples='D2s_v5, D2s_v4' }
    Medium = [ordered]@{ Vcpu=4; RamGB=16; MinOsDiskGB=128; MinProfileGB=30; Examples='D4s_v5, D4s_v4' }
    Heavy  = [ordered]@{ Vcpu=8; RamGB=32; MinOsDiskGB=128; MinProfileGB=30; Examples='D8s_v5, D8s_v4' }
  }
  CpuScalingFactorMin = 1.5
  CpuScalingFactorMax = 1.9
}

$DefaultSystemResourceReserve = 0.15   # MS docs: 15-20% virtualisation overhead
$DefaultCpuUtil = 0.80
$DefaultMemUtil = 0.80
#endregion

#region FSLogix IOPS model
function Get-FsLogixIopsRequirements {
  param([Parameter(Mandatory)][int]$ConcurrentUsers, [Parameter(Mandatory)][string]$Workload, [string]$HostPoolType='Pooled')
  $steadyPerUser = switch ($Workload) { 'Light' {10} 'Medium' {10} 'Heavy' {15} 'Power' {20} default {10} }
  $burstPerUser  = switch ($Workload) { 'Light' {50} 'Medium' {50} 'Heavy' {60} 'Power' {75} default {50} }
  $burstMult = if ($HostPoolType -eq 'Pooled') { 1.0 } else { 0.7 }
  [pscustomobject]@{
    SteadyStateIOPS = [int]($ConcurrentUsers * $steadyPerUser)
    BurstIOPS       = [int]($ConcurrentUsers * $burstPerUser * $burstMult)
    SteadyStateMBps = [double]([Math]::Round($ConcurrentUsers * 0.20, 2))
    BurstMBps       = [double]([Math]::Round($ConcurrentUsers * 0.50 * $burstMult, 2))
    PerUserSteadyIOPS = $steadyPerUser; PerUserBurstIOPS = $burstPerUser
    Note = 'MS docs: ~10 IOPS/user steady, ~50 IOPS/user burst.'
  }
}
function Get-FsLogixStorageTierRecommendation {
  param([Parameter(Mandatory)][int]$BurstIOPS, [Parameter(Mandatory)][int]$ConcurrentUsers, [string]$Workload='Medium')
  $tier = 'Azure Files Premium (SSD)'; $shareCount = 1
  $notes = [System.Collections.Generic.List[string]]::new()
  if ($Workload -eq 'Light' -and $ConcurrentUsers -lt 100 -and $BurstIOPS -lt 10000) {
    $tier = 'Azure Files Standard (HDD) may suffice'
    $notes.Add('Standard acceptable for small Light workloads. Validate logon latency.')
  } elseif ($BurstIOPS -le 100000) {
    $tier = 'Azure Files Premium (SSD)'
    # Share count: max of IOPS-based and user-count-based (MS: ~1000 concurrent users/share)
    $sharesFromIops = [int][Math]::Ceiling($BurstIOPS / 100000.0)
    $sharesFromUsers = [int][Math]::Ceiling($ConcurrentUsers / 1000.0)
    $shareCount = [Math]::Max(1, [Math]::Max($sharesFromIops, $sharesFromUsers))
    if ($shareCount -gt 1) { $notes.Add("Distribute across $shareCount Premium shares (~1000 users/share, 100K IOPS/share).") }
  } else { $tier = 'Azure NetApp Files'; $notes.Add('Burst >100K IOPS: Azure NetApp Files recommended.') }
  if ($ConcurrentUsers -ge 500 -and $shareCount -le 1) { $notes.Add("$ConcurrentUsers concurrent users: consider splitting across 2+ shares for logon storm resilience.") }
  $notes.Add('Premium SSD file shares recommended for production AVD. Provision capacity to meet IOPS needs.')
  $notes.Add('Validate: SMB latency <5ms, logon times, storage throttling in pilot.')
  [pscustomobject]@{ RecommendedTier=$tier; RecommendedShares=$shareCount; Notes=$notes.ToArray() }
}
#endregion

#region Disk recommendations
function Get-AvdDiskRecommendations {
  param([Parameter(Mandatory)]$Sizing, $AppOverhead=$null)
  $workload = [string]$Sizing.Workload; $hostPoolType = [string]$Sizing.HostPoolType; $peakUsers = [int]$Sizing.PeakConcurrentUsers
  $minOs = [int]$Sizing.Recommended.MinOsDiskGB

  $fslogix = Get-FsLogixIopsRequirements -ConcurrentUsers $peakUsers -Workload $workload -HostPoolType $hostPoolType
  $osIops = switch ($workload) { 'Light' { if($hostPoolType -eq 'Pooled'){1500}else{500} } 'Medium' { if($hostPoolType -eq 'Pooled'){2000}else{1000} }
    'Heavy' { if($hostPoolType -eq 'Pooled'){2500}else{1500} } 'Power' { if($hostPoolType -eq 'Pooled'){3000}else{2000} } default {1500} }
  $osMBps = switch ($workload) { 'Light' { if($hostPoolType -eq 'Pooled'){75}else{50} } 'Medium' { if($hostPoolType -eq 'Pooled'){100}else{75} }
    'Heavy' { if($hostPoolType -eq 'Pooled'){125}else{90} } 'Power' { if($hostPoolType -eq 'Pooled'){150}else{100} } default {80} }

  $osDiskSku = Get-OptimalOsDiskSku -MinSizeGiB $minOs -TargetIOPS $osIops -TargetMBps $osMBps
  $ephemeral = [ordered]@{ Recommended=($hostPoolType -eq 'Pooled'); Notes=@() }
  if ($hostPoolType -eq 'Pooled') { $ephemeral.Notes = @('Ephemeral OS disk for stateless pooled hosts.','Requires VM cache >= OS image size. No deallocate/snapshots/ASR.') }

  $storageTier = Get-FsLogixStorageTierRecommendation -BurstIOPS $fslogix.BurstIOPS -ConcurrentUsers $peakUsers -Workload $workload

  # Storage risk
  $riskLevel = 'Low'; $riskNotes = New-Object System.Collections.Generic.List[string]
  if ($workload -in @('Heavy','Power')) { $riskLevel = 'High'; $riskNotes.Add("$workload workload: Premium storage mandatory.") }
  if ($workload -eq 'Medium' -and $peakUsers -ge 200) { if($riskLevel -ne 'High'){$riskLevel='Medium'}; $riskNotes.Add("Medium + $peakUsers users: validate burst IOPS.") }
  if ($peakUsers -ge 500) { $riskLevel = 'High'; $riskNotes.Add("$peakUsers users: multiple shares required.") }
  if ($hostPoolType -eq 'Pooled' -and $peakUsers -ge 150) { if($riskLevel -eq 'Low'){$riskLevel='Medium'}; $riskNotes.Add("Pooled + ${peakUsers} users: plan for logon storms.") }
  if ($riskNotes.Count -eq 0) { $riskNotes.Add('No obvious storage risks. Validate with pilot.') }

  # Data disk (for databases/heavy apps)
  $dataDisk = $null
  if ($AppOverhead -and $AppOverhead.RequiresDataDisk) {
    $dataDisk = [pscustomobject]@{
      Required = $true
      Type     = ($AppOverhead.DataDiskType ?? 'Premium SSD v2')
      MinIOPS  = $AppOverhead.DataDiskMinIOPS
      MinMBps  = $AppOverhead.DataDiskMinMBps
      MinGB    = $AppOverhead.TotalDiskGB
      Notes    = @(
        "Database/app data disk required: $($AppOverhead.DataDiskType ?? 'Premium SSD v2')"
        "Min IOPS: $($AppOverhead.DataDiskMinIOPS), Min MB/s: $($AppOverhead.DataDiskMinMBps)"
        'Best practice: Separate disks for data, log, tempdb (SQL Server).'
        'Format with 64KB allocation unit size for SQL Server data files.'
      )
    }
  }

  [pscustomobject]@{
    SessionHostDisks = [pscustomobject]@{
      OsDisk = [pscustomobject]@{
        RecommendedType=$osDiskSku.Sku; SuggestedSizeGiB=$osDiskSku.SizeGiB
        ProvisionedIOPS=$osDiskSku.IOPS; ProvisionedMBps=$osDiskSku.MBps
        BurstIOPS=$osDiskSku.BurstIOPS; BurstMBps=$osDiskSku.BurstMBps; MinimumFromGuidelineGiB=$minOs
      }
      DataDisk = $dataDisk
      EphemeralOsDisk = [pscustomobject]$ephemeral
    }
    FsLogixStorage = [pscustomobject]@{
      RecommendedTier=$storageTier.RecommendedTier; RecommendedShares=$storageTier.RecommendedShares
      PerformanceTargets = [pscustomobject]@{
        SteadyStateIOPS=$fslogix.SteadyStateIOPS; BurstIOPS=$fslogix.BurstIOPS
        SteadyStateMBps=$fslogix.SteadyStateMBps; BurstMBps=$fslogix.BurstMBps
        PerUserSteadyIOPS=$fslogix.PerUserSteadyIOPS; PerUserBurstIOPS=$fslogix.PerUserBurstIOPS
      }
      StorageRisk = [pscustomobject]@{ Level=$riskLevel; Notes=$riskNotes.ToArray() }
    }
  }
}
#endregion

#region Azure integration
function Test-AzAvailable {
  (Get-Module -ListAvailable -Name Az.Accounts) -and (Get-Module -ListAvailable -Name Az.Compute)
}

function Write-AzModuleInstallHint {
  Write-UiWarning "Az PowerShell modules not found (optional feature).`n`nInstall:`n  Install-Module Az -Scope CurrentUser"
}

function Connect-AzIfNeeded {
  try {
    Import-Module Az.Accounts -ErrorAction Stop | Out-Null
    Import-Module Az.Compute  -ErrorAction Stop | Out-Null
    if (-not (Get-AzContext -ErrorAction SilentlyContinue)) {
      Connect-AzAccount -ErrorAction Stop | Out-Null
    }
    return $true
  } catch {
    Write-UiError "Azure sign-in failed: $($_.Exception.Message)"
    return $false
  }
}

function Get-AzVmSizesInLocation {
  <#
  .SYNOPSIS
    Returns VM sizes for a region. Tries Get-AzVMSize first, then falls back
    to Get-AzComputeResourceSku (required for newer Az module versions).
  #>
  param([Parameter(Mandatory)][string]$Location)

  Import-Module Az.Compute -ErrorAction Stop | Out-Null

  # --- Path 1: Get-AzVMSize (classic, may be deprecated in newer modules) ---
  try {
    $cmdVmSize = Get-Command -Name Get-AzVMSize -ErrorAction SilentlyContinue
    if ($cmdVmSize) {
      $vmSizes = Get-AzVMSize -Location $Location -ErrorAction Stop
      $mapped = foreach ($s in $vmSizes) {
        if (-not $s.Name) { continue }
        $cores = [int]$s.NumberOfCores
        $mem   = [int]$s.MemoryInMB
        if ($cores -le 0 -or $mem -le 0) { continue }
        [pscustomobject]@{
          Name          = [string]$s.Name
          NumberOfCores = $cores
          MemoryInMB    = $mem
        }
      }
      $mapped = $mapped | Sort-Object Name -Unique
      if ($mapped -and $mapped.Count -gt 0) { return $mapped }
    }
  } catch {
    # Silently fall through to fallback
  }

  # --- Path 2: Get-AzComputeResourceSku (newer, more reliable) ---
  function Get-CapabilityValue {
    param([Parameter(Mandatory)]$Sku, [Parameter(Mandatory)][string[]]$Names)
    foreach ($n in $Names) {
      $cap = $Sku.Capabilities | Where-Object { $_.Name -eq $n } | Select-Object -First 1
      if ($cap -and $cap.Value) { return $cap.Value }
    }
    return $null
  }

  $skusRaw = $null
  try {
    $cmd = Get-Command -Name Get-AzComputeResourceSku -ErrorAction Stop
    if ($cmd.Parameters.ContainsKey('Location')) {
      $skusRaw = Get-AzComputeResourceSku -Location $Location -ErrorAction Stop
    } else {
      $skusRaw = Get-AzComputeResourceSku -ErrorAction Stop |
        Where-Object { $_.Locations -contains $Location }
    }
  } catch {
    try {
      $skusRaw = Get-AzComputeResourceSku -ErrorAction Stop |
        Where-Object { $_.Locations -contains $Location }
    } catch {
      Write-UiError "Failed to query VM sizes for '${Location}': $($_.Exception.Message)"
      return @()
    }
  }

  if (-not $skusRaw) {
    Write-UiWarning "No resource SKUs returned for region '${Location}'. Check region name and subscription."
    return @()
  }

  $skus = $skusRaw | Where-Object { $_.ResourceType -eq 'virtualMachines' }

  $out = foreach ($sku in $skus) {
    # Skip restricted SKUs
    if ($sku.Restrictions) {
      $blocked = $sku.Restrictions |
        Where-Object { $_.ReasonCode -eq 'NotAvailableForSubscription' } |
        Select-Object -First 1
      if ($blocked) { continue }
    }

    $vCpuRaw = Get-CapabilityValue -Sku $sku -Names @('vCPUs', 'vCPUsAvailable')
    $memRaw  = Get-CapabilityValue -Sku $sku -Names @('MemoryGB', 'MemoryGBs')
    if (-not $vCpuRaw -or -not $memRaw) { continue }

    $vCpu = 0
    if (-not [int]::TryParse([string]$vCpuRaw, [ref]$vCpu)) { continue }

    $memGb = 0.0
    if (-not [double]::TryParse(
          [string]$memRaw,
          [System.Globalization.NumberStyles]::Float,
          [System.Globalization.CultureInfo]::InvariantCulture,
          [ref]$memGb)) { continue }

    if ($vCpu -le 0 -or $memGb -le 0) { continue }

    [pscustomobject]@{
      Name          = $sku.Name
      NumberOfCores = $vCpu
      MemoryInMB    = [int][Math]::Round($memGb * 1024)
    }
  }

  $result = $out | Sort-Object Name -Unique
  if (-not $result -or $result.Count -eq 0) {
    Write-UiWarning "No usable VM sizes found for '${Location}'. Verify region name and subscription permissions."
  }
  return $result
}
function Get-BestVmSize {
  param([Parameter(Mandatory)][object[]]$Sizes, [Parameter(Mandatory)][int]$MinVcpu, [Parameter(Mandatory)][double]$MinRamGB,
    [int]$MaxVcpu=0,
    [string]$Series='Any', [string]$Workload='Medium', [bool]$HasDatabase=$false, [bool]$RequiresGPU=$false)
  if ($MinVcpu -lt 1 -or $MinRamGB -le 0) { return $null }
  $ramMB = [int]([Math]::Ceiling($MinRamGB * 1024))

  # --- Step 1: AVD-compatible families only ---
  $avdFamilies = '^Standard_(D|E|F|NV|NC|B|M)\d'
  $avdFiltered = @($Sizes | Where-Object {
    $_.Name -match $avdFamilies -and
    $_.Name -notmatch '^Standard_(DC|EC|H|ND|A)\d' -and
    $_.Name -notmatch '^Standard_[A-Z]+p[a-z]*\d'
  })
  if ($avdFiltered.Count -lt 1) { $avdFiltered = @($Sizes) }

  # --- Step 2: Apply STRICT series filter FIRST (if user chose a specific series) ---
  $isStrictSeries = ($Series -ne 'Any')
  if ($isStrictSeries) {
    $avdFiltered = @($avdFiltered | Where-Object { $_.Name -match "^Standard_${Series}" })
    if ($avdFiltered.Count -lt 1) {
      # NO VMs of this series available in this region at all
      return [pscustomobject]@{
        _Error = $true
        _Message = "No $Series-series VMs available in this region. Change VM series to 'Any (auto)' or select a different region."
      }
    }
  }

  # --- Step 3: Find VMs that meet the EXACT calculated specs ---
  # Smallest VM that meets BOTH min vCPU AND min RAM
  $candidates = @($avdFiltered | Where-Object { $_.NumberOfCores -ge $MinVcpu -and $_.MemoryInMB -ge $ramMB })

  # Apply vCPU max range
  $rangeNote = $null
  if ($MaxVcpu -gt 0 -and $candidates.Count -gt 0) {
    $withinRange = @($candidates | Where-Object { $_.NumberOfCores -le $MaxVcpu })
    if ($withinRange.Count -gt 0) {
      $candidates = $withinRange
    } else {
      $rangeNote = "No VM within vCPU range ($MinVcpu-$MaxVcpu) meets RAM requirement ($([Math]::Round($MinRamGB,1)) GB). Range exceeded."
    }
  }

  # No candidates at all for this series
  if ($candidates.Count -lt 1) {
    if ($isStrictSeries) {
      return [pscustomobject]@{
        _Error = $true
        _Message = "No $Series-series VM found with >= $MinVcpu vCPU and >= $([Math]::Round($MinRamGB,1)) GB RAM. Try 'Any (auto)' or reduce requirements."
      }
    }
    return $null
  }

  # --- Step 4: Select the SMALLEST matching VM ---
  # Primary: fewest vCPUs (closest to calculated need)
  # Secondary: least RAM (don't over-provision)
  # Then: prefer v5+, prefer 's' suffix (premium storage)
  $preferred = @(); foreach ($x in $candidates) { $m = Get-VmNameMetadata -Name $x.Name; if ($m.Generation -ge 5) { $preferred += $x } }
  $pool = if ($preferred.Count -gt 0) { $preferred } else { $candidates }

  $ranked = foreach ($x in $pool) { $m = Get-VmNameMetadata -Name $x.Name
    [pscustomobject]@{
      Obj=$x; Cores=[int]$x.NumberOfCores; Mem=[int]$x.MemoryInMB
      # For 'Any' series: rank by family preference; for strict series: all same rank
      FamRank = if ($isStrictSeries) { 0 } else { (Get-AvdVmFamilyRank -Family $m.Family -Workload $Workload -HasDatabase $HasDatabase -RequiresGPU $RequiresGPU) }
      GenRank=(Get-GenerationRank -Generation $m.Generation)
      SufRank=(Get-SuffixRank -HasS $m.HasS -HasD $m.HasD)
      Name=[string]$x.Name
    }
  }

  # Sort: smallest cores first, then least RAM, then best family/generation
  $best = ($ranked | Sort-Object Cores, Mem, FamRank, GenRank, SufRank, Name | Select-Object -First 1).Obj

  if ($rangeNote -and $best) {
    $best | Add-Member -NotePropertyName RangeExceededNote -NotePropertyValue $rangeNote -Force
  }
  return $best
}
function Get-AzVmHourlyRetailPrice {
  param([Parameter(Mandatory)][string]$ArmRegionName, [Parameter(Mandatory)][string]$ArmSkuName, [string]$CurrencyCode='USD')
  $endpoint = 'https://prices.azure.com/api/retail/prices'
  $filterExpr = "serviceName eq 'Virtual Machines' and armRegionName eq '$ArmRegionName' and armSkuName eq '$ArmSkuName' and unitOfMeasure eq '1 Hour'"
  $uri = "${endpoint}?currencyCode='$CurrencyCode'&`$filter=$([uri]::EscapeDataString($filterExpr))"
  try { $all = New-Object System.Collections.Generic.List[object]; $next = $uri
    for ($i=0; $i -lt 3 -and $next; $i++) { $resp = Invoke-RestMethod -Method Get -Uri $next -TimeoutSec 20 -ErrorAction Stop
      if ($resp.Items) { foreach ($it in $resp.Items) { $all.Add($it) } }; $next = $resp.NextPageLink }
    if ($all.Count -lt 1) { return $null }
    $c = $all | Where-Object { $_.meterName -notmatch 'Spot' -and $_.productName -notmatch 'Spot' }
    if (-not $c -or $c.Count -lt 1) { $c = $all }
    $w = $c | Where-Object { $_.productName -match 'Windows' }
    $pool = if ($w -and $w.Count -gt 0) { $w } else { $c }
    $cons = $pool | Where-Object { $_.PSObject.Properties.Name -contains 'type' -and $_.type -eq 'Consumption' }
    if ($cons -and $cons.Count -gt 0) { $pool = $cons }
    $best = $pool | Sort-Object retailPrice | Select-Object -First 1
    if (-not $best) { return $null }
    [pscustomobject]@{ RetailPricePerHour=[double]$best.retailPrice; CurrencyCode=[string]$best.currencyCode
      ProductName=[string]$best.productName; MeterName=[string]$best.meterName; ArmSkuName=[string]$best.armSkuName
      PricingNote='Retail list price. No discounts/RI/savings/AHB applied.' }
  } catch { [pscustomobject]@{ Error=$true; Message=$_.Exception.Message; PricingNote='Price lookup failed.' } }
}
#endregion

#region Sizing calculation (with application overhead)
function Get-AvdSizing {
  param(
    [Parameter(Mandatory)][ValidateSet('Pooled','Personal')][string]$HostPoolType,
    [Parameter(Mandatory)][ValidateSet('Light','Medium','Heavy','Power')][string]$Workload,
    [Parameter(Mandatory)][int]$TotalUsers,
    [Parameter(Mandatory)][ValidateSet('Percent','User')][string]$ConcurrencyMode,
    [Parameter(Mandatory)][double]$ConcurrencyValue,
    [double]$PeakFactor=1.0, [double]$CpuTargetUtil=$DefaultCpuUtil, [double]$MemTargetUtil=$DefaultMemUtil,
    [Alias('VirtualizationOverhead')][double]$SystemResourceReserve=$DefaultSystemResourceReserve,
    [int]$MinVcpuPerHost=4, [int]$MaxVcpuPerHost=24, [int]$NPlusOneHosts=1, [double]$ExtraHeadroomPercent=0,
    [double]$ProfileContainerGB=30, [double]$ProfileGrowthPercent=20, [double]$ProfileOverheadPercent=10,
    [ValidateSet('BreadthFirst','DepthFirst')][string]$LoadBalancing='BreadthFirst',
    [int]$MaxSessionLimit=0,
    $AppOverhead = $null   # output of Get-ApplicationOverhead
  )

  if ($TotalUsers -lt 1) { throw "TotalUsers must be >= 1." }
  if ($CpuTargetUtil -le 0 -or $CpuTargetUtil -gt 1) { throw "CPU utilization must be (0,1]." }
  if ($MemTargetUtil -le 0 -or $MemTargetUtil -gt 1) { throw "Memory utilization must be (0,1]." }
  if ($SystemResourceReserve -lt 0 -or $SystemResourceReserve -gt 0.5) { throw "System resource reserve must be [0,0.5]." }
  if ($PeakFactor -lt 1) { throw "PeakFactor must be >= 1." }

  $warnings = [System.Collections.Generic.List[string]]::new()

  $concurrent = if ($ConcurrencyMode -eq 'Percent') {
    $pct = $ConcurrencyValue / 100.0
    if ($pct -le 0 -or $pct -gt 1) { throw "Concurrency percent 1-100." }
    Get-CeilingInt -Value ($TotalUsers * $pct)
  } else { if ($ConcurrencyValue -lt 1 -or $ConcurrencyValue -gt $TotalUsers) { throw "Concurrency must be 1..TotalUsers." }; [int]$ConcurrencyValue }

  $peakConcurrent = [Math]::Min((Get-CeilingInt -Value ($concurrent * $PeakFactor)), $TotalUsers)
  if ($ExtraHeadroomPercent -gt 0) { $peakConcurrent = Get-CeilingInt -Value ($peakConcurrent * (1 + $ExtraHeadroomPercent/100.0)) }

  # Application overhead warnings
  if ($AppOverhead) {
    foreach ($n in $AppOverhead.Notes) { $warnings.Add($n) }
    if ($AppOverhead.HasDatabaseEngine -and $HostPoolType -eq 'Pooled') {
      $warnings.Add("WARNING: Database engines on Pooled multi-session hosts is strongly discouraged. Use Personal host pool.")
    }
    if ($AppOverhead.RequiresGPU) { $warnings.Add("GPU-enabled NV-series VMs required for selected applications.") }
  }

  if ($HostPoolType -eq 'Pooled') {
    $g = $Guidelines.MultiSession.$Workload
    $usersPerVcpu = [double]$g.UsersPerVcpu; $minVcpuBP = [int]$g.MinVcpu; $minRamBP = [double]$g.MinRamGB

    if ($MinVcpuPerHost -lt 4) { $warnings.Add("Pooled: <4 vCPU not recommended (MS best practice: 4-24).") }

    # Determine preferred VM series based on app requirements
    # GPU ALWAYS wins (harder constraint: only NV/NC have GPUs)
    $preferredSeries = 'D'
    if ($AppOverhead) {
      if ($AppOverhead.RequiresGPU) {
        $preferredSeries = 'NV'
        if ($AppOverhead.HasDatabaseEngine) {
          $warnings.Add("GPU + Database: NV-series selected (GPU is hard requirement). Consider Azure SQL as managed DB service.")
        }
      } elseif ($AppOverhead.HasDatabaseEngine) {
        $preferredSeries = 'E'
      }
    }
    if ($Workload -eq 'Power' -and $preferredSeries -eq 'D' -and $AppOverhead -and $AppOverhead.RequiresGPU) { $preferredSeries = 'NV' }

    $candidates = ((4..24 | Where-Object { $_ % 4 -eq 0 }) + 6) | Sort-Object -Unique |
      Where-Object { $_ -ge $MinVcpuPerHost -and $_ -le $MaxVcpuPerHost }

    $best = $null
    foreach ($vcpu in $candidates) {
      if ($vcpu -lt $minVcpuBP) { continue }

      # CPU-based users per host
      $usersPerHostCpu = [Math]::Floor($vcpu * $usersPerVcpu * $CpuTargetUtil * (1 - $SystemResourceReserve))
      if ($usersPerHostCpu -lt 1) { continue }

      # RAM from preferred VM series (uses real Azure VM specs)
      $ramFromLookup = Get-VmSeriesRamGB -Series $preferredSeries -Vcpu $vcpu

      # Per-user RAM from GO-EUC research: Light=2GB, Medium=3-4GB, Heavy=5-6GB, Power=8GB
      $osOverheadGB = 4.0
      $perUserRamGB = 4.0
      if ($g.Contains('RamPerUserGB')) { $perUserRamGB = [double]$g['RamPerUserGB'] }

      # RAM-based max users: how many users fit in the VM's real RAM?
      $usableRamGB = ($ramFromLookup * $MemTargetUtil) - $osOverheadGB
      $usersPerHostRam = if ($perUserRamGB -gt 0 -and $usableRamGB -gt 0) {
        [Math]::Floor($usableRamGB / $perUserRamGB)
      } else { $usersPerHostCpu }

      # Effective users/host = min(CPU-based, RAM-based)
      $usersPerHost = [Math]::Min($usersPerHostCpu, [Math]::Max(1, $usersPerHostRam))
      $hostsForPeak = Get-CeilingInt -Value ($peakConcurrent / $usersPerHost)

      # RAM estimate: what this user count actually needs (NOT the VM lookup size)
      $ramFromUsers = $osOverheadGB + ($usersPerHost * $perUserRamGB)

      # Add application overhead
      $ramEstimated = $ramFromUsers
      if ($AppOverhead) {
        $ramEstimated += $AppOverhead.TotalRamOverheadGB
        # If DB needs 8:1 ratio, enforce it
        if ($AppOverhead.MinMemoryToVcoreRatio -gt 0) {
          $minRamForRatio = $vcpu * $AppOverhead.MinMemoryToVcoreRatio
          $ramEstimated = [Math]::Max($ramEstimated, $minRamForRatio)
        }
      }

      # Ensure minimum from MS baseline, but do NOT inflate to VM lookup size
      $ramEstimated = [Math]::Max($minRamBP, $ramEstimated)

      $ramProvisioned = $ramEstimated / $MemTargetUtil

      $opt = [pscustomobject]@{
        VcpuPerHost=$vcpu; UsersPerHost=$usersPerHost; HostsForPeak=$hostsForPeak
        RamGB_Estimated=[Math]::Round($ramEstimated,2); RamGB_Provisioned=[Math]::Round($ramProvisioned,2)
        MinOsDiskGB=[int]$g.MinOsDiskGB
      }
      if (-not $best) { $best = $opt }
      elseif ($opt.HostsForPeak -lt $best.HostsForPeak) { $best = $opt }
      elseif ($opt.HostsForPeak -eq $best.HostsForPeak -and $opt.VcpuPerHost -lt $best.VcpuPerHost) { $best = $opt }
      elseif ($opt.HostsForPeak -eq $best.HostsForPeak -and $opt.VcpuPerHost -eq $best.VcpuPerHost -and $opt.RamGB_Provisioned -lt $best.RamGB_Provisioned) { $best = $opt }
    }
    if (-not $best) { throw "No suitable pooled sizing found." }

    $hostsTotal = $best.HostsForPeak + [Math]::Max(0, $NPlusOneHosts)
    $profileBase = [Math]::Max([double]$g.MinProfileGB, $ProfileContainerGB)
    $plannedPerUser = $profileBase * (1 + $ProfileGrowthPercent/100.0) * (1 + $ProfileOverheadPercent/100.0)

    # Load Balancing + Max Session Limit
    $calcMaxSessionLimit = if ($MaxSessionLimit -gt 0) { $MaxSessionLimit } else { $best.UsersPerHost }

    # Depth-first: needs fewer hosts powered on at off-peak but same total capacity
    # Breadth-first: all hosts share load evenly = better UX but higher cost
    $lbNotes = [System.Collections.Generic.List[string]]::new()
    if ($LoadBalancing -eq 'DepthFirst') {
      $lbNotes.Add("Depth-first: fills hosts sequentially up to max session limit ($calcMaxSessionLimit).")
      $lbNotes.Add("  Cost-optimised: idle hosts can be deallocated by autoscale.")
      $lbNotes.Add("  Recommended: use with Autoscale Scaling Plan for cost savings.")
      $lbNotes.Add("  Set MaxSessionLimit to $calcMaxSessionLimit (= calculated users/host).")
      $warnings.Add("Depth-first: set MaxSessionLimit=$calcMaxSessionLimit. Do NOT use default 999999.")
    } else {
      $lbNotes.Add("Breadth-first: distributes sessions evenly across all powered-on hosts.")
      $lbNotes.Add("  Best UX: lower per-host load, better logon storm handling.")
      $lbNotes.Add("  Higher cost: all hosts must be powered on during peak.")
      $lbNotes.Add("  MaxSessionLimit acts as safety cap (recommended: $calcMaxSessionLimit).")
    }

    # Autoscale scaling plan recommendations
    $autoscale = [pscustomobject]@{
      RampUp = [pscustomobject]@{
        LoadBalancing = 'BreadthFirst'
        MinHostsPct   = 20
        CapacityThresholdPct = 60
        Note = 'Breadth-first during ramp-up to handle logon storms evenly.'
      }
      Peak = [pscustomobject]@{
        LoadBalancing = $LoadBalancing
        MinHostsPct   = [int][Math]::Ceiling(($best.HostsForPeak / [Math]::Max(1,$hostsTotal)) * 100)
        CapacityThresholdPct = 90
        Note = "Peak: all $($best.HostsForPeak) hosts active, $LoadBalancing balancing."
      }
      RampDown = [pscustomobject]@{
        LoadBalancing = 'DepthFirst'
        MinHostsPct   = 10
        CapacityThresholdPct = 90
        ForceLogoff   = $true
        WaitMinutes   = 30
        Note = 'Depth-first to consolidate sessions, allow idle hosts to deallocate.'
      }
      OffPeak = [pscustomobject]@{
        LoadBalancing = 'DepthFirst'
        MinHostsPct   = 0
        Note = 'Depth-first. Min 0% hosts = all can shut down if no sessions.'
      }
    }

    $result = [pscustomobject]@{
      Mode='MultiSession'; HostPoolType=$HostPoolType; Workload=$Workload; TotalUsers=$TotalUsers
      ConcurrentUsers=$concurrent; PeakConcurrentUsers=$peakConcurrent
      CpuTargetUtil=$CpuTargetUtil; MemTargetUtil=$MemTargetUtil; SystemResourceReserve=$SystemResourceReserve
      LoadBalancing=$LoadBalancing; MaxSessionLimit=$calcMaxSessionLimit; PreferredSeries=$preferredSeries
      Recommended=$best; RecommendedHostsTotal=$hostsTotal; Examples=$g.Examples
      FsLogix=[pscustomobject]@{ PlannedPerUserGB=[Math]::Round($plannedPerUser,2); PlannedTotalGB_AtPeak=[Math]::Round($peakConcurrent*$plannedPerUser,2) }
      Autoscale=$autoscale; LoadBalancingNotes=@([string[]]$lbNotes.ToArray())
      AppOverhead=$AppOverhead; Notes=@([string[]]$warnings.ToArray())
    }
    $result | Add-Member -NotePropertyName Disks -NotePropertyValue (Get-AvdDiskRecommendations -Sizing $result -AppOverhead $AppOverhead) -Force
    return $result
  }

  # Personal
  $origWorkload = $Workload
  if ($Workload -eq 'Power') { $warnings.Add("Power: using Heavy baseline for personal sizing."); $Workload = 'Heavy' }
  $g2 = $Guidelines.SingleSession.$Workload
  $vcpuBase = [int]$g2.Vcpu

  # Determine preferred VM series for personal
  $preferredSeriesP = 'D'
  if ($AppOverhead) {
    if ($AppOverhead.RequiresGPU) {
      $preferredSeriesP = 'NV'
      if ($AppOverhead.HasDatabaseEngine) {
        $warnings.Add("GPU + Database: NV-series selected (GPU is hard requirement). Consider Azure SQL as managed DB service.")
      }
    } elseif ($AppOverhead.HasDatabaseEngine) {
      $preferredSeriesP = 'E'
    }
  }
  if ($origWorkload -eq 'Power' -and $preferredSeriesP -eq 'D' -and $AppOverhead -and $AppOverhead.RequiresGPU) { $preferredSeriesP = 'NV' }

  # Add application overhead for personal
  if ($AppOverhead) {
    $vcpuBase = [Math]::Max($vcpuBase, $vcpuBase + [int][Math]::Ceiling($AppOverhead.TotalCpuOverhead))
    if ($AppOverhead.MinMemoryToVcoreRatio -gt 0) {
      $minRamForRatio = $vcpuBase * $AppOverhead.MinMemoryToVcoreRatio
    }
  }

  # RAM from preferred VM series (real Azure VM specs)
  $ramFromSeries = Get-VmSeriesRamGB -Series $preferredSeriesP -Vcpu $vcpuBase
  $ramEstimated = [Math]::Max([double]$g2.RamGB, $ramFromSeries)

  if ($AppOverhead) {
    $ramEstimated += $AppOverhead.TotalRamOverheadGB
    if ($AppOverhead.MinMemoryToVcoreRatio -gt 0) {
      $ramEstimated = [Math]::Max($ramEstimated, $vcpuBase * $AppOverhead.MinMemoryToVcoreRatio)
    }
  }

  $ramProvisioned = $ramEstimated / $MemTargetUtil
  $hostsForPeak = $peakConcurrent
  $hostsTotal = $hostsForPeak + [Math]::Max(0, $NPlusOneHosts)
  $profileBase2 = [Math]::Max([double]$g2.MinProfileGB, $ProfileContainerGB)
  $plannedPerUser2 = $profileBase2 * (1 + $ProfileGrowthPercent/100.0) * (1 + $ProfileOverheadPercent/100.0)

  $result2 = [pscustomobject]@{
    Mode='SingleSession'; HostPoolType=$HostPoolType; Workload=$Workload; TotalUsers=$TotalUsers
    ConcurrentUsers=$concurrent; PeakConcurrentUsers=$peakConcurrent
    CpuTargetUtil=$CpuTargetUtil; MemTargetUtil=$MemTargetUtil; SystemResourceReserve=$SystemResourceReserve
    LoadBalancing='Persistent'; MaxSessionLimit=1; PreferredSeries=$preferredSeriesP
    Recommended=[pscustomobject]@{
      VcpuPerHost=$vcpuBase; UsersPerHost=1; HostsForPeak=$hostsForPeak
      RamGB_Estimated=[Math]::Round($ramEstimated,2); RamGB_Provisioned=[Math]::Round($ramProvisioned,2)
      MinOsDiskGB=[int]$g2.MinOsDiskGB
    }
    RecommendedHostsTotal=$hostsTotal; Examples=$g2.Examples
    FsLogix=[pscustomobject]@{ PlannedPerUserGB=[Math]::Round($plannedPerUser2,2); PlannedTotalGB_AtPeak=[Math]::Round($peakConcurrent*$plannedPerUser2,2) }
    AppOverhead=$AppOverhead; Notes=@([string[]]$warnings.ToArray())
  }
  $result2 | Add-Member -NotePropertyName Disks -NotePropertyValue (Get-AvdDiskRecommendations -Sizing $result2 -AppOverhead $AppOverhead) -Force
  return $result2
}
#endregion

#region Results rendering
function Set-ResultsGrid {
  param([Parameter(Mandatory)]$Sizing, $VmPick, $VmPrice,
    [Parameter(Mandatory)][System.Windows.Controls.DataGrid]$GridResults,
    [Parameter(Mandatory)][System.Windows.Controls.TextBox]$TxtNotes,
    $HidePricing=$true)

  $rows = [System.Collections.Generic.List[object]]::new()
  $r = $Sizing.Recommended

  $rows.Add([pscustomobject]@{ Key='HostPoolType'; Value=$Sizing.HostPoolType })
  $rows.Add([pscustomobject]@{ Key='Mode'; Value=$Sizing.Mode })
  $rows.Add([pscustomobject]@{ Key='Workload'; Value=$Sizing.Workload })
  $rows.Add([pscustomobject]@{ Key='TotalUsers'; Value=$Sizing.TotalUsers })
  $rows.Add([pscustomobject]@{ Key='ConcurrentUsers'; Value=$Sizing.ConcurrentUsers })
  $rows.Add([pscustomobject]@{ Key='PeakConcurrentUsers'; Value=$Sizing.PeakConcurrentUsers })
  $rows.Add([pscustomobject]@{ Key='SystemResourceReserve'; Value="$($Sizing.SystemResourceReserve) (MS: 15-20%)" })
  $rows.Add([pscustomobject]@{ Key='UsersPerHost'; Value=$r.UsersPerHost })
  $rows.Add([pscustomobject]@{ Key='vCPU per Host'; Value=$r.VcpuPerHost })
  $rows.Add([pscustomobject]@{ Key='RAM GB (estimated)'; Value=$r.RamGB_Estimated })
  $rows.Add([pscustomobject]@{ Key='RAM GB (provisioned)'; Value=$r.RamGB_Provisioned })
  $rows.Add([pscustomobject]@{ Key='Hosts for Peak'; Value=$r.HostsForPeak })
  $rows.Add([pscustomobject]@{ Key='Hosts total (N+1)'; Value=$Sizing.RecommendedHostsTotal })
  $rows.Add([pscustomobject]@{ Key='Min OS Disk (GB)'; Value=$r.MinOsDiskGB })
  $rows.Add([pscustomobject]@{ Key='Preferred VM Series'; Value="$($Sizing.PreferredSeries)-series" })
  $rows.Add([pscustomobject]@{ Key='Guideline Examples'; Value=$Sizing.Examples })

  # Load Balancing
  $rows.Add([pscustomobject]@{ Key='--- LOAD BALANCING ---'; Value='' })
  $rows.Add([pscustomobject]@{ Key='Algorithm'; Value=$Sizing.LoadBalancing })
  $rows.Add([pscustomobject]@{ Key='Max Session Limit'; Value=$Sizing.MaxSessionLimit })
  if ($Sizing.Autoscale) {
    $rows.Add([pscustomobject]@{ Key='Ramp-Up LB'; Value="$($Sizing.Autoscale.RampUp.LoadBalancing) (min $($Sizing.Autoscale.RampUp.MinHostsPct)% hosts)" })
    $rows.Add([pscustomobject]@{ Key='Peak LB'; Value="$($Sizing.Autoscale.Peak.LoadBalancing) (min $($Sizing.Autoscale.Peak.MinHostsPct)% hosts)" })
    $rows.Add([pscustomobject]@{ Key='Ramp-Down LB'; Value="$($Sizing.Autoscale.RampDown.LoadBalancing) (min $($Sizing.Autoscale.RampDown.MinHostsPct)% hosts)" })
    $rows.Add([pscustomobject]@{ Key='Off-Peak LB'; Value="$($Sizing.Autoscale.OffPeak.LoadBalancing) (min $($Sizing.Autoscale.OffPeak.MinHostsPct)% hosts)" })
  }

  # Application overhead
  if ($Sizing.AppOverhead -and @($Sizing.AppOverhead.SelectedApps).Count -gt 0) {
    $rows.Add([pscustomobject]@{ Key='--- APPLICATIONS ---'; Value='' })
    $rows.Add([pscustomobject]@{ Key='Selected Apps'; Value=($Sizing.AppOverhead.SelectedApps -join ', ') })
    $rows.Add([pscustomobject]@{ Key='App CPU Overhead'; Value="$($Sizing.AppOverhead.TotalCpuOverhead) vCPU" })
    $rows.Add([pscustomobject]@{ Key='App RAM Overhead'; Value="$($Sizing.AppOverhead.TotalRamOverheadGB) GB" })
    if ($Sizing.AppOverhead.RequiresGPU) { $rows.Add([pscustomobject]@{ Key='GPU Required'; Value='Yes (NV-series)' }) }
    if ($Sizing.AppOverhead.HasDatabaseEngine) {
      $rows.Add([pscustomobject]@{ Key='Database Engine'; Value=($Sizing.AppOverhead.DatabaseApps -join ', ') })
      $rows.Add([pscustomobject]@{ Key='Min Mem:vCore Ratio'; Value="$($Sizing.AppOverhead.MinMemoryToVcoreRatio):1" })
    }
  }

  # Disks
  if ($Sizing.Disks) {
    $d = $Sizing.Disks
    $rows.Add([pscustomobject]@{ Key='--- OS DISK ---'; Value='' })
    $rows.Add([pscustomobject]@{ Key='OS Disk SKU'; Value="$($d.SessionHostDisks.OsDisk.RecommendedType) ($($d.SessionHostDisks.OsDisk.SuggestedSizeGiB) GiB)" })
    $rows.Add([pscustomobject]@{ Key='OS Disk IOPS'; Value="$($d.SessionHostDisks.OsDisk.ProvisionedIOPS) (burst: $($d.SessionHostDisks.OsDisk.BurstIOPS))" })

    if ($d.SessionHostDisks.DataDisk -and $d.SessionHostDisks.DataDisk.Required) {
      $rows.Add([pscustomobject]@{ Key='--- DATA DISK ---'; Value='' })
      $rows.Add([pscustomobject]@{ Key='Data Disk Type'; Value=$d.SessionHostDisks.DataDisk.Type })
      $rows.Add([pscustomobject]@{ Key='Data Disk Min IOPS'; Value=$d.SessionHostDisks.DataDisk.MinIOPS })
      $rows.Add([pscustomobject]@{ Key='Data Disk Min MB/s'; Value=$d.SessionHostDisks.DataDisk.MinMBps })
      $rows.Add([pscustomobject]@{ Key='Data Disk Min GB'; Value=$d.SessionHostDisks.DataDisk.MinGB })
    }

    $rows.Add([pscustomobject]@{ Key='--- FSLOGIX ---'; Value='' })
    $rows.Add([pscustomobject]@{ Key='Storage Tier'; Value=$d.FsLogixStorage.RecommendedTier })
    $pt = $d.FsLogixStorage.PerformanceTargets
    $rows.Add([pscustomobject]@{ Key='Steady IOPS'; Value="$($pt.SteadyStateIOPS) ($($pt.PerUserSteadyIOPS)/user)" })
    $rows.Add([pscustomobject]@{ Key='Burst IOPS'; Value="$($pt.BurstIOPS) ($($pt.PerUserBurstIOPS)/user)" })
    $rows.Add([pscustomobject]@{ Key='Storage Risk'; Value=$d.FsLogixStorage.StorageRisk.Level })
  }

  if ($VmPick) {
    $rows.Add([pscustomobject]@{ Key='--- AZURE VM ---'; Value='' })
    $rows.Add([pscustomobject]@{ Key='VM Size'; Value=$VmPick.Name })
    $rows.Add([pscustomobject]@{ Key='VM vCPU'; Value=$VmPick.NumberOfCores })
    $rows.Add([pscustomobject]@{ Key='VM RAM (GB)'; Value=[Math]::Round($VmPick.MemoryInMB/1024,2) })
  }
  if ($VmPrice -and -not $HidePricing -and -not ($VmPrice.PSObject.Properties.Name -contains 'Error' -and $VmPrice.Error)) {
    $rows.Add([pscustomobject]@{ Key='--- PRICING ---'; Value='' })
    $rows.Add([pscustomobject]@{ Key='Price/Hour'; Value="$($VmPrice.RetailPricePerHour) $($VmPrice.CurrencyCode)" })
    $monthly = [Math]::Round($VmPrice.RetailPricePerHour * 730, 2)
    $rows.Add([pscustomobject]@{ Key='Est. Monthly/Host'; Value="$monthly $($VmPrice.CurrencyCode)" })
    $rows.Add([pscustomobject]@{ Key='Est. Monthly Total'; Value="$([Math]::Round($monthly * $Sizing.RecommendedHostsTotal, 2)) $($VmPrice.CurrencyCode) ($($Sizing.RecommendedHostsTotal) hosts)" })
  }

  $GridResults.ItemsSource = $rows

  # Notes
  $notes = [System.Collections.Generic.List[string]]::new()

  # Application info
  if (-not $Sizing.AppOverhead -or @($Sizing.AppOverhead.SelectedApps).Count -eq 0) {
    $notes.Add('INFO: No applications selected on the "Applications" tab.')
    $notes.Add('  The sizing is based on the Microsoft baseline workload profile only.')
    $notes.Add('  For more accurate results, select the applications your users will run.')
    $notes.Add('  App overhead (CPU, RAM, disk, GPU) will be added to the baseline automatically.')
    $notes.Add('')
  } else {
    $notes.Add("APPLICATIONS ($(@($Sizing.AppOverhead.SelectedApps).Count) selected):")
    $notes.Add("  Additional overhead per host: +$($Sizing.AppOverhead.TotalCpuOverhead) vCPU, +$($Sizing.AppOverhead.TotalRamOverheadGB) GB RAM")
    if ($Sizing.AppOverhead.RequiresGPU) { $notes.Add('  GPU required: NV-series VMs will be preferred.') }
    if ($Sizing.AppOverhead.HasDatabaseEngine) { $notes.Add("  Database engines: $($Sizing.AppOverhead.DatabaseApps -join ', ') - E-series VMs preferred (8:1 mem:vCore).") }
    $notes.Add('')
  }

  if ($Sizing.Notes -and @($Sizing.Notes).Count -gt 0) { $notes.Add('WARNINGS:'); foreach ($n in $Sizing.Notes) { $notes.Add("  $n") }; $notes.Add('') }

  # Load Balancing notes
  if ($Sizing.LoadBalancingNotes -and @($Sizing.LoadBalancingNotes).Count -gt 0) {
    $notes.Add('LOAD BALANCING:'); foreach ($n in $Sizing.LoadBalancingNotes) { $notes.Add("  $n") }; $notes.Add('')
  }
  if ($Sizing.Autoscale) {
    $notes.Add('AUTOSCALE SCALING PLAN (recommended):')
    $notes.Add("  Ramp-Up:   $($Sizing.Autoscale.RampUp.LoadBalancing), min $($Sizing.Autoscale.RampUp.MinHostsPct)% hosts, cap $($Sizing.Autoscale.RampUp.CapacityThresholdPct)%")
    $notes.Add("  Peak:      $($Sizing.Autoscale.Peak.LoadBalancing), min $($Sizing.Autoscale.Peak.MinHostsPct)% hosts, cap $($Sizing.Autoscale.Peak.CapacityThresholdPct)%")
    $notes.Add("  Ramp-Down: $($Sizing.Autoscale.RampDown.LoadBalancing), min $($Sizing.Autoscale.RampDown.MinHostsPct)% hosts, force logoff after $($Sizing.Autoscale.RampDown.WaitMinutes) min")
    $notes.Add("  Off-Peak:  $($Sizing.Autoscale.OffPeak.LoadBalancing), min $($Sizing.Autoscale.OffPeak.MinHostsPct)% hosts")
    $notes.Add('')
  }

  if ($Sizing.Disks.SessionHostDisks.DataDisk -and $Sizing.Disks.SessionHostDisks.DataDisk.Required) {
    $notes.Add('DATA DISK NOTES:'); foreach ($n in $Sizing.Disks.SessionHostDisks.DataDisk.Notes) { $notes.Add("  $n") }; $notes.Add('') }
  $notes.Add('MS SIZING REFERENCE: https://learn.microsoft.com/en-us/windows-server/remote/remote-desktop-services/session-host-virtual-machine-sizing-guidelines')
  $notes.Add('Reminder: Validate with pilot workloads and monitoring.')
  $TxtNotes.Text = ($notes -join [Environment]::NewLine)
}
#endregion

#region XAML
$XamlString = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AVD Sizing Calculator" Height="920" Width="1200"
        WindowStartupLocation="CenterScreen">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" FontSize="18" FontWeight="SemiBold" Text="AVD Sizing Calculator v2.2"/>
    <TabControl Grid.Row="1" Margin="0,10,0,10" x:Name="Tabs">

      <!-- TAB 1: Workload -->
      <TabItem Header="Workload">
        <Grid Margin="10">
          <Grid.ColumnDefinitions><ColumnDefinition Width="380"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>

          <!-- LEFT: Input fields with descriptions -->
          <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Padding="0,0,12,0">
          <StackPanel>
            <TextBlock FontWeight="Bold" Text="Host pool type"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Pooled: multiple users share each VM (cost-efficient). Personal: 1:1 user-to-VM mapping (isolated, persistent)."/>
            <ComboBox x:Name="CmbHostPoolType" Margin="0,4,0,12" SelectedIndex="0">
              <ComboBoxItem Content="Pooled (multi-session)"/><ComboBoxItem Content="Personal (single-session)"/>
            </ComboBox>

            <TextBlock FontWeight="Bold" Text="Workload class"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Determines users/vCPU density and RAM/user. Based on Microsoft session host sizing guidelines."/>
            <ComboBox x:Name="CmbWorkload" Margin="0,4,0,4" SelectedIndex="1">
              <ComboBoxItem Content="Light — 6 users/vCPU, 2 GB RAM/user"/>
              <ComboBoxItem Content="Medium — 4 users/vCPU, 4 GB RAM/user"/>
              <ComboBoxItem Content="Heavy — 2 users/vCPU, 6 GB RAM/user"/>
              <ComboBoxItem Content="Power — 1 user/vCPU, 8 GB RAM/user"/>
            </ComboBox>
            <TextBlock Foreground="#888" FontSize="10" FontStyle="Italic" TextWrapping="Wrap" Margin="0,0,0,12" Text="Light: basic office/web. Medium: Office + Teams + LOB. Heavy: multi-app, analytics. Power: CAD/dev, GPU workloads."/>

            <TextBlock FontWeight="Bold" Text="Total users (named)"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Total number of named users in the organisation who will use AVD."/>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,12">
              <TextBox x:Name="TxtTotalUsers" Width="120" Text="100"/>
            </StackPanel>

            <TextBlock FontWeight="Bold" Text="Concurrency"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="How many users are active at the same time. Percent: e.g. 60 = 60% of total users. User: absolute number of concurrent users."/>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,12">
              <ComboBox x:Name="CmbConcurrencyMode" Width="120" SelectedIndex="0">
                <ComboBoxItem Content="Percent"/><ComboBoxItem Content="User"/>
              </ComboBox>
              <TextBox x:Name="TxtConcurrencyValue" Width="80" Margin="8,0,0,0" Text="60"/>
              <TextBlock x:Name="LblConcurrencyHint" Margin="8,4,0,0" Text="% or #"/>
            </StackPanel>

            <TextBlock FontWeight="Bold" Text="Peak factor"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Multiplier for peak spikes above normal concurrency. 1.0 = no extra peak. 1.2 = 20% spike capacity."/>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,12">
              <TextBox x:Name="TxtPeakFactor" Width="120" Text="1.0"/><TextBlock Margin="8,4,0,0" Foreground="#888" Text="1.0 = no spike buffer"/>
            </StackPanel>

            <Separator Margin="0,4,0,8"/>
            <TextBlock FontWeight="Bold" Text="N+1 redundancy"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Extra hosts for failover. 1 = one standby host. Set 0 to disable."/>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,8">
              <TextBox x:Name="TxtNPlusOne" Width="120" Text="1"/>
            </StackPanel>
            <TextBlock FontWeight="Bold" Text="Extra headroom"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Additional percentage added to peak concurrent users for future growth."/>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,8">
              <TextBox x:Name="TxtExtraHeadroomPct" Width="120" Text="0"/><TextBlock Margin="8,4,0,0" Foreground="#888" Text="%"/>
            </StackPanel>

            <Separator Margin="0,4,0,8"/>
            <TextBlock FontWeight="Bold" Text="Load Balancing (pooled only)"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Breadth-first: spreads sessions evenly (best UX, higher cost). Depth-first: fills one host before the next (saves cost, allows deallocation)."/>
            <ComboBox x:Name="CmbLoadBalancing" Margin="0,4,0,8" SelectedIndex="0">
              <ComboBoxItem Content="Breadth-first (best UX)"/>
              <ComboBoxItem Content="Depth-first (cost optimised)"/>
            </ComboBox>
            <TextBlock FontWeight="Bold" Text="Max session limit per host"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Limits sessions per VM. 0 = auto-calculated from users/host. Important for Depth-first (never leave at default 999999)."/>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,8">
              <TextBox x:Name="TxtMaxSessionLimit" Width="120" Text="0"/><TextBlock Margin="8,4,0,0" Foreground="#888" Text="0 = auto"/>
            </StackPanel>
          </StackPanel>
          </ScrollViewer>

          <!-- RIGHT: Tuning Parameters -->
          <Border Grid.Column="1" Padding="12" BorderBrush="#DDD" BorderThickness="1" CornerRadius="6" Background="#FAFAFA">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock FontSize="14" FontWeight="Bold" Text="Tuning Parameters"/>
              <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Margin="0,4,0,6" Text="Advanced settings. Default values follow Microsoft best practice — change only if you have specific requirements."/>
              <Separator Margin="0,4,0,8"/>
              <TextBlock FontWeight="Bold" Text="FSLogix profile storage"/>
              <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Profile container size per user. MS minimum: 30 GB. Growth: annual expansion. Overhead: VHD/metadata."/>
              <StackPanel Orientation="Horizontal" Margin="0,4,0,4"><TextBox x:Name="TxtProfileGB" Width="100" Text="30"/><TextBlock Margin="8,4,0,0" Text="GB/user (min 30)"/></StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,4"><TextBox x:Name="TxtProfileGrowthPct" Width="100" Text="20"/><TextBlock Margin="8,4,0,0" Text="growth %"/></StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,8"><TextBox x:Name="TxtProfileOverheadPct" Width="100" Text="10"/><TextBlock Margin="8,4,0,0" Text="overhead %"/></StackPanel>
              <Separator Margin="0,4,0,8"/>
              <TextBlock FontWeight="Bold" Text="CPU / RAM / System Reserve"/>
              <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="CPU target: max utilisation before throttling. Memory target: usable RAM after OS. System reserve: virtualisation overhead (MS: 15-20%)."/>
              <StackPanel Orientation="Horizontal" Margin="0,4,0,4"><TextBox x:Name="TxtCpuUtil" Width="100" Text="0.80"/><TextBlock Margin="8,4,0,0" Text="CPU target (0.80 = 80%)"/></StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,4"><TextBox x:Name="TxtMemUtil" Width="100" Text="0.80"/><TextBlock Margin="8,4,0,0" Text="Memory target (0.80 = 80%)"/></StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,8"><TextBox x:Name="TxtVirtOverhead" Width="100" Text="0.15"/><TextBlock Margin="8,4,0,0" Text="System Reserve (MS: 0.15-0.20)"/></StackPanel>
              <Separator Margin="0,4,0,8"/>
              <TextBlock FontWeight="Bold" Text="vCPU range per host (pooled)"/>
              <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Min: smallest VM to consider (MS: 4 vCPU). Max: largest VM allowed. MS recommends max 24 for multi-session. Set 128 for unrestricted."/>
              <StackPanel Orientation="Horizontal" Margin="0,4,0,4"><TextBox x:Name="TxtMinVcpuHost" Width="100" Text="8"/><TextBlock Margin="8,4,0,0" Text="min vCPU/host (MS: 4)"/></StackPanel>
              <StackPanel Orientation="Horizontal"><TextBox x:Name="TxtMaxVcpuHost" Width="100" Text="128"/><TextBlock Margin="8,4,0,0" Text="max vCPU/host (MS rec: 24)"/></StackPanel>
            </StackPanel>
            </ScrollViewer>
          </Border>
        </Grid>
      </TabItem>

      <!-- TAB 2: Applications -->
      <TabItem Header="Applications">
        <Grid Margin="10">
          <Grid.ColumnDefinitions><ColumnDefinition Width="350"/><ColumnDefinition Width="350"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>

          <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Padding="0,0,8,0">
          <StackPanel>
            <TextBlock FontSize="14" FontWeight="Bold" Text="Client Applications" Margin="0,0,0,4"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8" Text="Standard office and productivity apps. Each adds CPU + RAM overhead per concurrent user on multi-session hosts."/>
            <CheckBox x:Name="ChkOffice" Content="Microsoft 365 (Word/Excel/PPT)" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkTeams" Content="Microsoft Teams" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkBrowser" Content="Web Browser (Edge/Chrome)" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkOutlook" Content="Microsoft Outlook" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkPdf" Content="PDF Editor (Acrobat/Foxit)" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkErp" Content="ERP Client (SAP GUI / Dynamics)" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkPowerBi" Content="Power BI Desktop" Margin="0,4,0,0"/>
            <TextBlock Foreground="#888" FontSize="10" FontStyle="Italic" TextWrapping="Wrap" Margin="0,6,0,0" Text="Tip: Teams + Browser together add significant RAM. Consider Teams media optimisation (WebRTC redirect) to reduce host CPU."/>

            <TextBlock FontSize="14" FontWeight="Bold" Text="Development Tools" Margin="0,18,0,4"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8" Text="Dev tools increase CPU and RAM significantly. Docker requires nested virtualisation (Personal host pool only). Consider Personal pool for developers."/>
            <CheckBox x:Name="ChkVS" Content="Visual Studio" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkVSCode" Content="VS Code" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkDocker" Content="Docker Desktop" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkGit" Content="Git / Build Tools" Margin="0,4,0,0"/>
          </StackPanel>
          </ScrollViewer>

          <ScrollViewer Grid.Column="1" VerticalScrollBarVisibility="Auto" Padding="0,0,8,0">
          <StackPanel>
            <TextBlock FontSize="14" FontWeight="Bold" Text="Database Engines" Margin="0,0,0,4"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8" Text="Running a DB on a session host is atypical. The calculator will switch to E-series (8 GB/vCPU) and add a data disk. For production, use Azure SQL or a managed PaaS service instead."/>
            <CheckBox x:Name="ChkSqlExpress" Content="SQL Server Express/Developer" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkSqlStd" Content="SQL Server Standard/Enterprise" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkPostgres" Content="PostgreSQL" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkMySql" Content="MySQL / MariaDB" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkSqlite" Content="SQLite / MS Access (local DB)" Margin="0,4,0,0"/>

            <TextBlock FontSize="14" FontWeight="Bold" Text="CAD / GPU Applications" Margin="0,18,0,4"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8" Text="GPU apps require NV-series VMs with dedicated NVIDIA GPU. These VMs are significantly more expensive. Personal host pool is strongly recommended. ISV GPU certification may be needed."/>
            <CheckBox x:Name="ChkAutoCAD" Content="AutoCAD / AutoCAD LT" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkRevit" Content="Revit / 3ds Max" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkSolidWorks" Content="SolidWorks / CATIA" Margin="0,4,0,0"/>
            <CheckBox x:Name="ChkVideoEdit" Content="Video Editing (Premiere/DaVinci)" Margin="0,4,0,0"/>

            <Separator Margin="0,16,0,8"/>
            <TextBlock FontWeight="Bold" Text="Database data volume per host"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,6" Text="Size for the dedicated data disk (Premium SSD v2). Only used when a database engine is selected."/>
            <StackPanel Orientation="Horizontal">
              <TextBox x:Name="TxtDbDataGB" Width="120" HorizontalAlignment="Left" Text="50"/>
              <TextBlock Margin="8,4,0,0" Foreground="#888" Text="GB"/>
            </StackPanel>
          </StackPanel>
          </ScrollViewer>

          <Border Grid.Column="2" Padding="12" BorderBrush="#DDD" BorderThickness="1" CornerRadius="6" Background="#FAFAFA">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock FontSize="14" FontWeight="Bold" Text="How Applications Affect Sizing"/>
              <Separator Margin="0,8,0,8"/>
              <TextBlock TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold" Text="Overhead Calculation"/>
              <TextBlock Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11.5" Foreground="#444">
Each selected app adds per-user CPU and RAM overhead on top of the baseline workload. The calculator sums all overheads and reduces users/host accordingly.
              </TextBlock>

              <TextBlock TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold" Text="VM Series Auto-Selection"/>
              <TextBlock Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11.5" Foreground="#444">
The calculator automatically picks the best VM series:
• No special apps → D-series (general purpose)
• Database selected → E-series (8 GB/vCPU for SQL)
• GPU/CAD selected → NV-series (dedicated GPU)
• GPU + Database → NV-series wins (GPU is hard constraint). Tip: offload DB to Azure SQL.
              </TextBlock>

              <TextBlock TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold" Text="Database Engines"/>
              <TextBlock Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11.5" Foreground="#444">
• Switches to E-series (memory-optimised)
• Adds Premium SSD v2 data disk recommendation
• Personal host pool strongly recommended
• Pooled + DB will show a warning
              </TextBlock>

              <TextBlock TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold" Text="CAD / GPU Applications"/>
              <TextBlock Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11.5" Foreground="#444">
• Switches to NV-series (NVIDIA A10 GPU)
• NV-series VMs: 2-4 users/host typical
• Significantly higher cost per host
• Personal pool recommended for heavy 3D
              </TextBlock>
            </StackPanel>
            </ScrollViewer>
          </Border>
        </Grid>
      </TabItem>

      <!-- TAB 3: VM Template -->
      <TabItem Header="VM Template">
        <Grid Margin="10">
          <Grid.ColumnDefinitions><ColumnDefinition Width="470"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Padding="0,0,12,0">
          <StackPanel>
            <TextBlock FontWeight="Bold" Text="Azure region"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="The Azure region where session hosts will be deployed. VM availability and pricing vary by region. Use the display name (e.g. 'Switzerland North', 'West Europe')."/>
            <TextBox x:Name="TxtLocation" Margin="0,4,0,12" Text="Switzerland North"/>

            <TextBlock FontWeight="Bold" Text="VM series preference"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Select a specific VM family or leave 'Any' to let the calculator pick based on workload and app requirements. A specific series is strictly enforced — no fallback to other series."/>
            <ComboBox x:Name="CmbVmSeries" Margin="0,4,0,12" SelectedIndex="0">
              <ComboBoxItem Content="Any (auto)"/><ComboBoxItem Content="D (general purpose, 4 GB/vCPU)"/>
              <ComboBoxItem Content="E (memory, 8 GB/vCPU)"/><ComboBoxItem Content="F (compute, 2 GB/vCPU)"/>
              <ComboBoxItem Content="NV (GPU visualisation)"/><ComboBoxItem Content="NC (GPU compute)"/>
              <ComboBoxItem Content="B (burstable)"/>
            </ComboBox>

            <TextBlock FontWeight="Bold" Text="Currency"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Currency for VM pricing. Prices are fetched from the Azure Retail Prices API (list prices, no discounts or RI)."/>
            <ComboBox x:Name="CmbCurrency" Margin="0,4,0,8" SelectedIndex="2">
              <ComboBoxItem Content="USD"/><ComboBoxItem Content="EUR"/><ComboBoxItem Content="CHF"/>
            </ComboBox>
            <CheckBox x:Name="ChkHidePricing" Content="Hide pricing information" IsChecked="True" Margin="0,0,0,12"/>
            <TextBlock Foreground="#888" FontSize="10" FontStyle="Italic" TextWrapping="Wrap" Margin="0,0,0,12" Text="When checked, no pricing data is shown in the results grid or the exported HTML report. Uncheck to display Azure retail list prices."/>

            <Separator Margin="0,4,0,8"/>
            <TextBlock FontWeight="Bold" Text="Marketplace image"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,6" Text="The Windows image for session hosts. Click 'Azure Login' to authenticate, then 'Load SKUs' to discover available images. The default is Windows 11 Enterprise multi-session with M365 Apps."/>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,4"><TextBox x:Name="TxtPublisher" Width="240" Text="MicrosoftWindowsDesktop"/><TextBlock Margin="8,4,0,0" Foreground="#888" Text="publisher"/></StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,4"><TextBox x:Name="TxtOffer" Width="240" Text="office-365"/><TextBlock Margin="8,4,0,0" Foreground="#888" Text="offer"/></StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,4"><ComboBox x:Name="CmbSku" Width="320"/><TextBlock Margin="8,4,0,0" Foreground="#888" Text="SKU"/></StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,10"><TextBox x:Name="TxtVersion" Width="240" Text="latest"/><TextBlock Margin="8,4,0,0" Foreground="#888" Text="version"/></StackPanel>
            <StackPanel Orientation="Horizontal">
              <Button x:Name="BtnAzLogin" Content="Azure Login" Width="140" Margin="0,0,10,0"/>
              <Button x:Name="BtnDiscoverSkus" Content="Load SKUs" Width="140"/>
            </StackPanel>

            <Separator Margin="0,12,0,8"/>
            <TextBlock FontWeight="Bold" Text="Template JSON"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,6" Text="After clicking 'Pick VM from Azure' on the Results tab, the ARM template JSON for the selected VM, image and region is generated here. Copy this into your IaC deployment."/>
            <TextBox x:Name="TxtTemplateOut" Height="160" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"/>
          </StackPanel>
          </ScrollViewer>

          <Border Grid.Column="1" Padding="12" BorderBrush="#DDD" BorderThickness="1" CornerRadius="6" Background="#FAFAFA">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock FontSize="14" FontWeight="Bold" Text="VM Selection Logic"/>
              <Separator Margin="0,8,0,8"/>

              <TextBlock TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold" Text="Series Auto-Selection (when 'Any')"/>
              <TextBlock Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11.5" Foreground="#444">
The picker selects the optimal VM series based on your apps:
• Standard workloads → D-series (4 GB/vCPU)
• Database engines → E-series (8 GB/vCPU)
• GPU/CAD apps → NV-series (NVIDIA A10)
• GPU always takes priority over database
              </TextBlock>

              <TextBlock TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold" Text="Strict Series Mode"/>
              <TextBlock Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11.5" Foreground="#444">
When you select a specific series (e.g. 'E'), only VMs of that series are considered. If no VM of that series meets the calculated requirements, an error is shown — no silent fallback to another series.
              </TextBlock>

              <TextBlock TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold" Text="Smallest Matching VM"/>
              <TextBlock Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11.5" Foreground="#444">
The picker always selects the smallest VM that meets the calculated vCPU and RAM requirements. It prefers v5/v6 generations and Premium Storage ('s' suffix) capable VMs.
              </TextBlock>

              <TextBlock TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold" Text="AVD-Compatible Only"/>
              <TextBlock Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11.5" Foreground="#444">
Only VM families suitable for AVD host pools are considered: D, E, F, NV, NC, B, M.
Excluded: A (too weak), H (HPC), ND (AI training), L (storage), DC/EC (confidential), ARM-based.
              </TextBlock>

              <TextBlock TextWrapping="Wrap" FontSize="12" FontWeight="SemiBold" Text="RAM per VM Series"/>
              <TextBlock Margin="0,4,0,4" TextWrapping="Wrap" FontSize="11.5" FontFamily="Consolas" Foreground="#444">
D = 4 GB/vCPU    E = 8 GB/vCPU
F = 2 GB/vCPU   NV = 7 GB/vCPU
NC = 8 GB/vCPU    B = 4 GB/vCPU
M = 28 GB/vCPU
              </TextBlock>
            </StackPanel>
            </ScrollViewer>
          </Border>
        </Grid>
      </TabItem>

      <!-- TAB 4: Results -->
      <TabItem Header="Results">
        <Grid Margin="10">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
          <StackPanel Grid.Row="0" Margin="0,0,0,6">
            <TextBlock FontSize="14" FontWeight="Bold" Text="Sizing Results"/>
            <TextBlock Foreground="#666" FontSize="11" TextWrapping="Wrap" Text="Click 'Calculate' to compute host sizing. Then 'Pick VM from Azure' to find the best matching VM template with pricing. Use 'Export HTML Report' or 'Export JSON' to save results."/>
          </StackPanel>
          <DataGrid Grid.Row="1" x:Name="GridResults" AutoGenerateColumns="True" IsReadOnly="True" Margin="0,4,0,8"/>
          <TextBlock Grid.Row="2" Foreground="#666" FontSize="11" Text="Notes, warnings and recommendations:" Margin="0,0,0,2"/>
          <TextBox Grid.Row="3" x:Name="TxtNotes" Height="170" TextWrapping="Wrap" AcceptsReturn="True"
                   VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="11"/>
        </Grid>
      </TabItem>
    </TabControl>

    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="BtnReset" Content="Reset" Width="80" Margin="0,0,20,0" Background="#C62828" Foreground="White" FontWeight="Bold"/>
      <Button x:Name="BtnCalculate" Content="Calculate" Width="120" Margin="0,0,10,0" FontWeight="Bold"/>
      <Button x:Name="BtnPickVm" Content="Pick VM from Azure" Width="160" Margin="0,0,10,0"/>
      <Button x:Name="BtnExportJson" Content="Export JSON" Width="120" Margin="0,0,10,0"/>
      <Button x:Name="BtnExportReport" Content="Export HTML Report" Width="140" Margin="0,0,10,0"/>
      <Button x:Name="BtnClose" Content="Close" Width="100"/>
    </StackPanel>
  </Grid>
</Window>
"@
$XamlString = $XamlString -replace '&(?!amp;|lt;|gt;|quot;|apos;)', '&amp;'
#endregion

#region Build UI + Bind
$xmlReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($XamlString))
$Window = [Windows.Markup.XamlReader]::Load($xmlReader)
$Window.Title = "AVD Sizing Calculator v$ScriptVersion"

# Workload tab
$CmbHostPoolType = $Window.FindName('CmbHostPoolType'); $CmbWorkload = $Window.FindName('CmbWorkload')
$TxtTotalUsers = $Window.FindName('TxtTotalUsers'); $CmbConcurrencyMode = $Window.FindName('CmbConcurrencyMode')
$TxtConcurrencyValue = $Window.FindName('TxtConcurrencyValue'); $LblConcurrencyHint = $Window.FindName('LblConcurrencyHint')
$TxtPeakFactor = $Window.FindName('TxtPeakFactor'); $TxtNPlusOne = $Window.FindName('TxtNPlusOne')
$TxtExtraHeadroomPct = $Window.FindName('TxtExtraHeadroomPct')
$TxtProfileGB = $Window.FindName('TxtProfileGB'); $TxtProfileGrowthPct = $Window.FindName('TxtProfileGrowthPct')
$TxtProfileOverheadPct = $Window.FindName('TxtProfileOverheadPct')
$TxtCpuUtil = $Window.FindName('TxtCpuUtil'); $TxtMemUtil = $Window.FindName('TxtMemUtil'); $TxtVirtOverhead = $Window.FindName('TxtVirtOverhead')
$TxtMinVcpuHost = $Window.FindName('TxtMinVcpuHost'); $TxtMaxVcpuHost = $Window.FindName('TxtMaxVcpuHost')
$CmbLoadBalancing = $Window.FindName('CmbLoadBalancing'); $TxtMaxSessionLimit = $Window.FindName('TxtMaxSessionLimit')

# Applications tab - checkboxes
$ChkOffice = $Window.FindName('ChkOffice'); $ChkTeams = $Window.FindName('ChkTeams')
$ChkBrowser = $Window.FindName('ChkBrowser'); $ChkOutlook = $Window.FindName('ChkOutlook')
$ChkPdf = $Window.FindName('ChkPdf'); $ChkErp = $Window.FindName('ChkErp'); $ChkPowerBi = $Window.FindName('ChkPowerBi')
$ChkSqlExpress = $Window.FindName('ChkSqlExpress'); $ChkSqlStd = $Window.FindName('ChkSqlStd')
$ChkPostgres = $Window.FindName('ChkPostgres'); $ChkMySql = $Window.FindName('ChkMySql'); $ChkSqlite = $Window.FindName('ChkSqlite')
$ChkVS = $Window.FindName('ChkVS'); $ChkVSCode = $Window.FindName('ChkVSCode')
$ChkDocker = $Window.FindName('ChkDocker'); $ChkGit = $Window.FindName('ChkGit')
$ChkAutoCAD = $Window.FindName('ChkAutoCAD'); $ChkRevit = $Window.FindName('ChkRevit')
$ChkSolidWorks = $Window.FindName('ChkSolidWorks'); $ChkVideoEdit = $Window.FindName('ChkVideoEdit')
$TxtDbDataGB = $Window.FindName('TxtDbDataGB')

# VM Template tab
$TxtLocation = $Window.FindName('TxtLocation'); $CmbVmSeries = $Window.FindName('CmbVmSeries')
$CmbCurrency = $Window.FindName('CmbCurrency'); $ChkHidePricing = $Window.FindName('ChkHidePricing')
$TxtPublisher = $Window.FindName('TxtPublisher'); $TxtOffer = $Window.FindName('TxtOffer')
$CmbSku = $Window.FindName('CmbSku'); $TxtVersion = $Window.FindName('TxtVersion')
$BtnAzLogin = $Window.FindName('BtnAzLogin'); $BtnDiscoverSkus = $Window.FindName('BtnDiscoverSkus')
$TxtTemplateOut = $Window.FindName('TxtTemplateOut')

# Results
$GridResults = $Window.FindName('GridResults'); $TxtNotes = $Window.FindName('TxtNotes')
$BtnCalculate = $Window.FindName('BtnCalculate'); $BtnPickVm = $Window.FindName('BtnPickVm')
$BtnExportJson = $Window.FindName('BtnExportJson'); $BtnExportReport = $Window.FindName('BtnExportReport')
$BtnReset = $Window.FindName('BtnReset'); $BtnClose = $Window.FindName('BtnClose')

$script:LastSizing = $null; $script:LastVmPick = $null; $script:LastVmPrice = $null

# App checkbox mapping
$script:AppCheckboxMap = [ordered]@{
  'Microsoft 365 (Word/Excel/PPT)' = $ChkOffice; 'Microsoft Teams' = $ChkTeams
  'Web Browser (Edge/Chrome)' = $ChkBrowser; 'Microsoft Outlook' = $ChkOutlook
  'PDF Editor (Acrobat/Foxit)' = $ChkPdf; 'ERP Client (SAP GUI / Dynamics)' = $ChkErp
  'Power BI Desktop' = $ChkPowerBi; 'SQL Server Express/Developer' = $ChkSqlExpress
  'SQL Server Standard/Enterprise' = $ChkSqlStd; 'PostgreSQL' = $ChkPostgres
  'MySQL / MariaDB' = $ChkMySql; 'SQLite / MS Access (local DB)' = $ChkSqlite
  'Visual Studio' = $ChkVS; 'VS Code' = $ChkVSCode; 'Docker Desktop' = $ChkDocker
  'Git / Build Tools' = $ChkGit; 'AutoCAD / AutoCAD LT' = $ChkAutoCAD
  'Revit / 3ds Max' = $ChkRevit; 'SolidWorks / CATIA' = $ChkSolidWorks
  'Video Editing (Premiere/DaVinci)' = $ChkVideoEdit
}

function Get-SelectedApps {
  $selected = [System.Collections.Generic.List[string]]::new()
  foreach ($kv in $script:AppCheckboxMap.GetEnumerator()) {
    if ($kv.Value.IsChecked -eq $true) { $selected.Add($kv.Key) }
  }
  return @([string[]]$selected.ToArray())
}
#endregion

#region Events
$CmbConcurrencyMode.add_SelectionChanged({ $LblConcurrencyHint.Text = if ((Get-ComboText -Combo $CmbConcurrencyMode) -eq 'Percent') { '% (e.g. 60)' } else { '# (e.g. 25)' } })
$BtnAzLogin.add_Click({ if (-not (Test-AzAvailable)) { Write-AzModuleInstallHint; return }; [void](Connect-AzIfNeeded) })
$BtnDiscoverSkus.add_Click({
  try { if (-not (Test-AzAvailable)) { Write-AzModuleInstallHint; return }; if (-not (Connect-AzIfNeeded)) { return }
    $loc = ConvertTo-ArmRegionName -LocationText $TxtLocation.Text
    $skus = Get-AzVMImageSku -Location $loc -PublisherName $TxtPublisher.Text.Trim() -Offer $TxtOffer.Text.Trim()
    $CmbSku.Items.Clear(); foreach ($s in $skus) { [void]$CmbSku.Items.Add($s.Skus) }
    if ($CmbSku.Items.Count -gt 0) { $CmbSku.SelectedIndex = 0 }; Write-UiInfo "Loaded: $($CmbSku.Items.Count) SKUs"
  } catch { Write-UiError "Failed: $($_.Exception.Message)" }
})

$BtnCalculate.add_Click({
  try {
    $hostPoolType = if ((Get-ComboText -Combo $CmbHostPoolType) -like 'Personal*') { 'Personal' } else { 'Pooled' }
    $workloadRaw = Get-ComboText -Combo $CmbWorkload
    $workload = ($workloadRaw -split '\s*[\u2014—-]\s*')[0].Trim()  # Extract 'Light' from 'Light — 6 users/vCPU...'
    $totalUsers = ConvertTo-IntSafe -Text $TxtTotalUsers.Text -Default 0
    if ($totalUsers -lt 1) { throw "Total Users must be >= 1." }

    # Gather selected applications
    $selectedApps = Get-SelectedApps
    $appOverhead = $null
    if (@($selectedApps).Count -gt 0) {
      $usersPerHostEstimate = if ($hostPoolType -eq 'Pooled') {
        switch ($workload) { 'Light' { 8 } 'Medium' { 4 } 'Heavy' { 2 } 'Power' { 1 } default { 4 } }
      } else { 1 }
      $appOverhead = Get-ApplicationOverhead -SelectedApps $selectedApps -UsersPerHost $usersPerHostEstimate
    }
    # Parse load balancing selection
    $lbRaw = Get-ComboText -Combo $CmbLoadBalancing
    $lbAlgo = if ($lbRaw -like 'Depth*') { 'DepthFirst' } else { 'BreadthFirst' }
    $maxSessLimit = ConvertTo-IntSafe -Text $TxtMaxSessionLimit.Text -Default 0

    $script:LastSizing = Get-AvdSizing -HostPoolType $hostPoolType -Workload $workload -TotalUsers $totalUsers `
      -ConcurrencyMode (Get-ComboText -Combo $CmbConcurrencyMode) `
      -ConcurrencyValue (ConvertTo-DoubleSafe -Text $TxtConcurrencyValue.Text -Default 60) `
      -PeakFactor (ConvertTo-DoubleSafe -Text $TxtPeakFactor.Text -Default 1.0) `
      -CpuTargetUtil (ConvertTo-DoubleSafe -Text $TxtCpuUtil.Text -Default 0.80) `
      -MemTargetUtil (ConvertTo-DoubleSafe -Text $TxtMemUtil.Text -Default 0.80) `
      -SystemResourceReserve (ConvertTo-DoubleSafe -Text $TxtVirtOverhead.Text -Default 0.15) `
      -MinVcpuPerHost (ConvertTo-IntSafe -Text $TxtMinVcpuHost.Text -Default 8) `
      -MaxVcpuPerHost (ConvertTo-IntSafe -Text $TxtMaxVcpuHost.Text -Default 128) `
      -NPlusOneHosts (ConvertTo-IntSafe -Text $TxtNPlusOne.Text -Default 1) `
      -ExtraHeadroomPercent (ConvertTo-DoubleSafe -Text $TxtExtraHeadroomPct.Text -Default 0) `
      -ProfileContainerGB (ConvertTo-DoubleSafe -Text $TxtProfileGB.Text -Default 30) `
      -ProfileGrowthPercent (ConvertTo-DoubleSafe -Text $TxtProfileGrowthPct.Text -Default 20) `
      -ProfileOverheadPercent (ConvertTo-DoubleSafe -Text $TxtProfileOverheadPct.Text -Default 10) `
      -LoadBalancing $lbAlgo -MaxSessionLimit $maxSessLimit `
      -AppOverhead $appOverhead

    if ($script:LastSizing -is [System.Array]) { $script:LastSizing = $script:LastSizing | Select-Object -First 1 }

    $TxtTemplateOut.Text = ([ordered]@{
      sizing=[ordered]@{ hostPoolType=$script:LastSizing.HostPoolType; workload=$script:LastSizing.Workload
        peakUsers=$script:LastSizing.PeakConcurrentUsers; hostsTotal=$script:LastSizing.RecommendedHostsTotal
        vcpu=$script:LastSizing.Recommended.VcpuPerHost; ramGB=$script:LastSizing.Recommended.RamGB_Provisioned }
      osDisk=[ordered]@{ sku=$script:LastSizing.Disks.SessionHostDisks.OsDisk.RecommendedType
        sizeGiB=$script:LastSizing.Disks.SessionHostDisks.OsDisk.SuggestedSizeGiB }
      fsLogix=[ordered]@{ tier=$script:LastSizing.Disks.FsLogixStorage.RecommendedTier
        riskLevel=$script:LastSizing.Disks.FsLogixStorage.StorageRisk.Level }
      apps = $selectedApps
    } | ConvertTo-Json -Depth 10)

    $hidePricingVal = ($ChkHidePricing.IsChecked -eq $true)
    Set-ResultsGrid -Sizing $script:LastSizing -VmPick $script:LastVmPick -VmPrice $script:LastVmPrice `
      -GridResults $GridResults -TxtNotes $TxtNotes -HidePricing $hidePricingVal
    $Window.FindName('Tabs').SelectedIndex = 3
  } catch { Write-UiError "Calculation failed: $($_.Exception.Message)" }
})

$BtnPickVm.add_Click({
  try {
    if (-not $script:LastSizing) { Write-UiWarning 'Calculate first.'; return }
    if (-not (Test-AzAvailable)) { Write-AzModuleInstallHint; return }
    if (-not (Connect-AzIfNeeded)) { return }
    $loc = ConvertTo-ArmRegionName -LocationText $TxtLocation.Text
    $currency = Get-CurrencyCode -Combo $CmbCurrency
    $seriesRaw = Get-ComboText -Combo $CmbVmSeries
    $series = switch -Wildcard ($seriesRaw) { 'D*' {'D'} 'E*' {'E'} 'F*' {'F'} 'NV*' {'NV'} 'NC*' {'NC'} 'B*' {'B'} default {'Any'} }

    # Use PreferredSeries from sizing if user chose 'Any'
    $hasDb = $script:LastSizing.AppOverhead -and $script:LastSizing.AppOverhead.HasDatabaseEngine
    $needsGpu = $script:LastSizing.AppOverhead -and $script:LastSizing.AppOverhead.RequiresGPU
    if ($series -eq 'Any' -and $script:LastSizing.PreferredSeries) {
      $series = $script:LastSizing.PreferredSeries
    }

    $sizes = Get-AzVmSizesInLocation $loc
    if (-not $sizes -or $sizes.Count -lt 1) {
      Write-UiWarning "No VM sizes returned for region '$loc'.`n`nPlease check:`n- Is the region name correct? (e.g. 'Switzerland North')`n- Are you signed in? Click 'Azure Login' first.`n- Does your subscription have access to this region?"
      return
    }

    # Pass the user's vCPU max range to Get-BestVmSize
    $maxVcpuRange = ConvertTo-IntSafe -Text $TxtMaxVcpuHost.Text -Default 128
    $best = Get-BestVmSize -Sizes $sizes -MinVcpu $script:LastSizing.Recommended.VcpuPerHost `
      -MinRamGB $script:LastSizing.Recommended.RamGB_Provisioned -MaxVcpu $maxVcpuRange `
      -Series $series -Workload $script:LastSizing.Workload -HasDatabase $hasDb -RequiresGPU $needsGpu

    # Handle strict error (no VM found in requested series)
    if (-not $best) {
      Write-UiWarning "No AVD-compatible VM found (min $($script:LastSizing.Recommended.VcpuPerHost) vCPU, min $([Math]::Round($script:LastSizing.Recommended.RamGB_Provisioned,0)) GB RAM)."
      return
    }
    if ($best.PSObject.Properties.Name -contains '_Error' -and $best._Error) {
      Write-UiWarning $best._Message
      return
    }

    # Check if vCPU range was exceeded
    $rangeMsg = ''
    if ($best.PSObject.Properties.Name -contains 'RangeExceededNote' -and $best.RangeExceededNote) {
      $rangeMsg = "`n`nWARNING: $($best.RangeExceededNote)"
    }

    $script:LastVmPick = $best
    $script:LastVmPrice = Get-AzVmHourlyRetailPrice -ArmRegionName $loc -ArmSkuName $best.Name -CurrencyCode $currency
    $hidePricingVal2 = ($ChkHidePricing.IsChecked -eq $true)
    Set-ResultsGrid -Sizing $script:LastSizing -VmPick $script:LastVmPick -VmPrice $script:LastVmPrice `
      -GridResults $GridResults -TxtNotes $TxtNotes -HidePricing $hidePricingVal2
    $infoMsg = "Selected: $($best.Name) ($($best.NumberOfCores) vCPU, $([Math]::Round($best.MemoryInMB/1024,1)) GB)$rangeMsg"
    if ($rangeMsg) { Write-UiWarning $infoMsg } else { Write-UiInfo $infoMsg }
  } catch { Write-UiError "Failed: $($_.Exception.Message)" }
})

$BtnExportJson.add_Click({
  try { if (-not $script:LastSizing) { Write-UiWarning 'Calculate first.'; return }
    $data = [ordered]@{ sizing=$script:LastSizing; vm=$script:LastVmPick; price=$script:LastVmPrice; template=$TxtTemplateOut.Text
      exportedAt=(Get-Date).ToString('o'); version=$ScriptVersion }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new(); $dlg.Filter='JSON|*.json'; $dlg.FileName='avd-sizing.json'
    if ($dlg.ShowDialog()) { [IO.File]::WriteAllText($dlg.FileName, ($data | ConvertTo-Json -Depth 18), [Text.Encoding]::UTF8)
      Write-UiInfo "Exported: $($dlg.FileName)" }
  } catch { Write-UiError "Export failed: $($_.Exception.Message)" }
})

$BtnExportReport.add_Click({
  try {
    if (-not $script:LastSizing) { Write-UiWarning 'Calculate first.'; return }
    $s = $script:LastSizing; $r = $s.Recommended

    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'HTML Report|*.html'
    $dlg.FileName = "AVD-Sizing-Report-$(Get-Date -Format 'yyyy-MM-dd').html"
    if (-not $dlg.ShowDialog()) { return }
    $reportPath = $dlg.FileName

    # Helper to HTML-encode
    function esc($t) { [System.Web.HttpUtility]::HtmlEncode([string]$t) }

    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $region = esc $TxtLocation.Text

    # Build apps table rows
    $appsRows = ''
    if ($s.AppOverhead -and @($s.AppOverhead.SelectedApps).Count -gt 0) {
      foreach ($appName in $s.AppOverhead.SelectedApps) {
        if ($script:ApplicationCatalog.Contains($appName)) {
          $app = $script:ApplicationCatalog[$appName]
          $appsRows += "<tr><td>$(esc $appName)</td><td>$(esc $app.Category)</td><td>+$($app.CpuOverheadPerUser) vCPU</td><td>+$($app.RamMBPerUser) MB</td></tr>`n"
        }
      }
    }

    # Autoscale table
    $autoscaleHtml = ''
    if ($s.Autoscale) {
      $autoscaleHtml = @"
      <h3>Autoscale Scaling Plan (Recommended)</h3>
      <table>
        <thead><tr><th>Phase</th><th>Load Balancing</th><th>Min Hosts %</th><th>Capacity Threshold</th></tr></thead>
        <tbody>
          <tr><td>Ramp-Up</td><td>$($s.Autoscale.RampUp.LoadBalancing)</td><td>$($s.Autoscale.RampUp.MinHostsPct)%</td><td>$($s.Autoscale.RampUp.CapacityThresholdPct)%</td></tr>
          <tr><td>Peak</td><td>$($s.Autoscale.Peak.LoadBalancing)</td><td>$($s.Autoscale.Peak.MinHostsPct)%</td><td>$($s.Autoscale.Peak.CapacityThresholdPct)%</td></tr>
          <tr><td>Ramp-Down</td><td>$($s.Autoscale.RampDown.LoadBalancing)</td><td>$($s.Autoscale.RampDown.MinHostsPct)%</td><td>$($s.Autoscale.RampDown.CapacityThresholdPct)%</td></tr>
          <tr><td>Off-Peak</td><td>$($s.Autoscale.OffPeak.LoadBalancing)</td><td>$($s.Autoscale.OffPeak.MinHostsPct)%</td><td>&mdash;</td></tr>
        </tbody>
      </table>
"@
    }

    # VM + pricing section
    $hidePricing = ($ChkHidePricing.IsChecked -eq $true)
    $vmHtml = ''
    if ($script:LastVmPick) {
      $vmName = esc $script:LastVmPick.Name
      $vmCores = $script:LastVmPick.NumberOfCores
      $vmRam = [Math]::Round($script:LastVmPick.MemoryInMB/1024,1)
      $priceRows = ''
      $priceNote = ''
      if (-not $hidePricing -and $script:LastVmPrice -and -not ($script:LastVmPrice.PSObject.Properties.Name -contains 'Error' -and $script:LastVmPrice.Error)) {
        $monthly = [Math]::Round($script:LastVmPrice.RetailPricePerHour * 730, 2)
        $totalMonthly = [Math]::Round($monthly * $s.RecommendedHostsTotal, 2)
        $cur = esc $script:LastVmPrice.CurrencyCode
        $priceRows = @"
          <tr><td>Price/Hour</td><td>$($script:LastVmPrice.RetailPricePerHour) $cur</td></tr>
          <tr><td>Est. Monthly/Host</td><td>$monthly $cur</td></tr>
          <tr class="highlight"><td>Est. Monthly Total</td><td><strong>$totalMonthly $cur</strong> ($($s.RecommendedHostsTotal) hosts)</td></tr>
"@
        $priceNote = '<p class="note">Retail list price. No discounts, Reserved Instances, Savings Plans, or Azure Hybrid Benefit applied.</p>'
      }
      $vmHtml = @"
      <section>
        <h2><span class="num">6</span> Azure VM Selection</h2>
        <table>
          <tbody>
            <tr><td>VM Size</td><td><strong>$vmName</strong></td></tr>
            <tr><td>vCPU</td><td>$vmCores</td></tr>
            <tr><td>RAM</td><td>$vmRam GB</td></tr>
            $priceRows
          </tbody>
        </table>
        $priceNote
      </section>
"@
    }

    # Warnings
    $warningsHtml = ''
    if ($s.Notes -and @($s.Notes).Count -gt 0) {
      $items = ''; foreach ($n in $s.Notes) { $items += "<li>$(esc $n)</li>`n" }
      $warningsHtml = "<div class='warnings'><h3>Warnings</h3><ul>$items</ul></div>"
    }

    # Disk section
    $diskHtml = ''
    if ($s.Disks) {
      $d = $s.Disks
      $dataDiskHtml = ''
      if ($d.SessionHostDisks.DataDisk -and $d.SessionHostDisks.DataDisk.Required) {
        $dataDiskHtml = @"
        <h3>Data Disk</h3>
        <table>
          <tbody>
            <tr><td>Type</td><td>$($d.SessionHostDisks.DataDisk.Type)</td></tr>
            <tr><td>Min IOPS</td><td>$($d.SessionHostDisks.DataDisk.MinIOPS)</td></tr>
            <tr><td>Min Size</td><td>$($d.SessionHostDisks.DataDisk.MinGB) GB</td></tr>
          </tbody>
        </table>
"@
      }
      $diskHtml = @"
      <section>
        <h2><span class="num">5</span> Storage</h2>
        <div class="grid-2">
          <div>
            <h3>OS Disk</h3>
            <table>
              <tbody>
                <tr><td>SKU</td><td><strong>$($d.SessionHostDisks.OsDisk.RecommendedType)</strong></td></tr>
                <tr><td>Size</td><td>$($d.SessionHostDisks.OsDisk.SuggestedSizeGiB) GiB</td></tr>
                <tr><td>IOPS</td><td>$($d.SessionHostDisks.OsDisk.ProvisionedIOPS) (burst: $($d.SessionHostDisks.OsDisk.BurstIOPS))</td></tr>
              </tbody>
            </table>
            $dataDiskHtml
          </div>
          <div>
            <h3>FSLogix Profile Storage</h3>
            <table>
              <tbody>
                <tr><td>Tier</td><td><strong>$($d.FsLogixStorage.RecommendedTier)</strong></td></tr>
                <tr><td>Steady IOPS</td><td>$($d.FsLogixStorage.PerformanceTargets.SteadyStateIOPS)</td></tr>
                <tr><td>Burst IOPS</td><td>$($d.FsLogixStorage.PerformanceTargets.BurstIOPS)</td></tr>
                <tr><td>Per User</td><td>$($s.FsLogix.PlannedPerUserGB) GB</td></tr>
                <tr><td>Total at Peak</td><td>$($s.FsLogix.PlannedTotalGB_AtPeak) GB</td></tr>
                <tr><td>Risk Level</td><td>$($d.FsLogixStorage.StorageRisk.Level)</td></tr>
              </tbody>
            </table>
          </div>
        </div>
      </section>
"@
    }

    # Applications section
    $appSection = ''
    if ($appsRows) {
      $gpuLine = ''; if ($s.AppOverhead.RequiresGPU) { $gpuLine = '<span class="badge gpu">GPU Required</span>' }
      $dbLine = ''; if ($s.AppOverhead.HasDatabaseEngine) { $dbLine = "<span class='badge db'>Database: $($s.AppOverhead.DatabaseApps -join ', ')</span>" }
      $appSection = @"
      <section>
        <h2><span class="num">3</span> Applications</h2>
        $gpuLine $dbLine
        <table>
          <thead><tr><th>Application</th><th>Category</th><th>CPU/User</th><th>RAM/User</th></tr></thead>
          <tbody>$appsRows</tbody>
          <tfoot><tr class="highlight"><td colspan="2"><strong>Total Overhead/Host</strong></td><td><strong>+$($s.AppOverhead.TotalCpuOverhead) vCPU</strong></td><td><strong>+$([Math]::Round($s.AppOverhead.TotalRamOverheadGB * 1024)) MB</strong></td></tr></tfoot>
        </table>
      </section>
"@
    } else {
      $appSection = @"
      <section>
        <h2><span class="num">3</span> Applications</h2>
        <p class="note">No specific applications selected. Sizing based on Microsoft baseline workload profile.</p>
      </section>
"@
    }

    # LB description
    $lbDesc = if ($s.LoadBalancing -eq 'BreadthFirst') {
      'Breadth-first distributes sessions evenly across all available hosts. Best user experience but higher cost as all hosts must stay powered on.'
    } else {
      'Depth-first saturates one host at a time before moving to the next. Cost-optimised as idle hosts can be deallocated.'
    }

    # Build HTML
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AVD Sizing Report &mdash; $(esc $generated)</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=JetBrains+Mono:wght@400;500&display=swap');
  :root {
    --bg: #0f1117; --surface: #1a1d27; --surface2: #242835;
    --border: #2e3345; --text: #e1e4ed; --text2: #8b90a5;
    --accent: #3b82f6; --accent2: #60a5fa; --accent-glow: rgba(59,130,246,0.15);
    --green: #34d399; --amber: #fbbf24; --red: #f87171;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'DM Sans', sans-serif; background: var(--bg); color: var(--text); line-height: 1.6; padding: 0; }
  .container { max-width: 960px; margin: 0 auto; padding: 48px 32px; }

  /* Header */
  header { background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); border-bottom: 1px solid var(--border); padding: 48px 0; }
  header .container { display: flex; justify-content: space-between; align-items: flex-end; padding-top: 0; padding-bottom: 0; }
  header h1 { font-size: 28px; font-weight: 700; letter-spacing: -0.5px; }
  header h1 span { color: var(--accent2); }
  .meta { text-align: right; color: var(--text2); font-size: 13px; }
  .meta strong { color: var(--text); font-weight: 500; }

  /* Hero summary cards */
  .hero { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin: 32px 0; }
  .card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 20px; position: relative; overflow: hidden; }
  .card::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: var(--accent); }
  .card .label { font-size: 11px; text-transform: uppercase; letter-spacing: 1.2px; color: var(--text2); font-weight: 500; margin-bottom: 8px; }
  .card .value { font-size: 28px; font-weight: 700; font-family: 'JetBrains Mono', monospace; color: var(--accent2); }
  .card .sub { font-size: 12px; color: var(--text2); margin-top: 4px; }

  /* Sections */
  section { margin: 40px 0; }
  h2 { font-size: 20px; font-weight: 700; margin-bottom: 20px; padding-bottom: 12px; border-bottom: 1px solid var(--border); display: flex; align-items: center; gap: 12px; }
  h2 .num { background: var(--accent); color: white; width: 28px; height: 28px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: 14px; font-weight: 700; flex-shrink: 0; }
  h3 { font-size: 15px; font-weight: 600; color: var(--text2); margin: 16px 0 10px; text-transform: uppercase; letter-spacing: 0.8px; }

  /* Tables */
  table { width: 100%; border-collapse: collapse; margin: 12px 0 20px; }
  th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 1px; color: var(--text2); padding: 10px 14px; border-bottom: 2px solid var(--border); font-weight: 600; }
  td { padding: 10px 14px; border-bottom: 1px solid var(--border); font-size: 14px; }
  tr:hover td { background: var(--accent-glow); }
  tr.highlight td { background: var(--surface2); font-weight: 500; }
  tbody tr:last-child td { border-bottom: none; }
  tfoot td { border-top: 2px solid var(--border); border-bottom: none; }

  /* Badges */
  .badge { display: inline-block; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; margin: 4px 4px 12px 0; }
  .badge.gpu { background: rgba(249,115,22,0.15); color: #fb923c; border: 1px solid rgba(249,115,22,0.3); }
  .badge.db { background: rgba(168,85,247,0.15); color: #c084fc; border: 1px solid rgba(168,85,247,0.3); }
  .badge.series { background: var(--accent-glow); color: var(--accent2); border: 1px solid rgba(59,130,246,0.3); }

  /* Grid */
  .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }
  .grid-2 > div { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 20px; }

  /* Notes */
  .note { color: var(--text2); font-size: 13px; font-style: italic; margin: 8px 0; }
  .warnings { background: rgba(251,191,36,0.08); border: 1px solid rgba(251,191,36,0.2); border-radius: 10px; padding: 16px 20px; margin: 16px 0; }
  .warnings h3 { color: var(--amber); margin-top: 0; }
  .warnings ul { padding-left: 20px; }
  .warnings li { font-size: 13px; color: var(--amber); margin: 4px 0; }

  /* Footer */
  footer { margin-top: 48px; padding-top: 24px; border-top: 1px solid var(--border); color: var(--text2); font-size: 12px; }
  footer a { color: var(--accent2); text-decoration: none; }
  footer a:hover { text-decoration: underline; }

  @media print { body { background: white; color: #1a1a1a; } .card { border: 1px solid #ccc; } .card::before { background: #333; } table { font-size: 12px; } }
  @media (max-width: 768px) { .hero { grid-template-columns: 1fr 1fr; } .grid-2 { grid-template-columns: 1fr; } }
</style>
</head>
<body>
<header>
  <div class="container">
    <div>
      <h1>Azure Virtual Desktop<br><span>Sizing Report</span></h1>
    </div>
    <div class="meta">
      <strong>$generated</strong><br>
      Calculator v$ScriptVersion<br>
      Region: $region
    </div>
  </div>
</header>

<div class="container">
  <!-- Hero Summary -->
  <div class="hero">
    <div class="card">
      <div class="label">Total Users</div>
      <div class="value">$($s.TotalUsers)</div>
      <div class="sub">Peak: $($s.PeakConcurrentUsers) concurrent</div>
    </div>
    <div class="card">
      <div class="label">Session Hosts</div>
      <div class="value">$($s.RecommendedHostsTotal)</div>
      <div class="sub">incl. N+1 redundancy</div>
    </div>
    <div class="card">
      <div class="label">Per Host</div>
      <div class="value">$($r.VcpuPerHost)v / $($r.RamGB_Provisioned)G</div>
      <div class="sub">$($r.UsersPerHost) users/host</div>
    </div>
    <div class="card">
      <div class="label">Host Pool</div>
      <div class="value" style="font-size:20px">$(esc $s.HostPoolType)</div>
      <div class="sub">$(esc $s.Workload) &bull; $(esc $s.LoadBalancing)</div>
    </div>
  </div>

  <!-- 1. Executive Summary -->
  <section>
    <h2><span class="num">1</span> Executive Summary</h2>
    <span class="badge series">$($s.PreferredSeries)-series preferred</span>
    <table>
      <tbody>
        <tr><td>Host Pool Type</td><td><strong>$($s.HostPoolType)</strong> ($($s.Mode))</td></tr>
        <tr><td>Workload Class</td><td>$($s.Workload)</td></tr>
        <tr><td>Total Users</td><td>$($s.TotalUsers)</td></tr>
        <tr><td>Peak Concurrent</td><td>$($s.PeakConcurrentUsers)</td></tr>
        <tr><td>Users per Host</td><td>$($r.UsersPerHost)</td></tr>
        <tr><td>vCPU per Host</td><td>$($r.VcpuPerHost)</td></tr>
        <tr><td>RAM per Host</td><td>$($r.RamGB_Provisioned) GB (provisioned)</td></tr>
        <tr class="highlight"><td>Session Hosts Total</td><td><strong>$($s.RecommendedHostsTotal)</strong> (incl. N+1)</td></tr>
        <tr><td>Guideline Examples</td><td>$($s.Examples)</td></tr>
      </tbody>
    </table>
  </section>

  <!-- 2. Workload -->
  <section>
    <h2><span class="num">2</span> Workload Profile</h2>
    <table>
      <tbody>
        <tr><td>CPU Target Utilisation</td><td>$([int]($s.CpuTargetUtil*100))%</td></tr>
        <tr><td>Memory Target Utilisation</td><td>$([int]($s.MemTargetUtil*100))%</td></tr>
        <tr><td>System Reserve</td><td>$([int]($s.SystemResourceReserve*100))% (MS: 15&ndash;20%)</td></tr>
      </tbody>
    </table>
  </section>

  <!-- 3. Applications -->
  $appSection

  <!-- 4. Load Balancing -->
  <section>
    <h2><span class="num">4</span> Load Balancing</h2>
    <table>
      <tbody>
        <tr><td>Algorithm</td><td><strong>$($s.LoadBalancing)</strong></td></tr>
        <tr><td>Max Session Limit</td><td>$($s.MaxSessionLimit)</td></tr>
      </tbody>
    </table>
    <p class="note">$(esc $lbDesc)</p>
    $autoscaleHtml
  </section>

  <!-- 5. Storage -->
  $diskHtml

  <!-- 6. VM -->
  $vmHtml

  <!-- 7. Notes -->
  <section>
    <h2><span class="num">7</span> Notes &amp; Recommendations</h2>
    $warningsHtml
    <h3>Microsoft References</h3>
    <ul style="padding-left:20px; color: var(--text2); font-size:13px;">
      <li><a href="https://learn.microsoft.com/en-us/windows-server/remote/remote-desktop-services/session-host-virtual-machine-sizing-guidelines">Session Host VM Sizing Guidelines</a></li>
      <li><a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/host-pool-load-balancing">Host Pool Load Balancing</a></li>
      <li><a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/autoscale-create-assign-scaling-plan">Autoscale Scaling Plans</a></li>
    </ul>
  </section>

  <footer>
    Generated by AVD Sizing Calculator v$ScriptVersion &bull; All recommendations should be validated with <a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/insights">AVD Insights</a> monitoring and pilot deployments.
  </footer>
</div>
</body>
</html>
"@

    [IO.File]::WriteAllText($reportPath, $html, [Text.Encoding]::UTF8)
    Write-UiInfo "Report exported: $reportPath"

  } catch { Write-UiError "Report export failed: $($_.Exception.Message)" }
})

$BtnReset.add_Click({
  # Clear calculation state
  $script:LastSizing = $null; $script:LastVmPick = $null; $script:LastVmPrice = $null

  # Workload tab - reset to defaults
  $CmbHostPoolType.SelectedIndex = 0
  $CmbWorkload.SelectedIndex = 1       # Medium
  $TxtTotalUsers.Text = '100'
  $CmbConcurrencyMode.SelectedIndex = 0
  $TxtConcurrencyValue.Text = '60'
  $TxtPeakFactor.Text = '1.0'
  $TxtNPlusOne.Text = '1'
  $TxtExtraHeadroomPct.Text = '0'
  $CmbLoadBalancing.SelectedIndex = 0  # Breadth-first
  $TxtMaxSessionLimit.Text = '0'

  # Tuning parameters
  $TxtProfileGB.Text = '30'
  $TxtProfileGrowthPct.Text = '20'
  $TxtProfileOverheadPct.Text = '10'
  $TxtCpuUtil.Text = '0.80'
  $TxtMemUtil.Text = '0.80'
  $TxtVirtOverhead.Text = '0.15'
  $TxtMinVcpuHost.Text = '8'
  $TxtMaxVcpuHost.Text = '128'

  # Applications tab - uncheck all
  foreach ($kv in $script:AppCheckboxMap.GetEnumerator()) {
    $kv.Value.IsChecked = $false
  }
  if ($TxtDbDataGB) { $TxtDbDataGB.Text = '100' }

  # Azure tab
  $CmbVmSeries.SelectedIndex = 0      # Any (auto)
  $ChkHidePricing.IsChecked = $true    # Hide pricing by default
  $TxtTemplateOut.Text = ''

  # Results tab - clear
  $GridResults.ItemsSource = $null
  $TxtNotes.Text = ''

  Write-UiInfo 'All settings reset to defaults.'
})

$BtnClose.add_Click({ $Window.Close() })

$CmbSku.Items.Clear(); @('win11-24h2-avd-m365','win11-23h2-avd-m365','win11-24h2-avd','win11-23h2-avd') | ForEach-Object { [void]$CmbSku.Items.Add($_) }
$CmbSku.SelectedIndex = 0
#endregion

[void]$Window.ShowDialog()