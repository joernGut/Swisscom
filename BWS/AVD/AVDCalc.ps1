<#
================================================================================
  Author:  Jörn Gutting (optimised by Claude)
  Script:  Azure Virtual Desktop (AVD) Sizing Calculator (WPF GUI) - PowerShell 7
  Version: 2.3.1
  Date:    2025-02-19

  CHANGELOG
  =========

  v2.3.1 (2025-02-19)
  --------------------
  - I18N: Full translation of info panels (84 named labels, 116 string keys)
    * "How Apps Affect Sizing" panel: VM Auto-Selection, Database Engines,
      CAD/GPU — all translated (EN/DE/FR/IT)
    * "VM Selection Logic" panel: Series Auto-Selection, Strict Series Mode,
      Smallest Matching VM, AVD-Compatible Only, RAM per VM Series — all translated
  - DESIGN: Workload class description expanded to multi-line with details
    * Each class (Light/Medium/Heavy/Power) now shows: users/vCPU, RAM,
      typical use cases, and recommendations — fully translated in all 4 languages

  v2.3.0 (2025-02-19)
  --------------------
  - BREAKING: Removed -UserCosts parameter (Costs tab now solely under -Expert)
    * Usage: .\AVD-SizingCalculator.ps1              (Standard)
    * Usage: .\AVD-SizingCalculator.ps1 -Expert      (All features)
  - FIX: Apply-Language crash — $Tabs/$Window not in scope
    * Rewritten to use $script:Window.FindName() with null-checks
    * $script:Window set explicitly after XAML load
    * LangMap array syntax fixed
  - UPDATE: Application RAM values revised based on 2024 real-world data
    * Microsoft 365 (Word/Excel/PPT): 512 MB -> 768 MB (real-world 600-900 MB)
    * Microsoft Teams: 768 MB -> 1024 MB (New Teams 800-1200 MB in AVD)
    * Microsoft Outlook: 384 MB -> 512 MB (cached mode 400-600 MB)
    * Power BI Desktop: 2048 MB -> 3072 MB (2-4 GB with datasets)
    * VS Code: 1024 MB -> 768 MB (base ~600 MB, extensions add more)
    * Application notes updated with specific memory ranges
    * Sources: GO-EUC research, MS docs, GitHub issues, VDI community data

  v2.2.9 (2025-02-19)
  --------------------
  - NEW: Multi-language support (EN, DE, FR, IT)
    * Language dropdown in top-right corner of the window
    * All GUI labels, descriptions, buttons, tab headers translated
    * HTML report section titles translated based on selected language
    * Centralized string dictionary ($script:Strings) with Get-Str helper
    * Apply-Language function updates all UI elements on language change
    * Default language: English

  v2.2.8 (2025-02-19)
  --------------------
  - CHANGE: "User Costs" tab renamed to "Costs" — consolidated cost hub
    * Currency selector moved from VM Template tab into Costs tab
    * Tab now contains: Currency, AHB, Operating Schedule, Discounts, Calculate
    * All cost-related settings in a single tab (no longer split across tabs)
    * Currency no longer requires Expert mode (available whenever Costs tab is visible)
  - CLEANUP: Removed PnlCurrency wrapper — CmbCurrency lives directly in Costs tab
  - CLEANUP: Removed duplicate CmbCurrency FindName binding

  v2.2.7 (2025-02-19)
  --------------------
  - CHANGE: HTML report is now mode-aware (Standard vs Expert)
    * Standard mode: Sizing, workload, apps, load balancing, storage, VM selection
      (no pricing, no user costs, no discount information)
    * Expert mode: All standard sections PLUS:
      - VM pricing (list price per hour, monthly estimates)
      - Full User Cost Analysis section with operating schedule
      - Applied Discounts breakdown (CSP/EA, RI, additional)
      - Host costs (monthly/yearly with effective pricing)
      - Per-user costs (named, concurrent, daily, yearly)
      - Monthly and yearly savings from discounts
    * Section numbers adjust dynamically based on visible sections
    * Diagnostics data is never included in the report

  v2.2.6 (2025-02-19)
  --------------------
  - NEW: Diagnostics dialog upgraded from MessageBox to full WPF dialog
    * Disconnect button: Disconnect-AzAccount + Clear-AzContext (full session cleanup)
    * Login button: Connect-AzAccount -Force (ignores cached token, fresh login prompt)
    * Live connection status display with colour feedback (green/red/orange)
    * Buttons disabled when Az.Accounts not installed or not connected
    * Dark-themed dialog matching main window design

  v2.2.5 (2025-02-19)
  --------------------
  - FIX: Diagnostics button crash when LastVmPrice has no Error property
    * Used safe property-exists check consistent with other handlers
  - DESIGN: Window height increased from 940 to 1020px
    * Workload tab left panel no longer requires scrolling at typical resolution
    * Reduced field margins (12px to 8px) and shortened description texts
    * Separator margins tightened (8px to 4px)

  v2.2.4 (2025-02-19)
  --------------------
  - CHANGE: Currency selector moved to Expert mode only
    * Non-Expert mode defaults to CHF silently
    * CSP/partner users typically know their currency and use Expert mode
  - NEW: Discount fields in User Costs tab
    * CSP / EA / Partner discount (% off Azure list price)
    * Reserved Instance discount (1yr ~35%, 3yr ~55%)
    * Additional discount (Savings Plans, promos, custom agreements)
    * Discounts stack cumulatively, capped at 100%
    * Results grid shows: list price, effective price, savings breakdown
    * Notes section details applied discounts with yearly savings projection
    * Hint shown when no discounts entered (typical CSP/EA/RI ranges)

  v2.2.3 (2025-02-19)
  --------------------
  - DESIGN: TabItem headers now fully themed with ControlTemplate override
    * Default WPF TabItem ignores Foreground for header text — fixed with custom template
    * Inactive tabs: muted text (#a6adc8) on dark background (#1e1e2e)
    * Hover: background lightens (#313244), text turns blue (#89b4fa)
    * Selected: dark surface (#313244), blue top border (#89b4fa), bright text (#cdd6f4)
    * Rounded top corners (6px) for modern look

  v2.2.2 (2025-02-19)
  --------------------
  - DESIGN: Dark theme (Catppuccin Mocha) with full WPF style triggers
    * MouseOver highlights on TextBox, ComboBox, CheckBox, Button, TabItem, DataGridRow
    * Focus state for TextBox (blue border + lighter background)
    * ComboBoxItem hover (#89b4fa) and selection (#b4befe) with dark text
    * DataGridRow/Cell selection visible (#585b70) with light foreground
    * DataGridColumnHeader styled (dark background, light text)
    * Button hover (brightens) and pressed (darkens) states
    * TextBox selection brush visible on dark background
  - CHANGE: Pricing information moved to Expert mode only
    * Removed Hide Pricing checkbox from VM Template tab
    * Pricing in GUI results grid: visible only with -Expert
    * Pricing in HTML report: visible only with -Expert
  - CHANGE: ComboBox dropdown items now white background with dark text for readability

  v2.2.1 (2025-02-19)
  --------------------
  - NEW: -Expert parameter (Template JSON, Export JSON, Azure Login, Load SKUs, Diagnostics)
  - NEW: -UserCosts parameter with dedicated Cost per User tab
    * Operating hours/day, days/month, AHB toggle
    * Per named user, concurrent user, daily, hourly, yearly cost breakdown
    * Not-included items list and savings options
  - NEW: Diagnostics button (Expert only)
    * Az module status, Azure connection, region validation, calculator state
  - NEW: Dark theme (Catppuccin Mocha) UI overhaul
    * Colour-coded buttons: Calculate (blue), Pick VM (green), Reset (red), Diagnostics (orange)
    * Card-style info panels with rounded borders
    * DockPanel button bar (Reset left, actions right)
  - FIX: OS Disk minimum raised from 32 GB to 128 GB (256 GB for Power)
  - FIX: OS Disk picker prefers provisioned IOPS over burst (3-tier fallback)
  - FIX: FSLogix share-split considers user count (~1000 users/share)
  - FIX: RAM provisioned no longer inflated by VM lookup RAM
  - FIX: VM Picker strict series enforcement (no silent fallback)
  - FIX: VM Picker selects smallest matching VM (Cores then Mem sort)

  v2.2.0 (2025-02-18)
  --------------------
  - NEW: VM series-specific RAM catalogs (D/E/F/NV/NC/B/M)
  - NEW: GPU priority logic (GPU > Database > Standard)
  - NEW: AVD-compatible VM filtering (excludes HPC, ARM, Confidential)
  - NEW: Professional dark-mode HTML report
  - NEW: Load Balancing selection (Breadth-first / Depth-first)
  - NEW: Autoscale Scaling Plan recommendations
  - FIX: .Count on potentially unwrapped arrays (StrictMode safe)
  - FIX: Get-AzVmSizesInLocation fallback via Get-AzComputeResourceSku

  v2.1.0 (2025-02-17)
  --------------------
  - Initial multi-tab GUI (Workload, Applications, VM Template, Results)
  - Azure VM Picker with region-based sizing
  - Azure Retail Prices API integration
  - FSLogix IOPS model and storage tier recommendations
  - Application overhead model (Client, Database, Development, CAD/GPU)
  - HTML and JSON export

================================================================================
#>

#requires -Version 7.0

param(
  [switch]$Expert
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptVersion  = '2.3.1'
$ScriptBuildUtc = '2025-02-19T18:00:00Z'

#region Ensure STA for WPF
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  $staArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', "`"$PSCommandPath`"")
  if ($Expert) { $staArgs += '-Expert' }
  $proc = Start-Process -FilePath 'pwsh' -ArgumentList $staArgs -PassThru -Wait
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
    Category='Client'; CpuOverheadPerUser=0.3; RamMBPerUser=768; DiskIOPSPerUser=2
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Standard Office suite. Real-world AVD: 600-900 MB per user with typical multi-doc usage.'
  }
  'Microsoft Teams' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=0.5; RamMBPerUser=1024; DiskIOPSPerUser=3
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='New Teams (Teams 2.x) with media optimisation. 800-1200 MB typical in AVD multi-session.'
  }
  'Web Browser (Edge/Chrome)' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=0.4; RamMBPerUser=1024; DiskIOPSPerUser=2
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Modern browsers 800-1500 MB depending on tabs. Limit tabs via GPO. Consider per-site process isolation policy.'
  }
  'Microsoft Outlook' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=0.2; RamMBPerUser=512; DiskIOPSPerUser=3
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Cached mode: 400-600 MB, increases FSLogix profile size. Online mode: ~300 MB but slower UX.'
  }
  'PDF Editor (Acrobat/Foxit)' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=0.1; RamMBPerUser=256; DiskIOPSPerUser=1
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Minimal overhead for typical usage. Large PDFs can spike to 500+ MB.'
  }
  'ERP Client (SAP GUI / Dynamics)' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=0.3; RamMBPerUser=512; DiskIOPSPerUser=2
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Thin-client ERP access. SAP GUI: 300-600 MB. Backend processing on separate app servers.'
  }
  'Power BI Desktop' = [pscustomobject]@{
    Category='Client'; CpuOverheadPerUser=1.0; RamMBPerUser=3072; DiskIOPSPerUser=5
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Very memory-intensive. 2-4 GB per user with large datasets. Consider Power BI Service (web) instead.'
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
    Category='Development'; CpuOverheadPerUser=0.5; RamMBPerUser=768; DiskIOPSPerUser=5
    DiskGBPerUser=0; RequiresGPU=$false; RequiresDataDisk=$false
    Notes='Lightweight IDE. Base ~600 MB, extensions increase usage. Limit file watchers for VDI.'
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

#region Localization
$script:Lang = 'en'
$script:Strings = @{
  # --- Tab Headers ---
  'tab.workload'      = @{ en='Workload'; de='Workload'; fr='Charge'; it='Carico' }
  'tab.applications'  = @{ en='Applications'; de='Anwendungen'; fr='Applications'; it='Applicazioni' }
  'tab.vmtemplate'    = @{ en='VM Template'; de='VM-Vorlage'; fr='Modèle VM'; it='Modello VM' }
  'tab.results'       = @{ en='Results'; de='Ergebnisse'; fr='Résultats'; it='Risultati' }
  'tab.costs'         = @{ en='Costs'; de='Kosten'; fr='Coûts'; it='Costi' }

  # --- Workload Tab ---
  'wl.hostpooltype'       = @{ en='Host pool type'; de='Hostpool-Typ'; fr='Type de pool'; it='Tipo pool host' }
  'wl.hostpooltype.desc'  = @{ en='Pooled: shared VMs (cost-efficient). Personal: 1:1 mapping (isolated).'; de='Pooled: geteilte VMs (kosteneffizient). Personal: 1:1 Zuordnung (isoliert).'; fr='Poolé: VM partagées (économique). Personnel: mappage 1:1 (isolé).'; it='Pooled: VM condivise (economico). Personale: mappatura 1:1 (isolato).' }
  'wl.pooled'             = @{ en='Pooled (multi-session)'; de='Pooled (Multi-Session)'; fr='Poolé (multi-session)'; it='Pooled (multi-sessione)' }
  'wl.personal'           = @{ en='Personal (single-session)'; de='Personal (Single-Session)'; fr='Personnel (session unique)'; it='Personale (sessione singola)' }
  'wl.workloadclass'      = @{ en='Workload class'; de='Workload-Klasse'; fr='Classe de charge'; it='Classe di carico' }
  'wl.workloadclass.desc' = @{ en='Users/vCPU density and RAM/user based on MS sizing guidelines.'; de='Benutzer/vCPU-Dichte und RAM/Benutzer gemäss MS-Richtlinien.'; fr='Densité utilisateurs/vCPU et RAM/utilisateur selon les directives MS.'; it='Densità utenti/vCPU e RAM/utente secondo le linee guida MS.' }
  'wl.workloadhint'       = @{ en='Light (6 users/vCPU, 2 GB RAM): Basic office apps, web browsing, data entry. Minimal CPU/RAM per user.
Medium (4 users/vCPU, 4 GB RAM): Office 365, Teams, Outlook, line-of-business apps. Most common enterprise profile.
Heavy (2 users/vCPU, 6 GB RAM): Multi-app workflows, analytics tools, BI reporting, large Excel models.
Power (1 user/vCPU, 8 GB RAM): CAD/3D, software development, GPU workloads, video editing. Often paired with Personal pools.'; de='Light (6 Benutzer/vCPU, 2 GB RAM): Einfache Office-Apps, Webbrowsing, Dateneingabe. Minimaler CPU/RAM-Bedarf pro Benutzer.
Medium (4 Benutzer/vCPU, 4 GB RAM): Office 365, Teams, Outlook, Fachanwendungen. Häufigstes Enterprise-Profil.
Heavy (2 Benutzer/vCPU, 6 GB RAM): Multi-App-Workflows, Analyse-Tools, BI-Reporting, grosse Excel-Modelle.
Power (1 Benutzer/vCPU, 8 GB RAM): CAD/3D, Softwareentwicklung, GPU-Workloads, Videobearbeitung. Oft mit Personal Pools.'; fr='Léger (6 utilisateurs/vCPU, 2 Go RAM): Apps bureautiques simples, navigation web, saisie de données. CPU/RAM minimal par utilisateur.
Moyen (4 utilisateurs/vCPU, 4 Go RAM): Office 365, Teams, Outlook, apps métier. Profil entreprise le plus courant.
Lourd (2 utilisateurs/vCPU, 6 Go RAM): Workflows multi-apps, outils analytiques, rapports BI, gros modèles Excel.
Power (1 utilisateur/vCPU, 8 Go RAM): CAD/3D, développement logiciel, charges GPU, montage vidéo. Souvent avec des pools personnels.'; it='Leggero (6 utenti/vCPU, 2 GB RAM): App da ufficio semplici, navigazione web, inserimento dati. CPU/RAM minimi per utente.
Medio (4 utenti/vCPU, 4 GB RAM): Office 365, Teams, Outlook, app aziendali. Profilo enterprise più comune.
Pesante (2 utenti/vCPU, 6 GB RAM): Workflow multi-app, strumenti analitici, report BI, modelli Excel complessi.
Power (1 utente/vCPU, 8 GB RAM): CAD/3D, sviluppo software, carichi GPU, editing video. Spesso con pool personali.' }
  'wl.totalusers'         = @{ en='Total users (named)'; de='Gesamtbenutzer (benannt)'; fr='Utilisateurs totaux (nommés)'; it='Utenti totali (nominativi)' }
  'wl.totalusers.desc'    = @{ en='Total named users who will use AVD.'; de='Gesamtzahl benannter AVD-Benutzer.'; fr='Nombre total d''utilisateurs AVD nommés.'; it='Numero totale di utenti AVD nominativi.' }
  'wl.concurrency'        = @{ en='Concurrency'; de='Gleichzeitigkeit'; fr='Concurrence'; it='Concorrenza' }
  'wl.concurrency.desc'   = @{ en='Percent or absolute number of concurrent users.'; de='Prozent oder absolute Anzahl gleichzeitiger Benutzer.'; fr='Pourcentage ou nombre absolu d''utilisateurs simultanés.'; it='Percentuale o numero assoluto di utenti simultanei.' }
  'wl.peakfactor'         = @{ en='Peak factor'; de='Spitzenfaktor'; fr='Facteur de pointe'; it='Fattore di picco' }
  'wl.peakfactor.desc'    = @{ en='Spike multiplier. 1.0 = none, 1.2 = 20% extra.'; de='Spitzenmultiplikator. 1.0 = kein, 1.2 = 20% extra.'; fr='Multiplicateur de pointe. 1.0 = aucun, 1.2 = 20% supplémentaire.'; it='Moltiplicatore di picco. 1.0 = nessuno, 1.2 = 20% in più.' }
  'wl.nplus1'             = @{ en='N+1 redundancy'; de='N+1 Redundanz'; fr='Redondance N+1'; it='Ridondanza N+1' }
  'wl.nplus1.desc'        = @{ en='Extra failover hosts. 1 = one standby, 0 = off.'; de='Extra Failover-Hosts. 1 = ein Standby, 0 = aus.'; fr='Hôtes de basculement supplémentaires. 1 = un en attente, 0 = désactivé.'; it='Host di failover aggiuntivi. 1 = uno in standby, 0 = disattivato.' }
  'wl.headroom'           = @{ en='Extra headroom'; de='Zusätzlicher Puffer'; fr='Marge supplémentaire'; it='Margine aggiuntivo' }
  'wl.headroom.desc'      = @{ en='Growth buffer added to peak concurrent users.'; de='Wachstumspuffer auf Spitzen-Benutzer.'; fr='Tampon de croissance ajouté aux utilisateurs de pointe.'; it='Buffer di crescita aggiunto agli utenti di picco.' }
  'wl.loadbalancing'      = @{ en='Load Balancing (pooled only)'; de='Lastverteilung (nur Pooled)'; fr='Équilibrage de charge (poolé uniquement)'; it='Bilanciamento carico (solo pooled)' }
  'wl.loadbalancing.desc' = @{ en='Breadth-first: even spread. Depth-first: fill sequentially.'; de='Breadth-first: gleichmässig verteilen. Depth-first: sequenziell füllen.'; fr='Breadth-first: répartition égale. Depth-first: remplissage séquentiel.'; it='Breadth-first: distribuzione uniforme. Depth-first: riempimento sequenziale.' }
  'wl.breadthfirst'       = @{ en='Breadth-first (best UX)'; de='Breadth-first (beste UX)'; fr='Breadth-first (meilleure UX)'; it='Breadth-first (migliore UX)' }
  'wl.depthfirst'         = @{ en='Depth-first (cost optimised)'; de='Depth-first (kostenoptimiert)'; fr='Depth-first (coût optimisé)'; it='Depth-first (costo ottimizzato)' }
  'wl.maxsession'         = @{ en='Max session limit per host'; de='Max. Sitzungslimit pro Host'; fr='Limite max. sessions par hôte'; it='Limite max. sessioni per host' }
  'wl.maxsession.desc'    = @{ en='0 = auto from users/host. Set explicitly for Depth-first.'; de='0 = automatisch. Für Depth-first explizit setzen.'; fr='0 = auto. Définir explicitement pour Depth-first.'; it='0 = automatico. Impostare esplicitamente per Depth-first.' }

  # --- Tuning Panel ---
  'tune.title'          = @{ en='Tuning Parameters'; de='Tuning-Parameter'; fr='Paramètres de réglage'; it='Parametri di ottimizzazione' }
  'tune.desc'           = @{ en='Advanced settings. Defaults follow Microsoft best practice.'; de='Erweiterte Einstellungen. Standards folgen MS Best Practice.'; fr='Paramètres avancés. Valeurs par défaut selon les meilleures pratiques MS.'; it='Impostazioni avanzate. Valori predefiniti secondo le best practice MS.' }
  'tune.fslogix'        = @{ en='FSLogix profile storage'; de='FSLogix Profilspeicher'; fr='Stockage profil FSLogix'; it='Storage profilo FSLogix' }
  'tune.fslogix.desc'   = @{ en='Profile container size per user. MS default max: 30 GB.'; de='Profilcontainer-Grösse pro Benutzer. MS-Standard: 30 GB.'; fr='Taille du conteneur de profil par utilisateur. Par défaut MS: 30 Go.'; it='Dimensione contenitore profilo per utente. Default MS: 30 GB.' }
  'tune.cpuram'         = @{ en='CPU / RAM / System Reserve'; de='CPU / RAM / Systemreserve'; fr='CPU / RAM / Réserve système'; it='CPU / RAM / Riserva di sistema' }
  'tune.cpuram.desc'    = @{ en='Target utilisation and virtualisation overhead (MS: 15-20%).'; de='Zielauslastung und Virtualisierungs-Overhead (MS: 15-20%).'; fr='Utilisation cible et surcharge de virtualisation (MS: 15-20%).'; it='Utilizzo target e overhead di virtualizzazione (MS: 15-20%).' }
  'tune.vcpurange'      = @{ en='vCPU range per host (pooled)'; de='vCPU-Bereich pro Host (Pooled)'; fr='Plage vCPU par hôte (poolé)'; it='Range vCPU per host (pooled)' }
  'tune.vcpurange.desc' = @{ en='MS recommends max 24 for multi-session. 128 = unrestricted.'; de='MS empfiehlt max 24 für Multi-Session. 128 = unbegrenzt.'; fr='MS recommande max 24 pour multi-session. 128 = illimité.'; it='MS raccomanda max 24 per multi-sessione. 128 = illimitato.' }

  # --- Applications Tab ---
  'app.client'      = @{ en='Client Applications'; de='Client-Anwendungen'; fr='Applications clientes'; it='Applicazioni client' }
  'app.client.desc' = @{ en='Each adds CPU + RAM overhead per concurrent user on the session host.'; de='Jede fügt CPU + RAM-Overhead pro Benutzer hinzu.'; fr='Chacune ajoute une surcharge CPU + RAM par utilisateur.'; it='Ognuna aggiunge overhead CPU + RAM per utente.' }
  'app.dev'         = @{ en='Development Tools'; de='Entwicklungstools'; fr='Outils de développement'; it='Strumenti di sviluppo' }
  'app.db'          = @{ en='Database Engines'; de='Datenbank-Engines'; fr='Moteurs de base de données'; it='Motori database' }
  'app.db.desc'     = @{ en='Switches to E-series (8 GB/vCPU). Consider Azure SQL for production.'; de='Wechsel zu E-Serie (8 GB/vCPU). Azure SQL für Produktion empfohlen.'; fr='Passe à la série E (8 Go/vCPU). Considérez Azure SQL pour la production.'; it='Passa alla serie E (8 GB/vCPU). Considerare Azure SQL per la produzione.' }
  'app.gpu'         = @{ en='CAD / GPU Applications'; de='CAD / GPU-Anwendungen'; fr='Applications CAD / GPU'; it='Applicazioni CAD / GPU' }
  'app.gpu.desc'    = @{ en='Requires NV-series VMs (NVIDIA GPU). Significantly higher cost.'; de='Erfordert NV-Serie (NVIDIA GPU). Deutlich höhere Kosten.'; fr='Nécessite des VM NV-series (GPU NVIDIA). Coût nettement plus élevé.'; it='Richiede VM serie NV (GPU NVIDIA). Costo significativamente più alto.' }
  'app.dbvolume'    = @{ en='Database data volume per host'; de='Datenbank-Datenvolumen pro Host'; fr='Volume de données BD par hôte'; it='Volume dati database per host' }
  'app.dbvolume.desc' = @{ en='Dedicated data disk size when a database engine is selected.'; de='Datendisk-Grösse bei ausgewählter Datenbank-Engine.'; fr='Taille du disque de données quand un moteur BD est sélectionné.'; it='Dimensione disco dati quando è selezionato un motore database.' }
  'app.howaffects'  = @{ en='How Apps Affect Sizing'; de='Einfluss auf Dimensionierung'; fr='Impact sur le dimensionnement'; it='Impatto sul dimensionamento' }

  # --- Info Panel: How Apps Affect Sizing ---
  'info.vmauto'       = @{ en='VM Series Auto-Selection'; de='Automatische VM-Serienauswahl'; fr='Sélection auto. de série VM'; it='Selezione automatica serie VM' }
  'info.vmauto.desc'  = @{ en='No special apps: D-series. Database: E-series. GPU/CAD: NV-series. GPU always takes priority.'; de='Keine speziellen Apps: D-Serie. Datenbank: E-Serie. GPU/CAD: NV-Serie. GPU hat immer Priorität.'; fr='Pas d''apps spéciales: série D. Base de données: série E. GPU/CAD: série NV. Le GPU a toujours la priorité.'; it='Nessuna app speciale: serie D. Database: serie E. GPU/CAD: serie NV. La GPU ha sempre la priorità.' }
  'info.dbengines'       = @{ en='Database Engines'; de='Datenbank-Engines'; fr='Moteurs de base de données'; it='Motori database' }
  'info.dbengines.desc'  = @{ en='E-series, Premium SSD v2 data disk. Personal pool strongly recommended.'; de='E-Serie, Premium SSD v2 Datendisk. Personal Pool dringend empfohlen.'; fr='Série E, disque de données Premium SSD v2. Pool personnel fortement recommandé.'; it='Serie E, disco dati Premium SSD v2. Pool personale fortemente raccomandato.' }
  'info.cadgpu'       = @{ en='CAD / GPU'; de='CAD / GPU'; fr='CAD / GPU'; it='CAD / GPU' }
  'info.cadgpu.desc'  = @{ en='NV-series (NVIDIA A10). 2-4 users/host typical. Personal pool recommended for heavy 3D.'; de='NV-Serie (NVIDIA A10). 2-4 Benutzer/Host typisch. Personal Pool empfohlen für intensive 3D-Nutzung.'; fr='Série NV (NVIDIA A10). 2-4 utilisateurs/hôte typique. Pool personnel recommandé pour la 3D intensive.'; it='Serie NV (NVIDIA A10). 2-4 utenti/host tipico. Pool personale raccomandato per 3D intensivo.' }

  # --- Info Panel: VM Selection Logic ---
  'info.seriesauto'       = @{ en='Series Auto-Selection'; de='Automatische Serienauswahl'; fr='Sélection automatique de série'; it='Selezione automatica serie' }
  'info.seriesauto.desc'  = @{ en='Standard: D-series. Database: E-series. GPU/CAD: NV-series. GPU takes priority.'; de='Standard: D-Serie. Datenbank: E-Serie. GPU/CAD: NV-Serie. GPU hat Priorität.'; fr='Standard: série D. Base de données: série E. GPU/CAD: série NV. Le GPU a la priorité.'; it='Standard: serie D. Database: serie E. GPU/CAD: serie NV. La GPU ha priorità.' }
  'info.strictseries'       = @{ en='Strict Series Mode'; de='Strenge Serienauswahl'; fr='Mode série stricte'; it='Modalità serie rigorosa' }
  'info.strictseries.desc'  = @{ en='Specific series selected: only VMs of that series. No silent fallback.'; de='Spezifische Serie gewählt: nur VMs dieser Serie. Kein stiller Fallback.'; fr='Série spécifique sélectionnée: uniquement les VM de cette série. Pas de fallback silencieux.'; it='Serie specifica selezionata: solo VM di quella serie. Nessun fallback silenzioso.' }
  'info.smallestvm'       = @{ en='Smallest Matching VM'; de='Kleinste passende VM'; fr='Plus petite VM correspondante'; it='VM corrispondente più piccola' }
  'info.smallestvm.desc'  = @{ en='Picks smallest VM meeting calculated vCPU + RAM. Prefers v5/v6 with Premium Storage.'; de='Wählt kleinste VM die berechnete vCPU + RAM erfüllt. Bevorzugt v5/v6 mit Premium Storage.'; fr='Sélectionne la plus petite VM répondant aux vCPU + RAM calculés. Préfère v5/v6 avec Premium Storage.'; it='Seleziona la VM più piccola che soddisfa vCPU + RAM calcolati. Preferisce v5/v6 con Premium Storage.' }
  'info.avdcompat'       = @{ en='AVD-Compatible Only'; de='Nur AVD-kompatibel'; fr='Compatible AVD uniquement'; it='Solo compatibile AVD' }
  'info.avdcompat.desc'  = @{ en='D, E, F, NV, NC, B, M series. Excluded: A, H, ND, L, DC/EC, ARM.'; de='D, E, F, NV, NC, B, M Serien. Ausgeschlossen: A, H, ND, L, DC/EC, ARM.'; fr='Séries D, E, F, NV, NC, B, M. Exclues: A, H, ND, L, DC/EC, ARM.'; it='Serie D, E, F, NV, NC, B, M. Escluse: A, H, ND, L, DC/EC, ARM.' }
  'info.ramperseries'    = @{ en='RAM per VM Series'; de='RAM pro VM-Serie'; fr='RAM par série VM'; it='RAM per serie VM' }

  # --- VM Template Tab ---
  'vm.region'       = @{ en='Azure region'; de='Azure-Region'; fr='Région Azure'; it='Regione Azure' }
  'vm.region.desc'  = @{ en='Region where session hosts will be deployed. Affects VM availability and pricing.'; de='Region für die Bereitstellung. Beeinflusst VM-Verfügbarkeit und Preise.'; fr='Région de déploiement. Affecte la disponibilité et les prix.'; it='Regione di distribuzione. Influisce su disponibilità e prezzi.' }
  'vm.series'       = @{ en='VM series preference'; de='VM-Serien-Präferenz'; fr='Préférence de série VM'; it='Preferenza serie VM' }
  'vm.series.desc'  = @{ en='Specific series is strictly enforced — no silent fallback to other series.'; de='Gewählte Serie wird strikt erzwungen — kein Fallback.'; fr='La série choisie est strictement appliquée — pas de fallback.'; it='La serie scelta è applicata rigorosamente — nessun fallback.' }
  'vm.image'        = @{ en='Marketplace image'; de='Marketplace-Image'; fr='Image Marketplace'; it='Immagine Marketplace' }
  'vm.image.desc'   = @{ en='Windows image for session hosts. Azure Login + Load SKUs available in Expert mode.'; de='Windows-Image für Session Hosts. Azure Login + SKUs laden im Expertenmodus.'; fr='Image Windows pour les hôtes de session. Azure Login + SKUs en mode Expert.'; it='Immagine Windows per gli host di sessione. Azure Login + SKU in modalità Expert.' }
  'vm.logic'        = @{ en='VM Selection Logic'; de='VM-Auswahllogik'; fr='Logique de sélection VM'; it='Logica selezione VM' }

  # --- Results Tab ---
  'res.title'       = @{ en='Sizing Results'; de='Sizing-Ergebnisse'; fr='Résultats du dimensionnement'; it='Risultati dimensionamento' }
  'res.desc'        = @{ en='Click ''Calculate'' then ''Pick VM from Azure'' to find the best VM template with pricing.'; de='Klicken Sie ''Berechnen'' und dann ''VM von Azure wählen''.'; fr='Cliquez ''Calculer'' puis ''Choisir VM depuis Azure''.'; it='Clicca ''Calcola'' poi ''Scegli VM da Azure''.' }
  'res.notes'       = @{ en='Notes, warnings and recommendations:'; de='Hinweise, Warnungen und Empfehlungen:'; fr='Notes, avertissements et recommandations:'; it='Note, avvisi e raccomandazioni:' }

  # --- Costs Tab ---
  'cost.title'      = @{ en='Cost Analysis'; de='Kostenanalyse'; fr='Analyse des coûts'; it='Analisi dei costi' }
  'cost.desc'       = @{ en='Configure pricing, discounts and calculate per-user costs. Requires ''Pick VM from Azure'' on the Results tab first.'; de='Preise, Rabatte konfigurieren und Pro-Benutzer-Kosten berechnen. Zuerst ''VM von Azure wählen'' auf dem Ergebnis-Tab.'; fr='Configurer les prix, remises et calculer les coûts par utilisateur. ''Choisir VM depuis Azure'' d''abord.'; it='Configurare prezzi, sconti e calcolare i costi per utente. Prima ''Scegli VM da Azure''.' }
  'cost.pricing'    = @{ en='Pricing'; de='Preise'; fr='Tarification'; it='Prezzi' }
  'cost.currency'   = @{ en='Currency'; de='Währung'; fr='Devise'; it='Valuta' }
  'cost.currency.desc' = @{ en='Currency for Azure Retail Prices API.'; de='Währung für Azure Retail Prices API.'; fr='Devise pour l''API Azure Retail Prices.'; it='Valuta per l''API Azure Retail Prices.' }
  'cost.ahb'        = @{ en='Azure Hybrid Benefit'; de='Azure Hybrid Benefit'; fr='Azure Hybrid Benefit'; it='Azure Hybrid Benefit' }
  'cost.ahb.desc'   = @{ en='M365 E3/E5 or Win E3/E5 licenses cover Windows cost.'; de='M365 E3/E5 oder Win E3/E5 Lizenzen decken Windows-Kosten.'; fr='Les licences M365 E3/E5 ou Win E3/E5 couvrent le coût Windows.'; it='Le licenze M365 E3/E5 o Win E3/E5 coprono il costo Windows.' }
  'cost.ahb.check'  = @{ en='Azure Hybrid Benefit active'; de='Azure Hybrid Benefit aktiv'; fr='Azure Hybrid Benefit actif'; it='Azure Hybrid Benefit attivo' }
  'cost.schedule'   = @{ en='Operating Schedule'; de='Betriebszeiten'; fr='Horaires d''exploitation'; it='Orario operativo' }
  'cost.hours'      = @{ en='Operating hours per day'; de='Betriebsstunden pro Tag'; fr='Heures d''exploitation par jour'; it='Ore operative al giorno' }
  'cost.hours.desc' = @{ en='24 = always on, 10 = business hours only.'; de='24 = immer an, 10 = nur Geschäftszeiten.'; fr='24 = toujours actif, 10 = heures ouvrables uniquement.'; it='24 = sempre attivo, 10 = solo ore lavorative.' }
  'cost.days'       = @{ en='Operating days per month'; de='Betriebstage pro Monat'; fr='Jours d''exploitation par mois'; it='Giorni operativi al mese' }
  'cost.days.desc'  = @{ en='22 = work month, 30 = daily use.'; de='22 = Arbeitsmonat, 30 = täglich.'; fr='22 = mois ouvrable, 30 = utilisation quotidienne.'; it='22 = mese lavorativo, 30 = uso quotidiano.' }
  'cost.discounts'  = @{ en='Discounts'; de='Rabatte'; fr='Remises'; it='Sconti' }
  'cost.disc.desc'  = @{ en='Applied cumulatively to the Azure retail list price.'; de='Kumulativ auf den Azure-Listenpreis angewendet.'; fr='Appliquées cumulativement au prix catalogue Azure.'; it='Applicate cumulativamente al prezzo di listino Azure.' }
  'cost.csp'        = @{ en='CSP / EA / Partner discount'; de='CSP / EA / Partner-Rabatt'; fr='Remise CSP / EA / Partenaire'; it='Sconto CSP / EA / Partner' }
  'cost.csp.desc'   = @{ en='Negotiated discount (CSP margin, EA, MPA).'; de='Verhandelter Rabatt (CSP-Marge, EA, MPA).'; fr='Remise négociée (marge CSP, EA, MPA).'; it='Sconto negoziato (margine CSP, EA, MPA).' }
  'cost.ri'         = @{ en='Reserved Instance discount'; de='Reserved Instance-Rabatt'; fr='Remise instance réservée'; it='Sconto istanza riservata' }
  'cost.ri.desc'    = @{ en='1yr ~35%, 3yr ~55%.'; de='1 Jahr ~35%, 3 Jahre ~55%.'; fr='1 an ~35%, 3 ans ~55%.'; it='1 anno ~35%, 3 anni ~55%.' }
  'cost.additional'      = @{ en='Additional discount'; de='Zusätzlicher Rabatt'; fr='Remise supplémentaire'; it='Sconto aggiuntivo' }
  'cost.additional.desc' = @{ en='Savings Plans, promos, custom agreements.'; de='Savings Plans, Aktionen, Sondervereinbarungen.'; fr='Plans d''épargne, promos, accords personnalisés.'; it='Piani di risparmio, promozioni, accordi personalizzati.' }
  'cost.breakdown'  = @{ en='Cost Breakdown'; de='Kostenaufschlüsselung'; fr='Ventilation des coûts'; it='Ripartizione costi' }

  # --- Buttons ---
  'btn.calculate'   = @{ en='Calculate'; de='Berechnen'; fr='Calculer'; it='Calcola' }
  'btn.pickvm'      = @{ en='Pick VM from Azure'; de='VM von Azure wählen'; fr='Choisir VM depuis Azure'; it='Scegli VM da Azure' }
  'btn.exportjson'  = @{ en='Export JSON'; de='JSON exportieren'; fr='Exporter JSON'; it='Esporta JSON' }
  'btn.exportreport'= @{ en='Export HTML Report'; de='HTML-Bericht exportieren'; fr='Exporter rapport HTML'; it='Esporta report HTML' }
  'btn.close'       = @{ en='Close'; de='Schliessen'; fr='Fermer'; it='Chiudi' }
  'btn.reset'       = @{ en='Reset'; de='Zurücksetzen'; fr='Réinitialiser'; it='Reimposta' }
  'btn.diagnostics' = @{ en='Diagnostics'; de='Diagnose'; fr='Diagnostics'; it='Diagnostica' }
  'btn.calccosts'   = @{ en='Calculate Costs'; de='Kosten berechnen'; fr='Calculer les coûts'; it='Calcola costi' }
  'btn.azlogin'     = @{ en='Azure Login'; de='Azure Login'; fr='Connexion Azure'; it='Accesso Azure' }
  'btn.loadskus'    = @{ en='Load SKUs'; de='SKUs laden'; fr='Charger SKUs'; it='Carica SKU' }

  # --- Report ---
  'rpt.title'       = @{ en='Sizing Report'; de='Sizing-Bericht'; fr='Rapport de dimensionnement'; it='Report dimensionamento' }
  'rpt.execsummary' = @{ en='Executive Summary'; de='Zusammenfassung'; fr='Résumé exécutif'; it='Riepilogo' }
  'rpt.workload'    = @{ en='Workload Profile'; de='Workload-Profil'; fr='Profil de charge'; it='Profilo di carico' }
  'rpt.apps'        = @{ en='Applications'; de='Anwendungen'; fr='Applications'; it='Applicazioni' }
  'rpt.loadbal'     = @{ en='Load Balancing'; de='Lastverteilung'; fr='Équilibrage de charge'; it='Bilanciamento carico' }
  'rpt.storage'     = @{ en='Storage'; de='Speicher'; fr='Stockage'; it='Storage' }
  'rpt.vmselection' = @{ en='Azure VM Selection'; de='Azure VM-Auswahl'; fr='Sélection VM Azure'; it='Selezione VM Azure' }
  'rpt.usercosts'   = @{ en='User Cost Analysis'; de='Benutzer-Kostenanalyse'; fr='Analyse des coûts utilisateur'; it='Analisi costi utente' }
  'rpt.notes'       = @{ en='Notes &amp; Recommendations'; de='Hinweise &amp; Empfehlungen'; fr='Notes &amp; Recommandations'; it='Note &amp; Raccomandazioni' }
  'rpt.warnings'    = @{ en='Warnings'; de='Warnungen'; fr='Avertissements'; it='Avvisi' }
  'rpt.generated'   = @{ en='Generated by'; de='Erstellt mit'; fr='Généré par'; it='Generato da' }
  'rpt.validate'    = @{ en='All recommendations should be validated with'; de='Alle Empfehlungen sollten validiert werden mit'; fr='Toutes les recommandations doivent être validées avec'; it='Tutte le raccomandazioni devono essere validate con' }
  'rpt.monitoring'  = @{ en='monitoring and pilot deployments'; de='Monitoring und Pilotbereitstellungen'; fr='surveillance et déploiements pilotes'; it='monitoraggio e distribuzioni pilota' }
}

function Get-Str([string]$Key) {
  $entry = $script:Strings[$Key]
  if ($entry) { $val = $entry[$script:Lang]; if ($val) { return $val } else { return $entry['en'] } }
  return $Key
}

function Apply-Language {
  $w = $script:Window
  if (-not $w) { return }

  # Tab headers
  $tabs = $w.FindName('Tabs')
  if ($tabs) {
    $tabs.Items[0].Header = Get-Str 'tab.workload'
    $tabs.Items[1].Header = Get-Str 'tab.applications'
    $tabs.Items[2].Header = Get-Str 'tab.vmtemplate'
    $tabs.Items[3].Header = Get-Str 'tab.results'
    if ($tabs.Items.Count -gt 4) { $tabs.Items[4].Header = Get-Str 'tab.costs' }
  }

  # Buttons
  $b = $w.FindName('BtnCalculate');    if ($b) { $b.Content = Get-Str 'btn.calculate' }
  $b = $w.FindName('BtnPickVm');       if ($b) { $b.Content = Get-Str 'btn.pickvm' }
  $b = $w.FindName('BtnExportJson');   if ($b) { $b.Content = Get-Str 'btn.exportjson' }
  $b = $w.FindName('BtnExportReport'); if ($b) { $b.Content = Get-Str 'btn.exportreport' }
  $b = $w.FindName('BtnClose');        if ($b) { $b.Content = Get-Str 'btn.close' }
  $b = $w.FindName('BtnReset');        if ($b) { $b.Content = Get-Str 'btn.reset' }
  $b = $w.FindName('BtnDiagnostics');  if ($b) { $b.Content = Get-Str 'btn.diagnostics' }
  $b = $w.FindName('BtnCalcUserCosts');if ($b) { $b.Content = Get-Str 'btn.calccosts' }
  $b = $w.FindName('BtnAzLogin');      if ($b) { $b.Content = Get-Str 'btn.azlogin' }
  $b = $w.FindName('BtnDiscoverSkus'); if ($b) { $b.Content = Get-Str 'btn.loadskus' }

  # Helper: safe set text
  function Set-LblText([string]$Name, [string]$Key) {
    $el = $w.FindName($Name)
    if ($el) { $el.Text = Get-Str $Key }
  }

  # Workload tab
  Set-LblText 'LblHostPoolType'     'wl.hostpooltype'
  Set-LblText 'LblHostPoolTypeDesc' 'wl.hostpooltype.desc'
  Set-LblText 'LblWorkloadClass'    'wl.workloadclass'
  Set-LblText 'LblWorkloadClassDesc' 'wl.workloadclass.desc'
  Set-LblText 'LblWorkloadHint'     'wl.workloadhint'
  Set-LblText 'LblTotalUsers'       'wl.totalusers'
  Set-LblText 'LblTotalUsersDesc'   'wl.totalusers.desc'
  Set-LblText 'LblConcurrency'      'wl.concurrency'
  Set-LblText 'LblConcurrencyDesc'  'wl.concurrency.desc'
  Set-LblText 'LblPeakFactor'       'wl.peakfactor'
  Set-LblText 'LblPeakFactorDesc'   'wl.peakfactor.desc'
  Set-LblText 'LblNPlus1'           'wl.nplus1'
  Set-LblText 'LblNPlus1Desc'       'wl.nplus1.desc'
  Set-LblText 'LblHeadroom'         'wl.headroom'
  Set-LblText 'LblHeadroomDesc'     'wl.headroom.desc'
  Set-LblText 'LblLoadBal'          'wl.loadbalancing'
  Set-LblText 'LblLoadBalDesc'      'wl.loadbalancing.desc'
  Set-LblText 'LblMaxSession'       'wl.maxsession'
  Set-LblText 'LblMaxSessionDesc'   'wl.maxsession.desc'

  # Tuning panel
  Set-LblText 'LblTuneTitle'      'tune.title'
  Set-LblText 'LblTuneDesc'       'tune.desc'
  Set-LblText 'LblFsLogix'        'tune.fslogix'
  Set-LblText 'LblFsLogixDesc'    'tune.fslogix.desc'
  Set-LblText 'LblCpuRam'         'tune.cpuram'
  Set-LblText 'LblCpuRamDesc'     'tune.cpuram.desc'
  Set-LblText 'LblVcpuRange'      'tune.vcpurange'
  Set-LblText 'LblVcpuRangeDesc'  'tune.vcpurange.desc'

  # Applications tab
  Set-LblText 'LblClientApps'     'app.client'
  Set-LblText 'LblClientAppsDesc' 'app.client.desc'
  Set-LblText 'LblDevTools'       'app.dev'
  Set-LblText 'LblDbEngines'      'app.db'
  Set-LblText 'LblDbEnginesDesc'  'app.db.desc'
  Set-LblText 'LblGpuApps'        'app.gpu'
  Set-LblText 'LblGpuAppsDesc'    'app.gpu.desc'
  Set-LblText 'LblDbVolume'       'app.dbvolume'
  Set-LblText 'LblDbVolumeDesc'   'app.dbvolume.desc'
  Set-LblText 'LblHowAffects'     'app.howaffects'

  # Info panel: How Apps Affect Sizing
  Set-LblText 'LblInfoVmAuto'          'info.vmauto'
  Set-LblText 'LblInfoVmAutoDesc'      'info.vmauto.desc'
  Set-LblText 'LblInfoDbEngines'       'info.dbengines'
  Set-LblText 'LblInfoDbEnginesDesc'   'info.dbengines.desc'
  Set-LblText 'LblInfoCadGpu'          'info.cadgpu'
  Set-LblText 'LblInfoCadGpuDesc'      'info.cadgpu.desc'

  # VM Template tab
  Set-LblText 'LblRegion'         'vm.region'
  Set-LblText 'LblRegionDesc'     'vm.region.desc'
  Set-LblText 'LblVmSeries'       'vm.series'
  Set-LblText 'LblVmSeriesDesc'   'vm.series.desc'
  Set-LblText 'LblImage'          'vm.image'
  Set-LblText 'LblImageDesc'      'vm.image.desc'
  Set-LblText 'LblVmLogic'        'vm.logic'

  # Info panel: VM Selection Logic
  Set-LblText 'LblInfoSeriesAuto'       'info.seriesauto'
  Set-LblText 'LblInfoSeriesAutoDesc'   'info.seriesauto.desc'
  Set-LblText 'LblInfoStrictSeries'     'info.strictseries'
  Set-LblText 'LblInfoStrictSeriesDesc' 'info.strictseries.desc'
  Set-LblText 'LblInfoSmallestVm'       'info.smallestvm'
  Set-LblText 'LblInfoSmallestVmDesc'   'info.smallestvm.desc'
  Set-LblText 'LblInfoAvdCompat'        'info.avdcompat'
  Set-LblText 'LblInfoAvdCompatDesc'    'info.avdcompat.desc'
  Set-LblText 'LblInfoRamPerSeries'     'info.ramperseries'

  # Results tab
  Set-LblText 'LblResTitle'       'res.title'
  Set-LblText 'LblResDesc'        'res.desc'
  Set-LblText 'LblResNotes'       'res.notes'

  # Costs tab
  Set-LblText 'LblCostTitle'      'cost.title'
  Set-LblText 'LblCostDesc'       'cost.desc'
  Set-LblText 'LblPricing'        'cost.pricing'
  Set-LblText 'LblCurrency'       'cost.currency'
  Set-LblText 'LblCurrencyDesc'   'cost.currency.desc'
  Set-LblText 'LblAhb'            'cost.ahb'
  Set-LblText 'LblAhbDesc'        'cost.ahb.desc'
  Set-LblText 'LblSchedule'       'cost.schedule'
  Set-LblText 'LblHours'          'cost.hours'
  Set-LblText 'LblHoursDesc'      'cost.hours.desc'
  Set-LblText 'LblDays'           'cost.days'
  Set-LblText 'LblDaysDesc'       'cost.days.desc'
  Set-LblText 'LblDiscounts'      'cost.discounts'
  Set-LblText 'LblDiscDesc'       'cost.disc.desc'
  Set-LblText 'LblCsp'            'cost.csp'
  Set-LblText 'LblCspDesc'        'cost.csp.desc'
  Set-LblText 'LblRi'             'cost.ri'
  Set-LblText 'LblRiDesc'         'cost.ri.desc'
  Set-LblText 'LblAdditional'     'cost.additional'
  Set-LblText 'LblAdditionalDesc' 'cost.additional.desc'
  Set-LblText 'LblBreakdown'      'cost.breakdown'

  # CheckBox content
  $chk = $w.FindName('ChkAhb'); if ($chk) { $chk.Content = Get-Str 'cost.ahb.check' }
}
#endregion
#region XAML
$XamlString = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AVD Sizing Calculator" Height="1020" Width="1240"
        WindowStartupLocation="CenterScreen"
        Background="#1e1e2e" Foreground="#cdd6f4">
  <Window.Resources>
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="#cdd6f4"/>
    </Style>

    <!-- TextBox: dark bg, light text, visible selection + focus -->
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#313244"/>
      <Setter Property="Foreground" Value="#cdd6f4"/>
      <Setter Property="BorderBrush" Value="#585b70"/>
      <Setter Property="SelectionBrush" Value="#585b70"/>
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="CaretBrush" Value="#cdd6f4"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="BorderBrush" Value="#89b4fa"/>
        </Trigger>
        <Trigger Property="IsFocused" Value="True">
          <Setter Property="BorderBrush" Value="#89b4fa"/>
          <Setter Property="Background" Value="#3b3d52"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- ComboBox: dark text on system dropdown, blue hover border -->
    <Style TargetType="ComboBox">
      <Setter Property="Background" Value="#313244"/>
      <Setter Property="Foreground" Value="#1e1e2e"/>
      <Setter Property="BorderBrush" Value="#585b70"/>
      <Setter Property="Padding" Value="6,4"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="BorderBrush" Value="#89b4fa"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="ComboBoxItem">
      <Setter Property="Background" Value="White"/>
      <Setter Property="Foreground" Value="#1e1e2e"/>
      <Style.Triggers>
        <Trigger Property="IsHighlighted" Value="True">
          <Setter Property="Background" Value="#89b4fa"/>
          <Setter Property="Foreground" Value="#1e1e2e"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#b4befe"/>
          <Setter Property="Foreground" Value="#1e1e2e"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- CheckBox: light text, blue highlight on hover -->
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#cdd6f4"/>
      <Setter Property="Margin" Value="0,4,0,0"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Foreground" Value="#89b4fa"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- Button: hover brightens, pressed darkens -->
    <Style TargetType="Button">
      <Setter Property="Background" Value="#45475a"/>
      <Setter Property="Foreground" Value="#cdd6f4"/>
      <Setter Property="BorderBrush" Value="#585b70"/>
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#585b70"/>
          <Setter Property="BorderBrush" Value="#89b4fa"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
          <Setter Property="Background" Value="#313244"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- TabItem -->
    <!-- TabItem: full template override for dark theme headers -->
    <Style TargetType="TabItem">
      <Setter Property="Foreground" Value="#a6adc8"/>
      <Setter Property="Background" Value="#1e1e2e"/>
      <Setter Property="Padding" Value="14,6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="TabBorder" Background="{TemplateBinding Background}"
                    BorderBrush="#585b70" BorderThickness="1,1,1,0"
                    CornerRadius="6,6,0,0" Padding="{TemplateBinding Padding}" Margin="2,0,2,0">
              <ContentPresenter x:Name="TabText" ContentSource="Header"
                                HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="TabBorder" Property="Background" Value="#313244"/>
                <Setter Property="Foreground" Value="#89b4fa"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="TabBorder" Property="Background" Value="#313244"/>
                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#89b4fa"/>
                <Setter TargetName="TabBorder" Property="BorderThickness" Value="1,2,1,0"/>
                <Setter Property="Foreground" Value="#cdd6f4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- DataGrid: dark rows, readable headers + selection -->
    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="#313244"/>
      <Setter Property="Foreground" Value="#cdd6f4"/>
      <Setter Property="BorderBrush" Value="#585b70"/>
      <Setter Property="RowBackground" Value="#313244"/>
      <Setter Property="AlternatingRowBackground" Value="#3b3d52"/>
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#45475a"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#45475a"/>
      <Setter Property="Foreground" Value="#cdd6f4"/>
      <Setter Property="Padding" Value="8,4"/>
      <Setter Property="BorderBrush" Value="#585b70"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
    </Style>
    <Style TargetType="DataGridRow">
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#585b70"/>
          <Setter Property="Foreground" Value="#cdd6f4"/>
        </Trigger>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#45475a"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="BorderBrush" Value="Transparent"/>
      <Setter Property="Foreground" Value="#cdd6f4"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#585b70"/>
          <Setter Property="Foreground" Value="#cdd6f4"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="Separator">
      <Setter Property="Background" Value="#45475a"/>
      <Setter Property="Margin" Value="0,8,0,8"/>
    </Style>
  </Window.Resources>

  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <DockPanel Grid.Row="0" Margin="0,0,0,4" LastChildFill="False">
      <StackPanel DockPanel.Dock="Left" Orientation="Horizontal">
        <TextBlock FontSize="20" FontWeight="Bold" Foreground="#89b4fa" Text="AVD Sizing Calculator"/>
        <TextBlock FontSize="12" VerticalAlignment="Bottom" Margin="10,0,0,2" Foreground="#6c7086" Text="v2.3.1"/>
      </StackPanel>
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
        <TextBlock Foreground="#6c7086" FontSize="11" VerticalAlignment="Center" Margin="0,0,6,0" Text="Language:"/>
        <ComboBox x:Name="CmbLanguage" Width="110" SelectedIndex="0">
          <ComboBoxItem Content="English"/><ComboBoxItem Content="Deutsch"/><ComboBoxItem Content="Français"/><ComboBoxItem Content="Italiano"/>
        </ComboBox>
      </StackPanel>
    </DockPanel>

    <TabControl Grid.Row="1" Margin="0,8,0,8" x:Name="Tabs" Background="#1e1e2e" BorderBrush="#585b70">

      <!-- TAB 1: Workload -->
      <TabItem Header="Workload">
        <Grid Margin="10">
          <Grid.ColumnDefinitions><ColumnDefinition Width="380"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>

          <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Padding="0,0,12,0">
          <StackPanel>
            <TextBlock x:Name="LblHostPoolType" FontWeight="Bold" Text="Host pool type"/>
            <TextBlock x:Name="LblHostPoolTypeDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Text="Pooled: shared VMs (cost-efficient). Personal: 1:1 mapping (isolated)."/>
            <ComboBox x:Name="CmbHostPoolType" Margin="0,3,0,8" SelectedIndex="0">
              <ComboBoxItem Content="Pooled (multi-session)"/><ComboBoxItem Content="Personal (single-session)"/>
            </ComboBox>

            <TextBlock x:Name="LblWorkloadClass" FontWeight="Bold" Text="Workload class"/>
            <TextBlock x:Name="LblWorkloadClassDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Text="Users/vCPU density and RAM/user based on MS sizing guidelines."/>
            <ComboBox x:Name="CmbWorkload" Margin="0,3,0,2" SelectedIndex="1">
              <ComboBoxItem Content="Light &#x2014; 6 users/vCPU, 2 GB RAM/user"/>
              <ComboBoxItem Content="Medium &#x2014; 4 users/vCPU, 4 GB RAM/user"/>
              <ComboBoxItem Content="Heavy &#x2014; 2 users/vCPU, 6 GB RAM/user"/>
              <ComboBoxItem Content="Power &#x2014; 1 user/vCPU, 8 GB RAM/user"/>
            </ComboBox>
            <TextBlock x:Name="LblWorkloadHint" Foreground="#6c7086" FontSize="10" TextWrapping="Wrap" Margin="0,0,0,8" Text="Light (6 users/vCPU, 2 GB RAM): Basic office apps, web browsing, data entry. Minimal CPU/RAM per user.&#x0a;Medium (4 users/vCPU, 4 GB RAM): Office 365, Teams, Outlook, line-of-business apps. Most common enterprise profile.&#x0a;Heavy (2 users/vCPU, 6 GB RAM): Multi-app workflows, analytics tools, BI reporting, large Excel models.&#x0a;Power (1 user/vCPU, 8 GB RAM): CAD/3D, software development, GPU workloads, video editing. Often paired with Personal pools."/>

            <TextBlock x:Name="LblTotalUsers" FontWeight="Bold" Text="Total users (named)"/>
            <TextBlock x:Name="LblTotalUsersDesc" Foreground="#a6adc8" FontSize="11" Text="Total named users who will use AVD."/>
            <StackPanel Orientation="Horizontal" Margin="0,3,0,8"><TextBox x:Name="TxtTotalUsers" Width="120" Text="100"/></StackPanel>

            <TextBlock x:Name="LblConcurrency" FontWeight="Bold" Text="Concurrency"/>
            <TextBlock x:Name="LblConcurrencyDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Text="Percent or absolute number of concurrent users."/>
            <StackPanel Orientation="Horizontal" Margin="0,3,0,8">
              <ComboBox x:Name="CmbConcurrencyMode" Width="120" SelectedIndex="0"><ComboBoxItem Content="Percent"/><ComboBoxItem Content="User"/></ComboBox>
              <TextBox x:Name="TxtConcurrencyValue" Width="80" Margin="8,0,0,0" Text="60"/>
              <TextBlock x:Name="LblConcurrencyHint" Margin="8,4,0,0" Foreground="#6c7086" Text="% or #"/>
            </StackPanel>

            <TextBlock x:Name="LblPeakFactor" FontWeight="Bold" Text="Peak factor"/>
            <TextBlock x:Name="LblPeakFactorDesc" Foreground="#a6adc8" FontSize="11" Text="Spike multiplier. 1.0 = none, 1.2 = 20% extra."/>
            <StackPanel Orientation="Horizontal" Margin="0,3,0,8">
              <TextBox x:Name="TxtPeakFactor" Width="120" Text="1.0"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="1.0 = no spike buffer"/>
            </StackPanel>

            <Separator Margin="0,4,0,4"/>
            <TextBlock x:Name="LblNPlus1" FontWeight="Bold" Text="N+1 redundancy"/>
            <TextBlock x:Name="LblNPlus1Desc" Foreground="#a6adc8" FontSize="11" Text="Extra failover hosts. 1 = one standby, 0 = off."/>
            <StackPanel Orientation="Horizontal" Margin="0,3,0,6"><TextBox x:Name="TxtNPlusOne" Width="120" Text="1"/></StackPanel>

            <TextBlock x:Name="LblHeadroom" FontWeight="Bold" Text="Extra headroom"/>
            <TextBlock x:Name="LblHeadroomDesc" Foreground="#a6adc8" FontSize="11" Text="Growth buffer added to peak concurrent users."/>
            <StackPanel Orientation="Horizontal" Margin="0,3,0,6"><TextBox x:Name="TxtExtraHeadroomPct" Width="120" Text="0"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="%"/></StackPanel>

            <Separator Margin="0,4,0,4"/>
            <TextBlock x:Name="LblLoadBal" FontWeight="Bold" Text="Load Balancing (pooled only)"/>
            <TextBlock x:Name="LblLoadBalDesc" Foreground="#a6adc8" FontSize="11" Text="Breadth-first: even spread. Depth-first: fill sequentially."/>
            <ComboBox x:Name="CmbLoadBalancing" Margin="0,3,0,6" SelectedIndex="0">
              <ComboBoxItem Content="Breadth-first (best UX)"/><ComboBoxItem Content="Depth-first (cost optimised)"/>
            </ComboBox>
            <TextBlock x:Name="LblMaxSession" FontWeight="Bold" Text="Max session limit per host"/>
            <TextBlock x:Name="LblMaxSessionDesc" Foreground="#a6adc8" FontSize="11" Text="0 = auto from users/host. Set explicitly for Depth-first."/>
            <StackPanel Orientation="Horizontal" Margin="0,3,0,4"><TextBox x:Name="TxtMaxSessionLimit" Width="120" Text="0"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="0 = auto"/></StackPanel>
          </StackPanel>
          </ScrollViewer>

          <!-- RIGHT: Tuning Parameters -->
          <Border Grid.Column="1" Padding="14" BorderBrush="#585b70" BorderThickness="1" CornerRadius="8" Background="#313244">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock x:Name="LblTuneTitle" FontSize="14" FontWeight="Bold" Foreground="#89b4fa" Text="Tuning Parameters"/>
              <TextBlock x:Name="LblTuneDesc" Foreground="#a6adc8" FontSize="11" Margin="0,4,0,6" Text="Advanced settings. Defaults follow Microsoft best practice."/>
              <Separator/>
              <TextBlock x:Name="LblFsLogix" FontWeight="Bold" Text="FSLogix profile storage"/>
              <TextBlock x:Name="LblFsLogixDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Text="Profile container size per user. MS default max: 30 GB."/>
              <StackPanel Orientation="Horizontal" Margin="0,4,0,4"><TextBox x:Name="TxtProfileGB" Width="100" Text="30"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="GB/user"/></StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,4"><TextBox x:Name="TxtProfileGrowthPct" Width="100" Text="20"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="growth %"/></StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,8"><TextBox x:Name="TxtProfileOverheadPct" Width="100" Text="10"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="overhead %"/></StackPanel>
              <Separator/>
              <TextBlock x:Name="LblCpuRam" FontWeight="Bold" Text="CPU / RAM / System Reserve"/>
              <TextBlock x:Name="LblCpuRamDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Text="Target utilisation and virtualisation overhead (MS: 15-20%)."/>
              <StackPanel Orientation="Horizontal" Margin="0,4,0,4"><TextBox x:Name="TxtCpuUtil" Width="100" Text="0.80"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="CPU target (0.80 = 80%)"/></StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,4"><TextBox x:Name="TxtMemUtil" Width="100" Text="0.80"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="Memory target"/></StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,8"><TextBox x:Name="TxtVirtOverhead" Width="100" Text="0.15"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="System Reserve (0.15-0.20)"/></StackPanel>
              <Separator/>
              <TextBlock x:Name="LblVcpuRange" FontWeight="Bold" Text="vCPU range per host (pooled)"/>
              <TextBlock x:Name="LblVcpuRangeDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Text="MS recommends max 24 for multi-session. 128 = unrestricted."/>
              <StackPanel Orientation="Horizontal" Margin="0,4,0,4"><TextBox x:Name="TxtMinVcpuHost" Width="100" Text="8"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="min vCPU/host"/></StackPanel>
              <StackPanel Orientation="Horizontal"><TextBox x:Name="TxtMaxVcpuHost" Width="100" Text="128"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="max vCPU/host (MS rec: 24)"/></StackPanel>
            </StackPanel>
            </ScrollViewer>
          </Border>
        </Grid>
      </TabItem>

      <!-- TAB 2: Applications -->
      <TabItem Header="Applications">
        <Grid Margin="10">
          <Grid.ColumnDefinitions><ColumnDefinition Width="340"/><ColumnDefinition Width="340"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>

          <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Padding="0,0,8,0">
          <StackPanel>
            <TextBlock x:Name="LblClientApps" FontSize="13" FontWeight="Bold" Foreground="#89b4fa" Text="Client Applications" Margin="0,0,0,4"/>
            <TextBlock x:Name="LblClientAppsDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8" Text="Each adds CPU + RAM overhead per concurrent user on the session host."/>
            <CheckBox x:Name="ChkOffice" Content="Microsoft 365 (Word/Excel/PPT)"/>
            <CheckBox x:Name="ChkTeams" Content="Microsoft Teams"/>
            <CheckBox x:Name="ChkBrowser" Content="Web Browser (Edge/Chrome)"/>
            <CheckBox x:Name="ChkOutlook" Content="Microsoft Outlook"/>
            <CheckBox x:Name="ChkPdf" Content="PDF Editor (Acrobat/Foxit)"/>
            <CheckBox x:Name="ChkErp" Content="ERP Client (SAP GUI / Dynamics)"/>
            <CheckBox x:Name="ChkPowerBi" Content="Power BI Desktop"/>

            <TextBlock x:Name="LblDevTools" FontSize="13" FontWeight="Bold" Foreground="#89b4fa" Text="Development Tools" Margin="0,18,0,4"/>
            <TextBlock Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8" Text="Dev tools increase CPU/RAM. Docker requires Personal pool."/>
            <CheckBox x:Name="ChkVS" Content="Visual Studio"/>
            <CheckBox x:Name="ChkVSCode" Content="VS Code"/>
            <CheckBox x:Name="ChkDocker" Content="Docker Desktop"/>
            <CheckBox x:Name="ChkGit" Content="Git / Build Tools"/>
          </StackPanel>
          </ScrollViewer>

          <ScrollViewer Grid.Column="1" VerticalScrollBarVisibility="Auto" Padding="0,0,8,0">
          <StackPanel>
            <TextBlock x:Name="LblDbEngines" FontSize="13" FontWeight="Bold" Foreground="#89b4fa" Text="Database Engines" Margin="0,0,0,4"/>
            <TextBlock x:Name="LblDbEnginesDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8" Text="Switches to E-series (8 GB/vCPU). Consider Azure SQL for production."/>
            <CheckBox x:Name="ChkSqlExpress" Content="SQL Server Express/Developer"/>
            <CheckBox x:Name="ChkSqlStd" Content="SQL Server Standard/Enterprise"/>
            <CheckBox x:Name="ChkPostgres" Content="PostgreSQL"/>
            <CheckBox x:Name="ChkMySql" Content="MySQL / MariaDB"/>
            <CheckBox x:Name="ChkSqlite" Content="SQLite / MS Access (local DB)"/>

            <TextBlock x:Name="LblGpuApps" FontSize="13" FontWeight="Bold" Foreground="#89b4fa" Text="CAD / GPU Applications" Margin="0,18,0,4"/>
            <TextBlock x:Name="LblGpuAppsDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8" Text="Requires NV-series VMs (NVIDIA GPU). Significantly higher cost."/>
            <CheckBox x:Name="ChkAutoCAD" Content="AutoCAD / AutoCAD LT"/>
            <CheckBox x:Name="ChkRevit" Content="Revit / 3ds Max"/>
            <CheckBox x:Name="ChkSolidWorks" Content="SolidWorks / CATIA"/>
            <CheckBox x:Name="ChkVideoEdit" Content="Video Editing (Premiere/DaVinci)"/>

            <Separator Margin="0,16,0,8"/>
            <TextBlock x:Name="LblDbVolume" FontWeight="Bold" Text="Database data volume per host"/>
            <TextBlock x:Name="LblDbVolumeDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,6" Text="Dedicated data disk size when a database engine is selected."/>
            <StackPanel Orientation="Horizontal"><TextBox x:Name="TxtDbDataGB" Width="120" Text="50"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="GB"/></StackPanel>
          </StackPanel>
          </ScrollViewer>

          <Border Grid.Column="2" Padding="14" BorderBrush="#585b70" BorderThickness="1" CornerRadius="8" Background="#313244">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock x:Name="LblHowAffects" FontSize="14" FontWeight="Bold" Foreground="#89b4fa" Text="How Apps Affect Sizing"/>
              <Separator/>
              <TextBlock x:Name="LblInfoVmAuto" FontWeight="SemiBold" Text="VM Series Auto-Selection"/>
              <TextBlock x:Name="LblInfoVmAutoDesc" Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11" Foreground="#a6adc8" Text="No special apps: D-series. Database: E-series. GPU/CAD: NV-series. GPU always takes priority."/>
              <TextBlock x:Name="LblInfoDbEngines" FontWeight="SemiBold" Text="Database Engines"/>
              <TextBlock x:Name="LblInfoDbEnginesDesc" Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11" Foreground="#a6adc8" Text="E-series, Premium SSD v2 data disk. Personal pool strongly recommended."/>
              <TextBlock x:Name="LblInfoCadGpu" FontWeight="SemiBold" Text="CAD / GPU"/>
              <TextBlock x:Name="LblInfoCadGpuDesc" Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11" Foreground="#a6adc8" Text="NV-series (NVIDIA A10). 2-4 users/host typical. Personal pool recommended for heavy 3D."/>
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
            <TextBlock x:Name="LblRegion" FontWeight="Bold" Text="Azure region"/>
            <TextBlock x:Name="LblRegionDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Text="Region where session hosts will be deployed. Affects VM availability and pricing."/>
            <TextBox x:Name="TxtLocation" Margin="0,4,0,12" Text="Switzerland North"/>

            <TextBlock x:Name="LblVmSeries" FontWeight="Bold" Text="VM series preference"/>
            <TextBlock x:Name="LblVmSeriesDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Text="Specific series is strictly enforced &#x2014; no silent fallback to other series."/>
            <ComboBox x:Name="CmbVmSeries" Margin="0,4,0,12" SelectedIndex="0">
              <ComboBoxItem Content="Any (auto)"/><ComboBoxItem Content="D (general purpose, 4 GB/vCPU)"/>
              <ComboBoxItem Content="E (memory, 8 GB/vCPU)"/><ComboBoxItem Content="F (compute, 2 GB/vCPU)"/>
              <ComboBoxItem Content="NV (GPU visualisation)"/><ComboBoxItem Content="NC (GPU compute)"/>
              <ComboBoxItem Content="B (burstable)"/>
            </ComboBox>

            <Separator/>
            <TextBlock x:Name="LblImage" FontWeight="Bold" Text="Marketplace image"/>
            <TextBlock x:Name="LblImageDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,6" Text="Windows image for session hosts. Azure Login + Load SKUs available in Expert mode."/>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,4"><TextBox x:Name="TxtPublisher" Width="240" Text="MicrosoftWindowsDesktop"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="publisher"/></StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,4"><TextBox x:Name="TxtOffer" Width="240" Text="office-365"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="offer"/></StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,4"><ComboBox x:Name="CmbSku" Width="320"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="SKU"/></StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,10"><TextBox x:Name="TxtVersion" Width="240" Text="latest"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="version"/></StackPanel>
            <StackPanel x:Name="PnlAzureActions" Orientation="Horizontal" Visibility="Collapsed">
              <Button x:Name="BtnAzLogin" Content="Azure Login" Width="140" Margin="0,0,10,0" Background="#45475a" Foreground="#89b4fa"/>
              <Button x:Name="BtnDiscoverSkus" Content="Load SKUs" Width="140" Background="#45475a" Foreground="#89b4fa"/>
            </StackPanel>

            <StackPanel x:Name="PnlTemplateJson" Visibility="Collapsed">
            <Separator Margin="0,12,0,8"/>
            <TextBlock FontWeight="Bold" Text="Template JSON"/>
            <TextBlock Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,6" Text="ARM template JSON for the selected VM. Copy into your IaC deployment."/>
            <TextBox x:Name="TxtTemplateOut" Height="160" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"/>
            </StackPanel>
          </StackPanel>
          </ScrollViewer>

          <Border Grid.Column="1" Padding="14" BorderBrush="#585b70" BorderThickness="1" CornerRadius="8" Background="#313244">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock x:Name="LblVmLogic" FontSize="14" FontWeight="Bold" Foreground="#89b4fa" Text="VM Selection Logic"/>
              <Separator/>
              <TextBlock x:Name="LblInfoSeriesAuto" FontWeight="SemiBold" Text="Series Auto-Selection"/>
              <TextBlock x:Name="LblInfoSeriesAutoDesc" Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11" Foreground="#a6adc8" Text="Standard: D-series. Database: E-series. GPU/CAD: NV-series. GPU takes priority."/>
              <TextBlock x:Name="LblInfoStrictSeries" FontWeight="SemiBold" Text="Strict Series Mode"/>
              <TextBlock x:Name="LblInfoStrictSeriesDesc" Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11" Foreground="#a6adc8" Text="Specific series selected: only VMs of that series. No silent fallback."/>
              <TextBlock x:Name="LblInfoSmallestVm" FontWeight="SemiBold" Text="Smallest Matching VM"/>
              <TextBlock x:Name="LblInfoSmallestVmDesc" Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11" Foreground="#a6adc8" Text="Picks smallest VM meeting calculated vCPU + RAM. Prefers v5/v6 with Premium Storage."/>
              <TextBlock x:Name="LblInfoAvdCompat" FontWeight="SemiBold" Text="AVD-Compatible Only"/>
              <TextBlock x:Name="LblInfoAvdCompatDesc" Margin="0,4,0,8" TextWrapping="Wrap" FontSize="11" Foreground="#a6adc8" Text="D, E, F, NV, NC, B, M series. Excluded: A, H, ND, L, DC/EC, ARM."/>
              <Separator/>
              <TextBlock x:Name="LblInfoRamPerSeries" FontWeight="SemiBold" Text="RAM per VM Series" Margin="0,0,0,4"/>
              <TextBlock FontFamily="Consolas" FontSize="11" Foreground="#a6adc8" xml:space="preserve" Text="D  = 4 GB/vCPU   E  = 8 GB/vCPU&#x0a;F  = 2 GB/vCPU   NV = 7 GB/vCPU&#x0a;NC = 8 GB/vCPU   B  = 4 GB/vCPU&#x0a;M  = 28 GB/vCPU"/>
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
            <TextBlock x:Name="LblResTitle" FontSize="14" FontWeight="Bold" Foreground="#89b4fa" Text="Sizing Results"/>
            <TextBlock x:Name="LblResDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Text="Click 'Calculate' then 'Pick VM from Azure' to find the best VM template with pricing."/>
          </StackPanel>
          <DataGrid Grid.Row="1" x:Name="GridResults" AutoGenerateColumns="True" IsReadOnly="True" Margin="0,4,0,8"/>
          <TextBlock x:Name="LblResNotes" Grid.Row="2" Foreground="#6c7086" FontSize="11" Text="Notes, warnings and recommendations:" Margin="0,0,0,2"/>
          <TextBox Grid.Row="3" x:Name="TxtNotes" Height="170" TextWrapping="Wrap" AcceptsReturn="True"
                   VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="11" IsReadOnly="True"/>
        </Grid>
      </TabItem>

      <!-- TAB 5: Costs (visible with -Expert) -->
      <TabItem Header="Costs" x:Name="TabUserCosts" Visibility="Collapsed">
        <Grid Margin="10">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <StackPanel Grid.Row="0" Margin="0,0,0,8">
            <TextBlock x:Name="LblCostTitle" FontSize="14" FontWeight="Bold" Foreground="#89b4fa" Text="Cost Analysis"/>
            <TextBlock x:Name="LblCostDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Text="Configure pricing, discounts and calculate per-user costs. Requires 'Pick VM from Azure' on the Results tab first."/>
          </StackPanel>
          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="400"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>

            <StackPanel Grid.Column="0" Margin="0,0,16,0">
              <!-- Currency -->
              <TextBlock x:Name="LblPricing" FontSize="13" FontWeight="Bold" Foreground="#89b4fa" Text="Pricing" Margin="0,0,0,4"/>
              <TextBlock x:Name="LblCurrency" FontWeight="Bold" Text="Currency"/>
              <TextBlock x:Name="LblCurrencyDesc" Foreground="#a6adc8" FontSize="11" Text="Currency for Azure Retail Prices API."/>
              <ComboBox x:Name="CmbCurrency" Margin="0,3,0,8" SelectedIndex="2">
                <ComboBoxItem Content="USD"/><ComboBoxItem Content="EUR"/><ComboBoxItem Content="CHF"/>
              </ComboBox>

              <TextBlock x:Name="LblAhb" FontWeight="Bold" Text="Azure Hybrid Benefit"/>
              <TextBlock x:Name="LblAhbDesc" Foreground="#a6adc8" FontSize="11" Text="M365 E3/E5 or Win E3/E5 licenses cover Windows cost."/>
              <CheckBox x:Name="ChkAhb" Content="Azure Hybrid Benefit active" IsChecked="True" Margin="0,3,0,8"/>

              <Separator Margin="0,4,0,4"/>

              <!-- Operating Schedule -->
              <TextBlock x:Name="LblSchedule" FontSize="13" FontWeight="Bold" Foreground="#89b4fa" Text="Operating Schedule" Margin="0,0,0,4"/>
              <TextBlock x:Name="LblHours" FontWeight="Bold" Text="Operating hours per day"/>
              <TextBlock x:Name="LblHoursDesc" Foreground="#a6adc8" FontSize="11" Text="24 = always on, 10 = business hours only."/>
              <StackPanel Orientation="Horizontal" Margin="0,3,0,8"><TextBox x:Name="TxtOperatingHours" Width="100" Text="10"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="hrs/day (1-24)"/></StackPanel>

              <TextBlock x:Name="LblDays" FontWeight="Bold" Text="Operating days per month"/>
              <TextBlock x:Name="LblDaysDesc" Foreground="#a6adc8" FontSize="11" Text="22 = work month, 30 = daily use."/>
              <StackPanel Orientation="Horizontal" Margin="0,3,0,8"><TextBox x:Name="TxtOperatingDays" Width="100" Text="22"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="days/month (1-31)"/></StackPanel>

              <Separator Margin="0,4,0,4"/>

              <!-- Discounts -->
              <TextBlock x:Name="LblDiscounts" FontSize="13" FontWeight="Bold" Foreground="#fab387" Text="Discounts" Margin="0,0,0,4"/>
              <TextBlock x:Name="LblDiscDesc" Foreground="#a6adc8" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,6" Text="Applied cumulatively to the Azure retail list price."/>

              <TextBlock x:Name="LblCsp" FontWeight="Bold" Text="CSP / EA / Partner discount"/>
              <TextBlock x:Name="LblCspDesc" Foreground="#a6adc8" FontSize="11" Text="Negotiated discount (CSP margin, EA, MPA)."/>
              <StackPanel Orientation="Horizontal" Margin="0,3,0,8"><TextBox x:Name="TxtCspDiscount" Width="100" Text="0"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="% (e.g. 15)"/></StackPanel>

              <TextBlock x:Name="LblRi" FontWeight="Bold" Text="Reserved Instance discount"/>
              <TextBlock x:Name="LblRiDesc" Foreground="#a6adc8" FontSize="11" Text="1yr ~35%, 3yr ~55%."/>
              <StackPanel Orientation="Horizontal" Margin="0,3,0,8"><TextBox x:Name="TxtRiDiscount" Width="100" Text="0"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="% (e.g. 35)"/></StackPanel>

              <TextBlock x:Name="LblAdditional" FontWeight="Bold" Text="Additional discount"/>
              <TextBlock x:Name="LblAdditionalDesc" Foreground="#a6adc8" FontSize="11" Text="Savings Plans, promos, custom agreements."/>
              <StackPanel Orientation="Horizontal" Margin="0,3,0,8"><TextBox x:Name="TxtAdditionalDiscount" Width="100" Text="0"/><TextBlock Margin="8,4,0,0" Foreground="#6c7086" Text="%"/></StackPanel>

              <Separator Margin="0,4,0,4"/>
              <Button x:Name="BtnCalcUserCosts" Content="Calculate Costs" Width="200" HorizontalAlignment="Left" FontWeight="Bold"
                      Background="#89b4fa" Foreground="#1e1e2e"/>
            </StackPanel>

            <Border Grid.Column="1" Padding="14" BorderBrush="#585b70" BorderThickness="1" CornerRadius="8" Background="#313244">
              <StackPanel>
                <TextBlock x:Name="LblBreakdown" FontSize="13" FontWeight="Bold" Foreground="#89b4fa" Text="Cost Breakdown" Margin="0,0,0,8"/>
                <DataGrid x:Name="GridUserCosts" AutoGenerateColumns="True" IsReadOnly="True" Margin="0,0,0,8"/>
                <TextBox x:Name="TxtUserCostNotes" Height="140" TextWrapping="Wrap" AcceptsReturn="True"
                         VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="11" IsReadOnly="True"
                         Background="#1e1e2e" Foreground="#a6adc8" BorderBrush="#585b70"/>
              </StackPanel>
            </Border>
          </Grid>
          </ScrollViewer>
        </Grid>
      </TabItem>
    </TabControl>

    <!-- Bottom button bar -->
    <DockPanel Grid.Row="2" LastChildFill="False" Margin="0,4,0,0">
      <StackPanel DockPanel.Dock="Left" Orientation="Horizontal">
        <Button x:Name="BtnReset" Content="Reset" Width="80" Background="#f38ba8" Foreground="#1e1e2e" FontWeight="Bold"/>
        <Button x:Name="BtnDiagnostics" Content="Diagnostics" Width="110" Margin="10,0,0,0" Visibility="Collapsed" Background="#45475a" Foreground="#fab387"/>
      </StackPanel>
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="BtnCalculate" Content="Calculate" Width="120" FontWeight="Bold" Background="#89b4fa" Foreground="#1e1e2e"/>
        <Button x:Name="BtnPickVm" Content="Pick VM from Azure" Width="170" Margin="10,0,0,0" Background="#a6e3a1" Foreground="#1e1e2e" FontWeight="Bold"/>
        <Button x:Name="BtnExportJson" Content="Export JSON" Width="110" Margin="10,0,0,0"/>
        <Button x:Name="BtnExportReport" Content="Export HTML Report" Width="150" Margin="10,0,0,0"/>
        <Button x:Name="BtnClose" Content="Close" Width="80" Margin="10,0,0,0"/>
      </StackPanel>
    </DockPanel>
  </Grid>
</Window>
"@
$XamlString = $XamlString -replace '&(?!amp;|lt;|gt;|quot;|apos;|#x)', '&amp;'
#endregion

#region Build UI + Bind
$xmlReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($XamlString))
$Window = [Windows.Markup.XamlReader]::Load($xmlReader)
$script:Window = $Window
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
$TxtPublisher = $Window.FindName('TxtPublisher'); $TxtOffer = $Window.FindName('TxtOffer')
$CmbSku = $Window.FindName('CmbSku'); $TxtVersion = $Window.FindName('TxtVersion')
$BtnAzLogin = $Window.FindName('BtnAzLogin'); $BtnDiscoverSkus = $Window.FindName('BtnDiscoverSkus')
$TxtTemplateOut = $Window.FindName('TxtTemplateOut')

# Results
$GridResults = $Window.FindName('GridResults'); $TxtNotes = $Window.FindName('TxtNotes')
$BtnCalculate = $Window.FindName('BtnCalculate'); $BtnPickVm = $Window.FindName('BtnPickVm')
$BtnExportJson = $Window.FindName('BtnExportJson'); $BtnExportReport = $Window.FindName('BtnExportReport')
$BtnReset = $Window.FindName('BtnReset'); $BtnClose = $Window.FindName('BtnClose')

# Language selector
$CmbLanguage = $Window.FindName('CmbLanguage')
$script:LangMap = @('en','de','fr','it')
$CmbLanguage.add_SelectionChanged({
  $script:Lang = $script:LangMap[$CmbLanguage.SelectedIndex]
  Apply-Language
})

# Costs tab (visible with -Expert)
$TabUserCosts = $Window.FindName('TabUserCosts')
$CmbCurrency = $Window.FindName('CmbCurrency')
$ChkAhb = $Window.FindName('ChkAhb')
$TxtOperatingHours = $Window.FindName('TxtOperatingHours')
$TxtOperatingDays = $Window.FindName('TxtOperatingDays')
$TxtCspDiscount = $Window.FindName('TxtCspDiscount')
$TxtRiDiscount = $Window.FindName('TxtRiDiscount')
$TxtAdditionalDiscount = $Window.FindName('TxtAdditionalDiscount')
$BtnCalcUserCosts = $Window.FindName('BtnCalcUserCosts')
$GridUserCosts = $Window.FindName('GridUserCosts')
$TxtUserCostNotes = $Window.FindName('TxtUserCostNotes')
if ($Expert) { $TabUserCosts.Visibility = 'Visible' }

# Expert mode: show Template JSON, Export JSON, Azure Login/SKUs, Diagnostics
$PnlTemplateJson = $Window.FindName('PnlTemplateJson')
$PnlAzureActions = $Window.FindName('PnlAzureActions')
$BtnDiagnostics = $Window.FindName('BtnDiagnostics')
if ($Expert) {
  $PnlTemplateJson.Visibility = 'Visible'
  $PnlAzureActions.Visibility = 'Visible'
  $BtnExportJson.Visibility = 'Visible'
  $BtnDiagnostics.Visibility = 'Visible'
} else {
  $PnlTemplateJson.Visibility = 'Collapsed'
  $PnlAzureActions.Visibility = 'Collapsed'
  $BtnExportJson.Visibility = 'Collapsed'
  $BtnDiagnostics.Visibility = 'Collapsed'
}

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

    $hidePricingVal = (-not $Expert)
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
    $hidePricingVal2 = (-not $Expert)
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

    # VM + pricing section (Expert only shows pricing)
    $hidePricing = (-not $Expert)
    $vmHtml = ''
    $sectionNum = 6
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
          <tr><td>Price/Hour (list)</td><td>$($script:LastVmPrice.RetailPricePerHour) $cur</td></tr>
          <tr><td>Est. Monthly/Host (list)</td><td>$monthly $cur</td></tr>
          <tr class="highlight"><td>Est. Monthly Total (list)</td><td><strong>$totalMonthly $cur</strong> ($($s.RecommendedHostsTotal) hosts)</td></tr>
"@
        $priceNote = '<p class="note">Azure retail list price (pay-as-you-go). Discounts, RI, Savings Plans, and AHB not applied. See User Costs section below for effective pricing.</p>'
      }
      $vmHtml = @"
      <section>
        <h2><span class="num">$sectionNum</span> $(Get-Str 'rpt.vmselection')</h2>
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
      $sectionNum++
    }

    # User Costs section (Expert mode only)
    $userCostsHtml = ''
    if ($Expert -and $script:LastVmPrice -and -not ($script:LastVmPrice.PSObject.Properties.Name -contains 'Error' -and $script:LastVmPrice.Error)) {
      $ucHoursPerDay  = [Math]::Max(1, [Math]::Min(24, (ConvertTo-IntSafe -Text $TxtOperatingHours.Text -Default 10)))
      $ucDaysPerMonth = [Math]::Max(1, [Math]::Min(31, (ConvertTo-IntSafe -Text $TxtOperatingDays.Text -Default 22)))
      $ucHasAhb = ($ChkAhb.IsChecked -eq $true)
      $ucCspPct = [Math]::Max(0, [Math]::Min(100, (ConvertTo-DoubleSafe -Text $TxtCspDiscount.Text -Default 0)))
      $ucRiPct  = [Math]::Max(0, [Math]::Min(100, (ConvertTo-DoubleSafe -Text $TxtRiDiscount.Text -Default 0)))
      $ucAddPct = [Math]::Max(0, [Math]::Min(100, (ConvertTo-DoubleSafe -Text $TxtAdditionalDiscount.Text -Default 0)))
      $ucTotalDiscPct = [Math]::Min(100, $ucCspPct + $ucRiPct + $ucAddPct)
      $ucMultiplier = 1.0 - ($ucTotalDiscPct / 100.0)

      $ucListPrice = [double]$script:LastVmPrice.RetailPricePerHour
      $ucEffPrice  = [Math]::Round($ucListPrice * $ucMultiplier, 6)
      $ucCur = esc $script:LastVmPrice.CurrencyCode
      $ucMonthlyHrs = $ucHoursPerDay * $ucDaysPerMonth
      $ucHostsTotal = [int]$s.RecommendedHostsTotal
      $ucTotalUsers = [int]$s.TotalUsers
      $ucPeakUsers  = [int]$s.PeakConcurrentUsers

      $ucMonthlyPerHost = [Math]::Round($ucEffPrice * $ucMonthlyHrs, 2)
      $ucMonthlyAll     = [Math]::Round($ucMonthlyPerHost * $ucHostsTotal, 2)
      $ucPerNamed       = if ($ucTotalUsers -gt 0) { [Math]::Round($ucMonthlyAll / $ucTotalUsers, 2) } else { 0 }
      $ucPerConcurrent  = if ($ucPeakUsers -gt 0) { [Math]::Round($ucMonthlyAll / $ucPeakUsers, 2) } else { 0 }
      $ucPerDay         = if ($ucTotalUsers -gt 0 -and $ucDaysPerMonth -gt 0) { [Math]::Round($ucPerNamed / $ucDaysPerMonth, 2) } else { 0 }
      $ucYearlyPerUser  = [Math]::Round($ucPerNamed * 12, 2)
      $ucYearlyTotal    = [Math]::Round($ucMonthlyAll * 12, 2)

      $ucListMonthlyAll = [Math]::Round($ucListPrice * $ucMonthlyHrs * $ucHostsTotal, 2)
      $ucMonthlySavings = [Math]::Round($ucListMonthlyAll - $ucMonthlyAll, 2)
      $ucYearlySavings  = [Math]::Round($ucMonthlySavings * 12, 2)

      # Discount rows
      $discountRows = ''
      if ($ucTotalDiscPct -gt 0) {
        if ($ucCspPct -gt 0) { $discountRows += "<tr><td>CSP / EA / Partner</td><td>$ucCspPct %</td></tr>`n" }
        if ($ucRiPct -gt 0)  { $discountRows += "<tr><td>Reserved Instance</td><td>$ucRiPct %</td></tr>`n" }
        if ($ucAddPct -gt 0) { $discountRows += "<tr><td>Additional</td><td>$ucAddPct %</td></tr>`n" }
        $discountRows += @"
          <tr class="highlight"><td>Total Discount</td><td><strong>$ucTotalDiscPct %</strong></td></tr>
          <tr><td>Effective Price/Hour</td><td>$ucEffPrice $ucCur (was $ucListPrice $ucCur)</td></tr>
          <tr><td>Monthly Savings</td><td><strong>$ucMonthlySavings $ucCur</strong></td></tr>
          <tr><td>Yearly Savings</td><td><strong>$ucYearlySavings $ucCur</strong></td></tr>
"@
      }

      $discountSection = ''
      if ($ucTotalDiscPct -gt 0) {
        $discountSection = @"
        <h3>Applied Discounts</h3>
        <table><tbody>$discountRows</tbody></table>
"@
      }

      $userCostsHtml = @"
      <section>
        <h2><span class="num">$sectionNum</span> $(Get-Str 'rpt.usercosts')</h2>
        <h3>Operating Schedule</h3>
        <table>
          <tbody>
            <tr><td>Hours/Day</td><td>$ucHoursPerDay</td></tr>
            <tr><td>Days/Month</td><td>$ucDaysPerMonth</td></tr>
            <tr><td>Monthly Hours/Host</td><td>$ucMonthlyHrs</td></tr>
            <tr><td>Azure Hybrid Benefit</td><td>$(if($ucHasAhb){'Yes (Windows license covered)'}else{'No'})</td></tr>
          </tbody>
        </table>
        $discountSection
        <h3>Host Costs</h3>
        <table>
          <tbody>
            <tr><td>Hosts Total</td><td>$ucHostsTotal</td></tr>
            <tr><td>Monthly/Host</td><td>$ucMonthlyPerHost $ucCur</td></tr>
            <tr class="highlight"><td>Monthly All Hosts</td><td><strong>$ucMonthlyAll $ucCur</strong></td></tr>
            <tr><td>Yearly All Hosts</td><td>$ucYearlyTotal $ucCur</td></tr>
          </tbody>
        </table>
        <h3>Per-User Costs</h3>
        <table>
          <tbody>
            <tr><td>Total Named Users</td><td>$ucTotalUsers</td></tr>
            <tr><td>Peak Concurrent Users</td><td>$ucPeakUsers</td></tr>
            <tr class="highlight"><td>Monthly / Named User</td><td><strong>$ucPerNamed $ucCur</strong></td></tr>
            <tr><td>Monthly / Concurrent User</td><td>$ucPerConcurrent $ucCur</td></tr>
            <tr><td>Daily / Named User</td><td>$ucPerDay $ucCur</td></tr>
            <tr class="highlight"><td>Yearly / Named User</td><td><strong>$ucYearlyPerUser $ucCur</strong></td></tr>
          </tbody>
        </table>
        <p class="note">Compute costs only. Does not include FSLogix storage, OS disks, networking, licenses (unless AHB), AVD access rights, monitoring, or backup.</p>
      </section>
"@
      $sectionNum++
    }

    # Warnings
    $warningsHtml = ''
    if ($s.Notes -and @($s.Notes).Count -gt 0) {
      $items = ''; foreach ($n in $s.Notes) { $items += "<li>$(esc $n)</li>`n" }
      $warningsHtml = "<div class='warnings'><h3>$(Get-Str 'rpt.warnings')</h3><ul>$items</ul></div>"
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
        <h2><span class="num">5</span> $(Get-Str 'rpt.storage')</h2>
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
        <h2><span class="num">3</span> $(Get-Str 'rpt.apps')</h2>
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
        <h2><span class="num">3</span> $(Get-Str 'rpt.apps')</h2>
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
      <h1>Azure Virtual Desktop<br><span>$(Get-Str 'rpt.title')</span></h1>
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
    <h2><span class="num">1</span> $(Get-Str 'rpt.execsummary')</h2>
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
    <h2><span class="num">2</span> $(Get-Str 'rpt.workload')</h2>
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
    <h2><span class="num">4</span> $(Get-Str 'rpt.loadbal')</h2>
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

  <!-- User Costs (Expert only) -->
  $userCostsHtml

  <!-- Notes -->
  <section>
    <h2><span class="num">$sectionNum</span> $(Get-Str 'rpt.notes')</h2>
    $warningsHtml
    <h3>Microsoft References</h3>
    <ul style="padding-left:20px; color: var(--text2); font-size:13px;">
      <li><a href="https://learn.microsoft.com/en-us/windows-server/remote/remote-desktop-services/session-host-virtual-machine-sizing-guidelines">Session Host VM Sizing Guidelines</a></li>
      <li><a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/host-pool-load-balancing">Host Pool Load Balancing</a></li>
      <li><a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/autoscale-create-assign-scaling-plan">Autoscale Scaling Plans</a></li>
    </ul>
  </section>

  <footer>
    $(Get-Str 'rpt.generated') AVD Sizing Calculator v$ScriptVersion &bull; All recommendations should be validated with <a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/insights">AVD Insights</a> monitoring and pilot deployments.
  </footer>
</div>
</body>
</html>
"@

    [IO.File]::WriteAllText($reportPath, $html, [Text.Encoding]::UTF8)
    Write-UiInfo "Report exported: $reportPath"

  } catch { Write-UiError "Report export failed: $($_.Exception.Message)" }
})

# User Costs calculation
$BtnCalcUserCosts.add_Click({
  try {
    if (-not $script:LastSizing) { Write-UiWarning 'Calculate sizing first.'; return }
    if (-not $script:LastVmPick) { Write-UiWarning "Click 'Pick VM from Azure' first to get pricing data."; return }
    if (-not $script:LastVmPrice -or ($script:LastVmPrice.PSObject.Properties.Name -contains 'Error' -and $script:LastVmPrice.Error)) {
      Write-UiWarning 'VM pricing not available. Check Azure login and try Pick VM again.'; return
    }

    $s = $script:LastSizing; $r = $s.Recommended; $p = $script:LastVmPrice
    $hoursPerDay  = [Math]::Max(1, [Math]::Min(24, (ConvertTo-IntSafe -Text $TxtOperatingHours.Text -Default 10)))
    $daysPerMonth = [Math]::Max(1, [Math]::Min(31, (ConvertTo-IntSafe -Text $TxtOperatingDays.Text -Default 22)))
    $hasAhb = ($ChkAhb.IsChecked -eq $true)

    # Parse discount percentages (0-100, clamped)
    $cspPct = [Math]::Max(0, [Math]::Min(100, (ConvertTo-DoubleSafe -Text $TxtCspDiscount.Text -Default 0)))
    $riPct  = [Math]::Max(0, [Math]::Min(100, (ConvertTo-DoubleSafe -Text $TxtRiDiscount.Text -Default 0)))
    $addPct = [Math]::Max(0, [Math]::Min(100, (ConvertTo-DoubleSafe -Text $TxtAdditionalDiscount.Text -Default 0)))
    $totalDiscountPct = [Math]::Min(100, $cspPct + $riPct + $addPct)
    $discountMultiplier = 1.0 - ($totalDiscountPct / 100.0)

    $listPricePerHour = [double]$p.RetailPricePerHour
    $effectivePricePerHour = [Math]::Round($listPricePerHour * $discountMultiplier, 6)

    $currency = [string]$p.CurrencyCode
    $hostsTotal = [int]$s.RecommendedHostsTotal
    $usersPerHost = [int]$r.UsersPerHost
    $peakUsers = [int]$s.PeakConcurrentUsers
    $totalUsers = [int]$s.TotalUsers

    # Monthly hours per host
    $monthlyHoursPerHost = $hoursPerDay * $daysPerMonth

    # VM compute cost (with discounts applied)
    $monthlyPerHost = [Math]::Round($effectivePricePerHour * $monthlyHoursPerHost, 2)
    $monthlyAllHosts = [Math]::Round($monthlyPerHost * $hostsTotal, 2)

    # List price costs (for comparison)
    $monthlyPerHostList = [Math]::Round($listPricePerHour * $monthlyHoursPerHost, 2)
    $monthlyAllHostsList = [Math]::Round($monthlyPerHostList * $hostsTotal, 2)
    $monthlySavings = [Math]::Round($monthlyAllHostsList - $monthlyAllHosts, 2)

    # Per-user costs (two perspectives)
    $costPerConcurrentUser = if ($peakUsers -gt 0) { [Math]::Round($monthlyAllHosts / $peakUsers, 2) } else { 0 }
    $costPerNamedUser = if ($totalUsers -gt 0) { [Math]::Round($monthlyAllHosts / $totalUsers, 2) } else { 0 }
    $costPerUserPerDay = if ($totalUsers -gt 0 -and $daysPerMonth -gt 0) { [Math]::Round($costPerNamedUser / $daysPerMonth, 2) } else { 0 }
    $costPerUserPerHour = if ($totalUsers -gt 0 -and $monthlyHoursPerHost -gt 0) { [Math]::Round($monthlyAllHosts / $totalUsers / $monthlyHoursPerHost, 4) } else { 0 }

    # Yearly
    $yearlyPerNamedUser = [Math]::Round($costPerNamedUser * 12, 2)
    $yearlyTotal = [Math]::Round($monthlyAllHosts * 12, 2)

    # Build results grid
    $rows = [System.Collections.Generic.List[object]]::new()
    $rows.Add([pscustomobject]@{ Key='--- VM TEMPLATE ---'; Value='' })
    $rows.Add([pscustomobject]@{ Key='VM Size'; Value=$script:LastVmPick.Name })
    $rows.Add([pscustomobject]@{ Key='List Price/Hour'; Value="$listPricePerHour $currency" })
    if ($totalDiscountPct -gt 0) {
      $rows.Add([pscustomobject]@{ Key='Effective Price/Hour'; Value="$effectivePricePerHour $currency (-$totalDiscountPct%)" })
    }
    $rows.Add([pscustomobject]@{ Key='Operating hours/day'; Value=$hoursPerDay })
    $rows.Add([pscustomobject]@{ Key='Operating days/month'; Value=$daysPerMonth })
    $rows.Add([pscustomobject]@{ Key='Monthly hours/host'; Value=$monthlyHoursPerHost })
    $rows.Add([pscustomobject]@{ Key='Azure Hybrid Benefit'; Value=$(if($hasAhb){'Yes (Windows license covered)'}else{'No (Windows cost included in VM price)'}) })

    if ($totalDiscountPct -gt 0) {
      $rows.Add([pscustomobject]@{ Key='--- DISCOUNTS ---'; Value='' })
      if ($cspPct -gt 0) { $rows.Add([pscustomobject]@{ Key='CSP / EA / Partner'; Value="$cspPct %" }) }
      if ($riPct -gt 0)  { $rows.Add([pscustomobject]@{ Key='Reserved Instance'; Value="$riPct %" }) }
      if ($addPct -gt 0) { $rows.Add([pscustomobject]@{ Key='Additional discount'; Value="$addPct %" }) }
      $rows.Add([pscustomobject]@{ Key='Total discount'; Value="$totalDiscountPct %" })
      $rows.Add([pscustomobject]@{ Key='Monthly savings'; Value="$monthlySavings $currency" })
    }

    $rows.Add([pscustomobject]@{ Key='--- HOST COSTS ---'; Value='' })
    $rows.Add([pscustomobject]@{ Key='Hosts total'; Value=$hostsTotal })
    $rows.Add([pscustomobject]@{ Key='Monthly/Host'; Value="$monthlyPerHost $currency" })
    $rows.Add([pscustomobject]@{ Key='Monthly all Hosts'; Value="$monthlyAllHosts $currency" })
    if ($totalDiscountPct -gt 0) {
      $rows.Add([pscustomobject]@{ Key='Monthly all Hosts (list)'; Value="$monthlyAllHostsList $currency" })
    }
    $rows.Add([pscustomobject]@{ Key='Yearly all Hosts'; Value="$yearlyTotal $currency" })
    $rows.Add([pscustomobject]@{ Key='--- PER USER COSTS ---'; Value='' })
    $rows.Add([pscustomobject]@{ Key='Total named users'; Value=$totalUsers })
    $rows.Add([pscustomobject]@{ Key='Peak concurrent users'; Value=$peakUsers })
    $rows.Add([pscustomobject]@{ Key='Users/Host'; Value=$usersPerHost })
    $rows.Add([pscustomobject]@{ Key='Monthly/named user'; Value="$costPerNamedUser $currency" })
    $rows.Add([pscustomobject]@{ Key='Monthly/concurrent user'; Value="$costPerConcurrentUser $currency" })
    $rows.Add([pscustomobject]@{ Key='Daily/named user'; Value="$costPerUserPerDay $currency" })
    $rows.Add([pscustomobject]@{ Key='Hourly/named user'; Value="$costPerUserPerHour $currency" })
    $rows.Add([pscustomobject]@{ Key='Yearly/named user'; Value="$yearlyPerNamedUser $currency" })

    $GridUserCosts.ItemsSource = $rows

    # Notes
    $notes = [System.Text.StringBuilder]::new()
    [void]$notes.AppendLine("USER COST ANALYSIS")
    [void]$notes.AppendLine("==================")
    [void]$notes.AppendLine("VM: $($script:LastVmPick.Name) | $hostsTotal hosts | $effectivePricePerHour $currency/hr")
    [void]$notes.AppendLine("Schedule: ${hoursPerDay}h/day x ${daysPerMonth} days = $monthlyHoursPerHost hrs/month")
    if ($totalDiscountPct -gt 0) {
      [void]$notes.AppendLine("")
      [void]$notes.AppendLine("APPLIED DISCOUNTS:")
      if ($cspPct -gt 0) { [void]$notes.AppendLine("  CSP / EA / Partner:     $cspPct %") }
      if ($riPct -gt 0)  { [void]$notes.AppendLine("  Reserved Instance:      $riPct %") }
      if ($addPct -gt 0) { [void]$notes.AppendLine("  Additional:             $addPct %") }
      [void]$notes.AppendLine("  --------------------------------")
      [void]$notes.AppendLine("  Total discount:         $totalDiscountPct %")
      [void]$notes.AppendLine("  List price:             $listPricePerHour $currency/hr")
      [void]$notes.AppendLine("  Effective price:        $effectivePricePerHour $currency/hr")
      [void]$notes.AppendLine("  Monthly savings:        $monthlySavings $currency")
      [void]$notes.AppendLine("  Yearly savings:         $([Math]::Round($monthlySavings * 12, 2)) $currency")
    }
    [void]$notes.AppendLine("")
    [void]$notes.AppendLine("COMPUTE ONLY (after discounts, no storage/networking/licenses):")
    [void]$notes.AppendLine("  Monthly total: $monthlyAllHosts $currency")
    [void]$notes.AppendLine("  Per named user/month: $costPerNamedUser $currency")
    [void]$notes.AppendLine("  Per named user/year: $yearlyPerNamedUser $currency")
    [void]$notes.AppendLine("")
    [void]$notes.AppendLine("NOT INCLUDED:")
    [void]$notes.AppendLine("  - FSLogix profile storage (Azure Files Premium)")
    [void]$notes.AppendLine("  - OS managed disks")
    [void]$notes.AppendLine("  - Networking (VNet, NSG, Private Endpoints)")
    [void]$notes.AppendLine("  - Windows / M365 licenses (unless AHB)")
    [void]$notes.AppendLine("  - AVD access rights (M365 E3/E5/F3, Win E3/E5)")
    [void]$notes.AppendLine("  - Monitoring (Log Analytics, AVD Insights)")
    [void]$notes.AppendLine("  - Backup, DR, management tools")
    if (-not $hasAhb) {
      [void]$notes.AppendLine("")
      [void]$notes.AppendLine("TIP: Enable Azure Hybrid Benefit to save ~40% on Windows VMs")
      [void]$notes.AppendLine("  Requires: Windows Server SA or M365 E3/E5 licenses")
    }
    if ($totalDiscountPct -eq 0) {
      [void]$notes.AppendLine("")
      [void]$notes.AppendLine("TIP: Enter your CSP/EA discount to see effective pricing.")
      [void]$notes.AppendLine("  Typical CSP margins: 5-15%. EA discounts: 10-25%.")
      [void]$notes.AppendLine("  Combined with RI (1yr ~35%, 3yr ~55%) for maximum savings.")
    }

    $TxtUserCostNotes.Text = $notes.ToString()
    Write-UiInfo "User costs calculated: $costPerNamedUser $currency/user/month ($yearlyPerNamedUser $currency/user/year)$(if($totalDiscountPct -gt 0){" (incl. $totalDiscountPct% discount)"})"
  } catch { Write-UiError "User cost calculation failed: $($_.Exception.Message)" }
})

# Diagnostics handler (Expert mode only) — opens a dialog with info + Disconnect/Login buttons
$BtnDiagnostics.add_Click({
  try {
    # --- Gather diagnostics info ---
    $diag = [System.Text.StringBuilder]::new()
    [void]$diag.AppendLine("=== AVD SIZING CALCULATOR DIAGNOSTICS ===")
    [void]$diag.AppendLine("Version:      $ScriptVersion")
    [void]$diag.AppendLine("Build:        $ScriptBuildUtc")
    [void]$diag.AppendLine("Date:         $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$diag.AppendLine("PS Version:   $($PSVersionTable.PSVersion)")
    [void]$diag.AppendLine("OS:           $([System.Runtime.InteropServices.RuntimeInformation]::OSDescription)")
    [void]$diag.AppendLine("Thread:       $([System.Threading.Thread]::CurrentThread.ApartmentState)")
    [void]$diag.AppendLine("")

    # Az module status
    [void]$diag.AppendLine("--- Az PowerShell Modules ---")
    $azAccounts = Get-Module -ListAvailable -Name Az.Accounts -ErrorAction SilentlyContinue | Select-Object -First 1
    $azCompute  = Get-Module -ListAvailable -Name Az.Compute  -ErrorAction SilentlyContinue | Select-Object -First 1
    $azResources = Get-Module -ListAvailable -Name Az.Resources -ErrorAction SilentlyContinue | Select-Object -First 1
    [void]$diag.AppendLine("Az.Accounts:  $(if($azAccounts){"v$($azAccounts.Version) OK"}else{'NOT INSTALLED'})")
    [void]$diag.AppendLine("Az.Compute:   $(if($azCompute){"v$($azCompute.Version) OK"}else{'NOT INSTALLED'})")
    [void]$diag.AppendLine("Az.Resources: $(if($azResources){"v$($azResources.Version) OK"}else{'NOT INSTALLED (optional)'})")
    [void]$diag.AppendLine("")

    # Azure connection status
    $isLoggedIn = $false
    $accountId = ''
    [void]$diag.AppendLine("--- Azure Connection ---")
    try {
      if ($azAccounts) {
        Import-Module Az.Accounts -ErrorAction Stop | Out-Null
        $ctx = Get-AzContext -ErrorAction SilentlyContinue
        if ($ctx -and $ctx.Account) {
          $isLoggedIn = $true
          $accountId = $ctx.Account.Id
          [void]$diag.AppendLine("Logged in:    YES")
          [void]$diag.AppendLine("Account:      $($ctx.Account.Id)")
          [void]$diag.AppendLine("Subscription: $($ctx.Subscription.Name) ($($ctx.Subscription.Id))")
          [void]$diag.AppendLine("Tenant:       $($ctx.Tenant.Id)")
        } else {
          [void]$diag.AppendLine("Logged in:    NO")
        }
      } else {
        [void]$diag.AppendLine("Status:       Az modules not installed")
      }
    } catch {
      [void]$diag.AppendLine("Status:       Error: $($_.Exception.Message)")
    }
    [void]$diag.AppendLine("")

    # Region check
    [void]$diag.AppendLine("--- Region Check ---")
    $region = $TxtLocation.Text.Trim()
    [void]$diag.AppendLine("Configured:   $region")
    try {
      if ($azAccounts -and $isLoggedIn) {
        $loc = Get-AzLocation -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $region -or $_.Location -eq $region }
        if ($loc) { [void]$diag.AppendLine("Region found: $($loc.DisplayName) ($($loc.Location))") }
        else      { [void]$diag.AppendLine("Region:       NOT FOUND - check spelling") }
      } else {
        [void]$diag.AppendLine("Region:       Cannot verify (not logged in)")
      }
    } catch { [void]$diag.AppendLine("Region:       Error: $($_.Exception.Message)") }
    [void]$diag.AppendLine("")

    # Calculator state
    [void]$diag.AppendLine("--- Calculator State ---")
    [void]$diag.AppendLine("LastSizing:   $(if($script:LastSizing){'Present'}else{'Empty (run Calculate)'})")
    [void]$diag.AppendLine("LastVmPick:   $(if($script:LastVmPick){$script:LastVmPick.Name}else{'Empty (run Pick VM)'})")
    [void]$diag.AppendLine("LastVmPrice:  $(if($script:LastVmPrice -and -not ($script:LastVmPrice.PSObject.Properties.Name -contains 'Error' -and $script:LastVmPrice.Error)){"$($script:LastVmPrice.RetailPricePerHour) $($script:LastVmPrice.CurrencyCode)/hr"}else{'Empty or error'})")
    [void]$diag.AppendLine("")
    [void]$diag.AppendLine("--- Script Parameters ---")
    [void]$diag.AppendLine("Expert:       $Expert")

    # --- Build Diagnostics WPF Dialog ---
    $diagXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Diagnostics" Height="520" Width="620" WindowStartupLocation="CenterOwner"
        Background="#1e1e2e" Foreground="#cdd6f4" ResizeMode="NoResize">
  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBox Grid.Row="0" x:Name="TxtDiagInfo" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
             VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="11"
             Background="#313244" Foreground="#cdd6f4" BorderBrush="#585b70" Padding="8"/>
    <StackPanel Grid.Row="1" Margin="0,12,0,0">
      <TextBlock FontWeight="Bold" Foreground="#fab387" Text="Azure Account Actions" Margin="0,0,0,6"/>
      <TextBlock x:Name="TxtConnStatus" Foreground="#a6adc8" FontSize="11" Margin="0,0,0,8"/>
      <StackPanel Orientation="Horizontal">
        <Button x:Name="BtnDiagLogin" Content="Login with different account" Width="220" Padding="10,6"
                Background="#89b4fa" Foreground="#1e1e2e" FontWeight="Bold" Cursor="Hand" Margin="0,0,10,0"/>
        <Button x:Name="BtnDiagDisconnect" Content="Disconnect account" Width="180" Padding="10,6"
                Background="#f38ba8" Foreground="#1e1e2e" FontWeight="Bold" Cursor="Hand"/>
      </StackPanel>
    </StackPanel>
    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
      <Button x:Name="BtnDiagClose" Content="Close" Width="100" Padding="10,6"
              Background="#45475a" Foreground="#cdd6f4" Cursor="Hand"/>
    </StackPanel>
  </Grid>
</Window>
"@
    $diagReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($diagXaml))
    $diagWindow = [Windows.Markup.XamlReader]::Load($diagReader)
    $diagWindow.Owner = $Window

    $txtDiagInfo = $diagWindow.FindName('TxtDiagInfo')
    $txtConnStatus = $diagWindow.FindName('TxtConnStatus')
    $btnDiagLogin = $diagWindow.FindName('BtnDiagLogin')
    $btnDiagDisconnect = $diagWindow.FindName('BtnDiagDisconnect')
    $btnDiagClose = $diagWindow.FindName('BtnDiagClose')

    $txtDiagInfo.Text = $diag.ToString()

    if ($isLoggedIn) {
      $txtConnStatus.Text = "Connected as: $accountId"
      $btnDiagDisconnect.IsEnabled = $true
    } else {
      $txtConnStatus.Text = "Not connected to Azure"
      $btnDiagDisconnect.IsEnabled = $false
    }

    if (-not $azAccounts) {
      $btnDiagLogin.IsEnabled = $false
      $btnDiagDisconnect.IsEnabled = $false
      $txtConnStatus.Text = "Az.Accounts module not installed — run: Install-Module Az.Accounts -Scope CurrentUser"
    }

    # Disconnect button
    $btnDiagDisconnect.add_Click({
      try {
        $txtConnStatus.Text = "Disconnecting..."
        $btnDiagDisconnect.IsEnabled = $false
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
        Clear-AzContext -Force -ErrorAction SilentlyContinue | Out-Null
        $txtConnStatus.Text = "Disconnected. All Azure sessions cleared."
        $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#a6e3a1')
        $btnDiagLogin.Content = "Login to Azure"
        Write-UiInfo "Azure account disconnected."
      } catch {
        $txtConnStatus.Text = "Disconnect failed: $($_.Exception.Message)"
        $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#f38ba8')
        $btnDiagDisconnect.IsEnabled = $true
      }
    }.GetNewClosure())

    # Login button
    $btnDiagLogin.add_Click({
      try {
        $txtConnStatus.Text = "Opening Azure login prompt..."
        $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#fab387')
        # Force new login (ignores cached token)
        $newCtx = Connect-AzAccount -Force -ErrorAction Stop
        if ($newCtx -and $newCtx.Context -and $newCtx.Context.Account) {
          $newAcct = $newCtx.Context.Account.Id
          $newSub  = $newCtx.Context.Subscription.Name
          $txtConnStatus.Text = "Logged in as: $newAcct ($newSub)"
          $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#a6e3a1')
          $btnDiagDisconnect.IsEnabled = $true
          Write-UiInfo "Azure login successful: $newAcct"
        } else {
          $txtConnStatus.Text = "Login cancelled or failed."
          $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#f38ba8')
        }
      } catch {
        $txtConnStatus.Text = "Login failed: $($_.Exception.Message)"
        $txtConnStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#f38ba8')
      }
    }.GetNewClosure())

    # Close button
    $btnDiagClose.add_Click({ $diagWindow.Close() }.GetNewClosure())

    [void]$diagWindow.ShowDialog()

  } catch { Write-UiError "Diagnostics failed: $($_.Exception.Message)" }
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
  $TxtTemplateOut.Text = ''

  # Results tab - clear
  $GridResults.ItemsSource = $null
  $TxtNotes.Text = ''

  # Costs tab - clear
  $TxtOperatingHours.Text = '10'
  $TxtOperatingDays.Text = '22'
  $ChkAhb.IsChecked = $true
  $TxtCspDiscount.Text = '0'
  $TxtRiDiscount.Text = '0'
  $TxtAdditionalDiscount.Text = '0'
  $GridUserCosts.ItemsSource = $null
  $TxtUserCostNotes.Text = ''

  Write-UiInfo 'All settings reset to defaults.'
})

$BtnClose.add_Click({ $Window.Close() })

$CmbSku.Items.Clear(); @('win11-24h2-avd-m365','win11-23h2-avd-m365','win11-24h2-avd','win11-23h2-avd') | ForEach-Object { [void]$CmbSku.Items.Add($_) }
$CmbSku.SelectedIndex = 0
#endregion

[void]$Window.ShowDialog()