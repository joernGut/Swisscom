#requires -Version 5.1
<#
.SYNOPSIS
    Analyze an on-premises file server folder structure for SharePoint Online migration readiness.

.DESCRIPTION
    This script provides a Windows Forms GUI to:
      - Select a root folder on a file server.
      - Enter a target SharePoint library URL prefix.
      - Define the output report path (Excel .xlsx).
      - Analyze all files and folders for SharePoint/OneDrive path and naming limitations.
      - Export a beautified Excel report (.xlsx) using ImportExcel.
      - Show progress including elapsed time and estimated remaining time.

    All comments and output are in English by design.

.NOTES
    Requirements:
      - Windows PowerShell 5.1 (or PowerShell 7 on Windows with WinForms support).
      - Internet access for first run (automatic ImportExcel installation).
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ------------------------- Configuration -------------------------

# SharePoint / OneDrive limitations (decoded path)
$Global:MaxDecodedSharePointPathLength = 400      # Entire decoded path including file name
$Global:MaxSegmentLengthSharePoint      = 255     # Max per path segment (name)
$Global:MaxWindowsLegacyPathLength      = 260     # Legacy Windows MAX_PATH
$Global:MaxEncodedSharePointUrlLength   = 400     # Practical best-practice limit for encoded URL

# Invalid characters / reserved names for SharePoint / Windows
$Global:SpInvalidChars = @('"', '*', ':', '<', '>', '?', '/', '\', '|')

$Global:SpReservedNames = @(
    '.lock',
    'CON','PRN','AUX','NUL',
    'COM0','COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9',
    'LPT0','LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9',
    '_vti_',
    'desktop.ini'
)

# ------------------------- Helper: Ensure ImportExcel is installed -------------------------

function Ensure-ImportExcel {
    <#
    .SYNOPSIS
        Ensures that the ImportExcel module is installed and loaded.
        Tries to auto-install it from PSGallery to CurrentUser if missing.
    #>

    if (Get-Module -Name ImportExcel -ErrorAction SilentlyContinue) {
        return $true
    }

    $mod = Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1
    if ($mod) {
        try {
            Import-Module ImportExcel -ErrorAction Stop
            return $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Found ImportExcel module on disk, but could not load it:`n$($_.Exception.Message)",
                "ImportExcel error", 'OK', 'Error'
            ) | Out-Null
            return $false
        }
    }

    if (-not (Get-Command -Name Install-Module -ErrorAction SilentlyContinue)) {
        [System.Windows.Forms.MessageBox]::Show(
            "The cmdlet 'Install-Module' is not available. Please install PowerShellGet / use a more recent PowerShell version, then install ImportExcel manually.",
            "Install-Module not available", 'OK', 'Error'
        ) | Out-Null
            return $false
    }

    $answer = [System.Windows.Forms.MessageBox]::Show(
        "The 'ImportExcel' module is not installed.`n" +
        "It is required to export the report as .xlsx.`n`n" +
        "Do you want to install it now for the current user from PSGallery?",
        "ImportExcel required",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        return $false
    }

    try {
        # Ensure NuGet provider
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop
        }

        # Trust PSGallery if possible
        if (Get-Command -Name Set-PSRepository -ErrorAction SilentlyContinue) {
            try {
                Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction SilentlyContinue
            } catch { }
        }

        [System.Windows.Forms.MessageBox]::Show(
            "Installing 'ImportExcel' module for the current user. This may take a moment...",
            "Installing ImportExcel", 'OK', 'Information'
        ) | Out-Null

        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop

        [System.Windows.Forms.MessageBox]::Show(
            "'ImportExcel' module successfully installed and loaded.",
            "ImportExcel installed", 'OK', 'Information'
        ) | Out-Null

        return $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to install or load 'ImportExcel':`n$($_.Exception.Message)",
            "ImportExcel installation failed", 'OK', 'Error'
        ) | Out-Null
        return $false
    }
}

# ------------------------- Helper Functions: Path checks -------------------------

function Get-SharePointPathInfo {
    <#
    .SYNOPSIS
        Builds decoded and encoded SharePoint path information from a file system item.
    #>
    param(
        [Parameter(Mandatory)]
        [System.IO.FileSystemInfo]$Item,

        [Parameter(Mandatory)]
        [string]$RootPath,

        [Parameter(Mandatory)]
        [string]$UrlPrefix
    )

    $rootResolved = (Resolve-Path -LiteralPath $RootPath).ProviderPath.TrimEnd('\')
    $fullPath     = $Item.FullName

    # Build relative path (Windows -> SharePoint path segments)
    $relativePart = $fullPath.Substring($rootResolved.Length).TrimStart('\')
    if ([string]::IsNullOrEmpty($relativePart)) {
        $relativePart = $Item.Name
    }

    $segments = $relativePart -split '\\'

    # Build decoded path (for 400-character SharePoint limit)
    $spUri       = [System.Uri]$UrlPrefix
    $spBasePath  = $spUri.AbsolutePath.Trim('/')

    $decodedRelativePath = ($segments -join '/')
    if ([string]::IsNullOrEmpty($spBasePath)) {
        $decodedSpPathWithoutDomain = $decodedRelativePath
    } else {
        $decodedSpPathWithoutDomain = $spBasePath + '/' + $decodedRelativePath
    }

    # Build encoded path (realistic URL length including encoding)
    $encodedSegments = foreach ($seg in $segments) {
        [System.Uri]::EscapeDataString($seg)
    }
    $encodedRelativePath = ($encodedSegments -join '/')
    if ([string]::IsNullOrEmpty($spBasePath)) {
        $encodedSpPathWithoutDomain = $encodedRelativePath
    } else {
        $encodedSpPathWithoutDomain = $spBasePath + '/' + $encodedRelativePath
    }

    [pscustomobject]@{
        FullPath                        = $fullPath
        RelativePath                    = $relativePart -replace '\\','/'
        DecodedSharePointPathPart       = $decodedSpPathWithoutDomain
        DecodedSharePointPathLength     = $decodedSpPathWithoutDomain.Length
        EncodedSharePointPathPart       = $encodedSpPathWithoutDomain
        EncodedSharePointPathLength     = $encodedSpPathWithoutDomain.Length
        Segments                        = $segments
    }
}

function Test-SharePointItem {
    <#
    .SYNOPSIS
        Tests a file or folder against common SharePoint/OneDrive limitations.
    #>
    param(
        [Parameter(Mandatory)]
        [System.IO.FileSystemInfo]$Item,

        [Parameter(Mandatory)]
        [string]$RootPath,

        [Parameter(Mandatory)]
        [string]$UrlPrefix
    )

    $issues = New-Object System.Collections.Generic.List[string]

    $info = Get-SharePointPathInfo -Item $Item -RootPath $RootPath -UrlPrefix $UrlPrefix

    $fullPath  = $info.FullPath
    $segments  = $info.Segments
    $itemType  = if ($Item.PSIsContainer) { 'Folder' } else { 'File' }

    # --- Windows legacy path length check ---
    if ($fullPath.Length -gt $Global:MaxWindowsLegacyPathLength) {
        $issues.Add("Windows full path length ($($fullPath.Length)) exceeds legacy MAX_PATH $($Global:MaxWindowsLegacyPathLength). Some tools or older applications may fail.")
    }

    # --- SharePoint decoded path length check ---
    if ($info.DecodedSharePointPathLength -gt $Global:MaxDecodedSharePointPathLength) {
        $issues.Add("Decoded SharePoint path length ($($info.DecodedSharePointPathLength)) exceeds $($Global:MaxDecodedSharePointPathLength) characters (SharePoint/OneDrive limit).")
    }

    # --- Encoded URL length (best practice check) ---
    if ($info.EncodedSharePointPathLength -gt $Global:MaxEncodedSharePointUrlLength) {
        $issues.Add("Encoded SharePoint URL length ($($info.EncodedSharePointPathLength)) exceeds $($Global:MaxEncodedSharePointUrlLength) characters. Browsers/clients may have issues with such deep paths.")
    }

    # --- Per-segment checks (names, invalid chars, reserved names) ---
    foreach ($seg in $segments) {
        if ([string]::IsNullOrWhiteSpace($seg)) {
            continue
        }

        $segTrimmed = $seg.Trim()

        if ($segTrimmed.Length -gt $Global:MaxSegmentLengthSharePoint) {
            $issues.Add("Name segment '$segTrimmed' length ($($segTrimmed.Length)) exceeds $($Global:MaxSegmentLengthSharePoint) characters (SharePoint name limit).")
        }

        if ($seg.StartsWith(' ') -or $seg.EndsWith(' ')) {
            $issues.Add("Name segment '$seg' has leading or trailing spaces, which are not supported in SharePoint.")
        }

        if ($seg.StartsWith('.') -or $seg.EndsWith('.')) {
            $issues.Add("Name segment '$seg' starts or ends with a dot ('.'), which is not supported in SharePoint.")
        }

        if ($seg.StartsWith('~$')) {
            $issues.Add("Name segment '$seg' starts with '~$', which is blocked in SharePoint/OneDrive.")
        }

        if ($Global:SpReservedNames -contains $segTrimmed) {
            $issues.Add("Name segment '$segTrimmed' is a reserved name in Windows/SharePoint and cannot be used.")
        }

        foreach ($ch in $Global:SpInvalidChars) {
            if ($seg.Contains($ch)) {
                $issues.Add("Name segment '$seg' contains invalid character '$ch' for SharePoint/OneDrive.")
                break
            }
        }
    }

    # --- File name specific checks ---
    $name = $Item.Name
    if (-not $Item.PSIsContainer) {
        if ($name.Length -gt 128) {
            $issues.Add("File name '$name' is longer than 128 characters. Some migration tools and clients may not support this.")
        }
    }

    [pscustomobject]@{
        ItemType                        = $itemType
        FullPath                        = $info.FullPath
        RelativePath                    = $info.RelativePath
        DecodedSharePointPathPart       = $info.DecodedSharePointPathPart
        DecodedSharePointPathLength     = $info.DecodedSharePointPathLength
        EncodedSharePointPathPart       = $info.EncodedSharePointPathPart
        EncodedSharePointPathLength     = $info.EncodedSharePointPathLength
        WindowsFullPathLength           = $fullPath.Length
        Name                            = $Item.Name
        IssueSummary                    = ($issues -join ' | ')
        BlockingIssue                   = [bool]($issues.Count -gt 0)
    }
}

# ------------------------- GUI Setup -------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "SharePoint Migration Analyzer"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(800, 320)
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Root folder
$labelRoot = New-Object System.Windows.Forms.Label
$labelRoot.Text = "File server root folder:"
$labelRoot.AutoSize = $true
$labelRoot.Location = New-Object System.Drawing.Point(10, 20)

$textRoot = New-Object System.Windows.Forms.TextBox
$textRoot.Location = New-Object System.Drawing.Point(200, 18)
$textRoot.Size = New-Object System.Drawing.Size(470, 20)

$buttonBrowseRoot = New-Object System.Windows.Forms.Button
$buttonBrowseRoot.Text = "Browse..."
$buttonBrowseRoot.Location = New-Object System.Drawing.Point(680, 16)
$buttonBrowseRoot.Size = New-Object System.Drawing.Size(80, 24)

# SharePoint URL
$labelUrl = New-Object System.Windows.Forms.Label
$labelUrl.Text = "SharePoint library URL prefix:"
$labelUrl.AutoSize = $true
$labelUrl.Location = New-Object System.Drawing.Point(10, 60)

$textUrl = New-Object System.Windows.Forms.TextBox
$textUrl.Location = New-Object System.Drawing.Point(200, 58)
$textUrl.Size = New-Object System.Drawing.Size(560, 20)
$textUrl.Text = "https://tenant.sharepoint.com/sites/YourSite/Shared%20Documents"

# Output path
$labelOutput = New-Object System.Windows.Forms.Label
$labelOutput.Text = "Report output file (.xlsx):"
$labelOutput.AutoSize = $true
$labelOutput.Location = New-Object System.Drawing.Point(10, 100)

$textOutput = New-Object System.Windows.Forms.TextBox
$textOutput.Location = New-Object System.Drawing.Point(200, 98)
$textOutput.Size = New-Object System.Drawing.Size(470, 20)

$buttonBrowseOutput = New-Object System.Windows.Forms.Button
$buttonBrowseOutput.Text = "Browse..."
$buttonBrowseOutput.Location = New-Object System.Drawing.Point(680, 96)
$buttonBrowseOutput.Size = New-Object System.Drawing.Size(80, 24)

# Run button
$buttonRun = New-Object System.Windows.Forms.Button
$buttonRun.Text = "Analyze"
$buttonRun.Location = New-Object System.Drawing.Point(10, 140)
$buttonRun.Size = New-Object System.Drawing.Size(100, 30)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(200, 145)
$progressBar.Size = New-Object System.Drawing.Size(470, 20)
$progressBar.Minimum = 0
$progressBar.Maximum = 100

# Status label
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = "Status: Idle"
$labelStatus.AutoSize = $true
$labelStatus.Location = New-Object System.Drawing.Point(200, 175)

# Add controls
$form.Controls.AddRange(@(
    $labelRoot,   $textRoot,   $buttonBrowseRoot,
    $labelUrl,    $textUrl,
    $labelOutput, $textOutput, $buttonBrowseOutput,
    $buttonRun,   $progressBar, $labelStatus
))

# ------------------------- GUI Event Handlers -------------------------

$buttonBrowseRoot.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select the root folder of the file server structure"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textRoot.Text = $dlg.SelectedPath
    }
})

$buttonBrowseOutput.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title    = "Select report output file (.xlsx)"
    $dlg.Filter   = "Excel Workbook (*.xlsx)|*.xlsx"
    $dlg.FileName = ("SharePointPathAnalysis_{0:yyyyMMdd_HHmmss}.xlsx" -f (Get-Date))
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textOutput.Text = $dlg.FileName
    }
})

# Analyze button click
$buttonRun.Add_Click({
    $rootPath  = $textRoot.Text.Trim()
    $urlPrefix = $textUrl.Text.Trim()
    $output    = $textOutput.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($rootPath) -or -not (Test-Path -LiteralPath $rootPath)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid file server root folder.", "Input required", 'OK', 'Warning') | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($urlPrefix)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter the SharePoint library URL prefix.", "Input required", 'OK', 'Warning') | Out-Null
        return
    }

    try {
        [void][System.Uri]$urlPrefix
    } catch {
        [System.Windows.Forms.MessageBox]::Show("The SharePoint URL prefix is not a valid URL.", "Input error", 'OK', 'Warning') | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($output)) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $desktop   = [Environment]::GetFolderPath("Desktop")
        $output    = Join-Path $desktop "SharePointPathAnalysis_$timestamp.xlsx"
        $textOutput.Text = $output
    }

    # Ensure .xlsx extension
    $ext = [System.IO.Path]::GetExtension($output)
    if ([string]::IsNullOrWhiteSpace($ext) -or $ext -ne ".xlsx") {
        $output = [System.IO.Path]::ChangeExtension($output, ".xlsx")
        $textOutput.Text = $output
    }

    # Ensure ImportExcel is installed and loaded (auto-install if needed)
    if (-not (Ensure-ImportExcel)) {
        return
    }

    $buttonRun.Enabled = $false
    $progressBar.Value = 0
    $labelStatus.Text  = "Status: Collecting items..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $items = Get-ChildItem -LiteralPath $rootPath -Recurse -Force
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error while enumerating items: $($_.Exception.Message)", "Error", 'OK', 'Error') | Out-Null
        $buttonRun.Enabled = $true
        return
    }

    $total     = $items.Count
    $startTime = Get-Date
    $scanDate  = $startTime  # one timestamp for all results

    if ($total -eq 0) {
        $totalStr = "00:00:00"
        [System.Windows.Forms.MessageBox]::Show(
            "No items found in the selected root folder.`nTotal analysis time: $totalStr",
            "SharePoint Analysis", 'OK', 'Information'
        ) | Out-Null
        $labelStatus.Text = "Status: Completed - no items (Total: $totalStr)"
        $buttonRun.Enabled = $true
        return
    }

    $results = New-Object System.Collections.Generic.List[object]
    $index   = 0

    foreach ($item in $items) {
        $index++

        $result = Test-SharePointItem -Item $item -RootPath $rootPath -UrlPrefix $urlPrefix
        if ($result.BlockingIssue) {
            # Add context metadata for nicer report, including limits
            $result | Add-Member -NotePropertyName ScanRoot            -NotePropertyValue $rootPath                            -Force
            $result | Add-Member -NotePropertyName SharePointUrl       -NotePropertyValue $urlPrefix                           -Force
            $result | Add-Member -NotePropertyName ScanDate            -NotePropertyValue $scanDate                            -Force
            $result | Add-Member -NotePropertyName FileServerPathLimit -NotePropertyValue $Global:MaxWindowsLegacyPathLength   -Force
            $result | Add-Member -NotePropertyName SharePointPathLimit -NotePropertyValue $Global:MaxDecodedSharePointPathLength -Force

            $results.Add($result)
        }

        # Time estimation
        $elapsed    = (Get-Date) - $startTime
        $elapsedStr = $elapsed.ToString("hh\:mm\:ss")

        if ($index -gt 1 -and $elapsed.TotalSeconds -gt 1) {
            $avgSecondsPerItem = $elapsed.TotalSeconds / $index
            $remainingSeconds  = $avgSecondsPerItem * ($total - $index)
            $remaining         = [TimeSpan]::FromSeconds($remainingSeconds)
            $remainingStr      = $remaining.ToString("hh\:mm\:ss")
        } else {
            $remainingStr = "estimating..."
        }

        $percent = [int](($index / $total) * 100)
        $progressBar.Value = [Math]::Min([Math]::Max($percent, 0), 100)

        $statusText = "Scanning {0} of {1}: {2} | Elapsed: {3} | Est. remaining: {4}" -f `
                      $index, $total, $item.FullName, $elapsedStr, $remainingStr
        $labelStatus.Text = "Status: " + $statusText

        [System.Windows.Forms.Application]::DoEvents()
    }

    $endTime   = Get-Date
    $totalSpan = $endTime - $startTime
    $totalStr  = $totalSpan.ToString("hh\:mm\:ss")

    if (-not $results -or $results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No blocking issues found.`nTotal analysis time: $totalStr",
            "SharePoint Analysis", 'OK', 'Information'
        ) | Out-Null
        $labelStatus.Text = "Status: Completed - no blocking issues (Total: $totalStr)"
        $buttonRun.Enabled = $true
        $progressBar.Value = 0
        return
    }

    # Reorder and select columns for a clearer report
    $orderedResults = $results | Select-Object `
        ItemType,
        Name,
        RelativePath,
        FullPath,
        WindowsFullPathLength,
        DecodedSharePointPathLength,
        EncodedSharePointPathLength,
        FileServerPathLimit,
        SharePointPathLimit,
        IssueSummary,
        ScanRoot,
        SharePointUrl,
        ScanDate

    try {
        # Export with nice table formatting, title, filters, etc.
        $excelPkg = $orderedResults | Export-Excel -Path $output `
            -WorksheetName "BlockingItems" `
            -TableName "BlockingItems" `
            -TableStyle "Medium9" `
            -AutoSize `
            -AutoFilter `
            -BoldTopRow `
            -FreezeTopRow `
            -Title "SharePoint Migration Path Analysis" `
            -TitleBold `
            -TitleSize 16 `
            -PassThru

        $ws = $excelPkg.Workbook.Worksheets["BlockingItems"]

        # Last used row in sheet
        $lastRow = $ws.Dimension.End.Row

        # Conditional formatting only on data rows (we starten ab Zeile 2, darüber ist Titel/Leerzeile)
        # 1) WindowsFullPathLength > 260  (Spalte E)
        Add-ConditionalFormatting -Worksheet $ws `
            -Address "E2:E$lastRow" `
            -RuleType GreaterThan `
            -ConditionValue $Global:MaxWindowsLegacyPathLength `
            -ForeGroundColor Red `
            -Bold

        # 2) DecodedSharePointPathLength > 400  (Spalte F)
        Add-ConditionalFormatting -Worksheet $ws `
            -Address "F2:F$lastRow" `
            -RuleType GreaterThan `
            -ConditionValue $Global:MaxDecodedSharePointPathLength `
            -ForeGroundColor Red `
            -Bold

        # 3) EncodedSharePointPathLength > 400  (Spalte G)
        Add-ConditionalFormatting -Worksheet $ws `
            -Address "G2:G$lastRow" `
            -RuleType GreaterThan `
            -ConditionValue $Global:MaxEncodedSharePointUrlLength `
            -ForeGroundColor Red `
            -Bold

        # Save final formatting
        Export-Excel -ExcelPackage $excelPkg -AutoSize

        $msg = "Analysis completed. Blocking items: {0}.`nExcel report written to:`n{1}`nTotal analysis time: {2}" -f `
               $results.Count, $output, $totalStr
        [System.Windows.Forms.MessageBox]::Show($msg, "SharePoint Analysis", 'OK', 'Information') | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error while exporting to formatted Excel (.xlsx): $($_.Exception.Message)`nOutput file: $output",
            "Export error", 'OK', 'Error'
        ) | Out-Null
    }

    $labelStatus.Text = "Status: Completed (Total: $totalStr)"
    $progressBar.Value = 0
    $buttonRun.Enabled = $true
})

[void]$form.ShowDialog()
