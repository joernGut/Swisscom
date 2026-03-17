#requires -Version 5.1
<#
.SYNOPSIS
    Analyze an on-premises file server folder structure for SharePoint Online migration readiness.
.DESCRIPTION
    Rules based on Microsoft official documentation (support.microsoft.com/en-us/office/restrictions-and-limitations-in-onedrive-and-sharepoint):

    PATH LENGTH CHECKS (Blocking):
      - SharePoint Online decoded path > 400 chars  [hard limit, MS Docs]
      - Segment (folder/file name)     > 255 chars  [hard limit, MS Docs]
      - OneDrive sync path             > 400 chars  [relative path for sync clients]

    PATH LENGTH CHECKS (Warning):
      - Windows MAX_PATH               > 260 chars  [legacy apps; mitigated by LongPathsEnabled]
      - Excel desktop app path         > 218 chars  [Excel Win32 hard limit, KB 325573]
      - Word/PPT/Access desktop path   > 259 chars  [Office Win32 limit, KB 325573]
      - Folder depth                   > 10 levels  [best practice recommendation]

    NAME / CHARACTER CHECKS (Blocking):
      - Invalid characters: " * : < > ? / \ |     [blocked in SP Online & OneDrive]
      - Leading/trailing spaces                      [silently stripped, causes sync issues]
      - Leading/trailing dot                         [not supported in SP names]
      - Names starting with ~$                       [Office temp file prefix, blocked]
      - Names starting with _vti_                    [SharePoint internal reserved prefix]
      - Folder named exactly 'forms' at root level   [SP reserved folder name]
      - Single dot (.) as name                       [invalid in SP]
      - Windows/SP reserved device names             [CON, PRN, NUL, COM0-9, LPT0-9, etc.]
      - Reserved file names                          [.lock, desktop.ini, ~$.* patterns]

    NAME / CHARACTER CHECKS (Warning):
      - Names containing spaces                      [spaces = %20 in URLs, inflates path length]
      - Characters # and %                           [supported in SP Online but problematic
                                                      in older Office desktop apps and some tools]

    FILE CHECKS:
      - File size > 250 GB                           [Blocking - SP/OneDrive upload limit]
      - Temp file extensions (.tmp, .bak, .DS_Store) [Warning - typically unwanted in SP]

    NOTE: # and % are officially supported in SharePoint Online (enabled since 2017).
    They are flagged as Warning (not Blocking) because they can still cause issues with
    older Office desktop versions and third-party tools.

    NOTE: Encoded URL length has NO limit in SharePoint Online - only the decoded path counts.
.NOTES
    Requires: Windows PowerShell 5.1, ImportExcel (auto-installed)
    Design: Microsoft Fluent 2 / M365 Admin Center light theme
    Source: https://support.microsoft.com/en-us/office/restrictions-and-limitations-in-onedrive-and-sharepoint-64883a5d-228e-48f5-b3d2-eb39e07630fa
#>

Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==============================================================
#  CONFIGURATION
# ==============================================================

$script:Config = @{
    # Hard limits - Blocking violations
    MaxDecodedSharePointPathLength = 400    # MS Docs: decoded path incl. filename
    MaxSegmentLength               = 255    # MS Docs: per folder/file name component
    MaxOneDriveSyncPathLength      = 400    # MS Docs: relative path for OneDrive sync client

    # Soft limits - Warning violations
    MaxWindowsLegacyPathLength     = 260    # Windows MAX_PATH (legacy apps without LongPathsEnabled)
    MaxExcelDesktopPathLength      = 218    # Excel Win32 hard limit (KB 325573) - cannot be changed
    MaxOfficeDesktopPathLength     = 259    # Word/PPT/Access Win32 limit (KB 325573)
    MaxFolderDepth                 = 10     # Best practice: avoid deep nesting

    # File checks
    MaxFileSizeBytes               = 268435456000  # 250 GB - SP/OneDrive upload limit (Blocking)

    # Encoded URL: NO limit in SP Online - only decoded path counts (removed from checks)
}

# Characters that are BLOCKED in SP Online and OneDrive (Blocking)
$script:SpInvalidCharsBlocking = [char[]]@(
    '"', '*', ':', '<', '>', '?', '/', '\', '|'
)

# Characters supported in SP Online since 2017 but still problematic
# in older Office desktop apps and some migration tools (Warning only)
$script:SpInvalidCharsWarning = [char[]]@('#', '%')

# Reserved names: exact-match (case-insensitive)
$script:SpReservedNames = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        '.lock', 'desktop.ini', '.',
        'CON','PRN','AUX','NUL',
        'COM0','COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9',
        'LPT0','LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9'
    ),
    [System.StringComparer]::OrdinalIgnoreCase
)

# Temp/junk file extensions that typically should not be migrated (Warning)
$script:TempExtensions = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@('.tmp','.bak','.temp','.ds_store','.thumbs.db'),
    [System.StringComparer]::OrdinalIgnoreCase
)


# All rules enabled by default; GUI checkboxes update this set at runtime
$script:EnabledRules = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        'PATH-SP','PATH-OD','PATH-WIN','PATH-XL','PATH-OFF','DEPTH',
        'FILE-SIZE','FILE-TEMP',
        'NAME-LEN','NAME-TRIM','NAME-DOT','NAME-TILDE','NAME-TILDE-FOLDER',
        'NAME-VTI','NAME-FORMS','NAME-RESERVED',
        'CHAR-BLOCKED','CHAR-WARN','NAME-SPACES'
    ),
    [System.StringComparer]::OrdinalIgnoreCase
)

# ==============================================================
#  HELPER: Ensure ImportExcel module
# ==============================================================

function Initialize-ImportExcel {
    if (Get-Module -Name ImportExcel -ErrorAction SilentlyContinue) { return $true }
    $mod = Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1
    if ($mod) {
        try   { Import-Module ImportExcel -ErrorAction Stop; return $true }
        catch { Show-MsgError "Found ImportExcel on disk but could not load it:`n$($_.Exception.Message)" "ImportExcel Error"; return $false }
    }
    if (-not (Get-Command -Name Install-Module -ErrorAction SilentlyContinue)) {
        Show-MsgError "Install-Module is not available. Please install ImportExcel manually." "Module Not Found"
        return $false
    }
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "The 'ImportExcel' module is required but not installed.`n`nInstall it now for the current user from PSGallery?",
        "ImportExcel Required",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }
    try {
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop
        }
        try { Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop
        Show-MsgInfo "'ImportExcel' successfully installed and loaded." "Module Installed"
        return $true
    } catch {
        Show-MsgError "Failed to install ImportExcel:`n$($_.Exception.Message)" "Installation Failed"
        return $false
    }
}

function Show-MsgError ($msg,$title) { [System.Windows.Forms.MessageBox]::Show($msg,$title,'OK','Error')       | Out-Null }
function Show-MsgWarn  ($msg,$title) { [System.Windows.Forms.MessageBox]::Show($msg,$title,'OK','Warning')     | Out-Null }
function Show-MsgInfo  ($msg,$title) { [System.Windows.Forms.MessageBox]::Show($msg,$title,'OK','Information') | Out-Null }

# ==============================================================
#  CORE: Build SharePoint path information
# ==============================================================

function Get-SharePointPathInfo {
    param(
        [System.IO.FileSystemInfo] $Item,
        [string] $RootResolved,
        [string] $SpBasePath
    )
    $relativePart = $Item.FullName.Substring($RootResolved.Length).TrimStart('\')
    if ([string]::IsNullOrEmpty($relativePart)) { $relativePart = $Item.Name }
    [string[]]$segments = @($relativePart -split '\\')
    $decodedRel  = $segments -join '/'
    $decodedFull = if ([string]::IsNullOrEmpty($SpBasePath)) { $decodedRel } else { "$SpBasePath/$decodedRel" }
    $encodedSegs = $segments | ForEach-Object { [System.Uri]::EscapeDataString($_) }
    $encodedRel  = $encodedSegs -join '/'
    $encodedFull = if ([string]::IsNullOrEmpty($SpBasePath)) { $encodedRel } else { "$SpBasePath/$encodedRel" }
    [pscustomobject]@{
        FullPath                    = $Item.FullName
        RelativePath                = $relativePart -replace '\\','/'
        DecodedSharePointPathPart   = $decodedFull
        DecodedSharePointPathLength = $decodedFull.Length
        EncodedSharePointPathPart   = $encodedFull
        EncodedSharePointPathLength = $encodedFull.Length
        Segments                    = $segments
    }
}

function Test-SharePointItem {
    param(
        [System.IO.FileSystemInfo] $Item,
        [string] $RootPath,
        [string] $UrlPrefix
    )
    $issueList   = [System.Collections.Generic.List[object]]::new()
    $script:_sev = 0

    function Add-Issue ([string]$code,[string]$issue,[string]$explanation,[string]$action,[int]$sev=2) {
        [void]$issueList.Add([pscustomobject]@{ Code=$code; Issue=$issue; Explanation=$explanation; Action=$action; Sev=$sev })
        if ($sev -gt $script:_sev) { $script:_sev = $sev }
    }

    $spUri        = [System.Uri]$UrlPrefix
    $spBase       = $spUri.LocalPath.Trim('/')
    $rootResolved = $RootPath.TrimEnd('\')
    $info         = Get-SharePointPathInfo -Item $Item -RootResolved $rootResolved -SpBasePath $spBase
    $fullPath     = $info.FullPath
    $depth        = @($info.Segments | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count

    # BLOCKING: SharePoint decoded path limit (400 chars)
    if ($info.DecodedSharePointPathLength -gt $script:Config.MaxDecodedSharePointPathLength) {
        $over = $info.DecodedSharePointPathLength - $script:Config.MaxDecodedSharePointPathLength
        if ($script:EnabledRules.Contains('PATH-SP')) { Add-Issue 'PATH-SP' "SP decoded path too long: $($info.DecodedSharePointPathLength)/$($script:Config.MaxDecodedSharePointPathLength) chars" "SharePoint Online enforces a hard limit of $($script:Config.MaxDecodedSharePointPathLength) characters for the decoded path including library base path and file name. Files exceeding this limit cannot be uploaded or accessed in SharePoint Online." "Shorten folder names along the path, reduce nesting depth, or move the file closer to the library root. Must reduce by $over characters." 2 }
    }

    # BLOCKING: OneDrive sync relative path limit (400 chars)
    $relLen = $info.RelativePath.Length
    if ($relLen -gt $script:Config.MaxOneDriveSyncPathLength) {
        $over = $relLen - $script:Config.MaxOneDriveSyncPathLength
        if ($script:EnabledRules.Contains('PATH-OD')) { Add-Issue 'PATH-OD' "OneDrive sync path too long: $relLen/$($script:Config.MaxOneDriveSyncPathLength) chars" "The OneDrive sync client limits the relative file path to $($script:Config.MaxOneDriveSyncPathLength) characters. Files exceeding this cannot be synced to Windows or Mac and will show a sync error." "Shorten folder names or reduce nesting. Note: the local OneDrive root folder adds further characters on top of this limit. Must reduce by $over characters." 2 }
    }

    # WARNING: Windows MAX_PATH (260 chars)
    if ($fullPath.Length -gt $script:Config.MaxWindowsLegacyPathLength) {
        $over = $fullPath.Length - $script:Config.MaxWindowsLegacyPathLength
        if ($script:EnabledRules.Contains('PATH-WIN')) { Add-Issue 'PATH-WIN' "Windows path too long: $($fullPath.Length)/$($script:Config.MaxWindowsLegacyPathLength) chars" "Windows has a legacy MAX_PATH limit of $($script:Config.MaxWindowsLegacyPathLength) characters. Applications without long-path awareness will fail to open, copy, or process this file." "Option 1: Enable LongPathsEnabled in Windows (Win10 1607+). Option 2: Shorten path by renaming folders or reducing nesting. Must reduce by $over characters." 1 }
    }

    # WARNING: Folder nesting depth
    if ($Item.PSIsContainer -and $depth -gt $script:Config.MaxFolderDepth) {
        if ($script:EnabledRules.Contains('DEPTH')) { Add-Issue 'DEPTH' "Folder nesting too deep: $depth levels (max recommended: $($script:Config.MaxFolderDepth))" "Microsoft recommends keeping folder hierarchies to a maximum of $($script:Config.MaxFolderDepth) levels. Deeper structures increase path length risk, reduce usability, and degrade SharePoint search performance." "Flatten the folder structure before migration. Consolidate sub-folders with few items. Use SharePoint metadata columns instead of deep folder nesting for categorisation." 1 }
    }

    # File-only checks
    if (-not $Item.PSIsContainer) {
        $ext = [System.IO.Path]::GetExtension($Item.Name).ToLower()

        # WARNING: Excel desktop path limit (218 chars)
        if ($fullPath.Length -gt $script:Config.MaxExcelDesktopPathLength -and $ext -in @('.xlsx','.xlsm','.xlsb','.xls','.xlam','.xltx','.xltm')) {
            $over = $fullPath.Length - $script:Config.MaxExcelDesktopPathLength
            if ($script:EnabledRules.Contains('PATH-XL')) { Add-Issue 'PATH-XL' "Excel desktop path too long: $($fullPath.Length)/$($script:Config.MaxExcelDesktopPathLength) chars" "Microsoft Excel Win32 has a hard-coded path limit of $($script:Config.MaxExcelDesktopPathLength) characters (KB 325573). This cannot be changed. If users open this via the synced OneDrive folder, Excel will refuse to open or save it." "Shorten the path before migration. Move file to a shallower folder or rename parent folders. Must reduce by $over chars. Workaround: open via browser (Excel Online has no this limit)." 1 }
        }

        # WARNING: Word/PPT/Access desktop path limit (259 chars)
        if ($fullPath.Length -gt $script:Config.MaxOfficeDesktopPathLength -and $ext -in @('.docx','.docm','.doc','.pptx','.pptm','.ppt','.accdb','.mdb')) {
            $over = $fullPath.Length - $script:Config.MaxOfficeDesktopPathLength
            if ($script:EnabledRules.Contains('PATH-OFF')) { Add-Issue 'PATH-OFF' "Office desktop path too long: $($fullPath.Length)/$($script:Config.MaxOfficeDesktopPathLength) chars" "Word, PowerPoint, and Access Win32 apps have a path limit of $($script:Config.MaxOfficeDesktopPathLength) characters (KB 325573). Opening synced files whose local path exceeds this will fail." "Shorten path by renaming parent folders or reducing nesting. Must reduce by $over chars. Workaround: open via browser (Office Online has no this limit)." 1 }
        }

        # BLOCKING: File size > 250 GB
        $sz = ([System.IO.FileInfo]$Item).Length
        if ($sz -gt $script:Config.MaxFileSizeBytes) {
            if ($script:EnabledRules.Contains('FILE-SIZE')) { Add-Issue 'FILE-SIZE' "File too large: $([Math]::Round($sz/1GB,2)) GB (limit: 250 GB)" "SharePoint Online and OneDrive enforce a maximum single-file upload size of 250 GB. Files exceeding this cannot be uploaded by any method." "Split the file into smaller parts before migration, or archive to Azure Blob Storage. Evaluate whether the file needs to be in SharePoint at all." 2 }
        }

        # WARNING: Temp/junk file extensions
        if ($script:TempExtensions.Contains($ext)) {
            if ($script:EnabledRules.Contains('FILE-TEMP')) { Add-Issue 'FILE-TEMP' "Temp/system file: $($Item.Name)" "Files with extension '$ext' are temporary or OS system files. They have no business value in SharePoint. OneDrive may ignore or delete some after migration." "Exclude this file from migration. Add '$ext' to the exclusion list in your migration tool. Review whether any business-critical content is stored here before deleting." 1 }
        }
    }

    # Per-segment name checks
    $segIndex = 0
    foreach ($seg in $info.Segments) {
        if ([string]::IsNullOrWhiteSpace($seg)) { continue }
        $segIndex++
        $segT = $seg.Trim()

        # BLOCKING: segment too long
        if ($segT.Length -gt $script:Config.MaxSegmentLength) {
            $over = $segT.Length - $script:Config.MaxSegmentLength
            $segShow = if ($segT.Length -gt 40) { $segT.Substring(0,40) + '...' } else { $segT }
            if ($script:EnabledRules.Contains('NAME-LEN')) { Add-Issue 'NAME-LEN' "Name too long ($($segT.Length)/$($script:Config.MaxSegmentLength) chars): '$segShow'" "Each individual folder or file name is limited to $($script:Config.MaxSegmentLength) characters in SharePoint Online. This item cannot be created or synced." "Rename to a shorter name. Must reduce by $over characters. Use abbreviations or remove redundant words." 2 }
        }

        # BLOCKING: leading/trailing spaces
        if ($seg -ne $segT) {
            if ($script:EnabledRules.Contains('NAME-TRIM')) { Add-Issue 'NAME-TRIM' "Leading/trailing spaces in name: '$seg'" "SharePoint silently strips leading/trailing spaces during upload, causing a name mismatch and sync conflicts. OneDrive may create duplicate files." "Rename to remove the leading or trailing spaces before migration. Use PowerShell: Rename-Item -LiteralPath `$path -NewName `$name.Trim()" 2 }
        }

        # BLOCKING: leading or trailing dot
        if ($seg.StartsWith('.') -or $seg.EndsWith('.')) {
            if ($script:EnabledRules.Contains('NAME-DOT')) { Add-Issue 'NAME-DOT' "Name starts or ends with a dot: '$seg'" "SharePoint does not support names beginning or ending with a period. These items cannot be uploaded." "Rename to remove the leading or trailing dot. Hidden Unix/Mac files (e.g. .gitignore) should be excluded from migration." 2 }
        }

        # BLOCKING: ~$ prefix (Office temp lock files)
        if ($seg.StartsWith('~$')) {
            if ($script:EnabledRules.Contains('NAME-TILDE')) { Add-Issue 'NAME-TILDE' "Office temp lock file: '$seg'" "Names starting with ~dollar are temporary lock files created by Office while a document is open. These are blocked by OneDrive and SharePoint." "Delete this file - it is a temporary lock file, not a real document. Close the corresponding Office document first if still open." 2 }
        }

        # BLOCKING: ~ prefix on folders
        if ($Item.PSIsContainer -and $seg.StartsWith('~')) {
            if ($script:EnabledRules.Contains('NAME-TILDE-FOLDER')) { Add-Issue 'NAME-TILDE-FOLDER' "Folder name starts with tilde: '$seg'" "SharePoint does not allow folder names beginning with a tilde character." "Rename the folder to remove the leading tilde before migration." 2 }
        }

        # BLOCKING: _vti_ prefix
        if ($seg.ToLowerInvariant().StartsWith('_vti_')) {
            if ($script:EnabledRules.Contains('NAME-VTI')) { Add-Issue 'NAME-VTI' "Reserved SP prefix in name: '$seg'" "Names beginning with '_vti_' are reserved for SharePoint internal system use and are blocked in all SharePoint versions." "Rename the item to remove the '_vti_' prefix. Choose a business-meaningful alternative name." 2 }
        }

        # BLOCKING: 'forms' folder at library root level
        if ($segIndex -eq 1 -and $Item.PSIsContainer -and $seg.ToLowerInvariant() -eq 'forms') {
            if ($script:EnabledRules.Contains('NAME-FORMS')) { Add-Issue 'NAME-FORMS' "Reserved folder name 'forms' at library root" "SharePoint uses a 'forms' folder internally at root level of every document library. A user folder with this name at root level conflicts with SharePoint internals." "Rename to something other than 'forms' (e.g. 'Application-Forms'). Sub-folders named 'forms' at deeper levels are permitted." 2 }
        }

        # BLOCKING: Windows/SP reserved names
        if ($script:SpReservedNames.Contains($segT)) {
            if ($script:EnabledRules.Contains('NAME-RESERVED')) { Add-Issue 'NAME-RESERVED' "Reserved name: '$segT'" "Windows device names (CON, PRN, AUX, NUL, COM0-9, LPT0-9) and system files (desktop.ini, .lock) cannot be used as file or folder names and will fail during migration." "Rename before migration - add a prefix or suffix (e.g. 'NUL' -> 'NUL-data')." 2 }
        }

        # BLOCKING: invalid characters
        $foundBlockChar = $false
        foreach ($ch in $script:SpInvalidCharsBlocking) {
            if ($seg.IndexOf($ch) -ge 0) {
                if ($script:EnabledRules.Contains('CHAR-BLOCKED')) { Add-Issue 'CHAR-BLOCKED' "Blocked character '$ch' in name: '$seg'" "The character '$ch' is not allowed in SharePoint Online or OneDrive file/folder names. Items with this character cannot be uploaded." "Replace '$ch' with an allowed alternative: use '-' or '_' instead of ':', '*', '?', '|'; remove '<' '>'; replace '/' or '' with '-'. Use a bulk-rename tool for large item sets." 2 }
                $foundBlockChar = $true
                break
            }
        }

        # WARNING: # and % (allowed in SP Online but risky)
        if (-not $foundBlockChar) {
            foreach ($ch in $script:SpInvalidCharsWarning) {
                if ($seg.IndexOf($ch) -ge 0) {
                    if ($script:EnabledRules.Contains('CHAR-WARN')) { Add-Issue 'CHAR-WARN' "Character '$ch' may cause issues: '$seg'" "'$ch' is supported in SP Online since 2017 but causes problems with older Office desktop apps (pre-2016), some migration tools, and REST API calls." "Consider renaming: replace '#' with 'Nr' or 'No.', replace '%' with 'Pct'. Acceptable if all users are on modern M365 clients and no third-party tools access these files." 1 }
                    break
                }
            }
        }

        # WARNING: spaces in name
        if ($segT -eq $seg -and $seg.Contains(' ')) {
            $spCount = ([regex]::Matches($seg,' ')).Count
            $suggested = $seg -replace ' ','_'
            if ($script:EnabledRules.Contains('NAME-SPACES')) { Add-Issue 'NAME-SPACES' "Spaces in name inflate URL length: '$seg'" "Each space is encoded as '%20' in the URL (+2 chars each). This name with $spCount space(s) adds $($spCount * 2) extra characters to the encoded URL. Paths that look short can exceed limits when encoded." "Replace spaces with underscores or hyphens: '$suggested'. This is best-practice, not a hard requirement." 1 }
        }
    }

    # Build output object
    $sevLabel  = switch ($script:_sev) { 2 { 'Blocking' } 1 { 'Warning' } default { 'OK' } }
    $summary   = ($issueList | ForEach-Object { "[$($_.Code)] $($_.Issue)" }) -join ' | '
    $detailArr = for ($i = 0; $i -lt $issueList.Count; $i++) {
        $it = $issueList[$i]
        "--- Finding $($i+1) [$($it.Code)] ---`nISSUE: $($it.Issue)`nWHY:   $($it.Explanation)`nFIX:   $($it.Action)"
    }
    $details = $detailArr -join "`n`n"
    $actions = ($issueList | ForEach-Object { "[$($_.Code)] $($_.Action)" }) -join "`n"

    [pscustomobject]@{
        Severity                    = $sevLabel
        ItemType                    = if ($Item.PSIsContainer) { 'Folder' } else { 'File' }
        Name                        = $Item.Name
        RelativePath                = $info.RelativePath
        FullPath                    = $info.FullPath
        WindowsFullPathLength       = $fullPath.Length
        DecodedSharePointPathLength = $info.DecodedSharePointPathLength
        FolderDepth                 = $depth
        IssueCount                  = $issueList.Count
        IssueSummary                = $summary
        IssueDetails                = $details
        RecommendedActions          = $actions
        MigrationStatus             = $(
            $blockingCodes = @($issueList | Where-Object { $_.Sev -eq 2 } | ForEach-Object { $_.Code })
            $warningCodes  = @($issueList | Where-Object { $_.Sev -eq 1 } | ForEach-Object { $_.Code })
            # SP-blocking codes = hard limits that prevent upload to SharePoint
            $spBlockers = @('PATH-SP','PATH-OD','FILE-SIZE','NAME-LEN','NAME-TRIM','NAME-DOT','NAME-TILDE','NAME-TILDE-FOLDER','NAME-VTI','NAME-FORMS','NAME-RESERVED','CHAR-BLOCKED')
            $winOnlyCodes = @($blockingCodes | Where-Object { $_ -in $spBlockers })
            if ($blockingCodes.Count -eq 0 -and $warningCodes.Count -eq 0) {
                'SP: Ready'
            } elseif ($blockingCodes.Count -eq 0) {
                'SP: Ready (Warnings)'
            } elseif ($winOnlyCodes.Count -gt 0) {
                'SP: Blocked'
            } else {
                'SP: Blocked'
            }
        )
        WindowsCompatible           = $(if ($issueList.Count -eq 0) { 'Yes' } elseif (@($issueList | Where-Object { $_.Code -notin @('PATH-SP','PATH-OD','PATH-XL','PATH-OFF','CHAR-WARN','NAME-SPACES','FILE-TEMP','DEPTH') }).Count -gt 0) { 'Check Required' } else { 'Yes' })
        BlockingIssue               = [bool]($issueList.Count -gt 0)
    }
}

function Export-AnalysisReport {
    param(
        [System.Collections.Generic.List[object]] $Results,
        [string]   $OutputPath,
        [string]   $RootPath,
        [string]   $UrlPrefix,
        [DateTime] $ScanDate,
        [int]      $TotalScanned,
        [TimeSpan] $Duration
    )

    # ── Colour palette (matching Fluent / M365) ──────────────────
    $xlBlue       = [System.Drawing.ColorTranslator]::FromHtml('#0078d4')
    $xlBlueDark   = [System.Drawing.ColorTranslator]::FromHtml('#005a9e')
    $xlBlueLight  = [System.Drawing.ColorTranslator]::FromHtml('#deecf9')
    $xlRed        = [System.Drawing.ColorTranslator]::FromHtml('#a4262c')
    $xlRedLight   = [System.Drawing.ColorTranslator]::FromHtml('#fde7e9')
    $xlOrange     = [System.Drawing.ColorTranslator]::FromHtml('#d83b01')
    $xlOrangeLight= [System.Drawing.ColorTranslator]::FromHtml('#fff4ce')
    $xlGreen      = [System.Drawing.ColorTranslator]::FromHtml('#107c10')
    $xlGreenLight = [System.Drawing.ColorTranslator]::FromHtml('#dff6dd')
    $xlGrey       = [System.Drawing.ColorTranslator]::FromHtml('#605e5c')
    $xlGreyLight  = [System.Drawing.ColorTranslator]::FromHtml('#f3f2f1')
    $xlWhite      = [System.Drawing.Color]::White
    $xlBlack      = [System.Drawing.ColorTranslator]::FromHtml('#323130')

    # ── Helper: column-number to letter ─────────────────────────
    function ColLetter ([int]$n) {
        $s = ''
        while ($n -gt 0) { $n--; $s = [char](65+($n%26))+$s; $n=[int]($n/26) }
        $s
    }

    # Statistics
    # @() forces array - PS 5.1 returns bare object (no .Count) when Where-Object has 1 match
    $allResults    = @($Results)
    $blockingArr   = @($allResults | Where-Object { $_.Severity -eq 'Blocking' })
    $warningArr    = @($allResults | Where-Object { $_.Severity -eq 'Warning'  })
    $folderArr     = @($allResults | Where-Object { $_.ItemType -eq 'Folder'   })
    $fileArr       = @($allResults | Where-Object { $_.ItemType -eq 'File'     })
    $blockingCount = $blockingArr.Count
    $warningCount  = $warningArr.Count
    $totalIssues   = $allResults.Count
    $cleanCount    = $TotalScanned - $totalIssues
    $folderIssues  = $folderArr.Count
    $fileIssues    = $fileArr.Count
    $pctClean      = if ($TotalScanned -gt 0) { [Math]::Round($cleanCount  / $TotalScanned * 100, 1) } else { 0 }
    $pctIssue      = if ($TotalScanned -gt 0) { [Math]::Round($totalIssues / $TotalScanned * 100, 1) } else { 0 }

    # Categorise issues - no special chars in hashtable keys
    $catCounts = [ordered]@{
        'Path too long (Windows)'        = 0
        'Path too long (SharePoint)'     = 0
        'Invalid character'              = 0
        'Reserved name'                  = 0
        'Space or dot issue'             = 0
        'Reserved prefix'                = 0
        'File name too long'             = 0
        'File too large (250 GB limit)'  = 0
        'Other'                          = 0
    }
    foreach ($r in $allResults) {
        $s = $r.IssueSummary
        if     ($s -match '\[PATH\].*MAX_PATH|\[PATH\].*Windows path') { $catCounts['Path too long (Windows)']++ }
        elseif ($s -match '\[PATH\].*SP path|\[PATH\].*OneDrive|\[PATH\].*Excel|\[PATH\].*Office') { $catCounts['Path too long (SharePoint)']++ }
        elseif ($s -match '\[CHAR\]') { $catCounts['Invalid character']++ }
        elseif ($s -match 'reserved name|reserved prefix|_vti_|tilde') { $catCounts['Reserved name']++ }
        elseif ($s -match '\[NAME\].*spaces|\[NAME\].*dot|\[NAME\].*trailing') { $catCounts['Space or dot issue']++ }
        elseif ($s -match 'dollar|~dollar|_vti_') { $catCounts['Reserved prefix']++ }
        elseif ($s -match '\[SIZE\]') { $catCounts['File name too long']++ }
        elseif ($s -match '250 GB|\[SIZE\]') { $catCounts['File too large (250 GB limit)']++ }
        else                                   { $catCounts['Other']++ }
    }

    # ══════════════════════════════════════════════════════════
    #  SHEET 1 — DASHBOARD
    # ══════════════════════════════════════════════════════════

    # Delete existing file so ExcelPackage starts fresh (avoids "worksheet already exists" on re-scan)
    if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }
    $pkg = [OfficeOpenXml.ExcelPackage]::new([System.IO.FileInfo]$OutputPath)
    $wsDash = $pkg.Workbook.Worksheets.Add("Dashboard")

    # Helper: write styled cell
    function WC ([OfficeOpenXml.ExcelWorksheet]$ws,[int]$r,[int]$c,[object]$val,
                 [System.Drawing.Color]$bg,   [System.Drawing.Color]$fg,
                 [bool]$bold=$false,           [int]$size=10,
                 [string]$hAlign='Left') {
        $cell = $ws.Cells[$r,$c]
        $cell.Value = $val
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($bg)
        $cell.Style.Font.Color.SetColor($fg)
        $cell.Style.Font.Bold   = $bold
        $cell.Style.Font.Size   = $size
        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::$hAlign
        $cell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    }

    # Title banner  (A1:F1 merged)
    $wsDash.Cells["A1:F1"].Merge = $true
    WC $wsDash 1 1 "SharePoint Migration Analyzer - Scan Report" $xlBlue $xlWhite $true 16 'Left'
    $wsDash.Cells["A1"].Style.Indent = 1
    $wsDash.Row(1).Height = 32

    # Sub-title (scan metadata)  A2:F2
    $wsDash.Cells["A2:F2"].Merge = $true
    $metaLine = "Scanned: $($ScanDate.ToString('yyyy-MM-dd HH:mm'))   Root: $RootPath   Duration: $($Duration.ToString('hh\:mm\:ss'))"
    WC $wsDash 2 1 $metaLine $xlBlueDark $xlWhite $false 9 'Left'
    $wsDash.Cells["A2"].Style.Indent = 1
    $wsDash.Row(2).Height = 18

    # ── KPI tiles  (row 4..8, columns A-B / C-D / E-F) ──────────
    $wsDash.Row(3).Height = 8   # spacer

    function KpiTile ([int]$startCol,[string]$label,[object]$value,[System.Drawing.Color]$tileBg,[System.Drawing.Color]$tileFg) {
        $endCol = $startCol + 1
        $colL   = ColLetter $startCol
        $colR   = ColLetter $endCol
        # Label row
        $wsDash.Cells["${colL}4:${colR}4"].Merge = $true
        WC $wsDash 4 $startCol $label $tileBg $tileFg $false 8 'Center'
        $wsDash.Row(4).Height = 16
        # Value row
        $wsDash.Cells["${colL}5:${colR}8"].Merge = $true
        WC $wsDash 5 $startCol $value $tileBg $tileFg $true 28 'Center'
        for ($rr = 5; $rr -le 8; $rr++) { $wsDash.Row($rr).Height = 18 }
        # Border around tile
        $range = $wsDash.Cells["${colL}4:${colR}8"]
        $range.Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin, $xlWhite) | Out-Null
    }

    KpiTile 1 "TOTAL SCANNED"    $TotalScanned  $xlBlue     $xlWhite
    KpiTile 3 "ITEMS WITH ISSUES" $totalIssues   $xlOrange   $xlWhite
    KpiTile 5 "BLOCKING"          $blockingCount $xlRed      $xlWhite

    $wsDash.Row(9).Height = 8   # spacer

    # ── Secondary KPI row ────────────────────────────────────────
    function KpiSmall ([int]$startCol,[string]$label,[object]$value,[System.Drawing.Color]$bg,[System.Drawing.Color]$fg) {
        $endCol = $startCol + 1
        $colL   = ColLetter $startCol
        $colR   = ColLetter $endCol
        $wsDash.Cells["${colL}10:${colR}10"].Merge = $true
        WC $wsDash 10 $startCol $label $bg $fg $false 8 'Center'
        $wsDash.Row(10).Height = 14
        $wsDash.Cells["${colL}11:${colR}12"].Merge = $true
        WC $wsDash 11 $startCol $value $bg $fg $true 14 'Center'
        for ($rr = 11; $rr -le 12; $rr++) { $wsDash.Row($rr).Height = 14 }
        $range = $wsDash.Cells["${colL}10:${colR}12"]
        $range.Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin, $xlGreyLight) | Out-Null
    }

    KpiSmall 1 "Warnings"        $warningCount            $xlOrangeLight $xlOrange
    KpiSmall 3 "Clean Items"     "$cleanCount ($pctClean%)" $xlGreenLight  $xlGreen
    KpiSmall 5 "Folder Issues"   $folderIssues             $xlGreyLight   $xlGrey

    $wsDash.Row(13).Height = 8   # spacer

    # ── Issue breakdown table ─────────────────────────────────────
    WC $wsDash 14 1 "Issue Category Breakdown" $xlBlueDark $xlWhite $true 10 'Left'
    $wsDash.Cells["A14:F14"].Merge = $true
    $wsDash.Row(14).Height = 20

    # Header
    WC $wsDash 15 1 "Category"  $xlBlueLight $xlBlack $true 9 'Left'
    WC $wsDash 15 2 "Count"     $xlBlueLight $xlBlack $true 9 'Center'
    WC $wsDash 15 3 "% of Issues" $xlBlueLight $xlBlack $true 9 'Center'
    $wsDash.Row(15).Height = 16

    $row = 16
    foreach ($cat in $catCounts.GetEnumerator()) {
        if ($cat.Value -eq 0) { continue }
        $pct = if ($totalIssues -gt 0) { [Math]::Round($cat.Value / $totalIssues * 100, 1) } else { 0 }
        $rowBg = if ($row % 2 -eq 0) { $xlGreyLight } else { $xlWhite }
        WC $wsDash $row 1 $cat.Key   $rowBg $xlBlack $false 9 'Left'
        WC $wsDash $row 2 $cat.Value $rowBg $xlBlack $false 9 'Center'
        WC $wsDash $row 3 "$pct %"   $rowBg $xlBlack $false 9 'Center'
        $wsDash.Row($row).Height = 15
        $row++
    }

    # Border around breakdown table
    $wsDash.Cells["A15:C$($row-1)"].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin, $xlBlueLight) | Out-Null

    # ── Scan metadata block (bottom of dashboard) ────────────────
    $row++
    WC $wsDash $row 1 "Scan Configuration" $xlBlueDark $xlWhite $true 10 'Left'
    $wsDash.Cells["A${row}:F${row}"].Merge = $true
    $wsDash.Row($row).Height = 20
    $row++

    $metaRows = [ordered]@{
        'Scan Root'              = $RootPath
        'SharePoint URL'         = $UrlPrefix
        'Scan Date'              = $ScanDate.ToString("yyyy-MM-dd HH:mm:ss")
        'Scan Duration'          = $Duration.ToString("hh\:mm\:ss")
        'Total Scanned'          = $TotalScanned
        'Max SP Decoded Path'    = "$($script:Config.MaxDecodedSharePointPathLength) chars (SP Online hard limit)"
        'Max OneDrive Sync Path' = "$($script:Config.MaxOneDriveSyncPathLength) chars (OneDrive sync client)"
        'Max Windows Path'       = "$($script:Config.MaxWindowsLegacyPathLength) chars (legacy MAX_PATH)"
        'Max Excel Desktop Path' = "$($script:Config.MaxExcelDesktopPathLength) chars (Excel Win32, KB 325573)"
        'Max Office Desktop'     = "$($script:Config.MaxOfficeDesktopPathLength) chars (Word/PPT/Access Win32)"
        'Max Segment Length'     = "$($script:Config.MaxSegmentLength) chars (per folder/file name)"
        'Max Folder Depth'       = "$($script:Config.MaxFolderDepth) levels (best practice)"
        'Max File Size'          = '250 GB (SP/OneDrive upload limit)'
    }
    foreach ($kv in $metaRows.GetEnumerator()) {
        $rowBg = if ($row % 2 -eq 0) { $xlGreyLight } else { $xlWhite }
        WC $wsDash $row 1 $kv.Key   $rowBg $xlGrey  $true  9 'Left'
        WC $wsDash $row 2 $kv.Value $rowBg $xlBlack $false 9 'Left'
        $wsDash.Cells["B${row}:F${row}"].Merge = $true
        $wsDash.Row($row).Height = 15
        $row++
    }

    # Column widths for Dashboard
    $wsDash.Column(1).Width = 38
    $wsDash.Column(2).Width = 22
    $wsDash.Column(3).Width = 16
    $wsDash.Column(4).Width = 16
    $wsDash.Column(5).Width = 16
    $wsDash.Column(6).Width = 16

    # ══════════════════════════════════════════════════════════
    #  SHEET 2 — ISSUES & MIGRATION STATUS (merged)
    # ══════════════════════════════════════════════════════════

    # Pre-compute per-result values needed for the sheet
    $spReady    = @($allResults | Where-Object { $_.MigrationStatus -eq 'SP: Ready' }).Count
    $spWarning  = @($allResults | Where-Object { $_.MigrationStatus -eq 'SP: Ready (Warnings)' }).Count
    $spBlocked  = @($allResults | Where-Object { $_.MigrationStatus -eq 'SP: Blocked' }).Count
    $cleanItems = $TotalScanned - $allResults.Count

    $wsIss = $pkg.Workbook.Worksheets.Add("Issues")

    # ── Title ────────────────────────────────────────────────────
    $wsIss.Cells["A1:M1"].Merge = $true
    $wsIss.Cells["A1"].Value    = "Issues & Migration Status  -  SharePoint Online vs. Windows File Server"
    $wsIss.Cells["A1"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $wsIss.Cells["A1"].Style.Fill.BackgroundColor.SetColor($xlBlueDark)
    $wsIss.Cells["A1"].Style.Font.Color.SetColor($xlWhite)
    $wsIss.Cells["A1"].Style.Font.Bold = $true
    $wsIss.Cells["A1"].Style.Font.Size = 13
    $wsIss.Row(1).Height = 28

    # ── Legend ────────────────────────────────────────────────────
    $legend = "SP: Ready = migrierbar ohne Einschraenkungen.   SP: Ready (Warnings) = Migration moeglich, aber Probleme im Alltag moeglich (z.B. aeltere Office-Clients).   SP: Blocked = Item KANN NICHT zu SharePoint migriert werden - Behebung erforderlich.   Windows Only = Funktioniert auf dem Fileserver, ist aber mit SharePoint nicht kompatibel.   Severity: Blocking = Upload schlaegt fehl | Warning = Upload moeglich, aber Einschraenkungen"
    $wsIss.Cells["A2:M2"].Merge = $true
    $wsIss.Cells["A2"].Value    = $legend
    $wsIss.Cells["A2"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $wsIss.Cells["A2"].Style.Fill.BackgroundColor.SetColor($xlBlueLight)
    $wsIss.Cells["A2"].Style.Font.Color.SetColor($xlBlueDark)
    $wsIss.Cells["A2"].Style.Font.Italic = $true
    $wsIss.Cells["A2"].Style.WrapText = $true
    $wsIss.Row(2).Height = 48

    # ── Summary block ─────────────────────────────────────────────
    $wsIss.Cells["A3:M3"].Merge = $true
    $wsIss.Cells["A3"].Value    = "ZUSAMMENFASSUNG"
    $wsIss.Cells["A3"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $wsIss.Cells["A3"].Style.Fill.BackgroundColor.SetColor($xlBlueDark)
    $wsIss.Cells["A3"].Style.Font.Color.SetColor($xlWhite)
    $wsIss.Cells["A3"].Style.Font.Bold = $true
    $wsIss.Row(3).Height = 18

    $sumDefs = [ordered]@{
        "Gescannte Items total"                          = $TotalScanned
        "Bereit fuer SharePoint Online (gesamt)"         = "$($cleanItems + $spReady + $spWarning) von $TotalScanned ($([Math]::Round(($cleanItems+$spReady+$spWarning)/$TotalScanned*100,1))%)"
        "  davon: Vollstaendig kompatibel (keine Issues)" = "$cleanItems Items - keine Massnahmen noetig"
        "  davon: Kompatibel mit Warnungen"              = "$spWarning Items - Nutzung pruefen (z.B. aeltere Office-Clients, URL-Laenge)"
        "Geblockt - Behebung erforderlich"               = "$spBlocked Items ($([Math]::Round($spBlocked/$TotalScanned*100,1))%) - koennen NICHT migriert werden"
        "Nur Fileserver-kompatibel (Windows Only)"       = "$(@($allResults | Where-Object { $_.WindowsCompatible -eq 'Check Required' }).Count) Items benoetigen Pruefung"
    }
    $bgMap = @{
        "Bereit fuer SharePoint Online (gesamt)"          = $xlGreenLight
        "  davon: Vollstaendig kompatibel (keine Issues)" = $xlGreenLight
        "  davon: Kompatibel mit Warnungen"               = $xlOrangeLight
        "Geblockt - Behebung erforderlich"                = $xlRedLight
    }
    $sr = 4
    foreach ($kv in $sumDefs.GetEnumerator()) {
        $bg = if ($bgMap.ContainsKey($kv.Key)) { $bgMap[$kv.Key] } else { $xlWhite }
        $bold = -not $kv.Key.StartsWith(' ')
        foreach ($c in 1..13) {
            $wsIss.Cells[$sr,$c].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $wsIss.Cells[$sr,$c].Style.Fill.BackgroundColor.SetColor($bg)
        }
        $wsIss.Cells[$sr,1].Value = $kv.Key
        $wsIss.Cells[$sr,1].Style.Font.Bold = $bold
        $wsIss.Cells["B${sr}:M${sr}"].Merge = $true
        $wsIss.Cells[$sr,2].Value = $kv.Value
        $wsIss.Cells[$sr,2].Style.Font.Bold = $bold
        $wsIss.Row($sr).Height = 16
        $sr++
    }

    # ── Section header: item detail ───────────────────────────────
    $wsIss.Cells["A${sr}:M${sr}"].Merge = $true
    $wsIss.Cells["A${sr}"].Value    = "ITEM DETAIL  -  Alle Items mit Problemen (sortiert nach Schwere)"
    $wsIss.Cells["A${sr}"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $wsIss.Cells["A${sr}"].Style.Fill.BackgroundColor.SetColor($xlBlueDark)
    $wsIss.Cells["A${sr}"].Style.Font.Color.SetColor($xlWhite)
    $wsIss.Cells["A${sr}"].Style.Font.Bold = $true
    $wsIss.Row($sr).Height = 18
    $sr++

    # ── Column headers ────────────────────────────────────────────
    # Group A: Migration status (cols 1-3)   Group B: Location (cols 4-6)
    # Group C: Technical (cols 7-9)           Group D: Issue detail (cols 10-13)
    $colHeaders = @(
        @{ H='SP Migration Status';      W=26; Bg=$xlBlueDark  }
        @{ H='Windows Fileserver';       W=20; Bg=$xlBlueDark  }
        @{ H='Severity';                 W=11; Bg=$xlBlueDark  }
        @{ H='Item Type';                W=10; Bg=$xlBlue      }
        @{ H='Name';                     W=36; Bg=$xlBlue      }
        @{ H='Relative Path';            W=60; Bg=$xlBlue      }
        @{ H='SP Path Length (decoded)'; W=14; Bg=$xlBlue      }
        @{ H='Windows Path Length';      W=14; Bg=$xlBlue      }
        @{ H='Folder Depth';             W=12; Bg=$xlBlue      }
        @{ H='Issue Summary';            W=55; Bg=$xlOrange    }
        @{ H='What is the problem?';     W=70; Bg=$xlRed       }
        @{ H='Why does it matter?';      W=70; Bg=$xlRed       }
        @{ H='How to fix it?';           W=70; Bg=$xlGreen     }
    )
    for ($c = 1; $c -le $colHeaders.Count; $c++) {
        $hdr  = $colHeaders[$c-1]
        $cell = $wsIss.Cells[$sr,$c]
        $cell.Value = $hdr.H
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($hdr.Bg)
        $cell.Style.Font.Color.SetColor($xlWhite)
        $cell.Style.Font.Bold = $true
        $cell.Style.Font.Size = 9
        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $cell.Style.WrapText = $true
        $wsIss.Column($c).Width = $hdr.W
    }
    $wsIss.Row($sr).Height = 30
    $dataStartRow = $sr + 1
    $sr++

    # ── Data rows ─────────────────────────────────────────────────
    # Sort: Blocking first, then Warning; within each: by SP path length desc
    $sortedResults = @($allResults | Sort-Object @{E={if($_.Severity -eq 'Blocking'){0}else{1}}},@{E={-$_.DecodedSharePointPathLength}})

    foreach ($r in $sortedResults) {
        # Parse IssueDetails into What/Why/Fix per finding, then join
        $whatList = @(); $whyList = @(); $fixList = @()
        $findings = $r.IssueDetails -split '--- Finding \d+' | Where-Object { $_ -match 'ISSUE:' }
        foreach ($f in $findings) {
            $issue = if ($f -match 'ISSUE:\s*(.+?)(\n|WHY:)') { $Matches[1].Trim() } else { '' }
            $why   = if ($f -match 'WHY:\s*(.+?)(\n|FIX:)')   { $Matches[1].Trim() } else { '' }
            $fix   = if ($f -match 'FIX:\s*(.+)$')            { $Matches[1].Trim() } else { '' }
            if ($issue) { $whatList += $issue }
            if ($why)   { $whyList  += $why   }
            if ($fix)   { $fixList  += $fix   }
        }
        $whatText = $whatList -join "`n"
        $whyText  = $whyList  -join "`n"
        $fixText  = $fixList  -join "`n"

        $rowBg = switch ($r.Severity) {
            'Blocking' { $xlRedLight    }
            'Warning'  { $xlOrangeLight }
            default    { $xlWhite       }
        }

        $vals = @(
            $r.MigrationStatus,
            $r.WindowsCompatible,
            $r.Severity,
            $r.ItemType,
            $r.Name,
            $r.RelativePath,
            $r.DecodedSharePointPathLength,
            $r.WindowsFullPathLength,
            $r.FolderDepth,
            $r.IssueSummary,
            $whatText,
            $whyText,
            $fixText
        )
        for ($c = 1; $c -le $vals.Count; $c++) {
            $cell = $wsIss.Cells[$sr,$c]
            $cell.Value = $vals[$c-1]
            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cell.Style.Fill.BackgroundColor.SetColor($rowBg)
            $cell.Style.WrapText = ($c -ge 10)
        }
        # Colour SP Migration Status cell
        $statusCell = $wsIss.Cells[$sr,1]
        $statusCell.Style.Font.Bold = $true
        switch ($r.MigrationStatus) {
            'SP: Blocked'          { $statusCell.Style.Font.Color.SetColor($xlRed)    }
            'SP: Ready (Warnings)' { $statusCell.Style.Font.Color.SetColor($xlOrange) }
            default                { $statusCell.Style.Font.Color.SetColor($xlGreen)  }
        }
        # Colour Severity cell
        $sevCell = $wsIss.Cells[$sr,3]
        $sevCell.Style.Font.Bold = $true
        switch ($r.Severity) {
            'Blocking' { $sevCell.Style.Font.Color.SetColor($xlRed)    }
            'Warning'  { $sevCell.Style.Font.Color.SetColor($xlOrange) }
        }
        # Conditional: SP path length red if over limit
        if ($r.DecodedSharePointPathLength -gt $script:Config.MaxDecodedSharePointPathLength) {
            $wsIss.Cells[$sr,7].Style.Font.Color.SetColor($xlRed)
            $wsIss.Cells[$sr,7].Style.Font.Bold = $true
        }
        if ($r.WindowsFullPathLength -gt $script:Config.MaxWindowsLegacyPathLength) {
            $wsIss.Cells[$sr,8].Style.Font.Color.SetColor($xlOrange)
            $wsIss.Cells[$sr,8].Style.Font.Bold = $true
        }
        $wsIss.Row($sr).Height = 15
        $sr++
    }

    # ── Auto row height for wrapping text rows ────────────────────
    # Set taller rows for the detail columns
    for ($r2 = $dataStartRow; $r2 -lt $sr; $r2++) {
        $wsIss.Row($r2).Height = 60
    }

    # ── AutoFilter on header row (sort + filter dropdowns on all 13 columns) ───
    $lastDataRow  = $sr - 1
    $filterRange  = "A$($dataStartRow-1):M${lastDataRow}"
    $wsIss.Cells[$filterRange].AutoFilter = $true

    # ── Freeze panes: rows above data + first 3 status columns ────────────────
    $wsIss.View.FreezePanes($dataStartRow, 4)

    # ══════════════════════════════════════════════════════════
    #  SHEET 3 — LIMITS & INFO
    # ══════════════════════════════════════════════════════════

    $wsInfo = $pkg.Workbook.Worksheets.Add("Limits and Rules")

    # Title
    $wsInfo.Cells["A1:C1"].Merge = $true
    $wsInfo.Cells["A1"].Value = "SharePoint Migration - Limits and Rules Reference"
    $wsInfo.Cells["A1"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $wsInfo.Cells["A1"].Style.Fill.BackgroundColor.SetColor($xlBlue)
    $wsInfo.Cells["A1"].Style.Font.Color.SetColor($xlWhite)
    $wsInfo.Cells["A1"].Style.Font.Bold = $true
    $wsInfo.Cells["A1"].Style.Font.Size = 13
    $wsInfo.Row(1).Height = 28

    # Section helper
    function InfoSection ([int]$r,[string]$title) {
        $wsInfo.Cells["A${r}:C${r}"].Merge = $true
        $wsInfo.Cells["A$r"].Value = $title
        $wsInfo.Cells["A$r"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $wsInfo.Cells["A$r"].Style.Fill.BackgroundColor.SetColor($xlBlueDark)
        $wsInfo.Cells["A$r"].Style.Font.Color.SetColor($xlWhite)
        $wsInfo.Cells["A$r"].Style.Font.Bold = $true
        $wsInfo.Cells["A$r"].Style.Font.Size = 10
        $wsInfo.Row($r).Height = 18
    }
    function InfoRow ([int]$r,[string]$setting,[string]$value,[string]$note,[bool]$alt=$false) {
        $bg = if ($alt) { $xlGreyLight } else { $xlWhite }
        foreach ($c in 1..3) {
            $wsInfo.Cells[$r,$c].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $wsInfo.Cells[$r,$c].Style.Fill.BackgroundColor.SetColor($bg)
        }
        $wsInfo.Cells[$r,1].Value = $setting
        $wsInfo.Cells[$r,1].Style.Font.Bold = $true
        $wsInfo.Cells[$r,2].Value = $value
        $wsInfo.Cells[$r,3].Value = $note
        $wsInfo.Row($r).Height = 15
    }

    InfoSection 2 "Path Length Limits"
    $wsInfo.Cells["A3"].Value = "Setting";   $wsInfo.Cells["B3"].Value = "Limit";    $wsInfo.Cells["C3"].Value = "Notes"
    foreach ($c in 1..3) {
        $wsInfo.Cells[3,$c].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $wsInfo.Cells[3,$c].Style.Fill.BackgroundColor.SetColor($xlBlueLight)
        $wsInfo.Cells[3,$c].Style.Font.Bold = $true
    }
    InfoRow 4  "Windows MAX_PATH"             "260 chars"   "Full UNC path incl. filename. Older apps fail above this."        $false
    InfoRow 5  "SharePoint Decoded Path"      "400 chars"   "Full decoded URL path (server-relative). Hard SP limit."          $true
    InfoRow 6  "OneDrive Sync Relative Path"  "400 chars"   "Relative path for OneDrive sync client. Root folder adds ~80-120 chars extra on top." $false
    InfoRow 7  "Segment (folder/file name)"   "255 chars"   "Each individual name component max length."                      $true
    InfoRow 8  "Excel Win32 Desktop App"      "218 chars"   "Excel desktop app hard limit (KB 325573). Cannot be changed. Applies to full local path." $false
    InfoRow 9  "Word/PPT/Access Desktop App"   "259 chars"   "Office Win32 hard limit (KB 325573). Applies to full local path when synced." $true
    InfoRow 10 "Folder depth"               "Max 10 levels" "Best practice. Deeper nesting increases path length and reduces usability." $false
    InfoRow 11 "File size"                  "250 GB"        "Maximum single file upload size to SharePoint/OneDrive."          $true

    InfoSection 11 "Invalid Characters"
    $wsInfo.Cells["A12:C12"] | ForEach-Object {
        $_.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $_.Style.Fill.BackgroundColor.SetColor($xlBlueLight)
        $_.Style.Font.Bold = $true
    }
    $wsInfo.Cells["A12"].Value = "Character(s)"; $wsInfo.Cells["B12"].Value = "Reason"
    $wsInfo.Cells["A12:B12"].Merge = $false
    $wsInfo.Cells["B12:C12"].Merge = $true
    InfoRow 13 '"  *  :  <  >  ?'          "Blocked in all SP/OneDrive file/folder names"      "" $false
    InfoRow 14 '/  \  |'                   "Path separator characters, not allowed in names"   "" $true
    InfoRow 15 '#  %'                      "URL encoding conflicts (# = anchor, % = escape)"   "" $false
    InfoRow 16 "Leading/trailing spaces"   "Silently stripped by SP, causes sync issues"        "" $true
    InfoRow 17 "Leading/trailing dot (.)"  "Not supported in SharePoint names"                 "" $false
    InfoRow 18 "Starts with ~$"            "Office temporary file prefix, blocked by OneDrive" "" $true
    InfoRow 19 "Starts with _vti_"         "SharePoint internal reserved prefix"               "" $false

    InfoSection 21 "Reserved Names (Windows / SharePoint)"
    $wsInfo.Cells["A22"].Value = "CON, PRN, AUX, NUL, COM0-COM9, LPT0-LPT9, .lock, desktop.ini"
    $wsInfo.Cells["A22:C22"].Merge = $true
    $wsInfo.Cells["A22"].Style.Font.Italic = $true
    $wsInfo.Row(22).Height = 15

    # Column widths for Info sheet
    $wsInfo.Column(1).Width = 34
    $wsInfo.Column(2).Width = 52
    $wsInfo.Column(3).Width = 52

    # ── Sheet order: Dashboard first ────────────────────────────
    $pkg.Workbook.Worksheets.MoveToStart("Limits and Rules")
    $pkg.Workbook.Worksheets.MoveToStart("Issues")
    $pkg.Workbook.Worksheets.MoveToStart("Dashboard")

    # ── Save ────────────────────────────────────────────────────
    $pkg.Save()
    $pkg.Dispose()
}

# ==============================================================
#  GUI - DESIGN TOKENS
# ==============================================================

function hex ([string]$h) { [System.Drawing.ColorTranslator]::FromHtml($h) }

# Neutrals
$clrAppBg        = hex '#f0f2f5'
$clrSurface      = hex '#ffffff'
$clrSurfaceAlt   = hex '#f8f9fa'
$clrSurfaceHover = hex '#f3f4f6'
$clrBorder       = hex '#e1e4e8'
$clrBorderInput  = hex '#d0d7de'
$clrBorderFocus  = hex '#0078d4'
# Text
$clrText         = hex '#1a1a2e'
$clrTextSub      = hex '#57606a'
$clrTextDisabled = hex '#8c959f'
# Accent (Microsoft Blue)
$clrAccent       = hex '#0078d4'
$clrAccentDark   = hex '#005a9e'
$clrAccentLight  = hex '#e8f4fd'
$clrAccentMid    = hex '#b3d7f5'
$clrOnAccent     = hex '#ffffff'
# Semantic
$clrSuccess      = hex '#1a7f37'
$clrSuccessBg    = hex '#dafbe1'
$clrWarning      = hex '#bf4b08'
$clrWarningBg    = hex '#fff8e1'
$clrError        = hex '#cf222e'
$clrErrorBg      = hex '#ffebe9'
# Category colours for rules panel
$clrCatPath      = hex '#e8f4fd'
$clrCatName      = hex '#f0f8e8'
$clrCatChar      = hex '#fff4e8'
$clrCatFile      = hex '#f5e8ff'
$clrCatPathBdr   = hex '#b3d7f5'
$clrCatNameBdr   = hex '#b8e0a0'
$clrCatCharBdr   = hex '#f5c87a'
$clrCatFileBdr   = hex '#d4a8f0'
$clrCatPathTxt   = hex '#005a9e'
$clrCatNameTxt   = hex '#2d6a0a'
$clrCatCharTxt   = hex '#7d4200'
$clrCatFileTxt   = hex '#5a1e8c'

# ==============================================================
#  GUI - TYPOGRAPHY
# ==============================================================

$fntH1      = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$fntH2      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$fntH3      = New-Object System.Drawing.Font("Segoe UI",  9, [System.Drawing.FontStyle]::Bold)
$fntBody    = New-Object System.Drawing.Font("Segoe UI",  9, [System.Drawing.FontStyle]::Regular)
$fntBold    = New-Object System.Drawing.Font("Segoe UI",  9, [System.Drawing.FontStyle]::Bold)
$fntSmall   = New-Object System.Drawing.Font("Segoe UI",  8, [System.Drawing.FontStyle]::Regular)
$fntSmallB  = New-Object System.Drawing.Font("Segoe UI",  8, [System.Drawing.FontStyle]::Bold)
$fntCaption = New-Object System.Drawing.Font("Segoe UI",  7, [System.Drawing.FontStyle]::Bold)
$fntMono    = New-Object System.Drawing.Font("Consolas",  8, [System.Drawing.FontStyle]::Regular)

# ==============================================================
#  GUI - CONTROL FACTORIES
# ==============================================================

function New-FLabel {
    param(
        [string]$Text, [int]$X, [int]$Y,
        [System.Drawing.Font]$Font   = $fntBody,
        [System.Drawing.Color]$Color = $clrTextSub
    )
    $l           = New-Object System.Windows.Forms.Label
    $l.Text      = $Text
    $l.AutoSize  = $true
    $l.Location  = New-Object System.Drawing.Point($X,$Y)
    $l.Font      = $Font
    $l.ForeColor = $Color
    $l.BackColor = [System.Drawing.Color]::Transparent
    $l
}

function New-FTextBox {
    param([int]$X,[int]$Y,[int]$W,[int]$H=26)
    $t             = New-Object System.Windows.Forms.TextBox
    $t.Location    = New-Object System.Drawing.Point($X,$Y)
    $t.Size        = New-Object System.Drawing.Size($W,$H)
    $t.Font        = $fntBody
    $t.BackColor   = $clrSurface
    $t.ForeColor   = $clrText
    $t.BorderStyle = 'FixedSingle'
    $t
}

function New-FButton {
    param([string]$Text,[int]$X,[int]$Y,[int]$W,[int]$H=30,[bool]$Primary=$false)
    $b           = New-Object System.Windows.Forms.Button
    $b.Text      = $Text
    $b.Location  = New-Object System.Drawing.Point($X,$Y)
    $b.Size      = New-Object System.Drawing.Size($W,$H)
    $b.Font      = $fntBold
    $b.FlatStyle = 'Flat'
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    if ($Primary) {
        $b.BackColor                  = $clrAccent
        $b.ForeColor                  = $clrOnAccent
        $b.FlatAppearance.BorderSize  = 0
        $b.Add_MouseEnter({ $this.BackColor = $clrAccentDark })
        $b.Add_MouseLeave({ $this.BackColor = $clrAccent })
    } else {
        $b.BackColor                  = $clrSurface
        $b.ForeColor                  = $clrText
        $b.FlatAppearance.BorderSize  = 1
        $b.FlatAppearance.BorderColor = $clrBorderInput
        $b.Add_MouseEnter({ $this.BackColor = $clrSurfaceHover })
        $b.Add_MouseLeave({ $this.BackColor = $clrSurface })
    }
    $b
}

function New-HDivider {
    param([int]$X,[int]$Y,[int]$W,[int]$H=1)
    $p           = New-Object System.Windows.Forms.Panel
    $p.Location  = New-Object System.Drawing.Point($X,$Y)
    $p.Size      = New-Object System.Drawing.Size($W,$H)
    $p.BackColor = $clrBorder
    $p
}

# ==============================================================
#  GUI - TOOLTIP
# ==============================================================

$tip                = New-Object System.Windows.Forms.ToolTip
$tip.AutoPopDelay   = 9000
$tip.InitialDelay   = 500
$tip.ReshowDelay    = 200
$tip.ShowAlways     = $true

function Set-Tip ([System.Windows.Forms.Control]$ctrl,[string]$text) {
    $tip.SetToolTip($ctrl, $text)
}

# ==============================================================
#  GUI - MAIN FORM
# ==============================================================

$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "SharePoint Migration Analyzer"
$form.StartPosition   = "CenterScreen"
$form.Size            = New-Object System.Drawing.Size(960, 918)
$form.MinimumSize     = New-Object System.Drawing.Size(960, 918)
$form.MaximizeBox     = $false
$form.BackColor       = $clrAppBg
$form.ForeColor       = $clrText
$form.Font            = $fntBody
$form.FormBorderStyle = "FixedSingle"

# ── Navigation bar ────────────────────────────────────────────
$pnlNav           = New-Object System.Windows.Forms.Panel
$pnlNav.Size      = New-Object System.Drawing.Size(960, 64)
$pnlNav.Location  = New-Object System.Drawing.Point(0, 0)
$pnlNav.BackColor = $clrSurface
$pnlNav.add_Paint({
    param($s,$e)
    $pen = New-Object System.Drawing.Pen($clrBorder, 1)
    $e.Graphics.DrawLine($pen, 0, $s.Height-1, $s.Width, $s.Height-1)
    $pen.Dispose()
})

# SP icon — rounded square with gradient
$pnlIcon           = New-Object System.Windows.Forms.Panel
$pnlIcon.Size      = New-Object System.Drawing.Size(40, 40)
$pnlIcon.Location  = New-Object System.Drawing.Point(20, 12)
$pnlIcon.BackColor = $clrAccent
$pnlIcon.add_Paint({
    param($s,$e)
    $g  = $e.Graphics
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    # Gradient background
    $rect = New-Object System.Drawing.Rectangle(0,0,$s.Width,$s.Height)
    $grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect,
        (hex '#0091ff'),
        (hex '#0050a0'),
        [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal
    )
    $g.FillRectangle($grad, $rect)
    $grad.Dispose()
    # SP text
    $fnt = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $br  = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $sf  = New-Object System.Drawing.StringFormat
    $sf.Alignment     = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString("SP", $fnt, $br, (New-Object System.Drawing.RectangleF(0,0,40,40)), $sf)
    $fnt.Dispose(); $br.Dispose(); $sf.Dispose()
})

$lblAppTitle = New-FLabel "SharePoint Migration Analyzer" 72 10 $fntH2 $clrText
$lblAppSub   = New-FLabel "Analyzes file server folder structures for SharePoint Online / OneDrive migration compatibility" 72 34 $fntSmall $clrTextSub
$pnlNav.Controls.AddRange(@($pnlIcon, $lblAppTitle, $lblAppSub))

# Blue accent stripe
$pnlStripe           = New-Object System.Windows.Forms.Panel
$pnlStripe.Size      = New-Object System.Drawing.Size(960, 3)
$pnlStripe.Location  = New-Object System.Drawing.Point(0, 64)
$pnlStripe.BackColor = $clrAccent

# ── Scan Configuration card ───────────────────────────────────
# Layout:
#   Y= 0   card border
#   Y= 0   pnlCardHead (36px)
#   Y= 48  Row 1 label
#   Y= 64  Row 1 sub-label
#   Y= 82  Row 1 textbox
#   Y=118  Row 2 label
#   Y=134  Row 2 sub-label
#   Y=152  Row 2 textbox
#   Y=190  Row 3 label
#   Y=206  Row 3 sub-label
#   Y=224  Row 3 textbox
#   Y=262  Info banner (30px)
#   total card height = 300

$pnlCard           = New-Object System.Windows.Forms.Panel
$pnlCard.Size      = New-Object System.Drawing.Size(900, 318)
$pnlCard.Location  = New-Object System.Drawing.Point(20, 76)
$pnlCard.BackColor = $clrSurface
$pnlCard.add_Paint({
    param($s,$e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    # Shadow-like outer border
    $pen1 = New-Object System.Drawing.Pen($clrBorder, 1)
    $g.DrawRectangle($pen1, 0, 0, $s.Width-1, $s.Height-1)
    $pen1.Dispose()
    # Top accent line (3px inside card at top)
    $pen2 = New-Object System.Drawing.Pen($clrAccent, 3)
    $g.DrawLine($pen2, 0, 0, $s.Width, 0)
    $pen2.Dispose()
})

# Card header
$pnlCardHead           = New-Object System.Windows.Forms.Panel
$pnlCardHead.Size      = New-Object System.Drawing.Size(900, 36)
$pnlCardHead.Location  = New-Object System.Drawing.Point(0, 0)
$pnlCardHead.BackColor = $clrAccentLight
$pnlCardHead.add_Paint({
    param($s,$e)
    $pen = New-Object System.Drawing.Pen($clrAccentMid, 1)
    $e.Graphics.DrawLine($pen, 0, $s.Height-1, $s.Width, $s.Height-1)
    $pen.Dispose()
})

$lblCardIcon          = New-FLabel "   [1]  Scan Configuration" 0 9 $fntBold $clrAccent
$lblCardIcon.AutoSize = $false
$lblCardIcon.Size     = New-Object System.Drawing.Size(500, 20)
$lblCardHint          = New-FLabel "Enter paths below, then click Run Analysis" 510 11 $fntSmall $clrAccent
$pnlCardHead.Controls.AddRange(@($lblCardIcon, $lblCardHint))

# Row 1: Root folder
$lblRoot           = New-FLabel "File server root folder" 16 44 $fntBold $clrText
$lblRoot.AutoSize  = $false
$lblRoot.Size      = New-Object System.Drawing.Size(600, 18)
$lblRootSub            = New-FLabel "Top-level folder on the file server - all subfolders and files will be scanned recursively." 16 66 $fntSmall $clrTextSub
$lblRootSub.AutoSize   = $false
$lblRootSub.Size       = New-Object System.Drawing.Size(856, 14)
$txtRoot    = New-FTextBox 16 88 756 26
$btnRoot    = New-FButton "  Browse..." 780 87 104 28

# Row 2: SharePoint URL
$lblUrl           = New-FLabel "SharePoint Online / OneDrive library URL prefix" 16 124 $fntBold $clrText
$lblUrl.AutoSize  = $false
$lblUrl.Size      = New-Object System.Drawing.Size(700, 18)
$lblUrlSub            = New-FLabel "Used to calculate the decoded SharePoint path length per item. Example: https://contoso.sharepoint.com/sites/IT/Shared%20Documents" 16 144 $fntSmall $clrTextSub
$lblUrlSub.AutoSize   = $false
$lblUrlSub.Size       = New-Object System.Drawing.Size(856, 14)
$txtUrl      = New-FTextBox 16 164 868 26
$txtUrl.Text      = $placeholder
$txtUrl.ForeColor = $clrTextDisabled
$txtUrl.Add_Enter({
    if ($txtUrl.ForeColor -eq $clrTextDisabled) { $txtUrl.Text = ""; $txtUrl.ForeColor = $clrText }
})
$txtUrl.Add_Leave({
    if ([string]::IsNullOrWhiteSpace($txtUrl.Text)) {
        $txtUrl.Text      = $placeholder
        $txtUrl.ForeColor = $clrTextDisabled
    }
})

# Row 3: Output file
$lblOut           = New-FLabel "Report output file (.xlsx)" 16 196 $fntBold $clrText
$lblOut.AutoSize  = $false
$lblOut.Size      = New-Object System.Drawing.Size(500, 18)
$lblOutSub            = New-FLabel "Excel workbook with Dashboard, Issues list and Limits reference sheets. Existing file will be overwritten on re-scan." 16 216 $fntSmall $clrTextSub
$lblOutSub.AutoSize   = $false
$lblOutSub.Size       = New-Object System.Drawing.Size(856, 14)
$txtOut    = New-FTextBox 16 232 756 26
$btnOut    = New-FButton "  Browse..." 780 231 104 28
# Info banner
$pnlInfoBanner           = New-Object System.Windows.Forms.Panel
$pnlInfoBanner.Size      = New-Object System.Drawing.Size(868, 30)
$pnlInfoBanner.Location  = New-Object System.Drawing.Point(16, 278)
$pnlInfoBanner.BackColor = $clrAccentLight
$pnlInfoBanner.add_Paint({
    param($s,$e)
    $pen = New-Object System.Drawing.Pen($clrAccentMid, 1)
    $e.Graphics.DrawRectangle($pen, 0, 0, $s.Width-1, $s.Height-1)
    $pen.Dispose()
})
$lblInfoText          = New-FLabel "  i   All 19 rules enabled by default. Use the Rules panel below to skip individual checks. [B] = Blocking (upload fails)   [W] = Warning (issues expected in use)" 0 7 $fntSmall $clrAccentDark
$lblInfoText.AutoSize = $false
$lblInfoText.Size     = New-Object System.Drawing.Size(868, 18)
$pnlInfoBanner.Controls.Add($lblInfoText)

$pnlCard.Controls.AddRange(@(
    $pnlCardHead,
    $lblRoot, $lblRootSub, $txtRoot, $btnRoot,
    $lblUrl,  $lblUrlSub,  $txtUrl,
    $lblOut,  $lblOutSub,  $txtOut,  $btnOut,
    $pnlInfoBanner
))

# ── Rules panel ───────────────────────────────────────────────
# Layout inside panel (H=258):
#   Y=  0  pnlRulesHead (42px)
#   Y= 50  category headers row 1
#   Y= 66  checkbox rows (3x @ 22px = 66px)
#   Y=132  separator (14px)
#   Y=146  category headers row 2
#   Y=162  checkbox rows (3x @ 22px = 66px for chars+names, file rows fit in col2)
#   Y=248  bottom padding

$pnlRules           = New-Object System.Windows.Forms.Panel
$pnlRules.Size      = New-Object System.Drawing.Size(900, 232)
$pnlRules.Location  = New-Object System.Drawing.Point(20, 388)
$pnlRules.BackColor = $clrSurface
$pnlRules.add_Paint({
    param($s,$e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $pen1 = New-Object System.Drawing.Pen($clrBorder, 1)
    $g.DrawRectangle($pen1, 0, 0, $s.Width-1, $s.Height-1)
    $pen1.Dispose()
    $pen2 = New-Object System.Drawing.Pen($clrWarning, 3)
    $g.DrawLine($pen2, 0, 0, $s.Width, 0)
    $pen2.Dispose()
})

$pnlRulesHead           = New-Object System.Windows.Forms.Panel
$pnlRulesHead.Size      = New-Object System.Drawing.Size(900, 42)
$pnlRulesHead.Location  = New-Object System.Drawing.Point(0, 0)
$pnlRulesHead.BackColor = $clrWarningBg
$pnlRulesHead.add_Paint({
    param($s,$e)
    $pen = New-Object System.Drawing.Pen((hex '#f5d5b8'), 1)
    $e.Graphics.DrawLine($pen, 0, $s.Height-1, $s.Width, $s.Height-1)
    $pen.Dispose()
})
$lblRulesTitle          = New-FLabel "  [2]  Active Rules  -  Uncheck to skip individual checks before running the scan." 0 5 $fntBold $clrWarning
$lblRulesTitle.AutoSize = $false
$lblRulesTitle.Size     = New-Object System.Drawing.Size(900, 18)
$lblRulesSub            = New-FLabel "  [B] Blocking = item cannot be uploaded to SharePoint.     [W] Warning = upload succeeds but daily use may be affected." 0 24 $fntSmall $clrWarning
$lblRulesSub.AutoSize   = $false
$lblRulesSub.Size       = New-Object System.Drawing.Size(900, 16)
$pnlRulesHead.Controls.AddRange(@($lblRulesTitle, $lblRulesSub))

# ── Category card helper ──────────────────────────────────────
function New-CatPanel ([int]$x,[int]$y,[int]$w,[int]$h,[string]$title,[System.Drawing.Color]$bg,[System.Drawing.Color]$bdr,[System.Drawing.Color]$fg) {
    $p           = New-Object System.Windows.Forms.Panel
    $p.Size      = New-Object System.Drawing.Size($w, $h)
    $p.Location  = New-Object System.Drawing.Point($x, $y)
    $p.BackColor = $bg
    $p.Tag       = $bdr   # store border color in Tag - accessible in Paint handler via $s.Tag
    $p.add_Paint({
        param($s,$e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]$s.Tag, 1)
        $e.Graphics.DrawRectangle($pen, 0, 0, $s.Width-1, $s.Height-1)
        $pen.Dispose()
    })
    $lbl           = New-Object System.Windows.Forms.Label
    $lbl.Text      = "  $title"
    $lbl.Font      = $fntSmallB
    $lbl.ForeColor = $fg
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    $lbl.AutoSize  = $false
    $lbl.Size      = New-Object System.Drawing.Size($w, 18)
    $lbl.Location  = New-Object System.Drawing.Point(0, 3)
    $p.Controls.Add($lbl)
    $p
}

# Category headers - one wide header spans both columns of each group
# Left half  X=8   W=440  covers col1(X=8)   + col2(X=232)  = path / chars+files
# Right half X=456 W=436  covers col3(X=456)  + col4(X=680)  = name / prefix
$catPath = New-CatPanel   8 50 440 22 "PATH LENGTH CHECKS"    $clrCatPath $clrCatPathBdr $clrCatPathTxt
$catName = New-CatPanel 456 50 436 22 "NAME / PREFIX CHECKS"  $clrCatName $clrCatNameBdr $clrCatNameTxt

# Divider between group 1 and 2

# Row 2 headers - chars+files on left, name cont. on right
$catChar = New-CatPanel   8 138 220 22 "CHARACTER CHECKS"         $clrCatChar $clrCatCharBdr $clrCatCharTxt
$catFile = New-CatPanel 236 138 212 22 "FILE CHECKS"              $clrCatFile $clrCatFileBdr $clrCatFileTxt

# ── Checkbox factory ──────────────────────────────────────────
function New-RuleCheck ([string]$text,[int]$x,[int]$y,[string]$tag,[bool]$checked=$true) {
    $cb           = New-Object System.Windows.Forms.CheckBox
    $cb.Text      = $text
    $cb.Tag       = $tag
    $cb.Checked   = $checked
    $cb.Font      = $fntSmall
    $cb.ForeColor = $clrText
    $cb.BackColor = [System.Drawing.Color]::Transparent
    $cb.AutoSize  = $false
    $cb.Size      = New-Object System.Drawing.Size(216, 20)
    $cb.Location  = New-Object System.Drawing.Point($x, $y)
    $cb.FlatStyle = 'System'
    $cb.Add_CheckedChanged({
        if ($this.Checked) { [void]$script:EnabledRules.Add($this.Tag) }
        else               { [void]$script:EnabledRules.Remove($this.Tag) }
    })
    $cb
}

# Col 1  X=8    Col 2  X=232    Col 3  X=456    Col 4  X=680
# Row A  Y=74   Row B  Y=96     Row C  Y=118
# Row D  Y=172  Row E  Y=194    Row F  Y=216

# PATH LENGTH (col 1+2) — 3 rows at 20px spacing
$cbPathSP   = New-RuleCheck "SharePoint path >400 ch. [B]"       8  74 'PATH-SP'   $true
$cbPathOD   = New-RuleCheck "OneDrive sync path >400 [B]"        8  94 'PATH-OD'   $true
$cbPathWin  = New-RuleCheck "Windows MAX_PATH >260 [W]"          8 114 'PATH-WIN'  $true
$cbPathXL   = New-RuleCheck "Excel desktop >218 chars [W]"     232  74 'PATH-XL'   $true
$cbPathOff  = New-RuleCheck "Office desktop >259 chars [W]"    232  94 'PATH-OFF'  $true
$cbDepth    = New-RuleCheck "Folder nesting depth >10 [W]"     232 114 'DEPTH'     $true

# NAME / PREFIX (col 3+4) — 8 items at 16px spacing, no gap between row1 and cont.
$cbNameLen  = New-RuleCheck "Name segment >255 chars [B]"      456  74 'NAME-LEN'          $true
$cbNameTrim = New-RuleCheck "Leading/trailing spaces [B]"      456  90 'NAME-TRIM'         $true
$cbNameDot  = New-RuleCheck "Leading/trailing dot (.) [B]"     456 106 'NAME-DOT'          $true
$cbForms    = New-RuleCheck "Folder 'forms' at root [B]"       456 122 'NAME-FORMS'        $true
$cbTilde    = New-RuleCheck "~$ Office lock prefix [B]"        680  74 'NAME-TILDE'        $true
$cbVti      = New-RuleCheck "_vti_ reserved prefix [B]"        680  90 'NAME-VTI'          $true
$cbReserved = New-RuleCheck "Reserved names CON NUL [B]"       680 106 'NAME-RESERVED'     $true
$cbTildeF   = New-RuleCheck "Tilde (~) folder prefix [B]"      680 122 'NAME-TILDE-FOLDER' $true

# CHARACTERS (col 1+2) — below PATH rows, 20px spacing
$cbCharBlk  = New-RuleCheck "Blocked chars * : < > ? [B]"       8 164 'CHAR-BLOCKED' $true
$cbCharWarn = New-RuleCheck "Hash/percent (# %) [W]"             8 184 'CHAR-WARN'   $true
$cbSpaces   = New-RuleCheck "Spaces in names [W]"              232 164 'NAME-SPACES' $true

# FILES (col 1+2 continued)
$cbFileSize = New-RuleCheck "File size exceeds 250 GB [B]"     232 184 'FILE-SIZE'   $true
$cbFileTemp = New-RuleCheck "Temp files (.tmp .bak) [W]"       232 204 'FILE-TEMP'   $true

$pnlRules.Controls.AddRange(@(
    $pnlRulesHead,
    $catPath,$catName,
    $cbPathSP,$cbPathOD,$cbPathWin,$cbPathXL,$cbPathOff,$cbDepth,
    $cbNameLen,$cbNameTrim,$cbNameDot,$cbTilde,$cbVti,$cbForms,
    $catChar,$catFile,
    $cbCharBlk,$cbCharWarn,$cbSpaces,
    $cbFileSize,$cbFileTemp,
    $cbReserved,$cbTildeF
))

# ── Action row ────────────────────────────────────────────────
$pnlActions           = New-Object System.Windows.Forms.Panel
$pnlActions.Size      = New-Object System.Drawing.Size(900, 50)
$pnlActions.Location  = New-Object System.Drawing.Point(20, 628)
$pnlActions.BackColor = $clrAppBg

$btnRun    = New-FButton "  Run Analysis" 0   10 160 32 $true
$btnCancel = New-FButton "  Cancel"       168 10 110 32
$btnCancel.Enabled   = $false
$btnCancel.ForeColor = $clrTextDisabled
$btnCancel.FlatAppearance.BorderColor = $clrBorder

# Stat chips
function New-StatChip ([string]$label,[int]$x,[System.Drawing.Color]$dotColor,[System.Drawing.Color]$bg) {
    $panel           = New-Object System.Windows.Forms.Panel
    $panel.Size      = New-Object System.Drawing.Size(152, 32)
    $panel.Location  = New-Object System.Drawing.Point($x, 9)
    $panel.BackColor = $bg
    $panel.Tag       = $dotColor   # store border color in Tag - accessible in Paint handler via $s.Tag
    $panel.add_Paint({
        param($s,$e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]$s.Tag, 1)
        $e.Graphics.DrawRectangle($pen, 0, 0, $s.Width-1, $s.Height-1)
        $pen.Dispose()
    })
    $dot           = New-Object System.Windows.Forms.Panel
    $dot.Size      = New-Object System.Drawing.Size(8, 8)
    $dot.Location  = New-Object System.Drawing.Point(10, 12)
    $dot.BackColor = $dotColor
    $lbl           = New-Object System.Windows.Forms.Label
    $lbl.Text      = $label
    $lbl.Font      = $fntSmallB
    $lbl.ForeColor = $clrText
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    $lbl.AutoSize  = $false
    $lbl.Size      = New-Object System.Drawing.Size(130, 30)
    $lbl.Location  = New-Object System.Drawing.Point(24, 0)
    $lbl.TextAlign = "MiddleLeft"
    $panel.Controls.AddRange(@($dot,$lbl))
    @{ Panel=$panel; Label=$lbl }
}

$chipScanned  = New-StatChip "Scanned: -"   440 $clrBorderInput $clrSurface
$chipIssues   = New-StatChip "Issues: -"    600 $clrWarning     $clrWarningBg
$chipBlocking = New-StatChip "Blocking: -"  760 $clrError       $clrErrorBg

$pnlActions.Controls.AddRange(@(
    $btnRun, $btnCancel,
    $chipScanned.Panel, $chipIssues.Panel, $chipBlocking.Panel
))

# ── Progress area ─────────────────────────────────────────────
$pnlProgressArea           = New-Object System.Windows.Forms.Panel
$pnlProgressArea.Size      = New-Object System.Drawing.Size(900, 42)
$pnlProgressArea.Location  = New-Object System.Drawing.Point(20, 684)
$pnlProgressArea.BackColor = $clrAppBg

$progress          = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(0, 0)
$progress.Size     = New-Object System.Drawing.Size(900, 8)
$progress.Style    = 'Continuous'
$progress.Minimum  = 0
$progress.Maximum  = 1000
$progress.BackColor= $clrBorder
$progress.ForeColor= $clrAccent

$lblProgressText          = New-FLabel "Ready - select a folder and click Run Analysis." 0 12 $fntSmall $clrTextSub
$lblProgressText.AutoSize = $false
$lblProgressText.Size     = New-Object System.Drawing.Size(900, 18)
$pnlProgressArea.Controls.AddRange(@($progress, $lblProgressText))

# ── Divider ───────────────────────────────────────────────────
$div2 = New-HDivider 20 730 900

# ── Activity Log panel ────────────────────────────────────────
$pnlLog           = New-Object System.Windows.Forms.Panel
$pnlLog.Size      = New-Object System.Drawing.Size(900, 148)
$pnlLog.Location  = New-Object System.Drawing.Point(20, 734)
$pnlLog.BackColor = $clrSurface
$pnlLog.add_Paint({
    param($s,$e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $pen1 = New-Object System.Drawing.Pen($clrBorder, 1)
    $g.DrawRectangle($pen1, 0, 0, $s.Width-1, $s.Height-1)
    $pen1.Dispose()
    $pen2 = New-Object System.Drawing.Pen($clrSuccess, 3)
    $g.DrawLine($pen2, 0, 0, $s.Width, 0)
    $pen2.Dispose()
})

$pnlLogHead           = New-Object System.Windows.Forms.Panel
$pnlLogHead.Size      = New-Object System.Drawing.Size(900, 26)
$pnlLogHead.Location  = New-Object System.Drawing.Point(0, 0)
$pnlLogHead.BackColor = $clrSuccessBg
$pnlLogHead.add_Paint({
    param($s,$e)
    $pen = New-Object System.Drawing.Pen((hex '#a0ddb0'), 1)
    $e.Graphics.DrawLine($pen, 0, $s.Height-1, $s.Width, $s.Height-1)
    $pen.Dispose()
})
$lblLogHead          = New-FLabel "  [3]  Activity Log" 0 5 $fntBold $clrSuccess
$lblLogHead.AutoSize = $false
$lblLogHead.Size     = New-Object System.Drawing.Size(900, 18)
$pnlLogHead.Controls.Add($lblLogHead)

$rtLog             = New-Object System.Windows.Forms.RichTextBox
$rtLog.Location    = New-Object System.Drawing.Point(1, 28)
$rtLog.Size        = New-Object System.Drawing.Size(898, 118)
$rtLog.BackColor   = $clrSurface
$rtLog.ForeColor   = $clrText
$rtLog.Font        = $fntMono
$rtLog.ReadOnly    = $true
$rtLog.BorderStyle = 'None'
$rtLog.ScrollBars  = 'Vertical'
$rtLog.WordWrap    = $false
$pnlLog.Controls.AddRange(@($pnlLogHead, $rtLog))

# ── Status bar ────────────────────────────────────────────────
$pnlStatusBar           = New-Object System.Windows.Forms.Panel
$pnlStatusBar.Dock      = 'Bottom'
$pnlStatusBar.Size      = New-Object System.Drawing.Size(960, 26)
$pnlStatusBar.BackColor = $clrSurface
$pnlStatusBar.add_Paint({
    param($s,$e)
    $pen = New-Object System.Drawing.Pen($clrBorder, 1)
    $e.Graphics.DrawLine($pen, 0, 0, $s.Width, 0)
    $pen.Dispose()
})
$lblStatus = New-FLabel "  Ready" 0 5 $fntSmall $clrTextSub
$lblStatus.AutoSize = $false
$lblStatus.Size     = New-Object System.Drawing.Size(700, 18)
$pnlStatusBar.Controls.Add($lblStatus)

# ── Assemble form ─────────────────────────────────────────────
$form.Controls.AddRange(@(
    $pnlNav, $pnlStripe,
    $pnlCard,
    $pnlRules,
    $pnlActions,
    $pnlProgressArea,
    $div2,
    $pnlLog,
    $pnlStatusBar
))

# ==============================================================
#  GUI - TOOLTIPS
# ==============================================================

Set-Tip $txtRoot   "Enter or browse to the root folder on the file server. All subfolders and files will be scanned recursively."
Set-Tip $btnRoot   "Open folder browser to select the file server root folder."
Set-Tip $txtUrl    "Enter the full SharePoint Online library URL including the document library path. Used to calculate the decoded SharePoint path length for each item. Example: https://contoso.sharepoint.com/sites/IT/Shared%20Documents"
Set-Tip $txtOut    "Enter or browse to the output path for the Excel report (.xlsx). If the file already exists it will be overwritten."
Set-Tip $btnOut    "Open save dialog to choose the Excel report output file location."
Set-Tip $btnRun    "Start the migration readiness scan. All files and folders under the root folder will be checked against the active rules."
Set-Tip $btnCancel "Cancel the running scan. Items already scanned will appear in the report."
Set-Tip $cbPathSP   "PATH-SP [Blocking] SharePoint Online hard limit: 400 chars for the decoded server-relative path (library base + folders + filename). Items exceeding this cannot be uploaded."
Set-Tip $cbPathOD   "PATH-OD [Blocking] OneDrive sync client limit: 400 chars relative path. The local OneDrive root folder adds 80-120 chars on top - items may sync-fail even if under SP limit."
Set-Tip $cbPathWin  "PATH-WIN [Warning] Windows legacy MAX_PATH: 260 chars. Apps without long-path awareness fail above this. Mitigable via LongPathsEnabled registry key (Win10 1607+)."
Set-Tip $cbPathXL   "PATH-XL [Warning] Excel Win32 hard limit: 218 chars (KB 325573). Cannot be changed. Applies when files are opened from a synced OneDrive folder."
Set-Tip $cbPathOff  "PATH-OFF [Warning] Word / PowerPoint / Access Win32 limit: 259 chars (KB 325573). Applies when opening synced files from a local OneDrive or mapped SharePoint folder."
Set-Tip $cbDepth    "DEPTH [Warning] Microsoft recommends max 10 folder levels. Deeper nesting increases path-length risk, reduces usability and degrades SharePoint search performance."
Set-Tip $cbNameLen  "NAME-LEN [Blocking] Each individual name segment is limited to 255 chars in SharePoint Online. Items exceeding this cannot be created or synced."
Set-Tip $cbNameTrim "NAME-TRIM [Blocking] SharePoint silently strips leading/trailing spaces during upload, causing name mismatches and sync conflicts."
Set-Tip $cbNameDot  "NAME-DOT [Blocking] SharePoint does not support names beginning or ending with a dot. Items like .gitignore or folder. cannot be uploaded."
Set-Tip $cbTilde    "NAME-TILDE [Blocking] Names starting with ~$ are Office temporary lock files (created while a document is open). These are blocked by OneDrive and SharePoint."
Set-Tip $cbVti      "NAME-VTI [Blocking] Names starting with _vti_ are reserved for SharePoint internal system use and are blocked in all SharePoint versions."
Set-Tip $cbForms    "NAME-FORMS [Blocking] SharePoint uses a forms folder internally at root of every library. A user folder with this name at root level conflicts with SharePoint internals."
Set-Tip $cbCharBlk  "CHAR-BLOCKED [Blocking] Characters not allowed in SharePoint / OneDrive names: `" * : < > ? / \ |. Items containing these cannot be uploaded."
Set-Tip $cbCharWarn "CHAR-WARN [Warning] # and % are officially supported in SharePoint Online since 2017 but cause problems with older Office desktop apps (pre-2016) and some migration tools."
Set-Tip $cbSpaces   "NAME-SPACES [Warning] Spaces are encoded as %20 in URLs (+2 chars each). Paths that look short can exceed limits when encoded. Best practice: use underscores or hyphens."
Set-Tip $cbFileSize "FILE-SIZE [Blocking] SharePoint Online and OneDrive max single-file upload size: 250 GB. Files exceeding this cannot be uploaded by any method."
Set-Tip $cbFileTemp "FILE-TEMP [Warning] Temp/system files (.tmp .bak .temp .ds_store .thumbs.db) have no business value in SharePoint. Some are silently dropped or blocked by OneDrive."
Set-Tip $cbReserved "NAME-RESERVED [Blocking] Windows device names (CON PRN AUX NUL COM0-9 LPT0-9) and system files (desktop.ini .lock) cannot be used as file or folder names."
Set-Tip $cbTildeF   "NAME-TILDE-FOLDER [Blocking] SharePoint does not allow folder names beginning with a tilde (~) character."

# ==============================================================
#  GUI - LOG & STATUS HELPERS
# ==============================================================

function Write-Log ([string]$msg,[string]$level='info') {
    $ts  = (Get-Date).ToString("HH:mm:ss")
    $col = switch ($level) {
        'ok'    { $clrSuccess }
        'warn'  { $clrWarning }
        'error' { $clrError   }
        'blue'  { $clrAccent  }
        default { $clrTextSub }
    }
    $rtLog.SelectionStart  = $rtLog.TextLength
    $rtLog.SelectionLength = 0
    $rtLog.SelectionColor  = $clrTextDisabled
    $rtLog.AppendText("[${ts}]  ")
    $rtLog.SelectionColor  = $col
    $rtLog.AppendText("$msg`n")
    $rtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Status ([string]$msg) {
    $lblStatus.Text = $msg
    [System.Windows.Forms.Application]::DoEvents()
}

# ==============================================================
#  GUI - EVENTS
# ==============================================================

$btnRoot.Add_Click({
    $dlg             = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select the root folder of the file server structure to analyze"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtRoot.Text      = $dlg.SelectedPath
        $txtRoot.ForeColor = $clrText
    }
})

$btnOut.Add_Click({
    $dlg          = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title    = "Save analysis report"
    $dlg.Filter   = "Excel Workbook (*.xlsx)|*.xlsx"
    $dlg.FileName = "SharePointAnalysis_{0:yyyyMMdd_HHmmss}.xlsx" -f (Get-Date)
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOut.Text      = $dlg.FileName
        $txtOut.ForeColor = $clrText
    }
})

$script:CancelRequested = $false
$btnCancel.Add_Click({
    $script:CancelRequested  = $true
    $btnCancel.Enabled       = $false
    $btnCancel.ForeColor     = $clrTextDisabled
    Write-Log "Cancel requested - stopping after current item..." 'warn'
    Set-Status "Cancelling..."
})

$btnRun.Add_Click({

    $rootPath  = $txtRoot.Text.Trim()
    $urlPrefix = $txtUrl.Text.Trim()
    $output    = $txtOut.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($rootPath) -or -not (Test-Path -LiteralPath $rootPath)) {
        Show-MsgWarn "Please select a valid file server root folder." "Input Required"; return
    }
    if ([string]::IsNullOrWhiteSpace($urlPrefix) -or $urlPrefix -eq $placeholder) {
        Show-MsgWarn "Please enter the target SharePoint library URL prefix." "Input Required"; return
    }
    try { $null = [System.Uri]$urlPrefix } catch {
        Show-MsgWarn "The SharePoint URL is not a valid URL." "Input Error"; return
    }
    if ([string]::IsNullOrWhiteSpace($output)) {
        $ts      = Get-Date -Format "yyyyMMdd_HHmmss"
        $output  = Join-Path ([Environment]::GetFolderPath("Desktop")) "SharePointAnalysis_$ts.xlsx"
        $txtOut.Text      = $output
        $txtOut.ForeColor = $clrText
    }
    if ([System.IO.Path]::GetExtension($output) -ne ".xlsx") {
        $output      = [System.IO.Path]::ChangeExtension($output, ".xlsx")
        $txtOut.Text = $output
    }
    if (-not (Initialize-ImportExcel)) { return }

    $script:CancelRequested  = $false
    $btnRun.Enabled          = $false
    $btnCancel.Enabled       = $true
    $btnCancel.ForeColor     = $clrText
    $progress.Value          = 0
    $chipScanned.Label.Text  = "Scanned: -"
    $chipIssues.Label.Text   = "Issues: -"
    $chipBlocking.Label.Text = "Blocking: -"
    $rtLog.Clear()
    Set-Status "Collecting items..."
    Write-Log "Analysis started" 'blue'
    Write-Log "Root  : $rootPath"
    Write-Log "URL   : $urlPrefix"
    Write-Log "Output: $output"

    try {
        $items = Get-ChildItem -LiteralPath $rootPath -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
        Show-MsgError "Failed to enumerate items:`n$($_.Exception.Message)" "Error"
        $btnRun.Enabled      = $true
        $btnCancel.Enabled   = $false
        $btnCancel.ForeColor = $clrTextDisabled
        return
    }

    $total     = $items.Count
    $scanDate  = Get-Date
    $startTime = $scanDate

    Write-Log "Found $total items. Checking against SharePoint rules..." 'blue'

    if ($total -eq 0) {
        Show-MsgInfo "No items found in the selected folder." "Analysis Complete"
        Set-Status "Ready"
        $btnRun.Enabled      = $true
        $btnCancel.Enabled   = $false
        $btnCancel.ForeColor = $clrTextDisabled
        return
    }

    $results       = [System.Collections.Generic.List[object]]::new()
    $index         = 0
    $errCount      = 0
    $blockingCount = 0

    foreach ($item in $items) {
        if ($script:CancelRequested) { Write-Log "Scan cancelled after $index items." 'warn'; break }
        $index++

        try {
            $result = Test-SharePointItem -Item $item -RootPath $rootPath -UrlPrefix $urlPrefix
            if ($result.BlockingIssue) {
                [void]$results.Add($result)
                if ($result.Severity -eq 'Blocking') { $blockingCount++ }
                if ($results.Count -le 60) {
                    $icon = if ($result.Severity -eq 'Blocking') { '[BLOCKING]' } else { '[WARNING] ' }
                    $lvl  = if ($result.Severity -eq 'Blocking') { 'error' }     else { 'warn' }
                    Write-Log "$icon  $($result.RelativePath)" $lvl
                } elseif ($results.Count -eq 61) {
                    Write-Log "...log capped at 60 entries - all issues captured in the report." 'info'
                }
            }
        } catch {
            $errCount++
            if ($errCount -le 10) { Write-Log "ERROR on '$($item.FullName)': $($_.Exception.Message)" 'error' }
        }

        if ($index % 50 -eq 0 -or $index -eq $total) {
            $el  = (Get-Date) - $startTime
            $eta = if ($index -gt 1 -and $el.TotalSeconds -gt 0.5) {
                [TimeSpan]::FromSeconds($el.TotalSeconds / $index * ($total - $index)) |
                    ForEach-Object { $_.ToString("hh\:mm\:ss") }
            } else { '...' }
            $progress.Value              = [Math]::Min([Math]::Max([int](($index/$total)*1000),0),1000)
            $chipScanned.Label.Text      = "Scanned: $index"
            $chipIssues.Label.Text       = "Issues: $($results.Count)"
            $chipBlocking.Label.Text     = "Blocking: $blockingCount"
            $lblProgressText.Text        = "Scanning $index / $total  |  Elapsed: $($el.ToString('hh\:mm\:ss'))  |  ETA: $eta  |  Issues: $($results.Count)"
            Set-Status "Scanning...  ($index / $total)"
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    $duration    = (Get-Date) - $startTime
    $durationStr = $duration.ToString("hh\:mm\:ss")
    $progress.Value = 1000

    if ($results.Count -eq 0) {
        Write-Log "All $index items passed - no compatibility issues found." 'ok'
        $lblProgressText.Text = "Completed in $durationStr - all items are SharePoint-compatible."
        Set-Status "Done - no issues found."
        Show-MsgInfo "No compatibility issues found.`n`nScanned:  $index items`nDuration: $durationStr" "Analysis Complete"
        $btnRun.Enabled      = $true
        $btnCancel.Enabled   = $false
        $btnCancel.ForeColor = $clrTextDisabled
        $progress.Value      = 0
        return
    }

    Write-Log "$($results.Count) items with issues ($blockingCount blocking). Exporting Excel report..." 'blue'
    $lblProgressText.Text = "Exporting report..."
    Set-Status "Exporting report..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        Export-AnalysisReport `
            -Results      $results    -OutputPath  $output `
            -RootPath     $rootPath   -UrlPrefix   $urlPrefix `
            -ScanDate     $scanDate   -TotalScanned $index `
            -Duration     $duration

        Write-Log "Report saved: $output" 'ok'
        $lblProgressText.Text = "Completed in $durationStr  |  Report: $output"
        Set-Status "Done - $($results.Count) issues ($blockingCount blocking)."

        Show-MsgInfo ("Analysis complete.`n`n" +
            "  Scanned:   $index items`n" +
            "  Issues:    $($results.Count) items`n" +
            "  Blocking:  $blockingCount items`n" +
            "  Duration:  $durationStr`n`n" +
            "Report saved to:`n$output") "Analysis Complete"

    } catch {
        Write-Log "Export error: $($_.Exception.Message)" 'error'
        Show-MsgError "Failed to export report:`n$($_.Exception.Message)" "Export Error"
    }

    $btnRun.Enabled      = $true
    $btnCancel.Enabled   = $false
    $btnCancel.ForeColor = $clrTextDisabled
    $progress.Value      = 0
})

# ==============================================================
#  LAUNCH
# ==============================================================

[void]$form.ShowDialog()