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


# Rules toggle - all enabled by default; GUI checkboxes update this set
$script:EnabledRules = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        'PATH-SP','PATH-OD','PATH-WIN','PATH-XL','PATH-OFF',
        'FILE-SIZE','FILE-TEMP','DEPTH',
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
    #  SHEET 2 — ISSUES (detailed table)
    # ══════════════════════════════════════════════════════════

    # Select-Object on one line (PS 5.1 parser requirement)
    $detailData = $Results | Select-Object Severity, ItemType, Name, RelativePath, FullPath, WindowsFullPathLength, DecodedSharePointPathLength, FolderDepth, IssueCount, IssueSummary, IssueDetails, RecommendedActions

    $pkg = $detailData | Export-Excel -ExcelPackage $pkg `
        -WorksheetName "Issues" -TableName "IssueDetails" `
        -TableStyle "Medium2" -AutoFilter -BoldTopRow -FreezeTopRow `
        -PassThru

    $wsIss     = $pkg.Workbook.Worksheets["Issues"]
    $lastRow   = $wsIss.Dimension.End.Row
    $lastCol   = $wsIss.Dimension.End.Column
    $headerRow = 1

    # ── Style header row ─────────────────────────────────────────
    for ($c = 1; $c -le $lastCol; $c++) {
        $cell = $wsIss.Cells[$headerRow,$c]
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($xlBlueDark)
        $cell.Style.Font.Color.SetColor($xlWhite)
        $cell.Style.Font.Bold = $true
        $cell.Style.Font.Size = 10
        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    }
    $wsIss.Row($headerRow).Height = 20

    # ── Build column map ─────────────────────────────────────────
    $colMap = @{}
    for ($c = 1; $c -le $lastCol; $c++) {
        $h = $wsIss.Cells[$headerRow,$c].Text
        if (-not [string]::IsNullOrEmpty($h)) { $colMap[$h] = $c }
    }
    $dataStart = $headerRow + 1

    # ── Row-level background highlight by severity ───────────────
    for ($r = $dataStart; $r -le $lastRow; $r++) {
        $sev = $wsIss.Cells[$r, $colMap['Severity']].Text
        $rowBg = switch ($sev) {
            'Blocking' { $xlRedLight    }
            'Warning'  { $xlOrangeLight }
            default    { $xlWhite       }
        }
        for ($c = 1; $c -le $lastCol; $c++) {
            $cell = $wsIss.Cells[$r,$c]
            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cell.Style.Fill.BackgroundColor.SetColor($rowBg)
        }
        $wsIss.Row($r).Height = 15
    }

    # ── Bold + colour the Severity cell text ────────────────────
    if ($colMap['Severity']) {
        for ($r = $dataStart; $r -le $lastRow; $r++) {
            $cell = $wsIss.Cells[$r, $colMap['Severity']]
            $sev  = $cell.Text
            $cell.Style.Font.Bold = $true
            switch ($sev) {
                'Blocking' { $cell.Style.Font.Color.SetColor($xlRed)    }
                'Warning'  { $cell.Style.Font.Color.SetColor($xlOrange) }
            }
        }
    }

    # ── Conditional formatting: over-limit path lengths ──────────
    foreach ($entry in @(
        @{ Col='WindowsFullPathLength';       Limit=$script:Config.MaxWindowsLegacyPathLength     }
        @{ Col='DecodedSharePointPathLength'; Limit=$script:Config.MaxDecodedSharePointPathLength }
        @{ Col='DecodedSharePointPathLength'; Limit=$script:Config.MaxDecodedSharePointPathLength }
    )) {
        if ($colMap[$entry.Col]) {
            $col = ColLetter $colMap[$entry.Col]
            Add-ConditionalFormatting -Worksheet $wsIss `
                -Address "${col}${dataStart}:${col}${lastRow}" `
                -RuleType GreaterThan -ConditionValue $entry.Limit `
                -ForeGroundColor $xlRed -Bold
        }
    }

    # ── Column widths (Issues sheet) ────────────────────────────
    $colWidths = @{
        'Severity'                    = 11
        'ItemType'                    = 9
        'Name'                        = 36
        'RelativePath'                = 60
        'FullPath'                    = 70
        'WindowsFullPathLength'       = 14
        'DecodedSharePointPathLength' = 14
        'FolderDepth'                 = 11
        'IssueCount'                  = 9
        'IssueSummary'                = 55

        'IssueDetails'                = 90

        'RecommendedActions'          = 80
    }
    foreach ($kv in $colWidths.GetEnumerator()) {
        if ($colMap[$kv.Key]) {
            $wsIss.Column($colMap[$kv.Key]).Width = $kv.Value
        }
    }

    # Wrap text for multi-line columns
    foreach ($wrapCol in @('IssueSummary','IssueDetails','RecommendedActions')) {
        if ($colMap[$wrapCol]) {
            $col = ColLetter $colMap[$wrapCol]
            $wsIss.Cells["${col}${dataStart}:${col}${lastRow}"].Style.WrapText = $true
        }
    }

    # Freeze header row
    $wsIss.View.FreezePanes($dataStart, 1)

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
    $pkg.Workbook.Worksheets.MoveToStart("Issues")
    $pkg.Workbook.Worksheets.MoveToStart("Dashboard")

    # ── Save ────────────────────────────────────────────────────
    $pkg.Save()
    $pkg.Dispose()
}

# ==============================================================
#  GUI - FLUENT 2 / M365 COLOUR TOKENS
# ==============================================================

function hex ([string]$h) { [System.Drawing.ColorTranslator]::FromHtml($h) }

$clrAppBg        = hex '#f3f2f1'
$clrSurface      = hex '#ffffff'
$clrSurfaceAlt   = hex '#faf9f8'
$clrBorder       = hex '#edebe9'
$clrBorderInput  = hex '#c8c6c4'
$clrText         = hex '#323130'
$clrTextSub      = hex '#605e5c'
$clrTextDisabled = hex '#a19f9d'
$clrAccent       = hex '#0078d4'
$clrAccentHover  = hex '#106ebe'
$clrAccentLight  = hex '#deecf9'
$clrOnAccent     = hex '#ffffff'
$clrSuccess      = hex '#107c10'
$clrWarning      = hex '#d83b01'
$clrError        = hex '#a4262c'

# ==============================================================
#  GUI - TYPOGRAPHY
# ==============================================================

$fntH2      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$fntBody    = New-Object System.Drawing.Font("Segoe UI",  9, [System.Drawing.FontStyle]::Regular)
$fntBold    = New-Object System.Drawing.Font("Segoe UI",  9, [System.Drawing.FontStyle]::Bold)
$fntSmall   = New-Object System.Drawing.Font("Segoe UI",  8, [System.Drawing.FontStyle]::Regular)
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
    $l            = New-Object System.Windows.Forms.Label
    $l.Text       = $Text
    $l.AutoSize   = $true
    $l.Location   = New-Object System.Drawing.Point($X,$Y)
    $l.Font       = $Font
    $l.ForeColor  = $Color
    $l.BackColor  = [System.Drawing.Color]::Transparent
    $l
}

function New-FTextBox {
    param([int]$X,[int]$Y,[int]$W,[int]$H=28)
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
    $b          = New-Object System.Windows.Forms.Button
    $b.Text     = $Text
    $b.Location = New-Object System.Drawing.Point($X,$Y)
    $b.Size     = New-Object System.Drawing.Size($W,$H)
    $b.Font     = $fntBold
    $b.FlatStyle= 'Flat'
    $b.Cursor   = [System.Windows.Forms.Cursors]::Hand
    if ($Primary) {
        $b.BackColor                    = $clrAccent
        $b.ForeColor                    = $clrOnAccent
        $b.FlatAppearance.BorderSize    = 0
        $b.Add_MouseEnter({ $this.BackColor = $clrAccentHover })
        $b.Add_MouseLeave({ $this.BackColor = $clrAccent })
    } else {
        $b.BackColor                    = $clrSurface
        $b.ForeColor                    = $clrText
        $b.FlatAppearance.BorderSize    = 1
        $b.FlatAppearance.BorderColor   = $clrBorderInput
        $b.Add_MouseEnter({ $this.BackColor = $clrSurfaceAlt })
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
#  GUI - MAIN FORM
# ==============================================================

$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "SharePoint Migration Analyzer"
$form.StartPosition   = "CenterScreen"
$form.Size            = New-Object System.Drawing.Size(880, 780)
$form.MinimumSize     = New-Object System.Drawing.Size(880, 780)
$form.MaximizeBox     = $false
$form.BackColor       = $clrAppBg
$form.ForeColor       = $clrText
$form.Font            = $fntBody
$form.FormBorderStyle = "FixedSingle"

# Navigation bar
$pnlNav           = New-Object System.Windows.Forms.Panel
$pnlNav.Size      = New-Object System.Drawing.Size(880, 56)
$pnlNav.Location  = New-Object System.Drawing.Point(0, 0)
$pnlNav.BackColor = $clrSurface

# SP icon tile
$pnlIcon           = New-Object System.Windows.Forms.Panel
$pnlIcon.Size      = New-Object System.Drawing.Size(32, 32)
$pnlIcon.Location  = New-Object System.Drawing.Point(16, 12)
$pnlIcon.BackColor = $clrAccent
$pnlIcon.Add_Paint({
    param($s,$e)
    $g   = $e.Graphics
    $g.SmoothingMode        = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint    = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $fnt = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $br  = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $sf  = New-Object System.Drawing.StringFormat
    $sf.Alignment     = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString("SP", $fnt, $br, (New-Object System.Drawing.RectangleF(0,0,32,32)), $sf)
    $fnt.Dispose(); $br.Dispose(); $sf.Dispose()
})

$lblAppTitle = New-FLabel "SharePoint Migration Analyzer" 58 9  $fntH2    $clrText
$lblAppSub   = New-FLabel "File server path compatibility check for SharePoint Online" 60 31 $fntSmall $clrTextSub
$pnlNav.Controls.AddRange(@($pnlIcon, $lblAppTitle, $lblAppSub))

# Blue accent stripe
$pnlStripe           = New-Object System.Windows.Forms.Panel
$pnlStripe.Size      = New-Object System.Drawing.Size(880, 3)
$pnlStripe.Location  = New-Object System.Drawing.Point(0, 56)
$pnlStripe.BackColor = $clrAccent

# Configuration card
$pnlCard           = New-Object System.Windows.Forms.Panel
$pnlCard.Size      = New-Object System.Drawing.Size(840, 238)
$pnlCard.Location  = New-Object System.Drawing.Point(20, 72)
$pnlCard.BackColor = $clrSurface
$pnlCard.add_Paint({
    param($s,$e)
    $pen  = New-Object System.Drawing.Pen($clrBorder, 1)
    $rect = New-Object System.Drawing.Rectangle(0, 0, ($s.Width-1), ($s.Height-1))
    $e.Graphics.DrawRectangle($pen, $rect)
    $pen.Dispose()
})

$pnlCardHead           = New-Object System.Windows.Forms.Panel
$pnlCardHead.Size      = New-Object System.Drawing.Size(840, 30)
$pnlCardHead.Location  = New-Object System.Drawing.Point(0, 0)
$pnlCardHead.BackColor = $clrAccentLight

$lblCardTitle          = New-FLabel "  Scan Configuration" 0 7 $fntBold $clrAccent
$lblCardTitle.AutoSize = $false
$lblCardTitle.Size     = New-Object System.Drawing.Size(840, 20)
$pnlCardHead.Controls.Add($lblCardTitle)

# Row 1: Root folder
$lblRoot = New-FLabel "File server root folder" 14 44 $fntBold $clrText
$txtRoot = New-FTextBox 14 63 698
$btnRoot = New-FButton "Browse..." 718 62 108 28

# Row 2: SharePoint URL
$lblUrl      = New-FLabel "SharePoint library URL prefix" 14 102 $fntBold $clrText
$txtUrl      = New-FTextBox 14 121 812
$placeholder = "https://tenant.sharepoint.com/sites/YourSite/Shared%20Documents"
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
$lblOut = New-FLabel "Report output file (.xlsx)" 14 160 $fntBold $clrText
$txtOut = New-FTextBox 14 179 698
$btnOut = New-FButton "Browse..." 718 178 108 28

# Info banner
$pnlInfoBanner           = New-Object System.Windows.Forms.Panel
$pnlInfoBanner.Size      = New-Object System.Drawing.Size(812, 22)
$pnlInfoBanner.Location  = New-Object System.Drawing.Point(14, 213)
$pnlInfoBanner.BackColor = $clrAccentLight

$lblInfoText          = New-FLabel "  Checks: blocked chars (* : < > ? | ...), reserved names, _vti_ / ~$ prefixes, path lengths, Excel 218 limit, folder depth, temp files" 0 3 $fntSmall $clrAccent
$lblInfoText.AutoSize = $false
$lblInfoText.Size     = New-Object System.Drawing.Size(812, 18)
$pnlInfoBanner.Controls.Add($lblInfoText)

$pnlCard.Controls.AddRange(@(
    $pnlCardHead,
    $lblRoot, $txtRoot, $btnRoot,
    $lblUrl,  $txtUrl,
    $lblOut,  $txtOut,  $btnOut,
    $pnlInfoBanner
))

# Rules panel
$pnlRules           = New-Object System.Windows.Forms.Panel
$pnlRules.Size      = New-Object System.Drawing.Size(840, 140)
$pnlRules.Location  = New-Object System.Drawing.Point(20, 318)
$pnlRules.BackColor = $clrSurface
$pnlRules.add_Paint({
    param($s,$e)
    $pen  = New-Object System.Drawing.Pen($clrBorder, 1)
    $rect = New-Object System.Drawing.Rectangle(0, 0, ($s.Width-1), ($s.Height-1))
    $e.Graphics.DrawRectangle($pen, $rect)
    $pen.Dispose()
})

$pnlRulesHead           = New-Object System.Windows.Forms.Panel
$pnlRulesHead.Size      = New-Object System.Drawing.Size(840, 24)
$pnlRulesHead.Location  = New-Object System.Drawing.Point(0, 0)
$pnlRulesHead.BackColor = $clrAccentLight

$lblRulesTitle          = New-FLabel "  Active Rules  (uncheck to skip a check)" 0 5 $fntBold $clrAccent
$lblRulesTitle.AutoSize = $false
$lblRulesTitle.Size     = New-Object System.Drawing.Size(840, 18)
$pnlRulesHead.Controls.Add($lblRulesTitle)

function New-RuleCheck ([string]$text,[int]$x,[int]$y,[string]$tag) {
    $cb           = New-Object System.Windows.Forms.CheckBox
    $cb.Text      = $text
    $cb.Tag       = $tag
    $cb.Checked   = $true
    $cb.Font      = $fntSmall
    $cb.ForeColor = $clrText
    $cb.BackColor = [System.Drawing.Color]::Transparent
    $cb.AutoSize  = $false
    $cb.Size      = New-Object System.Drawing.Size(196, 18)
    $cb.Location  = New-Object System.Drawing.Point($x, $y)
    $cb.FlatStyle = 'System'
    $cb.Add_CheckedChanged({
        if ($this.Checked) { [void]$script:EnabledRules.Add($this.Tag) }
        else               { [void]$script:EnabledRules.Remove($this.Tag) }
    })
    $cb
}

$cbPathSP   = New-RuleCheck "SP path >400 chars [B]"         8  32 'PATH-SP'
$cbPathOD   = New-RuleCheck "OneDrive path >400 [B]"       208  32 'PATH-OD'
$cbPathWin  = New-RuleCheck "Windows MAX_PATH [W]"          408  32 'PATH-WIN'
$cbPathXL   = New-RuleCheck "Excel 218 limit [W]"           608  32 'PATH-XL'
$cbPathOff  = New-RuleCheck "Office 259 limit [W]"            8  54 'PATH-OFF'
$cbDepth    = New-RuleCheck "Folder depth >10 [W]"          208  54 'DEPTH'
$cbFileSize = New-RuleCheck "File size >250 GB [B]"         408  54 'FILE-SIZE'
$cbFileTemp = New-RuleCheck "Temp file extensions [W]"      608  54 'FILE-TEMP'
$cbNameLen  = New-RuleCheck "Name >255 chars [B]"             8  76 'NAME-LEN'
$cbNameTrim = New-RuleCheck "Lead/trail spaces [B]"         208  76 'NAME-TRIM'
$cbNameDot  = New-RuleCheck "Lead/trail dot [B]"            408  76 'NAME-DOT'
$cbTilde    = New-RuleCheck "~dollar lock prefix [B]"       608  76 'NAME-TILDE'
$cbVti      = New-RuleCheck "_vti_ prefix [B]"                8  98 'NAME-VTI'
$cbForms    = New-RuleCheck "forms folder [B]"              208  98 'NAME-FORMS'
$cbReserved = New-RuleCheck "Reserved names [B]"            408  98 'NAME-RESERVED'
$cbTildeF   = New-RuleCheck "Tilde folder prefix [B]"       608  98 'NAME-TILDE-FOLDER'
$cbCharBlk  = New-RuleCheck "Blocked chars *:<>? [B]"         8 118 'CHAR-BLOCKED'
$cbCharWarn = New-RuleCheck "Hash/percent chars [W]"        208 118 'CHAR-WARN'
$cbSpaces   = New-RuleCheck "Spaces in names [W]"           408 118 'NAME-SPACES'

$pnlRules.Controls.AddRange(@(
    $pnlRulesHead,
    $cbPathSP,$cbPathOD,$cbPathWin,$cbPathXL,
    $cbPathOff,$cbDepth,$cbFileSize,$cbFileTemp,
    $cbNameLen,$cbNameTrim,$cbNameDot,$cbTilde,
    $cbVti,$cbForms,$cbReserved,$cbTildeF,
    $cbCharBlk,$cbCharWarn,$cbSpaces
))


# Action row
$pnlActions           = New-Object System.Windows.Forms.Panel
$pnlActions.Size      = New-Object System.Drawing.Size(840, 46)
$pnlActions.Location  = New-Object System.Drawing.Point(20, 465)
$pnlActions.BackColor = $clrAppBg

$btnRun    = New-FButton "> Run Analysis" 0   8 154 30 $true
$btnCancel = New-FButton "x Cancel"       162 8 106 30
$btnCancel.Enabled   = $false
$btnCancel.ForeColor = $clrTextDisabled
$btnCancel.FlatAppearance.BorderColor = $clrBorder

# Stat chips
function New-StatChip ([string]$label,[int]$x,[System.Drawing.Color]$dotColor) {
    $panel           = New-Object System.Windows.Forms.Panel
    $panel.Size      = New-Object System.Drawing.Size(138, 30)
    $panel.Location  = New-Object System.Drawing.Point($x, 8)
    $panel.BackColor = $clrSurface
    $panel.add_Paint({
        param($s,$e)
        $pen = New-Object System.Drawing.Pen($clrBorder,1)
        $e.Graphics.DrawRectangle($pen,0,0,$s.Width-1,$s.Height-1)
        $pen.Dispose()
    })
    $dot           = New-Object System.Windows.Forms.Panel
    $dot.Size      = New-Object System.Drawing.Size(8,8)
    $dot.Location  = New-Object System.Drawing.Point(10,11)
    $dot.BackColor = $dotColor
    $lbl           = New-Object System.Windows.Forms.Label
    $lbl.Text      = $label
    $lbl.Font      = $fntSmall
    $lbl.ForeColor = $clrText
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    $lbl.AutoSize  = $false
    $lbl.Size      = New-Object System.Drawing.Size(116,28)
    $lbl.Location  = New-Object System.Drawing.Point(22,0)
    $lbl.TextAlign = "MiddleLeft"
    $panel.Controls.AddRange(@($dot,$lbl))
    @{ Panel=$panel; Label=$lbl }
}

$chipScanned  = New-StatChip "Scanned: -"  430 $clrTextSub
$chipIssues   = New-StatChip "Issues: -"   576 $clrWarning
$chipBlocking = New-StatChip "Blocking: -" 722 $clrError

$pnlActions.Controls.AddRange(@(
    $btnRun, $btnCancel,
    $chipScanned.Panel, $chipIssues.Panel, $chipBlocking.Panel
))

# Progress area
$pnlProgressArea           = New-Object System.Windows.Forms.Panel
$pnlProgressArea.Size      = New-Object System.Drawing.Size(840, 38)
$pnlProgressArea.Location  = New-Object System.Drawing.Point(20, 517)
$pnlProgressArea.BackColor = $clrAppBg

$progress          = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(0, 0)
$progress.Size     = New-Object System.Drawing.Size(840, 10)
$progress.Style    = 'Continuous'
$progress.Minimum  = 0
$progress.Maximum  = 1000
$progress.BackColor= $clrBorder
$progress.ForeColor= $clrAccent

$lblProgressText          = New-FLabel "Ready - select a folder and click Run Analysis." 0 14 $fntSmall $clrTextSub
$lblProgressText.AutoSize = $false
$lblProgressText.Size     = New-Object System.Drawing.Size(840, 18)

$pnlProgressArea.Controls.AddRange(@($progress, $lblProgressText))

# Divider
$div2 = New-HDivider 20 561 840

# Log panel
$pnlLog           = New-Object System.Windows.Forms.Panel
$pnlLog.Size      = New-Object System.Drawing.Size(840, 168)
$pnlLog.Location  = New-Object System.Drawing.Point(20, 565)
$pnlLog.BackColor = $clrSurface
$pnlLog.add_Paint({
    param($s,$e)
    $pen = New-Object System.Drawing.Pen($clrBorder,1)
    $e.Graphics.DrawRectangle($pen,0,0,$s.Width-1,$s.Height-1)
    $pen.Dispose()
})

$pnlLogHead           = New-Object System.Windows.Forms.Panel
$pnlLogHead.Size      = New-Object System.Drawing.Size(840, 24)
$pnlLogHead.Location  = New-Object System.Drawing.Point(0, 0)
$pnlLogHead.BackColor = $clrSurfaceAlt
$pnlLogHead.add_Paint({
    param($s,$e)
    $pen = New-Object System.Drawing.Pen($clrBorder,1)
    $e.Graphics.DrawLine($pen,0,$s.Height-1,$s.Width,$s.Height-1)
    $pen.Dispose()
})

$lblLogHead          = New-FLabel "  Activity Log" 0 5 $fntCaption $clrTextSub
$lblLogHead.AutoSize = $false
$lblLogHead.Size     = New-Object System.Drawing.Size(840,18)
$pnlLogHead.Controls.Add($lblLogHead)

$rtLog             = New-Object System.Windows.Forms.RichTextBox
$rtLog.Location    = New-Object System.Drawing.Point(1, 25)
$rtLog.Size        = New-Object System.Drawing.Size(838, 142)
$rtLog.BackColor   = $clrSurface
$rtLog.ForeColor   = $clrText
$rtLog.Font        = $fntMono
$rtLog.ReadOnly    = $true
$rtLog.BorderStyle = 'None'
$rtLog.ScrollBars  = 'Vertical'
$rtLog.WordWrap    = $false

$pnlLog.Controls.AddRange(@($pnlLogHead, $rtLog))

# Status bar
$pnlStatusBar           = New-Object System.Windows.Forms.Panel
$pnlStatusBar.Dock      = 'Bottom'
$pnlStatusBar.Size      = New-Object System.Drawing.Size(880, 24)
$pnlStatusBar.BackColor = $clrSurface
$pnlStatusBar.add_Paint({
    param($s,$e)
    $pen = New-Object System.Drawing.Pen($clrBorder,1)
    $e.Graphics.DrawLine($pen,0,0,$s.Width,0)
    $pen.Dispose()
})
$lblStatus = New-FLabel "Ready" 10 4 $fntSmall $clrTextSub
$pnlStatusBar.Controls.Add($lblStatus)

# Assemble form
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