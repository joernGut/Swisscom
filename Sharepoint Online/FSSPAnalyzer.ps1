#parameter for script
param (
    [Parameter(Mandatory = $true)]
    [string]$StartPathParam,

    [Parameter(Mandatory = $true)]
    [string]$ExcelOutputParam,

    [Parameter(Mandatory = $true)]
    [string]$SharePointBaseUrlParam
)

# Requires: ImportExcel module
Install-Module -Name ImportExcel -Force


# Variables
# File server path to analyze
$StartPath = $StartPathParam

# Output Excel file path
$ExcelOutput = $ExcelOutputParam

# Base SharePoint URL for the document library (adjust as needed)
$SharePointBaseUrl = $SharePointBaseUrlParam


# Define special characters (excluding common safe filename characters)
$SpecialChars = '[^a-zA-Z0-9\-_.\\:\s]'

# Store results
$Results = @()

# Function to URL encode strings
function EncodeUrl {
    param([string]$String)
    return [System.Net.WebUtility]::UrlEncode($String)
}

# Function to build the SharePoint-style URL for a file/folder
function Get-SharePointUrl {
    param(
        [string]$FullLocalPath,
        [string]$StartPath,
        [string]$BaseUrl
    )

    # Get relative path from base
    $relativePath = $FullLocalPath.Substring($StartPath.Length).TrimStart('\')
    # Replace backslashes with forward slashes for SharePoint URL format
    $relativePathWeb = $relativePath -replace '\\', '/'
    return "$BaseUrl/$relativePathWeb"
}

Write-Host "Analyzing paths from: $StartPath`nPlease wait..."

# Recursively scan files and folders
Get-ChildItem -Path $StartPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
    try {
        $FullPath = $_.FullName

        $criteria = @()

        if ($FullPath.Length -gt 255) {
            $criteria += "Path > 255 characters"
        }

        if ($FullPath -match $SpecialChars) {
            $criteria += "Contains special characters"
        }

        # Generate SharePoint-style URL
        $spUrl = Get-SharePointUrl -FullLocalPath $FullPath -StartPath $StartPath -BaseUrl $SharePointBaseUrl
        $EncodedUrl = EncodeUrl $spUrl
        if ($EncodedUrl.Length -gt 400) {
            $criteria += "Encoded SharePoint URL > 400 characters"
        }

        if ($criteria.Count -gt 0) {
            $Results += [PSCustomObject]@{
                'FullPath'       = $FullPath
                'SharePointURL'  = $spUrl
                'IssuesDetected' = ($criteria -join ', ')
            }
        }
    }
    catch {
        Write-Warning "Error with item: $_.FullName"
    }
}

# Output to Excel
if ($Results.Count -gt 0) {
    Write-Host "`nIssues found: $($Results.Count). Exporting to Excel..."
    $Results | Export-Excel -Path $ExcelOutput -WorksheetName "PathAnalysis" -AutoSize -TableName "PathIssues"
    Write-Host "Excel file saved to: $ExcelOutput"
} else {
    Write-Host "No issues found."
}
