# Test Script: Teams Files Settings
# Zeigt die tatsächlichen Werte der Cloud Storage Provider

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Teams Files Settings - Diagnose" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Prüfe ob Teams-Verbindung besteht
try {
    $testConfig = Get-CsTeamsClientConfiguration -ErrorAction Stop
    Write-Host "✓ Teams Verbindung aktiv" -ForegroundColor Green
} catch {
    Write-Host "✗ Keine Teams Verbindung!" -ForegroundColor Red
    Write-Host "Bitte zuerst verbinden: Connect-MicrosoftTeams" -ForegroundColor Yellow
    exit
}

Write-Host ""
Write-Host "Cmdlet: Get-CsTeamsClientConfiguration" -ForegroundColor Yellow
Write-Host ""

# Hole Configuration
$config = Get-CsTeamsClientConfiguration

Write-Host "Cloud Storage Provider Status:" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Gray

# Citrix Files
Write-Host ""
Write-Host "1. CITRIX FILES" -ForegroundColor White
Write-Host "   Property: AllowCitrixContentSharing" -ForegroundColor Gray
Write-Host "   Wert:     $($config.AllowCitrixContentSharing)" -ForegroundColor Yellow
if ($null -eq $config.AllowCitrixContentSharing) {
    Write-Host "   Typ:      NULL (nicht gesetzt)" -ForegroundColor Gray
} else {
    Write-Host "   Typ:      $($config.AllowCitrixContentSharing.GetType().Name)" -ForegroundColor Gray
}
Write-Host "   = `$null:  $($null -eq $config.AllowCitrixContentSharing)" -ForegroundColor Gray
Write-Host "   = `$true:  $($config.AllowCitrixContentSharing -eq $true)" -ForegroundColor Gray
Write-Host "   = `$false: $($config.AllowCitrixContentSharing -eq $false)" -ForegroundColor Gray
if ($config.AllowCitrixContentSharing -eq $true) {
    Write-Host "   Status:   EINGESCHALTET (Enabled)" -ForegroundColor Red
} else {
    Write-Host "   Status:   AUSGESCHALTET (Disabled)" -ForegroundColor Green
}

# Dropbox
Write-Host ""
Write-Host "2. DROPBOX" -ForegroundColor White
Write-Host "   Property: AllowDropBox" -ForegroundColor Gray
Write-Host "   Wert:     $($config.AllowDropBox)" -ForegroundColor Yellow
if ($null -eq $config.AllowDropBox) {
    Write-Host "   Typ:      NULL (nicht gesetzt)" -ForegroundColor Gray
} else {
    Write-Host "   Typ:      $($config.AllowDropBox.GetType().Name)" -ForegroundColor Gray
}
Write-Host "   = `$null:  $($null -eq $config.AllowDropBox)" -ForegroundColor Gray
Write-Host "   = `$true:  $($config.AllowDropBox -eq $true)" -ForegroundColor Gray
Write-Host "   = `$false: $($config.AllowDropBox -eq $false)" -ForegroundColor Gray
if ($config.AllowDropBox -eq $true) {
    Write-Host "   Status:   EINGESCHALTET (Enabled)" -ForegroundColor Red
} else {
    Write-Host "   Status:   AUSGESCHALTET (Disabled)" -ForegroundColor Green
}

# Box
Write-Host ""
Write-Host "3. BOX" -ForegroundColor White
Write-Host "   Property: AllowBox" -ForegroundColor Gray
Write-Host "   Wert:     $($config.AllowBox)" -ForegroundColor Yellow
if ($null -eq $config.AllowBox) {
    Write-Host "   Typ:      NULL (nicht gesetzt)" -ForegroundColor Gray
} else {
    Write-Host "   Typ:      $($config.AllowBox.GetType().Name)" -ForegroundColor Gray
}
Write-Host "   = `$null:  $($null -eq $config.AllowBox)" -ForegroundColor Gray
Write-Host "   = `$true:  $($config.AllowBox -eq $true)" -ForegroundColor Gray
Write-Host "   = `$false: $($config.AllowBox -eq $false)" -ForegroundColor Gray
if ($config.AllowBox -eq $true) {
    Write-Host "   Status:   EINGESCHALTET (Enabled)" -ForegroundColor Red
} else {
    Write-Host "   Status:   AUSGESCHALTET (Disabled)" -ForegroundColor Green
}

# Google Drive
Write-Host ""
Write-Host "4. GOOGLE DRIVE" -ForegroundColor White
Write-Host "   Property: AllowGoogleDrive" -ForegroundColor Gray
Write-Host "   Wert:     $($config.AllowGoogleDrive)" -ForegroundColor Yellow
if ($null -eq $config.AllowGoogleDrive) {
    Write-Host "   Typ:      NULL (nicht gesetzt)" -ForegroundColor Gray
} else {
    Write-Host "   Typ:      $($config.AllowGoogleDrive.GetType().Name)" -ForegroundColor Gray
}
Write-Host "   = `$null:  $($null -eq $config.AllowGoogleDrive)" -ForegroundColor Gray
Write-Host "   = `$true:  $($config.AllowGoogleDrive -eq $true)" -ForegroundColor Gray
Write-Host "   = `$false: $($config.AllowGoogleDrive -eq $false)" -ForegroundColor Gray
if ($config.AllowGoogleDrive -eq $true) {
    Write-Host "   Status:   EINGESCHALTET (Enabled)" -ForegroundColor Red
} else {
    Write-Host "   Status:   AUSGESCHALTET (Disabled)" -ForegroundColor Green
}

# Egnyte
Write-Host ""
Write-Host "5. EGNYTE" -ForegroundColor White
Write-Host "   Property: AllowEgnyte" -ForegroundColor Gray
Write-Host "   Wert:     $($config.AllowEgnyte)" -ForegroundColor Yellow
if ($null -eq $config.AllowEgnyte) {
    Write-Host "   Typ:      NULL (nicht gesetzt)" -ForegroundColor Gray
} else {
    Write-Host "   Typ:      $($config.AllowEgnyte.GetType().Name)" -ForegroundColor Gray
}
Write-Host "   = `$null:  $($null -eq $config.AllowEgnyte)" -ForegroundColor Gray
Write-Host "   = `$true:  $($config.AllowEgnyte -eq $true)" -ForegroundColor Gray
Write-Host "   = `$false: $($config.AllowEgnyte -eq $false)" -ForegroundColor Gray
if ($config.AllowEgnyte -eq $true) {
    Write-Host "   Status:   EINGESCHALTET (Enabled)" -ForegroundColor Red
} else {
    Write-Host "   Status:   AUSGESCHALTET (Disabled)" -ForegroundColor Green
}

Write-Host ""
Write-Host "================================================" -ForegroundColor Gray
Write-Host ""
Write-Host "SOLL-WERT (BWS-Standard):" -ForegroundColor Yellow
Write-Host "  Alle Provider müssen AUSGESCHALTET sein" -ForegroundColor White
Write-Host "  Das bedeutet: Wert = `$false oder `$null" -ForegroundColor Gray
Write-Host ""

# Zusammenfassung
$enabledProviders = @()
if ($config.AllowCitrixContentSharing -eq $true) { $enabledProviders += "Citrix Files" }
if ($config.AllowDropBox -eq $true) { $enabledProviders += "Dropbox" }
if ($config.AllowBox -eq $true) { $enabledProviders += "Box" }
if ($config.AllowGoogleDrive -eq $true) { $enabledProviders += "Google Drive" }
if ($config.AllowEgnyte -eq $true) { $enabledProviders += "Egnyte" }

if ($enabledProviders.Count -eq 0) {
    Write-Host "✓ COMPLIANCE: Alle Cloud Storage Provider sind ausgeschaltet" -ForegroundColor Green
} else {
    Write-Host "✗ NON-COMPLIANCE: Folgende Provider sind eingeschaltet:" -ForegroundColor Red
    foreach ($provider in $enabledProviders) {
        Write-Host "  - $provider" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Vollständige Configuration:" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Gray
Write-Host ""
Write-Host "Alle Properties der CsTeamsClientConfiguration:" -ForegroundColor Yellow
$config | Format-List *

Write-Host ""
Write-Host "Suche nach 'Allow' Properties:" -ForegroundColor Yellow
$config | Get-Member -MemberType Properties | Where-Object { $_.Name -like "*Allow*" } | Format-Table Name, Definition

Write-Host ""
Write-Host "Suche nach 'File' Properties:" -ForegroundColor Yellow  
$config | Get-Member -MemberType Properties | Where-Object { $_.Name -like "*File*" } | Format-Table Name, Definition

Write-Host ""
Write-Host "Suche nach 'Storage' oder 'Content' Properties:" -ForegroundColor Yellow
$config | Get-Member -MemberType Properties | Where-Object { $_.Name -like "*Storage*" -or $_.Name -like "*Content*" } | Format-Table Name, Definition