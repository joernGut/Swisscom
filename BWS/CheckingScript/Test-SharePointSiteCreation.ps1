# SharePoint Site Creation - Property Diagnose
# Findet die korrekte Property für "Users can create SharePoint sites"

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  SharePoint Site Creation - Diagnose" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Prüfe SharePoint-Verbindung
try {
    $tenant = Get-SPOTenant -ErrorAction Stop
    Write-Host "✓ SharePoint Online Verbindung aktiv" -ForegroundColor Green
} catch {
    Write-Host "✗ Keine SharePoint-Verbindung!" -ForegroundColor Red
    Write-Host "Bitte zuerst verbinden:" -ForegroundColor Yellow
    Write-Host "  Connect-SPOService -Url https://TENANT-admin.sharepoint.com" -ForegroundColor Gray
    exit
}

Write-Host ""
Write-Host "Suche nach Site Creation Properties..." -ForegroundColor Yellow
Write-Host ""

# Alle Properties mit "Site" im Namen
Write-Host "1. Properties mit 'Site' im Namen:" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Gray
$tenant | Get-Member -MemberType Properties | Where-Object { $_.Name -like "*Site*" } | ForEach-Object {
    $propName = $_.Name
    $propValue = $tenant.$propName
    Write-Host "   $propName = $propValue" -ForegroundColor White
}

Write-Host ""
Write-Host "2. Properties mit 'Create' oder 'Creation' im Namen:" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Gray
$tenant | Get-Member -MemberType Properties | Where-Object { $_.Name -like "*Create*" } | ForEach-Object {
    $propName = $_.Name
    $propValue = $tenant.$propName
    Write-Host "   $propName = $propValue" -ForegroundColor White
}

Write-Host ""
Write-Host "3. Properties mit 'Self' im Namen:" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Gray
$tenant | Get-Member -MemberType Properties | Where-Object { $_.Name -like "*Self*" } | ForEach-Object {
    $propName = $_.Name
    $propValue = $tenant.$propName
    Write-Host "   $propName = $propValue" -ForegroundColor White
}

Write-Host ""
Write-Host "4. Properties mit 'Deny' im Namen:" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Gray
$tenant | Get-Member -MemberType Properties | Where-Object { $_.Name -like "*Deny*" } | ForEach-Object {
    $propName = $_.Name
    $propValue = $tenant.$propName
    Write-Host "   $propName = $propValue" -ForegroundColor White
}

Write-Host ""
Write-Host "5. Häufige Properties für Site Creation:" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Gray

# Liste bekannter Properties
$knownProps = @(
    "SelfServiceSiteCreationDisabled",
    "DenyAddAndCustomizePages",
    "ShowPeoplePickerSuggestionsForGuestUsers",
    "RequireAcceptingAccountMatchInvitedAccount",
    "SharingAllowedDomainList",
    "SharingBlockedDomainList",
    "EnableAIPIntegration"
)

foreach ($prop in $knownProps) {
    try {
        $value = $tenant.$prop
        if ($null -ne $value) {
            Write-Host "   $prop = $value" -ForegroundColor White
        }
    } catch {
        # Property existiert nicht
    }
}

Write-Host ""
Write-Host "6. Alle Tenant Properties (Vollständige Liste):" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Gray
$tenant | Format-List * | Out-String | Write-Host

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "WICHTIG - Was Sie jetzt tun müssen:" -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Gehen Sie zum SharePoint Admin Center:" -ForegroundColor White
Write-Host "   https://TENANT-admin.sharepoint.com" -ForegroundColor Gray
Write-Host ""
Write-Host "2. Navigieren Sie zu: Settings > Site Creation" -ForegroundColor White
Write-Host ""
Write-Host "3. Schauen Sie sich den aktuellen Zustand an:" -ForegroundColor White
Write-Host "   ☑ Users can create SharePoint sites = AKTIVIERT" -ForegroundColor Yellow
Write-Host "   ODER" -ForegroundColor Gray
Write-Host "   ☐ Users can create SharePoint sites = DEAKTIVIERT" -ForegroundColor Green
Write-Host ""
Write-Host "4. Vergleichen Sie die Werte oben mit der Einstellung!" -ForegroundColor White
Write-Host ""
Write-Host "5. ÄNDERN Sie die Einstellung im Admin Center" -ForegroundColor White
Write-Host "   (von aktiviert zu deaktiviert oder umgekehrt)" -ForegroundColor Gray
Write-Host ""
Write-Host "6. Führen Sie dieses Script ERNEUT aus" -ForegroundColor White
Write-Host ""
Write-Host "7. Schauen Sie welche Property sich geändert hat!" -ForegroundColor White
Write-Host "   Das ist die richtige Property für Site Creation!" -ForegroundColor Green
Write-Host ""
Write-Host "================================================" -ForegroundColor Gray
Write-Host ""
Write-Host "Bitte teilen Sie mir mit:" -ForegroundColor Yellow
Write-Host "  - Welche Property sich ändert" -ForegroundColor White
Write-Host "  - Von welchem Wert zu welchem Wert" -ForegroundColor White
Write-Host ""
