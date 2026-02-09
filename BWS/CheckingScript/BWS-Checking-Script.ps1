#Parameters

#BCID - Business Continuity ID
param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$BCID = "0000",
    
    [Parameter(Mandatory=$true)]
    [string]$ResourceGroupName,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportReport
)

# Naming Convention / String building
##-Azure
###--Storage Accounts
###---BWS Factory
$bwsStorAccFactory = "sa" + $BCID.ToLower() + "bwsfactorynch0"
###---BWS Inventory
$bwsStorAccInventory = "sa" + $BCID.ToLower() + "inventorynch0"
###---BWS Management Consoles
$bwsStorAccmgmtconsoles = "sa" + $BCID.ToLower() + "mgmtconsolesnch0"
###---BWS AD Backup
$bwsStorAccADBKP = "sa" + $BCID.ToLower() + "adbkpnch0"
###--VM
$bwsVMDomContrl = $BCID.ToLower() + "-S00"
$bwsVMDomContrlvDisk = "osdisk-" + $BCID.ToLower() + "-s00-nch-0"
$bwsVMNicDomContrl = "nic-" + $BCID.ToLower() + "-s00-nch-0"
###--Azure Vaults
###---Azure Key Vaults - BWS Factory
$bwsKeyVaultFactory = "kv-" + $BCID.ToLower() + "-bwsfactory-nch-0"
###--Azure Key Vaults - BWS Partners
$bwsKeyVaultPartner = "kv-" + $BCID.ToLower() + "-partners-nch-0"
###--vNets
$bwsvNETDefault = "vnet-" + $BCID.ToLower() + "-bws-nch-0"
###--Azure Gateways
$bwsAzVirtGW = "vpng-" + $BCID.ToLower() + "-bwsbns-nch-0"
$bwsLocNwGW = "lgw-" + $BCID.ToLower() + "-bwsbns-nch-0"
###--NSGs
$bwsNetAdds = "nsg-" + $BCID.ToLower() + "-snetadds-nch-0"
$bwsNetLoad = "nsg-" + $BCID.ToLower() + "-snetworkload-nch-0"
###--Public IPs
$bwsBnsPublicIP = "pip-" + $BCID.ToLower() + "-bwsbns-nch-0"
$bwsInetOutboundS00 = "pip-" + $BCID.ToLower() + "-internet-BBE0-S00-nch-0"
###--BNS/EC Connections
$bwsConBwsBnsEC = "s2sp1-" + $BCID.ToLower() + "-bwsbns-nch-0"
###--Automation Accounts
$bwsAutoAcc = "aa-" + $BCID.ToLower() + "-vmautomation-nch-0"
###--Managed identities
$bwsMI = "mi-" + $BCID.ToLower() + "-bwsfactory-nch-0"


#============================================================================
# BWS Base Check
#============================================================================

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  BWS Base Check - BCID: $BCID" -ForegroundColor Cyan
Write-Host "  Resource Group: $ResourceGroupName" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

## Azure Ressourcen-Definition
$azureResourcesToCheck = @(
    # Storage Accounts
    @{Name = $bwsStorAccFactory; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "BWS Factory"},
    @{Name = $bwsStorAccInventory; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "BWS Inventory"},
    @{Name = $bwsStorAccmgmtconsoles; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "Management Consoles"},
    @{Name = $bwsStorAccADBKP; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "AD Backup"},
    
    # Virtual Machines
    @{Name = $bwsVMDomContrl; Type = "Microsoft.Compute/virtualMachines"; Category = "Virtual Machine"; SubCategory = "Domain Controller"},
    @{Name = $bwsVMDomContrlvDisk; Type = "Microsoft.Compute/disks"; Category = "Virtual Machine"; SubCategory = "OS Disk"},
    @{Name = $bwsVMNicDomContrl; Type = "Microsoft.Network/networkInterfaces"; Category = "Virtual Machine"; SubCategory = "Network Interface"},
    
    # Azure Key Vaults
    @{Name = $bwsKeyVaultFactory; Type = "Microsoft.KeyVault/vaults"; Category = "Azure Vault"; SubCategory = "BWS Factory"},
    @{Name = $bwsKeyVaultPartner; Type = "Microsoft.KeyVault/vaults"; Category = "Azure Vault"; SubCategory = "BWS Partners"},
    
    # Virtual Networks
    @{Name = $bwsvNETDefault; Type = "Microsoft.Network/virtualNetworks"; Category = "vNet"; SubCategory = "Default vNet"},
    
    # Azure Gateways
    @{Name = $bwsAzVirtGW; Type = "Microsoft.Network/virtualNetworkGateways"; Category = "Azure Gateway"; SubCategory = "VPN Gateway"},
    @{Name = $bwsLocNwGW; Type = "Microsoft.Network/localNetworkGateways"; Category = "Azure Gateway"; SubCategory = "Local Network Gateway"},
    
    # Network Security Groups
    @{Name = $bwsNetAdds; Type = "Microsoft.Network/networkSecurityGroups"; Category = "NSG"; SubCategory = "ADDS Subnet"},
    @{Name = $bwsNetLoad; Type = "Microsoft.Network/networkSecurityGroups"; Category = "NSG"; SubCategory = "Workload Subnet"},
    
    # Public IPs
    @{Name = $bwsBnsPublicIP; Type = "Microsoft.Network/publicIPAddresses"; Category = "Public IP"; SubCategory = "BNS"},
    @{Name = $bwsInetOutboundS00; Type = "Microsoft.Network/publicIPAddresses"; Category = "Public IP"; SubCategory = "Internet Outbound S00"},
    
    # BNS/EC Connections
    @{Name = $bwsConBwsBnsEC; Type = "Microsoft.Network/connections"; Category = "BNS/EC Connection"; SubCategory = "S2S VPN"},
    
    # Automation Accounts
    @{Name = $bwsAutoAcc; Type = "Microsoft.Automation/automationAccounts"; Category = "Automation Account"; SubCategory = "VM Automation"},
    
    # Managed Identities
    @{Name = $bwsMI; Type = "Microsoft.ManagedIdentity/userAssignedIdentities"; Category = "Managed Identity"; SubCategory = "BWS Factory"}
)

# Prüfung durchführen
$foundResources = @()
$missingResources = @()
$errorResources = @()

Write-Host "Prüfe Azure Ressourcen..." -ForegroundColor Yellow
Write-Host ""

foreach ($resource in $azureResourcesToCheck) {
    $displayName = "$($resource.Category) - $($resource.SubCategory)"
    Write-Host "  [$($resource.Category)] " -NoNewline -ForegroundColor Gray
    Write-Host "$($resource.Name)" -NoNewline
    
    try {
        $azResource = Get-AzResource -ResourceGroupName $ResourceGroupName -Name $resource.Name -ResourceType $resource.Type -ErrorAction SilentlyContinue
        
        if ($azResource) {
            Write-Host " ✓" -ForegroundColor Green
            $foundResources += [PSCustomObject]@{
                Category = $resource.Category
                SubCategory = $resource.SubCategory
                Name = $resource.Name
                Type = $resource.Type
                Status = "Found"
                Location = $azResource.Location
                ResourceId = $azResource.ResourceId
            }
        } else {
            Write-Host " ✗ FEHLT" -ForegroundColor Red
            $missingResources += [PSCustomObject]@{
                Category = $resource.Category
                SubCategory = $resource.SubCategory
                Name = $resource.Name
                Type = $resource.Type
                Status = "Missing"
            }
        }
    } catch {
        Write-Host " ⚠ FEHLER" -ForegroundColor Yellow
        $errorResources += [PSCustomObject]@{
            Category = $resource.Category
            SubCategory = $resource.SubCategory
            Name = $resource.Name
            Type = $resource.Type
            Status = "Error"
            ErrorMessage = $_.Exception.Message
        }
    }
}

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  ZUSAMMENFASSUNG" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  Gesamt:    $($azureResourcesToCheck.Count) Ressourcen" -ForegroundColor White
Write-Host "  Gefunden:  $($foundResources.Count)" -ForegroundColor Green
Write-Host "  Fehlend:   $($missingResources.Count)" -ForegroundColor Red
Write-Host "  Fehler:    $($errorResources.Count)" -ForegroundColor Yellow
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

# Detaillierte Auflistung fehlender Ressourcen
if ($missingResources.Count -gt 0) {
    Write-Host "FEHLENDE RESSOURCEN:" -ForegroundColor Red
    Write-Host ""
    $missingResources | Format-Table Category, SubCategory, Name -AutoSize
    Write-Host ""
}

# Detaillierte Auflistung von Fehlern
if ($errorResources.Count -gt 0) {
    Write-Host "RESSOURCEN MIT FEHLERN:" -ForegroundColor Yellow
    Write-Host ""
    $errorResources | Format-Table Category, SubCategory, Name, ErrorMessage -AutoSize
    Write-Host ""
}

#============================================================================
# BWS Report erstellen
#============================================================================

$reportData = [PSCustomObject]@{
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    BCID = $BCID
    ResourceGroup = $ResourceGroupName
    TotalResources = $azureResourcesToCheck.Count
    FoundCount = $foundResources.Count
    MissingCount = $missingResources.Count
    ErrorCount = $errorResources.Count
    AllResourcesExist = ($missingResources.Count -eq 0 -and $errorResources.Count -eq 0)
    FoundResources = $foundResources
    MissingResources = $missingResources
    ErrorResources = $errorResources
}

# Optional: Report exportieren
if ($ExportReport) {
    $reportPath = "BWS_Check_Report_$BCID_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $reportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Host "Report exportiert nach: $reportPath" -ForegroundColor Green
    Write-Host ""
}

# Rückgabe des Report-Objekts
return $reportData


#============================================================================
# M365 Check (Platzhalter für zukünftige Erweiterung)
#============================================================================

## M365
### Intune
# TODO: Intune-Checks implementieren

#============================================================================
# BWS TA Check (Platzhalter für zukünftige Erweiterung)
#============================================================================

# TODO: TA-Checks implementieren