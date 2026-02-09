#Parameters

#BCID - Business Continuity ID
param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$BCID = "0000",
    
    [Parameter(Mandatory=$false)]
    [string]$SubscriptionId,
    
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
# Set Subscription Context
#============================================================================

if ($SubscriptionId) {
    Write-Host "Setting subscription context to: $SubscriptionId" -ForegroundColor Yellow
    try {
        Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
        Write-Host "Subscription context set successfully" -ForegroundColor Green
    } catch {
        Write-Host "Error setting subscription context: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
} else {
    $currentContext = Get-AzContext
    if ($currentContext) {
        Write-Host "Using current subscription: $($currentContext.Subscription.Name) ($($currentContext.Subscription.Id))" -ForegroundColor Yellow
    } else {
        Write-Host "No subscription context found. Please login with Connect-AzAccount or specify -SubscriptionId" -ForegroundColor Red
        return
    }
}

#============================================================================
# BWS Base Check
#============================================================================

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  BWS Base Check - BCID: $BCID" -ForegroundColor Cyan
Write-Host "  Searching across entire subscription" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

## Azure Resource Definition
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

# Perform Check
$foundResources = @()
$missingResources = @()
$errorResources = @()

Write-Host "Checking Azure Resources across subscription..." -ForegroundColor Yellow
Write-Host ""

foreach ($resource in $azureResourcesToCheck) {
    $displayName = "$($resource.Category) - $($resource.SubCategory)"
    Write-Host "  [$($resource.Category)] " -NoNewline -ForegroundColor Gray
    Write-Host "$($resource.Name)" -NoNewline
    
    try {
        # Search across entire subscription without specifying ResourceGroupName
        $azResource = Get-AzResource -Name $resource.Name -ResourceType $resource.Type -ErrorAction SilentlyContinue
        
        if ($azResource) {
            Write-Host " ✓" -ForegroundColor Green
            $foundResources += [PSCustomObject]@{
                Category = $resource.Category
                SubCategory = $resource.SubCategory
                Name = $resource.Name
                Type = $resource.Type
                Status = "Found"
                Location = $azResource.Location
                ResourceGroupName = $azResource.ResourceGroupName
                ResourceId = $azResource.ResourceId
            }
        } else {
            Write-Host " ✗ MISSING" -ForegroundColor Red
            $missingResources += [PSCustomObject]@{
                Category = $resource.Category
                SubCategory = $resource.SubCategory
                Name = $resource.Name
                Type = $resource.Type
                Status = "Missing"
            }
        }
    } catch {
        Write-Host " ⚠ ERROR" -ForegroundColor Yellow
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
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  Total:     $($azureResourcesToCheck.Count) Resources" -ForegroundColor White
Write-Host "  Found:     $($foundResources.Count)" -ForegroundColor Green
Write-Host "  Missing:   $($missingResources.Count)" -ForegroundColor Red
Write-Host "  Errors:    $($errorResources.Count)" -ForegroundColor Yellow
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

# Detailed listing of found resources with their resource groups
if ($foundResources.Count -gt 0) {
    Write-Host "FOUND RESOURCES:" -ForegroundColor Green
    Write-Host ""
    $foundResources | Format-Table Category, SubCategory, Name, ResourceGroupName, Location -AutoSize
    Write-Host ""
}

# Detailed listing of missing resources
if ($missingResources.Count -gt 0) {
    Write-Host "MISSING RESOURCES:" -ForegroundColor Red
    Write-Host ""
    $missingResources | Format-Table Category, SubCategory, Name -AutoSize
    Write-Host ""
}

# Detailed listing of errors
if ($errorResources.Count -gt 0) {
    Write-Host "RESOURCES WITH ERRORS:" -ForegroundColor Yellow
    Write-Host ""
    $errorResources | Format-Table Category, SubCategory, Name, ErrorMessage -AutoSize
    Write-Host ""
}

#============================================================================
# BWS Report Generation
#============================================================================

$currentContext = Get-AzContext
$reportData = [PSCustomObject]@{
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    BCID = $BCID
    SubscriptionId = $currentContext.Subscription.Id
    SubscriptionName = $currentContext.Subscription.Name
    TotalResources = $azureResourcesToCheck.Count
    FoundCount = $foundResources.Count
    MissingCount = $missingResources.Count
    ErrorCount = $errorResources.Count
    AllResourcesExist = ($missingResources.Count -eq 0 -and $errorResources.Count -eq 0)
    FoundResources = $foundResources
    MissingResources = $missingResources
    ErrorResources = $errorResources
}

# Optional: Export report
if ($ExportReport) {
    $reportPath = "BWS_Check_Report_$BCID_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $reportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Host "Report exported to: $reportPath" -ForegroundColor Green
    Write-Host ""
}

# Return report object
return $reportData


#============================================================================
# M365 Check (Placeholder for future expansion)
#============================================================================

## M365
### Intune
# TODO: Implement Intune checks

#============================================================================
# BWS TA Check (Placeholder for future expansion)
#============================================================================

# TODO: Implement TA checks