#Parameters

#BCID - Business Continuity ID
param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$BCID = "0000",
    
    [Parameter(Mandatory=$false)]
    [string]$SubscriptionId,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportReport,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipIntune
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
###---AVD Storage Accounts
$bwsStorAccAVD1 = "sa" + $BCID.ToLower() + "avd0nch0"
$bwsStorAccAVD2 = "sa" + $BCID.ToLower() + "avd1nch0"
$bwsStorAccAVDBKP1 = "sa" + $BCID.ToLower() + "avd0bkpnch0"
$bwsStorAccAVDBKP2 = "sa" + $BCID.ToLower() + "avd0bkpnch1"
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
###--Azure Virtual Desktop
###---Azure Virtual Desktop Hosts Farm 1
$bwsAVDHost0 = $BCID.ToLower() + "-AVD0-0"
$bwsAVDHost1 = $BCID.ToLower() + "-AVD0-1"
$bwsAVDHost2 = $BCID.ToLower() + "-AVD0-2"
$bwsAVDHost3 = $BCID.ToLower() + "-AVD0-3"
$bwsAVDHost4 = $BCID.ToLower() + "-AVD0-4"
$bwsAVDHost5 = $BCID.ToLower() + "-AVD0-5"
$bwsAVDHost6 = $BCID.ToLower() + "-AVD0-6"
$bwsAVDHost7 = $BCID.ToLower() + "-AVD0-7"
$bwsAVDHost8 = $BCID.ToLower() + "-AVD0-8"
$bwsAVDHost9 = $BCID.ToLower() + "-AVD0-9"
###---Azure Virtual Desktop Hosts Farm 2
$bwsAVDHost0 = $BCID.ToLower() + "-AVD1-0"
$bwsAVDHost1 = $BCID.ToLower() + "-AVD1-1"
$bwsAVDHost2 = $BCID.ToLower() + "-AVD1-2"
$bwsAVDHost3 = $BCID.ToLower() + "-AVD1-3"
$bwsAVDHost4 = $BCID.ToLower() + "-AVD1-4"
$bwsAVDHost5 = $BCID.ToLower() + "-AVD1-5"
$bwsAVDHost6 = $BCID.ToLower() + "-AVD1-6"
$bwsAVDHost7 = $BCID.ToLower() + "-AVD1-7"
$bwsAVDHost8 = $BCID.ToLower() + "-AVD1-8"
$bwsAVDHost9 = $BCID.ToLower() + "-AVD1-9"
###---Avd Availability Sets
$bwsAVDAvSet0 = "avail" + $BCID.ToLower() + "-avd0-nch-0"
$bwsAVDAvSet1 = "avail" + $BCID.ToLower() + "-avd1-nch-0"
###--- Host Pools
$bwsAVDHostPool0 = "vdpool" + $BCID.ToLower() + "-avd0-nch-0"
$bwsAVDHostPool1 = "vdpool" + $BCID.ToLower() + "-avd1-nch-1"

##-Intune Standard Policies Definition
$intuneStandardPolicies = @(
    "STD - Autopilot - Hybrid Domain Join",
    "STD - Autopilot - Skip User ESP",
    "STD - AVD Hosts -  Standard",
    "STD - AVD Users - Standard",
    "STD - MacOS - Defender for Endpoint  - Common settings",
    "STD - MacOS - Defender for Endpoint  - Full Disk Access",
    "STD - MacOS - Defender for Endpoint - Background Service permissions",
    "STD - MacOS - Defender for Endpoint - Extensions approval",
    "STD - MacOS - Defender for Endpoint - Network Filter",
    "STD - MacOS - Defender for Endpoint - Onboarding",
    "STD - MacOS - Defender for Endpoint - UI Notification permissions",
    "STD - MacOS Computers - Bitlocker silent enable",
    "STD - MacOS Computers - Standard",
    "STD - Office security baseline policies for BWS - Users",
    "STD - Windows Computers - Bitlocker silent enable",
    "STD - Windows Computers - Defender Additional Configuration",
    "STD - Windows Computers - Defender Onboarding",
    "STD - Windows Computers - Device Health",
    "STD - Windows Computers - Edge",
    "STD - Windows Computers - OneDrive",
    "STD - Windows Computers - Standard",
    "STD - Windows Computers - Windows Updates",
    "STD - Windows LAPS",
    "STD - Windows Users - Standard",
    "STD - Windows Users - Windows Hello for Business",
    "STD - Windows Users - Windows Hello for Business Cloud Trust"
)


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
Write-Host "  AZURE RESOURCES SUMMARY" -ForegroundColor Cyan
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
# M365 Intune Check
#============================================================================

$intuneFoundPolicies = @()
$intuneMissingPolicies = @()
$intuneErrorPolicies = @()
$intuneCheckPerformed = $false

if (-not $SkipIntune) {
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  INTUNE POLICY CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Check if Microsoft.Graph.Intune module is available
    try {
        Write-Host "Checking Microsoft Graph authentication..." -ForegroundColor Yellow
        
        # Try to get Graph connection status
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Not connected to Microsoft Graph. Attempting to connect..." -ForegroundColor Yellow
            Write-Host "Please authenticate when prompted..." -ForegroundColor Yellow
            
            try {
                # Connect with required scopes for Intune
                Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All", "DeviceManagementManagedDevices.Read.All" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Skipping Intune policy check. Use -SkipIntune to suppress this section." -ForegroundColor Yellow
                $SkipIntune = $true
            }
        } else {
            Write-Host "Already connected to Microsoft Graph as: $($graphContext.Account)" -ForegroundColor Green
        }
        
        if (-not $SkipIntune) {
            $intuneCheckPerformed = $true
            Write-Host ""
            Write-Host "Checking Intune Policies..." -ForegroundColor Yellow
            Write-Host ""
            
            # Get all Intune policies using Microsoft Graph
            try {
                # Get Device Configuration Policies
                $deviceConfigs = Get-MgDeviceManagementDeviceConfiguration -All -ErrorAction SilentlyContinue
                
                # Get Device Compliance Policies
                $compliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All -ErrorAction SilentlyContinue
                
                # Get Configuration Policies (Settings Catalog)
                $configPolicies = Get-MgDeviceManagementConfigurationPolicy -All -ErrorAction SilentlyContinue
                
                # Combine all policies
                $allIntunePolicies = @()
                if ($deviceConfigs) { $allIntunePolicies += $deviceConfigs }
                if ($compliancePolicies) { $allIntunePolicies += $compliancePolicies }
                if ($configPolicies) { $allIntunePolicies += $configPolicies }
                
                Write-Host "Found $($allIntunePolicies.Count) total Intune policies" -ForegroundColor Cyan
                Write-Host ""
                
                # Check each required policy
                foreach ($requiredPolicy in $intuneStandardPolicies) {
                    Write-Host "  [Intune Policy] " -NoNewline -ForegroundColor Gray
                    Write-Host "$requiredPolicy" -NoNewline
                    
                    # Search for policy by display name (case-insensitive, partial match)
                    $foundPolicy = $allIntunePolicies | Where-Object { 
                        $_.DisplayName -like "*$requiredPolicy*" -or 
                        $_.DisplayName -eq $requiredPolicy 
                    } | Select-Object -First 1
                    
                    if ($foundPolicy) {
                        Write-Host " ✓" -ForegroundColor Green
                        $intuneFoundPolicies += [PSCustomObject]@{
                            PolicyName = $requiredPolicy
                            ActualName = $foundPolicy.DisplayName
                            PolicyId = $foundPolicy.Id
                            Status = "Found"
                        }
                    } else {
                        Write-Host " ✗ MISSING" -ForegroundColor Red
                        $intuneMissingPolicies += [PSCustomObject]@{
                            PolicyName = $requiredPolicy
                            Status = "Missing"
                        }
                    }
                }
                
            } catch {
                Write-Host "Error retrieving Intune policies: $($_.Exception.Message)" -ForegroundColor Red
                $intuneErrorPolicies += [PSCustomObject]@{
                    Error = "Failed to retrieve policies"
                    Message = $_.Exception.Message
                }
            }
            
            Write-Host ""
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host "  INTUNE POLICIES SUMMARY" -ForegroundColor Cyan
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host "  Total:     $($intuneStandardPolicies.Count) Required Policies" -ForegroundColor White
            Write-Host "  Found:     $($intuneFoundPolicies.Count)" -ForegroundColor Green
            Write-Host "  Missing:   $($intuneMissingPolicies.Count)" -ForegroundColor Red
            Write-Host "  Errors:    $($intuneErrorPolicies.Count)" -ForegroundColor Yellow
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host ""
            
            # Detailed listing of found policies
            if ($intuneFoundPolicies.Count -gt 0) {
                Write-Host "FOUND INTUNE POLICIES:" -ForegroundColor Green
                Write-Host ""
                $intuneFoundPolicies | Format-Table PolicyName, ActualName -AutoSize
                Write-Host ""
            }
            
            # Detailed listing of missing policies
            if ($intuneMissingPolicies.Count -gt 0) {
                Write-Host "MISSING INTUNE POLICIES:" -ForegroundColor Red
                Write-Host ""
                $intuneMissingPolicies | Format-Table PolicyName -AutoSize
                Write-Host ""
            }
            
            # Detailed listing of errors
            if ($intuneErrorPolicies.Count -gt 0) {
                Write-Host "INTUNE POLICY ERRORS:" -ForegroundColor Yellow
                Write-Host ""
                $intuneErrorPolicies | Format-Table Error, Message -AutoSize
                Write-Host ""
            }
        }
        
    } catch {
        Write-Host "Error during Intune check: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Intune check skipped. Use -SkipIntune to suppress this section." -ForegroundColor Yellow
    }
} else {
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  INTUNE CHECK SKIPPED (use without -SkipIntune to enable)" -ForegroundColor Yellow
    Write-Host "======================================================" -ForegroundColor Cyan
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
    
    # Azure Resources
    AzureTotalResources = $azureResourcesToCheck.Count
    AzureFoundCount = $foundResources.Count
    AzureMissingCount = $missingResources.Count
    AzureErrorCount = $errorResources.Count
    AzureAllResourcesExist = ($missingResources.Count -eq 0 -and $errorResources.Count -eq 0)
    AzureFoundResources = $foundResources
    AzureMissingResources = $missingResources
    AzureErrorResources = $errorResources
    
    # Intune Policies
    IntuneCheckPerformed = $intuneCheckPerformed
    IntuneTotalPolicies = $intuneStandardPolicies.Count
    IntuneFoundCount = $intuneFoundPolicies.Count
    IntuneMissingCount = $intuneMissingPolicies.Count
    IntuneErrorCount = $intuneErrorPolicies.Count
    IntuneAllPoliciesExist = ($intuneMissingPolicies.Count -eq 0 -and $intuneErrorPolicies.Count -eq 0)
    IntuneFoundPolicies = $intuneFoundPolicies
    IntuneMissingPolicies = $intuneMissingPolicies
    IntuneErrorPolicies = $intuneErrorPolicies
    
    # Overall Status
    OverallStatus = (
        ($missingResources.Count -eq 0 -and $errorResources.Count -eq 0) -and 
        (-not $intuneCheckPerformed -or ($intuneMissingPolicies.Count -eq 0 -and $intuneErrorPolicies.Count -eq 0))
    )
}

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  OVERALL SUMMARY" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  BCID: $BCID" -ForegroundColor White
Write-Host "  Subscription: $($currentContext.Subscription.Name)" -ForegroundColor White
Write-Host ""
Write-Host "  Azure Resources:" -ForegroundColor White
Write-Host "    Total:   $($azureResourcesToCheck.Count)" -ForegroundColor White
Write-Host "    Found:   $($foundResources.Count)" -ForegroundColor $(if ($foundResources.Count -eq $azureResourcesToCheck.Count) { "Green" } else { "Yellow" })
Write-Host "    Missing: $($missingResources.Count)" -ForegroundColor $(if ($missingResources.Count -eq 0) { "Green" } else { "Red" })

if ($intuneCheckPerformed) {
    Write-Host ""
    Write-Host "  Intune Policies:" -ForegroundColor White
    Write-Host "    Total:   $($intuneStandardPolicies.Count)" -ForegroundColor White
    Write-Host "    Found:   $($intuneFoundPolicies.Count)" -ForegroundColor $(if ($intuneFoundPolicies.Count -eq $intuneStandardPolicies.Count) { "Green" } else { "Yellow" })
    Write-Host "    Missing: $($intuneMissingPolicies.Count)" -ForegroundColor $(if ($intuneMissingPolicies.Count -eq 0) { "Green" } else { "Red" })
}

Write-Host ""
Write-Host "  Overall Status: " -NoNewline -ForegroundColor White
if ($reportData.OverallStatus) {
    Write-Host "✓ PASSED" -ForegroundColor Green
} else {
    Write-Host "✗ ISSUES FOUND" -ForegroundColor Red
}
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

# Optional: Export report
if ($ExportReport) {
    $reportPath = "BWS_Check_Report_$BCID`_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $reportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Host "Report exported to: $reportPath" -ForegroundColor Green
    Write-Host ""
}

# Return report object
return $reportData


#============================================================================
# BWS TA Check (Placeholder for future expansion)
#============================================================================

# TODO: Implement TA checks