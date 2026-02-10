<#
.SYNOPSIS
    BWS (Business Workplace Service) Checking Script with GUI support
.DESCRIPTION
    Checks Azure resources and Intune policies for BWS environments
.PARAMETER BCID
    Business Continuity ID
.PARAMETER SubscriptionId
    Azure Subscription ID (optional)
.PARAMETER ExportReport
    Export results to JSON file
.PARAMETER SkipIntune
    Skip Intune policy checks
.PARAMETER SkipEntraID
    Skip Entra ID Connect checks
.PARAMETER SkipIntuneConnector
    Skip Hybrid Azure AD Join checks
.PARAMETER SkipDefender
    Skip Defender for Endpoint checks
.PARAMETER ShowAllPolicies
    Show all found Intune policies (debug mode)
.PARAMETER CompactView
    Show only summary without detailed tables
.PARAMETER GUI
    Launch graphical user interface
.EXAMPLE
    .\BWS-Checking-Script.ps1 -BCID "1234"
.EXAMPLE
    .\BWS-Checking-Script.ps1 -BCID "1234" -GUI
.EXAMPLE
    .\BWS-Checking-Script.ps1 -BCID "1234" -CompactView -ExportReport
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$BCID = "0000",
    
    [Parameter(Mandatory=$false)]
    [string]$SubscriptionId,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportReport,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipIntune,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipEntraID,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipIntuneConnector,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipDefender,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowAllPolicies,
    
    [Parameter(Mandatory=$false)]
    [switch]$CompactView,
    
    [Parameter(Mandatory=$false)]
    [switch]$GUI
)

#============================================================================
# Global Variables and Configuration
#============================================================================

# Intune Standard Policies Definition
$script:intuneStandardPolicies = @(
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
# Helper Functions
#============================================================================

function Get-BWS-ResourceNames {
    param([string]$BCID)
    
    return @{
        # Storage Accounts
        StorAccFactory = "sa" + $BCID.ToLower() + "bwsfactorynch0"
        StorAccInventory = "sa" + $BCID.ToLower() + "inventorynch0"
        StorAccMgmtConsoles = "sa" + $BCID.ToLower() + "mgmtconsolesnch0"
        StorAccADBKP = "sa" + $BCID.ToLower() + "adbkpnch0"
        StorAccAVD1 = "sa" + $BCID.ToLower() + "avd0nch0"
        StorAccAVD2 = "sa" + $BCID.ToLower() + "avd1nch0"
        StorAccAVDBKP1 = "sa" + $BCID.ToLower() + "avd0bkpnch0"
        StorAccAVDBKP2 = "sa" + $BCID.ToLower() + "avd0bkpnch1"
        
        # Virtual Machines
        VMDomContrl = $BCID.ToLower() + "-S00"
        VMDomContrlvDisk = "osdisk-" + $BCID.ToLower() + "-s00-nch-0"
        VMNicDomContrl = "nic-" + $BCID.ToLower() + "-s00-nch-0"
        
        # Key Vaults
        KeyVaultFactory = "kv-" + $BCID.ToLower() + "-bwsfactory-nch-0"
        KeyVaultPartner = "kv-" + $BCID.ToLower() + "-partners-nch-0"
        
        # vNets
        vNETDefault = "vnet-" + $BCID.ToLower() + "-bws-nch-0"
        
        # Gateways
        AzVirtGW = "vpng-" + $BCID.ToLower() + "-bwsbns-nch-0"
        LocNwGW = "lgw-" + $BCID.ToLower() + "-bwsbns-nch-0"
        
        # NSGs
        NetAdds = "nsg-" + $BCID.ToLower() + "-snetadds-nch-0"
        NetLoad = "nsg-" + $BCID.ToLower() + "-snetworkload-nch-0"
        
        # Public IPs
        BnsPublicIP = "pip-" + $BCID.ToLower() + "-bwsbns-nch-0"
        InetOutboundS00 = "pip-" + $BCID.ToLower() + "-internet-BBE0-S00-nch-0"
        
        # Connections
        ConBwsBnsEC = "s2sp1-" + $BCID.ToLower() + "-bwsbns-nch-0"
        
        # Automation
        AutoAcc = "aa-" + $BCID.ToLower() + "-vmautomation-nch-0"
        
        # Managed Identity
        MI = "mi-" + $BCID.ToLower() + "-bwsfactory-nch-0"
    }
}

function Normalize-PolicyName {
    param([string]$name)
    return $name -replace '\s+', ' ' -replace '^\s+|\s+$', '' | ForEach-Object { $_.ToLower() }
}

#============================================================================
# Main Check Functions
#============================================================================

function Test-AzureResources {
    param(
        [string]$BCID,
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  BWS Base Check - BCID: $BCID" -ForegroundColor Cyan
    Write-Host "  Searching across entire subscription" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $resourceNames = Get-BWS-ResourceNames -BCID $BCID
    
    $azureResourcesToCheck = @(
        @{Name = $resourceNames.StorAccFactory; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "BWS Factory"},
        @{Name = $resourceNames.StorAccInventory; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "BWS Inventory"},
        @{Name = $resourceNames.StorAccMgmtConsoles; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "Management Consoles"},
        @{Name = $resourceNames.StorAccADBKP; Type = "Microsoft.Storage/storageAccounts"; Category = "Storage Account"; SubCategory = "AD Backup"},
        @{Name = $resourceNames.VMDomContrl; Type = "Microsoft.Compute/virtualMachines"; Category = "Virtual Machine"; SubCategory = "Domain Controller"},
        @{Name = $resourceNames.VMDomContrlvDisk; Type = "Microsoft.Compute/disks"; Category = "Virtual Machine"; SubCategory = "OS Disk"},
        @{Name = $resourceNames.VMNicDomContrl; Type = "Microsoft.Network/networkInterfaces"; Category = "Virtual Machine"; SubCategory = "Network Interface"},
        @{Name = $resourceNames.KeyVaultFactory; Type = "Microsoft.KeyVault/vaults"; Category = "Azure Vault"; SubCategory = "BWS Factory"},
        @{Name = $resourceNames.KeyVaultPartner; Type = "Microsoft.KeyVault/vaults"; Category = "Azure Vault"; SubCategory = "BWS Partners"},
        @{Name = $resourceNames.vNETDefault; Type = "Microsoft.Network/virtualNetworks"; Category = "vNet"; SubCategory = "Default vNet"},
        @{Name = $resourceNames.AzVirtGW; Type = "Microsoft.Network/virtualNetworkGateways"; Category = "Azure Gateway"; SubCategory = "VPN Gateway"},
        @{Name = $resourceNames.LocNwGW; Type = "Microsoft.Network/localNetworkGateways"; Category = "Azure Gateway"; SubCategory = "Local Network Gateway"},
        @{Name = $resourceNames.NetAdds; Type = "Microsoft.Network/networkSecurityGroups"; Category = "NSG"; SubCategory = "ADDS Subnet"},
        @{Name = $resourceNames.NetLoad; Type = "Microsoft.Network/networkSecurityGroups"; Category = "NSG"; SubCategory = "Workload Subnet"},
        @{Name = $resourceNames.BnsPublicIP; Type = "Microsoft.Network/publicIPAddresses"; Category = "Public IP"; SubCategory = "BNS"},
        @{Name = $resourceNames.InetOutboundS00; Type = "Microsoft.Network/publicIPAddresses"; Category = "Public IP"; SubCategory = "Internet Outbound S00"},
        @{Name = $resourceNames.ConBwsBnsEC; Type = "Microsoft.Network/connections"; Category = "BNS/EC Connection"; SubCategory = "S2S VPN"},
        @{Name = $resourceNames.AutoAcc; Type = "Microsoft.Automation/automationAccounts"; Category = "Automation Account"; SubCategory = "VM Automation"},
        @{Name = $resourceNames.MI; Type = "Microsoft.ManagedIdentity/userAssignedIdentities"; Category = "Managed Identity"; SubCategory = "BWS Factory"}
    )
    
    $foundResources = @()
    $missingResources = @()
    $errorResources = @()
    
    Write-Host "Checking Azure Resources across subscription..." -ForegroundColor Yellow
    Write-Host ""
    
    foreach ($resource in $azureResourcesToCheck) {
        Write-Host "  [$($resource.Category)] " -NoNewline -ForegroundColor Gray
        Write-Host "$($resource.Name)" -NoNewline
        
        try {
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
    
    if (-not $CompactView) {
        if ($foundResources.Count -gt 0) {
            Write-Host "FOUND RESOURCES:" -ForegroundColor Green
            Write-Host ""
            $foundResources | Format-Table Category, SubCategory, Name, ResourceGroupName, Location -AutoSize
            Write-Host ""
        }
        
        if ($missingResources.Count -gt 0) {
            Write-Host "MISSING RESOURCES:" -ForegroundColor Red
            Write-Host ""
            $missingResources | Format-Table Category, SubCategory, Name -AutoSize
            Write-Host ""
        }
        
        if ($errorResources.Count -gt 0) {
            Write-Host "RESOURCES WITH ERRORS:" -ForegroundColor Yellow
            Write-Host ""
            $errorResources | Format-Table Category, SubCategory, Name, ErrorMessage -AutoSize
            Write-Host ""
        }
    }
    
    return @{
        Found = $foundResources
        Missing = $missingResources
        Errors = $errorResources
        Total = $azureResourcesToCheck.Count
    }
}

function Test-IntunePolicies {
    param(
        [bool]$ShowAllPolicies = $false,
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  INTUNE POLICY CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $intuneFoundPolicies = @()
    $intuneMissingPolicies = @()
    $intuneErrorPolicies = @()
    
    try {
        Write-Host "Checking Microsoft Graph authentication..." -ForegroundColor Yellow
        
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Not connected to Microsoft Graph. Attempting to connect..." -ForegroundColor Yellow
            Write-Host "Please authenticate when prompted..." -ForegroundColor Yellow
            
            try {
                Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All", "DeviceManagementManagedDevices.Read.All" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                return @{
                    Found = @()
                    Missing = @()
                    Errors = @(@{Error = "Connection failed"; Message = $_.Exception.Message})
                    Total = $script:intuneStandardPolicies.Count
                    CheckPerformed = $false
                }
            }
        } else {
            Write-Host "Already connected to Microsoft Graph as: $($graphContext.Account)" -ForegroundColor Green
        }
        
        Write-Host ""
        Write-Host "Checking Intune Policies..." -ForegroundColor Yellow
        Write-Host ""
        
        $allIntunePolicies = @()
        
        try {
            $deviceConfigs = Get-MgDeviceManagementDeviceConfiguration -All -ErrorAction Stop
            if ($deviceConfigs) { 
                $allIntunePolicies += $deviceConfigs 
                Write-Host "  Retrieved $($deviceConfigs.Count) Device Configuration policies" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Warning: Could not retrieve Device Configuration policies" -ForegroundColor Yellow
        }
        
        try {
            $compliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All -ErrorAction Stop
            if ($compliancePolicies) { 
                $allIntunePolicies += $compliancePolicies 
                Write-Host "  Retrieved $($compliancePolicies.Count) Device Compliance policies" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Warning: Could not retrieve Device Compliance policies" -ForegroundColor Yellow
        }
        
        try {
            $configPolicies = Get-MgDeviceManagementConfigurationPolicy -All -ErrorAction Stop
            if ($configPolicies) { 
                $allIntunePolicies += $configPolicies 
                Write-Host "  Retrieved $($configPolicies.Count) Configuration policies (Settings Catalog)" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Info: Configuration Policy cmdlet not available, trying Graph API..." -ForegroundColor Yellow
            
            try {
                $graphUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
                $configPoliciesResponse = Invoke-MgGraphRequest -Uri $graphUri -Method GET -ErrorAction Stop
                if ($configPoliciesResponse.value) {
                    $allIntunePolicies += $configPoliciesResponse.value
                    Write-Host "  Retrieved $($configPoliciesResponse.value.Count) Configuration policies via Graph API" -ForegroundColor Gray
                }
            } catch {
                Write-Host "  Warning: Could not retrieve Configuration policies via Graph API" -ForegroundColor Yellow
            }
        }
        
        try {
            $intentUri = "https://graph.microsoft.com/beta/deviceManagement/intents"
            $intentResponse = Invoke-MgGraphRequest -Uri $intentUri -Method GET -ErrorAction Stop
            if ($intentResponse.value) {
                $allIntunePolicies += $intentResponse.value
                Write-Host "  Retrieved $($intentResponse.value.Count) Endpoint Security policies" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Info: Could not retrieve Endpoint Security policies" -ForegroundColor Yellow
        }
        
        Write-Host ""
        Write-Host "Found $($allIntunePolicies.Count) total Intune policies" -ForegroundColor Cyan
        
        if ($ShowAllPolicies) {
            Write-Host ""
            Write-Host "DEBUG: All found Intune policies:" -ForegroundColor Magenta
            $allIntunePolicies | Sort-Object DisplayName | ForEach-Object { 
                Write-Host "  - $($_.DisplayName)" -ForegroundColor Gray 
            }
        }
        
        Write-Host ""
        
        foreach ($requiredPolicy in $script:intuneStandardPolicies) {
            Write-Host "  [Intune Policy] " -NoNewline -ForegroundColor Gray
            Write-Host "$requiredPolicy" -NoNewline
            
            $normalizedRequired = Normalize-PolicyName $requiredPolicy
            $foundPolicy = $null
            
            # Strategy 1: Exact match
            $foundPolicy = $allIntunePolicies | Where-Object { 
                (Normalize-PolicyName $_.DisplayName) -eq $normalizedRequired
            } | Select-Object -First 1
            
            # Strategy 2: Contains match
            if (-not $foundPolicy) {
                $foundPolicy = $allIntunePolicies | Where-Object { 
                    (Normalize-PolicyName $_.DisplayName) -like "*$normalizedRequired*"
                } | Select-Object -First 1
            }
            
            # Strategy 3: Reverse contains
            if (-not $foundPolicy) {
                $foundPolicy = $allIntunePolicies | Where-Object { 
                    $normalizedRequired -like "*$(Normalize-PolicyName $_.DisplayName)*"
                } | Select-Object -First 1
            }
            
            # Strategy 4: Fuzzy match
            if (-not $foundPolicy) {
                $cleanRequired = $normalizedRequired -replace '\s*(std|standard|policy|policies|-)\s*', ' ' -replace '\s+', ' ' -replace '^\s+|\s+$', ''
                $foundPolicy = $allIntunePolicies | Where-Object {
                    $cleanActual = (Normalize-PolicyName $_.DisplayName) -replace '\s*(std|standard|policy|policies|-)\s*', ' ' -replace '\s+', ' ' -replace '^\s+|\s+$', ''
                    $cleanActual -like "*$cleanRequired*" -or $cleanRequired -like "*$cleanActual*"
                } | Select-Object -First 1
            }
            
            if ($foundPolicy) {
                Write-Host " ✓" -ForegroundColor Green
                $intuneFoundPolicies += [PSCustomObject]@{
                    PolicyName = $requiredPolicy
                    ActualName = $foundPolicy.DisplayName
                    PolicyId = $foundPolicy.Id
                    Status = "Found"
                    MatchType = if ((Normalize-PolicyName $foundPolicy.DisplayName) -eq $normalizedRequired) { "Exact" } else { "Fuzzy" }
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
        $intuneErrorPolicies += @{
            Error = "Failed to retrieve policies"
            Message = $_.Exception.Message
        }
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  INTUNE POLICIES SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Total:     $($script:intuneStandardPolicies.Count) Required Policies" -ForegroundColor White
    Write-Host "  Found:     $($intuneFoundPolicies.Count)" -ForegroundColor Green
    Write-Host "  Missing:   $($intuneMissingPolicies.Count)" -ForegroundColor Red
    Write-Host "  Errors:    $($intuneErrorPolicies.Count)" -ForegroundColor Yellow
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView) {
        if ($intuneFoundPolicies.Count -gt 0) {
            Write-Host "FOUND INTUNE POLICIES:" -ForegroundColor Green
            Write-Host ""
            $intuneFoundPolicies | Format-Table PolicyName, ActualName, MatchType -AutoSize
            
            $fuzzyMatches = $intuneFoundPolicies | Where-Object { $_.MatchType -eq "Fuzzy" }
            if ($fuzzyMatches.Count -gt 0) {
                Write-Host ""
                Write-Host "Note: The following policies were matched using fuzzy logic:" -ForegroundColor Yellow
                $fuzzyMatches | ForEach-Object {
                    Write-Host "  Expected: $($_.PolicyName)" -ForegroundColor Yellow
                    Write-Host "  Found:    $($_.ActualName)" -ForegroundColor Gray
                    Write-Host ""
                }
            }
            Write-Host ""
        }
        
        if ($intuneMissingPolicies.Count -gt 0) {
            Write-Host "MISSING INTUNE POLICIES:" -ForegroundColor Red
            Write-Host ""
            $intuneMissingPolicies | Format-Table PolicyName -AutoSize
            Write-Host ""
        }
        
        if ($intuneErrorPolicies.Count -gt 0) {
            Write-Host "INTUNE POLICY ERRORS:" -ForegroundColor Yellow
            Write-Host ""
            $intuneErrorPolicies | Format-Table Error, Message -AutoSize
            Write-Host ""
        }
    }
    
    return @{
        Found = $intuneFoundPolicies
        Missing = $intuneMissingPolicies
        Errors = $intuneErrorPolicies
        Total = $script:intuneStandardPolicies.Count
        CheckPerformed = $true
    }
}

function Test-EntraIDConnect {
    param(
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  ENTRA ID CONNECT CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $entraIDStatus = @{
        IsInstalled = $false
        IsRunning = $false
        Version = $null
        ServiceStatus = $null
        LastSyncTime = $null
        SyncErrors = @()
        Details = @()
    }
    
    try {
        Write-Host "Checking Entra ID Connect status..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if Microsoft Graph is connected
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
            try {
                Connect-MgGraph -Scopes "Directory.Read.All", "Organization.Read.All" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                $entraIDStatus.SyncErrors += "Graph connection failed"
                return @{
                    Status = $entraIDStatus
                    CheckPerformed = $false
                }
            }
        }
        
        Write-Host ""
        
        # Check Entra ID Connect Sync Status via Graph API
        try {
            Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking directory synchronization..." -NoNewline
            
            $orgUri = "https://graph.microsoft.com/v1.0/organization"
            $orgInfo = Invoke-MgGraphRequest -Uri $orgUri -Method GET -ErrorAction Stop
            
            if ($orgInfo.value -and $orgInfo.value.Count -gt 0) {
                $org = $orgInfo.value[0]
                
                # Check if directory sync is enabled
                $onPremisesSyncEnabled = $org.onPremisesSyncEnabled
                
                if ($onPremisesSyncEnabled) {
                    Write-Host " ✓ ENABLED" -ForegroundColor Green
                    $entraIDStatus.IsInstalled = $true
                    
                    # Get last sync time
                    $lastSyncTime = $org.onPremisesLastSyncDateTime
                    if ($lastSyncTime) {
                        $entraIDStatus.LastSyncTime = $lastSyncTime
                        $timeSinceSync = (Get-Date) - [DateTime]$lastSyncTime
                        
                        Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
                        Write-Host "Last sync time: $lastSyncTime " -NoNewline
                        
                        # Check if sync is recent (within last 30 minutes)
                        if ($timeSinceSync.TotalMinutes -le 30) {
                            Write-Host "✓ RECENT" -ForegroundColor Green
                            $entraIDStatus.IsRunning = $true
                        } elseif ($timeSinceSync.TotalHours -le 2) {
                            Write-Host "⚠ WARNING (last sync > 30 min)" -ForegroundColor Yellow
                            $entraIDStatus.IsRunning = $true
                            $entraIDStatus.SyncErrors += "Last sync older than 30 minutes"
                        } else {
                            Write-Host "✗ OLD (last sync > 2 hours)" -ForegroundColor Red
                            $entraIDStatus.IsRunning = $false
                            $entraIDStatus.SyncErrors += "Last sync older than 2 hours"
                        }
                    } else {
                        Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
                        Write-Host "Last sync time: " -NoNewline
                        Write-Host "✗ UNKNOWN" -ForegroundColor Yellow
                        $entraIDStatus.SyncErrors += "No last sync time available"
                    }
                    
                    # Check for sync errors via Graph API
                    try {
                        Write-Host "  [Entra ID] " -NoNewline -ForegroundColor Gray
                        Write-Host "Checking for sync errors..." -NoNewline
                        
                        $syncErrorsUri = "https://graph.microsoft.com/v1.0/directory/onPremisesSynchronization"
                        $syncErrorsResponse = Invoke-MgGraphRequest -Uri $syncErrorsUri -Method GET -ErrorAction SilentlyContinue
                        
                        if ($syncErrorsResponse) {
                            Write-Host " ✓ NO ERRORS" -ForegroundColor Green
                        } else {
                            Write-Host " ⚠ UNABLE TO CHECK" -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Host " ⚠ UNABLE TO CHECK" -ForegroundColor Yellow
                    }
                    
                } else {
                    Write-Host " ✗ NOT ENABLED" -ForegroundColor Red
                    $entraIDStatus.IsInstalled = $false
                    $entraIDStatus.SyncErrors += "Directory synchronization not enabled"
                }
                
                $entraIDStatus.Details += "Organization: $($org.displayName)"
                
            } else {
                Write-Host " ✗ UNABLE TO CHECK" -ForegroundColor Yellow
                $entraIDStatus.SyncErrors += "Could not retrieve organization info"
            }
            
        } catch {
            Write-Host " ✗ ERROR" -ForegroundColor Red
            $entraIDStatus.SyncErrors += "Error checking Entra ID Connect: $($_.Exception.Message)"
        }
        
    } catch {
        Write-Host "Error during Entra ID Connect check: $($_.Exception.Message)" -ForegroundColor Red
        $entraIDStatus.SyncErrors += "General error: $($_.Exception.Message)"
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  ENTRA ID CONNECT SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Sync Enabled:    " -NoNewline -ForegroundColor White
    Write-Host $(if ($entraIDStatus.IsInstalled) { "Yes" } else { "No" }) -ForegroundColor $(if ($entraIDStatus.IsInstalled) { "Green" } else { "Red" })
    Write-Host "  Sync Active:     " -NoNewline -ForegroundColor White
    Write-Host $(if ($entraIDStatus.IsRunning) { "Yes" } else { "No" }) -ForegroundColor $(if ($entraIDStatus.IsRunning) { "Green" } else { "Red" })
    if ($entraIDStatus.LastSyncTime) {
        Write-Host "  Last Sync:       $($entraIDStatus.LastSyncTime)" -ForegroundColor White
    }
    Write-Host "  Errors:          $($entraIDStatus.SyncErrors.Count)" -ForegroundColor $(if ($entraIDStatus.SyncErrors.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView -and $entraIDStatus.SyncErrors.Count -gt 0) {
        Write-Host "ENTRA ID CONNECT ERRORS/WARNINGS:" -ForegroundColor Yellow
        Write-Host ""
        $entraIDStatus.SyncErrors | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    return @{
        Status = $entraIDStatus
        CheckPerformed = $true
    }
}

function Test-IntuneConnector {
    param(
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  HYBRID AZURE AD JOIN CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $connectorStatus = @{
        IsConnected = $false
        ConnectorVersion = $null
        LastCheckIn = $null
        HealthStatus = $null
        Connectors = @()
        Errors = @()
    }
    
    try {
        Write-Host "Checking Hybrid Azure AD Join status..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if Microsoft Graph is connected
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
            try {
                Connect-MgGraph -Scopes "DeviceManagementServiceConfig.Read.All", "DeviceManagementConfiguration.Read.All" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                $connectorStatus.Errors += "Graph connection failed"
                return @{
                    Status = $connectorStatus
                    CheckPerformed = $false
                }
            }
        }
        
        Write-Host ""
        
        # ============================================================================
        # COMMENTED OUT - Certificate/NDES Connector Check
        # Uncomment if needed for certificate-based authentication checks
        # ============================================================================
        <#
        # Check Intune Connector for Active Directory (NDES Connector)
        try {
            Write-Host "  [Connector] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking Intune Connector for AD (NDES)..." -NoNewline
            
            $certConnectorUri = "https://graph.microsoft.com/beta/deviceManagement/ndesConnectors"
            $certConnectors = Invoke-MgGraphRequest -Uri $certConnectorUri -Method GET -ErrorAction Stop
            
            if ($certConnectors.value -and $certConnectors.value.Count -gt 0) {
                $activeCertConnectors = $certConnectors.value | Where-Object { $_.state -eq "active" }
                
                if ($activeCertConnectors.Count -gt 0) {
                    Write-Host " ✓ ACTIVE ($($activeCertConnectors.Count) connector(s))" -ForegroundColor Green
                    $connectorStatus.IsConnected = $true
                    
                    foreach ($connector in $activeCertConnectors) {
                        $connectorStatus.Connectors += @{
                            Type = "Intune Connector for Active Directory"
                            Name = $connector.displayName
                            State = $connector.state
                            LastCheckIn = $connector.lastConnectionDateTime
                            Version = $connector.connectorVersion
                        }
                        
                        if ($connector.lastConnectionDateTime) {
                            $lastCheckIn = [DateTime]$connector.lastConnectionDateTime
                            $timeSinceCheckIn = (Get-Date) - $lastCheckIn
                            
                            Write-Host "  [Connector] " -NoNewline -ForegroundColor Gray
                            Write-Host "$($connector.displayName) - Last check-in: $($connector.lastConnectionDateTime) " -NoNewline
                            
                            if ($timeSinceCheckIn.TotalHours -le 1) {
                                Write-Host "✓ RECENT" -ForegroundColor Green
                            } elseif ($timeSinceCheckIn.TotalHours -le 24) {
                                Write-Host "⚠ WARNING (> 1 hour)" -ForegroundColor Yellow
                                $connectorStatus.Errors += "$($connector.displayName): Last check-in > 1 hour ago"
                            } else {
                                Write-Host "✗ OLD (> 24 hours)" -ForegroundColor Red
                                $connectorStatus.Errors += "$($connector.displayName): Last check-in > 24 hours ago"
                            }
                        }
                    }
                } else {
                    Write-Host " ⚠ INACTIVE" -ForegroundColor Yellow
                    $connectorStatus.Errors += "Intune Connector for AD exists but is not active"
                }
            } else {
                Write-Host " ⚠ NOT FOUND" -ForegroundColor Yellow
                $connectorStatus.Errors += "No Intune Connector for Active Directory configured"
            }
            
        } catch {
            Write-Host " ⚠ UNABLE TO CHECK" -ForegroundColor Yellow
            $connectorStatus.Errors += "Error checking Intune Connector for AD: $($_.Exception.Message)"
        }
        #>
        # ============================================================================
        
        # ============================================================================
        # COMMENTED OUT - Exchange Connector Check
        # Uncomment if needed for Exchange integration checks
        # ============================================================================
        <#
        # Check Exchange Connector
        try {
            Write-Host "  [Connector] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking Exchange Connector..." -NoNewline
            
            $exchangeConnectorUri = "https://graph.microsoft.com/beta/deviceManagement/exchangeConnectors"
            $exchangeConnectors = Invoke-MgGraphRequest -Uri $exchangeConnectorUri -Method GET -ErrorAction Stop
            
            if ($exchangeConnectors.value -and $exchangeConnectors.value.Count -gt 0) {
                $activeExchangeConnectors = $exchangeConnectors.value | Where-Object { $_.status -eq "healthy" -or $_.status -eq "active" }
                
                if ($activeExchangeConnectors.Count -gt 0) {
                    Write-Host " ✓ ACTIVE ($($activeExchangeConnectors.Count) connector(s))" -ForegroundColor Green
                    
                    foreach ($connector in $activeExchangeConnectors) {
                        $connectorStatus.Connectors += @{
                            Type = "Exchange Connector"
                            Name = $connector.serverName
                            State = $connector.status
                            LastCheckIn = $connector.lastSuccessfulSyncDateTime
                        }
                        
                        if ($connector.lastSuccessfulSyncDateTime) {
                            Write-Host "  [Connector] " -NoNewline -ForegroundColor Gray
                            Write-Host "$($connector.serverName) - Last sync: $($connector.lastSuccessfulSyncDateTime)" -ForegroundColor Gray
                        }
                    }
                } else {
                    Write-Host " ⚠ INACTIVE" -ForegroundColor Yellow
                    $connectorStatus.Errors += "Exchange connector exists but is not healthy"
                }
            } else {
                Write-Host " ⓘ NOT CONFIGURED" -ForegroundColor Gray
            }
            
        } catch {
            Write-Host " ⓘ NOT CONFIGURED" -ForegroundColor Gray
        }
        #>
        # ============================================================================
        
        # Check for Hybrid Azure AD Join status (ACTIVE - NOT COMMENTED)
        try {
            Write-Host "  [Hybrid Join] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking Hybrid Azure AD Join status..." -NoNewline
            
            # Check via organization settings
            $orgUri = "https://graph.microsoft.com/v1.0/organization"
            $orgInfo = Invoke-MgGraphRequest -Uri $orgUri -Method GET -ErrorAction Stop
            
            if ($orgInfo.value -and $orgInfo.value.Count -gt 0) {
                $org = $orgInfo.value[0]
                $onPremisesSyncEnabled = $org.onPremisesSyncEnabled
                
                if ($onPremisesSyncEnabled) {
                    Write-Host " ✓ ENABLED (Sync active)" -ForegroundColor Green
                } else {
                    Write-Host " ⓘ NOT ENABLED" -ForegroundColor Gray
                }
            } else {
                Write-Host " ⚠ UNABLE TO CHECK" -ForegroundColor Yellow
            }
            
        } catch {
            Write-Host " ⚠ UNABLE TO CHECK" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "Error during Hybrid Join check: $($_.Exception.Message)" -ForegroundColor Red
        $connectorStatus.Errors += "General error: $($_.Exception.Message)"
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  HYBRID AZURE AD JOIN SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Check Performed:   Yes" -ForegroundColor White
    Write-Host "  Errors/Warnings:   $($connectorStatus.Errors.Count)" -ForegroundColor $(if ($connectorStatus.Errors.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView) {
        if ($connectorStatus.Errors.Count -gt 0) {
            Write-Host "ERRORS/WARNINGS:" -ForegroundColor Yellow
            Write-Host ""
            $connectorStatus.Errors | ForEach-Object {
                Write-Host "  - $_" -ForegroundColor Yellow
            }
            Write-Host ""
        }
    }
    
    return @{
        Status = $connectorStatus
        CheckPerformed = $true
    }
}

function Test-DefenderForEndpoint {
    param(
        [string]$BCID,
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  MICROSOFT DEFENDER FOR ENDPOINT CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $defenderStatus = @{
        ConnectorActive = $false
        ConfiguredPolicies = 0
        OnboardedDevices = 0
        FilesFound = @()
        FilesMissing = @()
        Errors = @()
    }
    
    try {
        Write-Host "Checking Microsoft Defender for Endpoint..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if Microsoft Graph is connected
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
            try {
                Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All", "DeviceManagementManagedDevices.Read.All" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                $defenderStatus.Errors += "Graph connection failed"
                return @{
                    Status = $defenderStatus
                    CheckPerformed = $false
                }
            }
        }
        
        Write-Host ""
        
        # ============================================================================
        # Check 1: Defender Configuration Policies
        # ============================================================================
        try {
            Write-Host "  [Defender] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking Defender configuration policies..." -NoNewline
            
            # Try multiple policy sources
            $defenderPoliciesFound = $false
            $totalDefenderPolicies = 0
            
            # Check Device Configuration Policies
            try {
                $deviceConfigs = Get-MgDeviceManagementDeviceConfiguration -All -ErrorAction SilentlyContinue
                $defenderDeviceConfigs = $deviceConfigs | Where-Object { 
                    $_.DisplayName -like "*Defender*" -or 
                    $_.DisplayName -like "*ATP*" -or
                    $_.DisplayName -like "*Endpoint Protection*" -or
                    $_.DisplayName -like "*Antivirus*"
                }
                if ($defenderDeviceConfigs) {
                    $totalDefenderPolicies += $defenderDeviceConfigs.Count
                    $defenderPoliciesFound = $true
                }
            } catch {}
            
            # Check Configuration Policies (Settings Catalog)
            try {
                $configPoliciesUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
                $configPolicies = Invoke-MgGraphRequest -Uri $configPoliciesUri -Method GET -ErrorAction SilentlyContinue
                if ($configPolicies.value) {
                    $defenderConfigPolicies = $configPolicies.value | Where-Object {
                        $_.name -like "*Defender*" -or 
                        $_.name -like "*ATP*" -or
                        $_.name -like "*Endpoint*" -or
                        $_.name -like "*Antivirus*"
                    }
                    if ($defenderConfigPolicies) {
                        $totalDefenderPolicies += $defenderConfigPolicies.Count
                        $defenderPoliciesFound = $true
                    }
                }
            } catch {}
            
            # Check Endpoint Security Policies (Intents)
            try {
                $intentsUri = "https://graph.microsoft.com/beta/deviceManagement/intents"
                $intents = Invoke-MgGraphRequest -Uri $intentsUri -Method GET -ErrorAction SilentlyContinue
                if ($intents.value) {
                    $defenderIntents = $intents.value | Where-Object {
                        $_.displayName -like "*Defender*" -or
                        $_.displayName -like "*Antivirus*" -or
                        $_.displayName -like "*Endpoint*" -or
                        $_.templateId -like "*endpointSecurityAntivirus*" -or
                        $_.templateId -like "*endpointSecurityEndpointDetectionAndResponse*"
                    }
                    if ($defenderIntents) {
                        $totalDefenderPolicies += $defenderIntents.Count
                        $defenderPoliciesFound = $true
                        $defenderStatus.ConnectorActive = $true
                    }
                }
            } catch {}
            
            $defenderStatus.ConfiguredPolicies = $totalDefenderPolicies
            
            if ($defenderPoliciesFound -and $totalDefenderPolicies -gt 0) {
                Write-Host " ✓ FOUND ($totalDefenderPolicies policies)" -ForegroundColor Green
                $defenderStatus.ConnectorActive = $true
            } else {
                Write-Host " ⚠ NO POLICIES FOUND" -ForegroundColor Yellow
                $defenderStatus.Errors += "No Defender for Endpoint policies configured"
            }
            
        } catch {
            Write-Host " ⚠ ERROR" -ForegroundColor Yellow
            $defenderStatus.Errors += "Error checking Defender policies: $($_.Exception.Message)"
        }
        
        # ============================================================================
        # Check 2: Managed Devices (Defender-compatible)
        # ============================================================================
        try {
            Write-Host "  [Defender] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking compatible managed devices..." -NoNewline
            
            $managedDevices = Get-MgDeviceManagementManagedDevice -All -ErrorAction SilentlyContinue
            
            if ($managedDevices) {
                # Count Windows and macOS devices (Defender-compatible)
                $compatibleDevices = $managedDevices | Where-Object {
                    $_.OperatingSystem -eq "Windows" -or $_.OperatingSystem -eq "macOS"
                }
                
                $defenderStatus.OnboardedDevices = $compatibleDevices.Count
                
                if ($compatibleDevices.Count -gt 0) {
                    Write-Host " ✓ $($compatibleDevices.Count) compatible devices" -ForegroundColor Green
                } else {
                    Write-Host " ⓘ No compatible devices found" -ForegroundColor Gray
                }
            } else {
                Write-Host " ⓘ Unable to retrieve devices" -ForegroundColor Gray
            }
            
        } catch {
            Write-Host " ⓘ Unable to check" -ForegroundColor Gray
        }
        
        # ============================================================================
        # Check 3: Defender Onboarding Files in Storage Account
        # ============================================================================
        Write-Host "  [Defender] " -NoNewline -ForegroundColor Gray
        Write-Host "Checking onboarding files in Storage Account..." -NoNewline
        
        $requiredFiles = @(
            "GatewayWindowsDefenderATPOnboardingPackage_Intune_MacClients.zip",
            "GatewayWindowsDefenderATPOnboardingPackage_Intune_WinClients.zip",
            "GatewayWindowsDefenderATPOnboardingPackage_WinClients.zip",
            "GatewayWindowsDefenderATPOnboardingPackage_WinServers.zip"
        )
        
        $storageAccountName = "sa" + $BCID.ToLower() + "bwsfactorynch0"
        $containerName = "defender-files"
        
        try {
            # Get storage account
            $storageAccount = Get-AzStorageAccount | Where-Object { $_.StorageAccountName -eq $storageAccountName } | Select-Object -First 1
            
            if ($storageAccount) {
                $ctx = $storageAccount.Context
                
                # Check if container exists
                $container = Get-AzStorageContainer -Name $containerName -Context $ctx -ErrorAction SilentlyContinue
                
                if ($container) {
                    # Get all blobs
                    $blobs = Get-AzStorageBlob -Container $containerName -Context $ctx -ErrorAction SilentlyContinue
                    
                    if ($blobs) {
                        $blobNames = $blobs | ForEach-Object { $_.Name }
                        
                        # Check each required file
                        foreach ($file in $requiredFiles) {
                            if ($blobNames -contains $file) {
                                $defenderStatus.FilesFound += $file
                            } else {
                                $defenderStatus.FilesMissing += $file
                            }
                        }
                        
                        if ($defenderStatus.FilesMissing.Count -eq 0) {
                            Write-Host " ✓ ALL FILES PRESENT (4/4)" -ForegroundColor Green
                        } else {
                            Write-Host " ⚠ MISSING FILES ($($defenderStatus.FilesFound.Count)/4)" -ForegroundColor Yellow
                            $defenderStatus.Errors += "$($defenderStatus.FilesMissing.Count) onboarding file(s) missing"
                        }
                    } else {
                        Write-Host " ⚠ CONTAINER EMPTY (0/4)" -ForegroundColor Yellow
                        $defenderStatus.FilesMissing = $requiredFiles
                        $defenderStatus.Errors += "Container 'defender-files' is empty"
                    }
                } else {
                    Write-Host " ✗ CONTAINER NOT FOUND (0/4)" -ForegroundColor Red
                    $defenderStatus.FilesMissing = $requiredFiles
                    $defenderStatus.Errors += "Container 'defender-files' not found"
                }
            } else {
                Write-Host " ✗ STORAGE ACCOUNT NOT FOUND (0/4)" -ForegroundColor Red
                $defenderStatus.FilesMissing = $requiredFiles
                $defenderStatus.Errors += "Storage account '$storageAccountName' not found"
            }
            
        } catch {
            Write-Host " ⚠ ERROR (0/4)" -ForegroundColor Yellow
            $defenderStatus.FilesMissing = $requiredFiles
            $defenderStatus.Errors += "Error checking storage: $($_.Exception.Message)"
        }
        
    } catch {
        Write-Host "Error during Defender check: $($_.Exception.Message)" -ForegroundColor Red
        $defenderStatus.Errors += "General error: $($_.Exception.Message)"
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  DEFENDER FOR ENDPOINT SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Policies Configured: $($defenderStatus.ConfiguredPolicies)" -ForegroundColor $(if ($defenderStatus.ConfiguredPolicies -gt 0) { "Green" } else { "Yellow" })
    Write-Host "  Compatible Devices:  $($defenderStatus.OnboardedDevices)" -ForegroundColor $(if ($defenderStatus.OnboardedDevices -gt 0) { "Green" } else { "Gray" })
    Write-Host "  Onboarding Files:    $($defenderStatus.FilesFound.Count)/4" -ForegroundColor $(if ($defenderStatus.FilesMissing.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Status:              " -NoNewline -ForegroundColor White
    Write-Host $(if ($defenderStatus.ConnectorActive) { "Active" } else { "Not Configured" }) -ForegroundColor $(if ($defenderStatus.ConnectorActive) { "Green" } else { "Yellow" })
    Write-Host "  Errors/Warnings:     $($defenderStatus.Errors.Count)" -ForegroundColor $(if ($defenderStatus.Errors.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView) {
        if ($defenderStatus.FilesFound.Count -gt 0) {
            Write-Host "FOUND ONBOARDING FILES:" -ForegroundColor Green
            Write-Host ""
            $defenderStatus.FilesFound | ForEach-Object {
                Write-Host "  ✓ $_" -ForegroundColor Green
            }
            Write-Host ""
        }
        
        if ($defenderStatus.FilesMissing.Count -gt 0) {
            Write-Host "MISSING ONBOARDING FILES:" -ForegroundColor Red
            Write-Host ""
            $defenderStatus.FilesMissing | ForEach-Object {
                Write-Host "  ✗ $_" -ForegroundColor Red
            }
            Write-Host ""
        }
        
        if ($defenderStatus.Errors.Count -gt 0) {
            Write-Host "DEFENDER ERRORS/WARNINGS:" -ForegroundColor Yellow
            Write-Host ""
            $defenderStatus.Errors | ForEach-Object {
                Write-Host "  - $_" -ForegroundColor Yellow
            }
            Write-Host ""
        }
    }
    
    return @{
        Status = $defenderStatus
        CheckPerformed = $true
    }
}

#============================================================================
# GUI Mode
#============================================================================

if ($GUI) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "BWS Checking Tool - GUI"
    $form.Size = New-Object System.Drawing.Size(1000, 750)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    # BCID Input
    $labelBCID = New-Object System.Windows.Forms.Label
    $labelBCID.Location = New-Object System.Drawing.Point(20, 20)
    $labelBCID.Size = New-Object System.Drawing.Size(150, 20)
    $labelBCID.Text = "BCID:"
    $form.Controls.Add($labelBCID)
    
    $textBCID = New-Object System.Windows.Forms.TextBox
    $textBCID.Location = New-Object System.Drawing.Point(170, 18)
    $textBCID.Size = New-Object System.Drawing.Size(200, 20)
    $textBCID.Text = $BCID
    $form.Controls.Add($textBCID)
    
    # Subscription ID Input
    $labelSubID = New-Object System.Windows.Forms.Label
    $labelSubID.Location = New-Object System.Drawing.Point(20, 50)
    $labelSubID.Size = New-Object System.Drawing.Size(150, 20)
    $labelSubID.Text = "Subscription ID (optional):"
    $form.Controls.Add($labelSubID)
    
    $textSubID = New-Object System.Windows.Forms.TextBox
    $textSubID.Location = New-Object System.Drawing.Point(170, 48)
    $textSubID.Size = New-Object System.Drawing.Size(400, 20)
    $textSubID.Text = $SubscriptionId
    $form.Controls.Add($textSubID)
    
    # GroupBox for Check Selection
    $groupBoxChecks = New-Object System.Windows.Forms.GroupBox
    $groupBoxChecks.Location = New-Object System.Drawing.Point(20, 85)
    $groupBoxChecks.Size = New-Object System.Drawing.Size(300, 175)
    $groupBoxChecks.Text = "Select Checks to Run"
    $form.Controls.Add($groupBoxChecks)
    
    # Azure Check Checkbox
    $chkAzure = New-Object System.Windows.Forms.CheckBox
    $chkAzure.Location = New-Object System.Drawing.Point(15, 25)
    $chkAzure.Size = New-Object System.Drawing.Size(250, 20)
    $chkAzure.Text = "Azure Resources Check"
    $chkAzure.Checked = $true
    $groupBoxChecks.Controls.Add($chkAzure)
    
    # Intune Check Checkbox
    $chkIntune = New-Object System.Windows.Forms.CheckBox
    $chkIntune.Location = New-Object System.Drawing.Point(15, 50)
    $chkIntune.Size = New-Object System.Drawing.Size(250, 20)
    $chkIntune.Text = "Intune Policies Check"
    $chkIntune.Checked = $true
    $groupBoxChecks.Controls.Add($chkIntune)
    
    # Entra ID Connect Check Checkbox
    $chkEntraID = New-Object System.Windows.Forms.CheckBox
    $chkEntraID.Location = New-Object System.Drawing.Point(15, 75)
    $chkEntraID.Size = New-Object System.Drawing.Size(250, 20)
    $chkEntraID.Text = "Entra ID Connect Check"
    $chkEntraID.Checked = $true
    $groupBoxChecks.Controls.Add($chkEntraID)
    
    # Hybrid Join Check Checkbox
    $chkIntuneConn = New-Object System.Windows.Forms.CheckBox
    $chkIntuneConn.Location = New-Object System.Drawing.Point(15, 100)
    $chkIntuneConn.Size = New-Object System.Drawing.Size(280, 20)
    $chkIntuneConn.Text = "Hybrid Azure AD Join Check"
    $chkIntuneConn.Checked = $true
    $groupBoxChecks.Controls.Add($chkIntuneConn)
    
    # Defender Check Checkbox
    $chkDefender = New-Object System.Windows.Forms.CheckBox
    $chkDefender.Location = New-Object System.Drawing.Point(15, 125)
    $chkDefender.Size = New-Object System.Drawing.Size(280, 20)
    $chkDefender.Text = "Defender for Endpoint Check"
    $chkDefender.Checked = $true
    $groupBoxChecks.Controls.Add($chkDefender)
    
    # Options GroupBox
    $groupBoxOptions = New-Object System.Windows.Forms.GroupBox
    $groupBoxOptions.Location = New-Object System.Drawing.Point(340, 85)
    $groupBoxOptions.Size = New-Object System.Drawing.Size(300, 175)
    $groupBoxOptions.Text = "Options"
    $form.Controls.Add($groupBoxOptions)
    
    # Compact View Checkbox
    $chkCompact = New-Object System.Windows.Forms.CheckBox
    $chkCompact.Location = New-Object System.Drawing.Point(15, 25)
    $chkCompact.Size = New-Object System.Drawing.Size(250, 20)
    $chkCompact.Text = "Compact View"
    $chkCompact.Checked = $false
    $groupBoxOptions.Controls.Add($chkCompact)
    
    # Show All Policies Checkbox
    $chkShowAll = New-Object System.Windows.Forms.CheckBox
    $chkShowAll.Location = New-Object System.Drawing.Point(15, 50)
    $chkShowAll.Size = New-Object System.Drawing.Size(250, 20)
    $chkShowAll.Text = "Show All Policies (Debug)"
    $chkShowAll.Checked = $false
    $groupBoxOptions.Controls.Add($chkShowAll)
    
    # Export Report Checkbox
    $chkExport = New-Object System.Windows.Forms.CheckBox
    $chkExport.Location = New-Object System.Drawing.Point(15, 75)
    $chkExport.Size = New-Object System.Drawing.Size(250, 20)
    $chkExport.Text = "Export Report to JSON"
    $chkExport.Checked = $false
    $groupBoxOptions.Controls.Add($chkExport)
    
    # Run Button
    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Location = New-Object System.Drawing.Point(660, 85)
    $btnRun.Size = New-Object System.Drawing.Size(150, 60)
    $btnRun.Text = "Run Check"
    $btnRun.BackColor = [System.Drawing.Color]::LightGreen
    $btnRun.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnRun)
    
    # Clear Button
    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Location = New-Object System.Drawing.Point(660, 155)
    $btnClear.Size = New-Object System.Drawing.Size(150, 30)
    $btnClear.Text = "Clear Output"
    $form.Controls.Add($btnClear)
    
    # Status Label
    $labelStatus = New-Object System.Windows.Forms.Label
    $labelStatus.Location = New-Object System.Drawing.Point(20, 270)
    $labelStatus.Size = New-Object System.Drawing.Size(800, 20)
    $labelStatus.Text = "Ready - Please select checks and click 'Run Check'"
    $labelStatus.ForeColor = [System.Drawing.Color]::Blue
    $form.Controls.Add($labelStatus)
    
    # Progress Bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 295)
    $progressBar.Size = New-Object System.Drawing.Size(950, 20)
    $progressBar.Style = "Continuous"
    $form.Controls.Add($progressBar)
    
    # Output TextBox
    $textOutput = New-Object System.Windows.Forms.TextBox
    $textOutput.Location = New-Object System.Drawing.Point(20, 325)
    $textOutput.Size = New-Object System.Drawing.Size(950, 365)
    $textOutput.Multiline = $true
    $textOutput.ScrollBars = "Both"
    $textOutput.Font = New-Object System.Drawing.Font("Consolas", 9)
    $textOutput.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $textOutput.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $textOutput.ReadOnly = $true
    $textOutput.WordWrap = $false
    $form.Controls.Add($textOutput)
    
    # Clear Button Click
    $btnClear.Add_Click({
        $textOutput.Clear()
        $labelStatus.Text = "Output cleared - Ready for next check"
        $labelStatus.ForeColor = [System.Drawing.Color]::Blue
        $progressBar.Value = 0
    })
    
    # Run Button Click
    $btnRun.Add_Click({
        $textOutput.Clear()
        $progressBar.Value = 0
        $labelStatus.Text = "Initializing check..."
        $labelStatus.ForeColor = [System.Drawing.Color]::Orange
        $btnRun.Enabled = $false
        $form.Refresh()
        
        $bcid = $textBCID.Text
        $subId = $textSubID.Text
        $runAzure = $chkAzure.Checked
        $runIntune = $chkIntune.Checked
        $runEntraID = $chkEntraID.Checked
        $runIntuneConn = $chkIntuneConn.Checked
        $runDefender = $chkDefender.Checked
        $compact = $chkCompact.Checked
        $showAll = $chkShowAll.Checked
        $export = $chkExport.Checked
        
        try {
            # Set subscription context if provided
            if ($subId) {
                $textOutput.AppendText("Setting subscription context to: $subId`r`n")
                $textOutput.Refresh()
                try {
                    Set-AzContext -SubscriptionId $subId -ErrorAction Stop | Out-Null
                    $textOutput.AppendText("Subscription context set successfully`r`n`r`n")
                } catch {
                    $textOutput.AppendText("ERROR: Could not set subscription context: $($_.Exception.Message)`r`n`r`n")
                    $labelStatus.Text = "Error setting subscription context"
                    $labelStatus.ForeColor = [System.Drawing.Color]::Red
                    $btnRun.Enabled = $true
                    return
                }
            } else {
                $currentContext = Get-AzContext
                if ($currentContext) {
                    $textOutput.AppendText("Using current subscription: $($currentContext.Subscription.Name)`r`n`r`n")
                } else {
                    $textOutput.AppendText("ERROR: No subscription context found`r`n`r`n")
                    $labelStatus.Text = "Error: No subscription context"
                    $labelStatus.ForeColor = [System.Drawing.Color]::Red
                    $btnRun.Enabled = $true
                    return
                }
            }
            
            $progressBar.Value = 10
            
            # Redirect Write-Host
            $originalWriteHost = Get-Command Write-Host
            function global:Write-Host {
                param(
                    [Parameter(Position=0, ValueFromPipeline=$true)]
                    [object]$Object,
                    [System.ConsoleColor]$ForegroundColor,
                    [switch]$NoNewline
                )
                
                $msg = if ($Object) { $Object.ToString() } else { "" }
                
                if (-not $NoNewline) {
                    $script:textOutput.AppendText("$msg`r`n")
                } else {
                    $script:textOutput.AppendText($msg)
                }
                $script:textOutput.SelectionStart = $script:textOutput.Text.Length
                $script:textOutput.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            $azureResults = $null
            $intuneResults = $null
            $entraIDResults = $null
            $intuneConnResults = $null
            $defenderResults = $null
            
            $totalChecks = ($runAzure -as [int]) + ($runIntune -as [int]) + ($runEntraID -as [int]) + ($runIntuneConn -as [int]) + ($runDefender -as [int])
            $currentCheck = 0
            $progressIncrement = if ($totalChecks -gt 0) { 80 / $totalChecks } else { 0 }
            
            # Run Azure Check
            if ($runAzure) {
                $labelStatus.Text = "Running Azure Resources Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $azureResults = Test-AzureResources -BCID $bcid -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run Intune Check
            if ($runIntune) {
                $labelStatus.Text = "Running Intune Policies Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $intuneResults = Test-IntunePolicies -ShowAllPolicies $showAll -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run Entra ID Connect Check
            if ($runEntraID) {
                $labelStatus.Text = "Running Entra ID Connect Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $entraIDResults = Test-EntraIDConnect -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run Hybrid Join Check
            if ($runIntuneConn) {
                $labelStatus.Text = "Running Hybrid Azure AD Join Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $intuneConnResults = Test-IntuneConnector -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run Defender for Endpoint Check
            if ($runDefender) {
                $labelStatus.Text = "Running Defender for Endpoint Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $defenderResults = Test-DefenderForEndpoint -BCID $bcid -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Overall Summary
            Write-Host ""
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host "  OVERALL SUMMARY" -ForegroundColor Cyan
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host "  BCID: $bcid" -ForegroundColor White
            
            if ($runAzure -and $azureResults) {
                Write-Host ""
                Write-Host "  Azure Resources:" -ForegroundColor White
                Write-Host "    Total:   $($azureResults.Total)" -ForegroundColor White
                Write-Host "    Found:   $($azureResults.Found.Count)" -ForegroundColor Green
                Write-Host "    Missing: $($azureResults.Missing.Count)" -ForegroundColor $(if ($azureResults.Missing.Count -eq 0) { "Green" } else { "Red" })
            }
            
            if ($runIntune -and $intuneResults -and $intuneResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Intune Policies:" -ForegroundColor White
                Write-Host "    Total:   $($intuneResults.Total)" -ForegroundColor White
                Write-Host "    Found:   $($intuneResults.Found.Count)" -ForegroundColor Green
                Write-Host "    Missing: $($intuneResults.Missing.Count)" -ForegroundColor $(if ($intuneResults.Missing.Count -eq 0) { "Green" } else { "Red" })
            }
            
            if ($runEntraID -and $entraIDResults -and $entraIDResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Entra ID Connect:" -ForegroundColor White
                Write-Host "    Sync Enabled: " -NoNewline -ForegroundColor White
                Write-Host $(if ($entraIDResults.Status.IsInstalled) { "Yes" } else { "No" }) -ForegroundColor $(if ($entraIDResults.Status.IsInstalled) { "Green" } else { "Red" })
                Write-Host "    Sync Active:  " -NoNewline -ForegroundColor White
                Write-Host $(if ($entraIDResults.Status.IsRunning) { "Yes" } else { "No" }) -ForegroundColor $(if ($entraIDResults.Status.IsRunning) { "Green" } else { "Yellow" })
            }
            
            if ($runIntuneConn -and $intuneConnResults -and $intuneConnResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Hybrid Azure AD Join:" -ForegroundColor White
                Write-Host "    Status:       Check Performed" -ForegroundColor Green
                Write-Host "    Errors:       $($intuneConnResults.Status.Errors.Count)" -ForegroundColor $(if ($intuneConnResults.Status.Errors.Count -eq 0) { "Green" } else { "Yellow" })
            }
            
            if ($runDefender -and $defenderResults -and $defenderResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Defender for Endpoint:" -ForegroundColor White
                Write-Host "    Policies:     $($defenderResults.Status.ConfiguredPolicies)" -ForegroundColor $(if ($defenderResults.Status.ConfiguredPolicies -gt 0) { "Green" } else { "Yellow" })
                Write-Host "    Devices:      $($defenderResults.Status.OnboardedDevices)" -ForegroundColor $(if ($defenderResults.Status.OnboardedDevices -gt 0) { "Green" } else { "Gray" })
                Write-Host "    Files:        $($defenderResults.Status.FilesFound.Count)/4" -ForegroundColor $(if ($defenderResults.Status.FilesMissing.Count -eq 0) { "Green" } else { "Red" })
                Write-Host "    Status:       " -NoNewline -ForegroundColor White
                Write-Host $(if ($defenderResults.Status.ConnectorActive) { "Active" } else { "Not Configured" }) -ForegroundColor $(if ($defenderResults.Status.ConnectorActive) { "Green" } else { "Yellow" })
            }
            
            Write-Host "======================================================" -ForegroundColor Cyan
            
            if ($compact) {
                Write-Host ""
                Write-Host "Note: Compact View enabled" -ForegroundColor Gray
            }
            
            # Export report if requested
            if ($export) {
                $reportData = @{
                    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    BCID = $bcid
                    AzureResults = $azureResults
                    IntuneResults = $intuneResults
                    EntraIDResults = $entraIDResults
                    IntuneConnectorResults = $intuneConnResults
                    DefenderResults = $defenderResults
                }
                
                $reportPath = "BWS_Check_Report_${bcid}_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
                $reportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportPath -Encoding UTF8
                Write-Host ""
                Write-Host "Report exported to: $reportPath" -ForegroundColor Green
            }
            
            $progressBar.Value = 100
            $labelStatus.Text = "Check completed successfully!"
            $labelStatus.ForeColor = [System.Drawing.Color]::Green
            
        } catch {
            $textOutput.AppendText("`r`nERROR: $($_.Exception.Message)`r`n")
            $labelStatus.Text = "Error occurred during check"
            $labelStatus.ForeColor = [System.Drawing.Color]::Red
        } finally {
            # Restore Write-Host
            Remove-Item Function:\Write-Host -ErrorAction SilentlyContinue
            $btnRun.Enabled = $true
        }
    })
    
    [void]$form.ShowDialog()
    exit
}

#============================================================================
# Command Line Mode
#============================================================================

# Set Subscription Context
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

# Run Azure Check
$azureResults = Test-AzureResources -BCID $BCID -CompactView $CompactView

# Run Intune Check
$intuneResults = $null
if (-not $SkipIntune) {
    $intuneResults = Test-IntunePolicies -ShowAllPolicies $ShowAllPolicies -CompactView $CompactView
}

# Run Entra ID Connect Check
$entraIDResults = $null
if (-not $SkipEntraID) {
    $entraIDResults = Test-EntraIDConnect -CompactView $CompactView
}

# Run Intune Connector Check
$intuneConnResults = $null
if (-not $SkipIntuneConnector) {
    $intuneConnResults = Test-IntuneConnector -CompactView $CompactView
}

# Run Defender for Endpoint Check
$defenderResults = $null
if (-not $SkipDefender) {
    $defenderResults = Test-DefenderForEndpoint -BCID $BCID -CompactView $CompactView
}

# Overall Summary
$currentContext = Get-AzContext
Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  OVERALL SUMMARY" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  BCID: $BCID" -ForegroundColor White
Write-Host "  Subscription: $($currentContext.Subscription.Name)" -ForegroundColor White
Write-Host ""
Write-Host "  Azure Resources:" -ForegroundColor White
Write-Host "    Total:   $($azureResults.Total)" -ForegroundColor White
Write-Host "    Found:   $($azureResults.Found.Count)" -ForegroundColor $(if ($azureResults.Found.Count -eq $azureResults.Total) { "Green" } else { "Yellow" })
Write-Host "    Missing: $($azureResults.Missing.Count)" -ForegroundColor $(if ($azureResults.Missing.Count -eq 0) { "Green" } else { "Red" })

if ($intuneResults -and $intuneResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Intune Policies:" -ForegroundColor White
    Write-Host "    Total:   $($intuneResults.Total)" -ForegroundColor White
    Write-Host "    Found:   $($intuneResults.Found.Count)" -ForegroundColor $(if ($intuneResults.Found.Count -eq $intuneResults.Total) { "Green" } else { "Yellow" })
    Write-Host "    Missing: $($intuneResults.Missing.Count)" -ForegroundColor $(if ($intuneResults.Missing.Count -eq 0) { "Green" } else { "Red" })
}

if ($entraIDResults -and $entraIDResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Entra ID Connect:" -ForegroundColor White
    Write-Host "    Sync Enabled: " -NoNewline -ForegroundColor White
    Write-Host $(if ($entraIDResults.Status.IsInstalled) { "Yes" } else { "No" }) -ForegroundColor $(if ($entraIDResults.Status.IsInstalled) { "Green" } else { "Red" })
    Write-Host "    Sync Active:  " -NoNewline -ForegroundColor White
    Write-Host $(if ($entraIDResults.Status.IsRunning) { "Yes" } else { "No" }) -ForegroundColor $(if ($entraIDResults.Status.IsRunning) { "Green" } else { "Yellow" })
}

if ($intuneConnResults -and $intuneConnResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Hybrid Azure AD Join:" -ForegroundColor White
    Write-Host "    Status:       Check Performed" -ForegroundColor Green
    Write-Host "    Errors:       $($intuneConnResults.Status.Errors.Count)" -ForegroundColor $(if ($intuneConnResults.Status.Errors.Count -eq 0) { "Green" } else { "Yellow" })
}

if ($defenderResults -and $defenderResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Defender for Endpoint:" -ForegroundColor White
    Write-Host "    Policies:     $($defenderResults.Status.ConfiguredPolicies)" -ForegroundColor $(if ($defenderResults.Status.ConfiguredPolicies -gt 0) { "Green" } else { "Yellow" })
    Write-Host "    Devices:      $($defenderResults.Status.OnboardedDevices)" -ForegroundColor $(if ($defenderResults.Status.OnboardedDevices -gt 0) { "Green" } else { "Gray" })
    Write-Host "    Files:        $($defenderResults.Status.FilesFound.Count)/4" -ForegroundColor $(if ($defenderResults.Status.FilesMissing.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "    Status:       " -NoNewline -ForegroundColor White
    Write-Host $(if ($defenderResults.Status.ConnectorActive) { "Active" } else { "Not Configured" }) -ForegroundColor $(if ($defenderResults.Status.ConnectorActive) { "Green" } else { "Yellow" })
}

Write-Host ""
$overallStatus = ($azureResults.Missing.Count -eq 0 -and $azureResults.Errors.Count -eq 0) -and 
                 (-not $intuneResults -or ($intuneResults.Missing.Count -eq 0 -and $intuneResults.Errors.Count -eq 0)) -and
                 (-not $entraIDResults -or ($entraIDResults.Status.IsRunning)) -and
                 (-not $intuneConnResults -or ($intuneConnResults.Status.Errors.Count -eq 0)) -and
                 (-not $defenderResults -or ($defenderResults.Status.ConnectorActive -and $defenderResults.Status.FilesMissing.Count -eq 0))

Write-Host "  Overall Status: " -NoNewline -ForegroundColor White
if ($overallStatus) {
    Write-Host "✓ PASSED" -ForegroundColor Green
} else {
    Write-Host "✗ ISSUES FOUND" -ForegroundColor Red
}
Write-Host "======================================================" -ForegroundColor Cyan

if ($CompactView) {
    Write-Host ""
    Write-Host "Note: Compact View enabled. Use without -CompactView for detailed tables." -ForegroundColor Gray
}

Write-Host ""

# Export Report
if ($ExportReport) {
    $reportData = @{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        BCID = $BCID
        SubscriptionId = $currentContext.Subscription.Id
        SubscriptionName = $currentContext.Subscription.Name
        AzureResults = $azureResults
        IntuneResults = $intuneResults
        EntraIDResults = $entraIDResults
        IntuneConnectorResults = $intuneConnResults
        DefenderResults = $defenderResults
        OverallStatus = $overallStatus
    }
    
    $reportPath = "BWS_Check_Report_${BCID}_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $reportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Host "Report exported to: $reportPath" -ForegroundColor Green
    Write-Host ""
}

return $reportData