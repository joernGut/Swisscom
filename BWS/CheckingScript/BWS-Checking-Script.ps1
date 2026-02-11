<#
.SYNOPSIS
    BWS (Business Workplace Service) Checking Script with GUI support
.DESCRIPTION
    Checks Azure resources and Intune policies for BWS environments
.PARAMETER BCID
    Business Continuity ID
.PARAMETER CustomerName
    Name of the customer (optional, used in HTML report)
.PARAMETER SubscriptionId
    Azure Subscription ID (optional)
.PARAMETER ExportReport
    Export results to HTML file
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
.NOTES
    Version: 2.1.0
    Datum: 2025-02-11
    Autor: BWS PowerShell Script
.EXAMPLE
    .\BWS-Checking-Script.ps1 -BCID "1234" -CustomerName "Contoso AG"
.EXAMPLE
    .\BWS-Checking-Script.ps1 -BCID "1234" -CustomerName "Contoso AG" -GUI
.EXAMPLE
    .\BWS-Checking-Script.ps1 -BCID "1234" -CustomerName "Contoso AG" -ExportReport
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$BCID = "0000",
    
    [Parameter(Mandatory=$false)]
    [string]$CustomerName = "",
    
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
    [switch]$SkipSoftware,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipSharePoint,
    
    [Parameter(Mandatory=$false)]
    [string]$SharePointUrl,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipTeams,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("HTML", "PDF", "Both")]
    [string]$ExportFormat = "HTML",
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowAllPolicies,
    
    [Parameter(Mandatory=$false)]
    [switch]$CompactView,
    
    [Parameter(Mandatory=$false)]
    [switch]$GUI
)

# Script Version
$script:Version = "2.1.0"

#============================================================================
# Global Variables and Configuration
#============================================================================

# PowerShell Version Check
$psVersion = $PSVersionTable.PSVersion.Major
$psEdition = $PSVersionTable.PSEdition

if ($psVersion -ge 7 -or $psEdition -eq "Core") {
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Yellow
    Write-Host "  BWS-Checking-Script v$script:Version" -ForegroundColor Cyan
    Write-Host "  ⚠ WARNUNG: PowerShell Version Inkompatibilität" -ForegroundColor Yellow
    Write-Host "======================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Sie verwenden: PowerShell $($PSVersionTable.PSVersion) ($psEdition)" -ForegroundColor Yellow
    Write-Host "Empfohlen:     PowerShell 5.1 (Desktop)" -ForegroundColor Green
    Write-Host ""
    Write-Host "WICHTIG:" -ForegroundColor Red
    Write-Host "  Der SharePoint-Check funktioniert NUR in PowerShell 5.1!" -ForegroundColor Red
    Write-Host "  Das Modul 'Microsoft.Online.SharePoint.PowerShell'" -ForegroundColor Red
    Write-Host "  wird in PowerShell 7/Core NICHT unterstützt." -ForegroundColor Red
    Write-Host ""
    Write-Host "6 von 7 Checks funktionieren in PowerShell 7:" -ForegroundColor Yellow
    Write-Host "  ✓ Azure Resources" -ForegroundColor Green
    Write-Host "  ✓ Intune Policies" -ForegroundColor Green
    Write-Host "  ✓ Entra ID Connect" -ForegroundColor Green
    Write-Host "  ✓ Hybrid Azure AD Join" -ForegroundColor Green
    Write-Host "  ✓ Defender for Endpoint" -ForegroundColor Green
    Write-Host "  ✓ BWS Software Packages" -ForegroundColor Green
    Write-Host "  ✗ SharePoint Configuration (FEHLT)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Optionen:" -ForegroundColor Cyan
    Write-Host "  1. Script in PowerShell 5.1 neu starten (EMPFOHLEN)" -ForegroundColor White
    Write-Host "     → Schließen Sie diese Konsole" -ForegroundColor Gray
    Write-Host "     → Öffnen Sie 'Windows PowerShell' (nicht PowerShell 7)" -ForegroundColor Gray
    Write-Host "     → Führen Sie das Script erneut aus" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  2. SharePoint-Check überspringen" -ForegroundColor White
    Write-Host "     → Fügen Sie -SkipSharePoint Parameter hinzu" -ForegroundColor Gray
    Write-Host "     → Beispiel: .\BWS-Checking-Script.ps1 -BCID '1234' -SkipSharePoint" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  3. Mit Login-GUI arbeiten" -ForegroundColor White
    Write-Host "     → Starten Sie: .\Azure-M365-Login-GUI.ps1" -ForegroundColor Gray
    Write-Host "     → Klicken Sie auf 'PowerShell 5.1' Button (Blau)" -ForegroundColor Gray
    Write-Host "     → Führen Sie das Script in der neuen Konsole aus" -ForegroundColor Gray
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Yellow
    Write-Host ""
    
    # Frage ob fortfahren
    if (-not $SkipSharePoint) {
        $continue = Read-Host "Trotzdem fortfahren? SharePoint-Check wird fehlschlagen. (J/N)"
        if ($continue -ne "J" -and $continue -ne "j" -and $continue -ne "Y" -and $continue -ne "y") {
            Write-Host ""
            Write-Host "Script abgebrochen. Bitte verwenden Sie PowerShell 5.1." -ForegroundColor Yellow
            Write-Host ""
            exit
        }
        Write-Host ""
        Write-Host "Fahre fort ohne SharePoint-Check Unterstützung..." -ForegroundColor Yellow
        Write-Host ""
    }
}

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

function Test-BWSSoftwarePackages {
    param(
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  BWS STANDARD SOFTWARE PACKAGES CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $softwareStatus = @{
        Total = 7
        Found = @()
        Missing = @()
        Errors = @()
    }
    
    # Define required BWS software packages
    $requiredSoftware = @(
        "7-Zip",
        "Adobe Reader",
        "Chocolatey",
        "Cisco AnyConnect",
        "beyond Trust Remote support",
        "Microsoft 365 Apps for Windows 10 and later",
        "UpdateChocoSoftware"
    )
    
    try {
        Write-Host "Checking BWS Standard Software Packages..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if Microsoft Graph is connected
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        
        if (-not $graphContext) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
            try {
                Connect-MgGraph -Scopes "DeviceManagementApps.Read.All" -ErrorAction Stop
                Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            } catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                $softwareStatus.Errors += "Graph connection failed"
                return @{
                    Status = $softwareStatus
                    CheckPerformed = $false
                }
            }
        }
        
        Write-Host ""
        
        # Get all Intune Win32 Apps
        try {
            Write-Host "  [Software] Retrieving Win32 Apps from Intune..." -ForegroundColor Gray
            $win32Apps = Get-MgDeviceAppManagementMobileApp -All -Filter "isof('microsoft.graph.win32LobApp')" -ErrorAction SilentlyContinue
            Write-Host "  [Software] Found $($win32Apps.Count) Win32 Apps" -ForegroundColor Gray
        } catch {
            Write-Host "  [Software] Error retrieving Win32 Apps: $($_.Exception.Message)" -ForegroundColor Yellow
            $win32Apps = @()
        }
        
        # Get all Microsoft Store Apps
        try {
            Write-Host "  [Software] Retrieving Microsoft Store Apps from Intune..." -ForegroundColor Gray
            $storeApps = Get-MgDeviceAppManagementMobileApp -All -Filter "isof('microsoft.graph.winGetApp')" -ErrorAction SilentlyContinue
            Write-Host "  [Software] Found $($storeApps.Count) Store Apps" -ForegroundColor Gray
        } catch {
            Write-Host "  [Software] Error retrieving Store Apps: $($_.Exception.Message)" -ForegroundColor Yellow
            $storeApps = @()
        }
        
        # Get Microsoft 365 Apps
        try {
            Write-Host "  [Software] Retrieving Microsoft 365 Apps from Intune..." -ForegroundColor Gray
            # Try with filter first
            $m365Apps = Get-MgDeviceAppManagementMobileApp -All -Filter "isof('microsoft.graph.officeSuiteApp')" -ErrorAction SilentlyContinue
            
            # If filter doesn't work, get all apps and filter manually
            if (-not $m365Apps -or $m365Apps.Count -eq 0) {
                $allMobileApps = Get-MgDeviceAppManagementMobileApp -All -ErrorAction SilentlyContinue
                $m365Apps = $allMobileApps | Where-Object { 
                    $_.'@odata.type' -eq '#microsoft.graph.officeSuiteApp' -or
                    $_.DisplayName -like '*Microsoft 365 Apps*' -or
                    $_.DisplayName -like '*Office 365*'
                }
            }
            
            Write-Host "  [Software] Found $($m365Apps.Count) Office Suite Apps" -ForegroundColor Gray
        } catch {
            Write-Host "  [Software] Error retrieving Microsoft 365 Apps: $($_.Exception.Message)" -ForegroundColor Yellow
            $m365Apps = @()
        }
        
        Write-Host ""
        
        # Combine all apps
        $allApps = @()
        if ($win32Apps) { $allApps += $win32Apps }
        if ($storeApps) { $allApps += $storeApps }
        if ($m365Apps) { $allApps += $m365Apps }
        
        Write-Host "Total apps in Intune: $($allApps.Count)" -ForegroundColor White
        Write-Host ""
        
        # Check each required software
        foreach ($software in $requiredSoftware) {
            Write-Host "  [Software] " -NoNewline -ForegroundColor Gray
            Write-Host "Checking for '$software'..." -NoNewline
            
            # Search for the software with improved matching logic
            # Try exact match first, then partial matches
            $foundApp = $null
            
            # Try 1: Exact match (case-insensitive)
            $foundApp = $allApps | Where-Object { 
                $_.DisplayName -eq $software
            } | Select-Object -First 1
            
            # Try 2: Case-insensitive partial match
            if (-not $foundApp) {
                $foundApp = $allApps | Where-Object { 
                    $_.DisplayName -like "*$software*"
                } | Select-Object -First 1
            }
            
            # Try 3: Split software name and match individual words (for complex names)
            if (-not $foundApp) {
                $words = $software -split '\s+'
                foreach ($word in $words) {
                    if ($word.Length -gt 3) {  # Only use meaningful words
                        $foundApp = $allApps | Where-Object { 
                            $_.DisplayName -like "*$word*"
                        } | Select-Object -First 1
                        
                        if ($foundApp) {
                            # Verify it's a good match by checking if at least 2 words match
                            $matchCount = 0
                            foreach ($w in $words) {
                                if ($foundApp.DisplayName -like "*$w*") {
                                    $matchCount++
                                }
                            }
                            if ($matchCount -ge 2 -or $words.Count -eq 1) {
                                break
                            } else {
                                $foundApp = $null
                            }
                        }
                    }
                }
            }
            
            if ($foundApp) {
                Write-Host " ✓ FOUND" -ForegroundColor Green
                $matchType = "Partial"
                if ($foundApp.DisplayName -eq $software) {
                    $matchType = "Exact"
                } elseif ($foundApp.DisplayName -like "*$software*") {
                    $matchType = "Partial"
                } else {
                    $matchType = "Fuzzy"
                }
                
                $softwareStatus.Found += @{
                    SoftwareName = $software
                    ActualName = $foundApp.DisplayName
                    AppId = $foundApp.Id
                    MatchType = $matchType
                }
            } else {
                Write-Host " ✗ MISSING" -ForegroundColor Red
                $softwareStatus.Missing += @{
                    SoftwareName = $software
                }
            }
        }
        
    } catch {
        Write-Host "Error during software package check: $($_.Exception.Message)" -ForegroundColor Red
        $softwareStatus.Errors += "General error: $($_.Exception.Message)"
    }
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  BWS SOFTWARE PACKAGES SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Total Required:  $($softwareStatus.Total)" -ForegroundColor White
    Write-Host "  Found:           $($softwareStatus.Found.Count)" -ForegroundColor $(if ($softwareStatus.Found.Count -eq $softwareStatus.Total) { "Green" } else { "Yellow" })
    Write-Host "  Missing:         $($softwareStatus.Missing.Count)" -ForegroundColor $(if ($softwareStatus.Missing.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Errors:          $($softwareStatus.Errors.Count)" -ForegroundColor $(if ($softwareStatus.Errors.Count -eq 0) { "Green" } else { "Yellow" })
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $CompactView) {
        if ($softwareStatus.Found.Count -gt 0) {
            Write-Host "FOUND SOFTWARE PACKAGES:" -ForegroundColor Green
            Write-Host ""
            foreach ($app in $softwareStatus.Found) {
                Write-Host "  ✓ $($app.SoftwareName)" -ForegroundColor Green
                Write-Host "    Actual Name: $($app.ActualName)" -ForegroundColor Gray
                Write-Host "    Match Type:  $($app.MatchType)" -ForegroundColor Gray
                Write-Host ""
            }
        }
        
        if ($softwareStatus.Missing.Count -gt 0) {
            Write-Host "MISSING SOFTWARE PACKAGES:" -ForegroundColor Red
            Write-Host ""
            foreach ($app in $softwareStatus.Missing) {
                Write-Host "  ✗ $($app.SoftwareName)" -ForegroundColor Red
            }
            Write-Host ""
        }
        
        if ($softwareStatus.Errors.Count -gt 0) {
            Write-Host "ERRORS/WARNINGS:" -ForegroundColor Yellow
            Write-Host ""
            $softwareStatus.Errors | ForEach-Object {
                Write-Host "  - $_" -ForegroundColor Yellow
            }
            Write-Host ""
        }
    }
    
    return @{
        Status = $softwareStatus
        CheckPerformed = $true
    }
}

function Test-SharePointConfiguration {
    param(
        [bool]$CompactView = $false,
        [string]$SharePointUrl = ""
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  SHAREPOINT CONFIGURATION CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $spConfig = @{
        Settings = @{
            SharePointExternalSharing = $null
            OneDriveExternalSharing = $null
            SiteCreation = $null
            LegacyAuthBlocked = $null
            TenantUrl = $null
            ConnectionMethod = $null
        }
        Compliant = $false
        Errors = @()
        CheckPerformed = $false
    }
    
    try {
        Write-Host "Checking SharePoint configuration..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if SPO Management Shell is available (preferred) or PnP PowerShell
        $spoModuleAvailable = $false
        $moduleType = $null
        
        if (Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell") {
            $spoModuleAvailable = $true
            $moduleType = "SPO"
        } elseif (Get-Module -ListAvailable -Name "PnP.PowerShell") {
            $spoModuleAvailable = $true
            $moduleType = "PnP.PowerShell"
        }
        
        if ($spoModuleAvailable) {
            Write-Host "  [SharePoint] Using $moduleType module" -ForegroundColor Gray
            
            # Check if already connected or need to connect
            $needsConnection = $false
            $tenant = $null
            
            try {
                if ($moduleType -eq "SPO") {
                    $tenant = Get-SPOTenant -ErrorAction Stop
                } else {
                    $tenant = Get-PnPTenant -ErrorAction Stop
                }
            } catch {
                $needsConnection = $true
            }
            
            # If not connected and URL provided, try to connect
            if ($needsConnection -and $SharePointUrl) {
                Write-Host "  [SharePoint] Not connected, attempting connection to: $SharePointUrl" -ForegroundColor Yellow
                
                try {
                    if ($moduleType -eq "SPO") {
                        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
                        Connect-SPOService -Url $SharePointUrl -ErrorAction Stop
                        Write-Host "  [SharePoint] Connected successfully" -ForegroundColor Green
                        $tenant = Get-SPOTenant -ErrorAction Stop
                    } else {
                        Connect-PnPOnline -Url $SharePointUrl -Interactive -ErrorAction Stop
                        Write-Host "  [SharePoint] Connected successfully (PnP)" -ForegroundColor Green
                        $tenant = Get-PnPTenant -ErrorAction Stop
                    }
                    $needsConnection = $false
                } catch {
                    Write-Host "  [SharePoint] Connection failed: $($_.Exception.Message)" -ForegroundColor Red
                    $spConfig.Errors += "Failed to connect to SharePoint: $($_.Exception.Message)"
                }
            }
            
            # ALL CHECKS MUST BE INSIDE THIS if ($tenant) BLOCK!
            if ($tenant) {
                $spConfig.CheckPerformed = $true
                $spConfig.Settings.ConnectionMethod = $moduleType
                
                # Store Tenant URL
                if ($tenant.RootSiteUrl) {
                    $spConfig.Settings.TenantUrl = $tenant.RootSiteUrl
                }
                
                Write-Host ""
                
                # ============================================================
                # CHECK 1: External Sharing (SharePoint and OneDrive)
                # ============================================================
                try {
                    Write-Host "  [SharePoint] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking External Sharing settings..." -NoNewline
                    
                    # SharePoint External Sharing (SharingCapability)
                    $spSharingCapability = $tenant.SharingCapability
                    
                    # OneDrive External Sharing (OneDriveSharingCapability)
                    $odSharingCapability = $tenant.OneDriveSharingCapability
                    
                    # Check SharePoint - Should be "Anyone" (3)
                    if ($spSharingCapability -eq 3 -or $spSharingCapability -eq "ExistingExternalUserSharingOnly") {
                        Write-Host " ✓ SharePoint: Anyone" -ForegroundColor Green
                        $spConfig.Settings.SharePointExternalSharing = "Anyone"
                    } else {
                        Write-Host " ⚠ SharePoint: $spSharingCapability" -ForegroundColor Yellow
                        $spConfig.Settings.SharePointExternalSharing = $spSharingCapability.ToString()
                        $spConfig.Errors += "SharePoint External Sharing should be 'Anyone'"
                    }
                    
                    # Check OneDrive - Should be "Only people in your organization" (0)
                    Write-Host "  [OneDrive]    " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking OneDrive External Sharing..." -NoNewline
                    
                    if ($odSharingCapability -eq 0 -or $odSharingCapability -eq "Disabled") {
                        Write-Host " ✓ Only Organization" -ForegroundColor Green
                        $spConfig.Settings.OneDriveExternalSharing = "Disabled"
                    } else {
                        Write-Host " ⚠ $odSharingCapability" -ForegroundColor Yellow
                        $spConfig.Settings.OneDriveExternalSharing = $odSharingCapability.ToString()
                        $spConfig.Errors += "OneDrive External Sharing should be 'Disabled'"
                    }
                    
                } catch {
                    Write-Host " ⚠ ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $spConfig.Settings.SharePointExternalSharing = "Error"
                    $spConfig.Settings.OneDriveExternalSharing = "Error"
                    $spConfig.Errors += "Error checking external sharing: $($_.Exception.Message)"
                }
                
                # ============================================================
                # CHECK 2: Site Creation
                # ============================================================
                try {
                    Write-Host "  [SharePoint] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Site Creation settings..." -NoNewline
                    
                    $restrictionEnabled = $null
                    
                    # Use Get-SPORestrictedSiteCreation cmdlet
                    if ($moduleType -eq "SPO") {
                        try {
                            Write-Host "" # New line for debug
                            Write-Host "  [DEBUG] Calling Get-SPORestrictedSiteCreation..." -ForegroundColor Cyan
                            
                            $restrictedInfo = Get-SPORestrictedSiteCreation -ErrorAction Stop
                            
                            Write-Host "  [DEBUG] Cmdlet executed successfully" -ForegroundColor Cyan
                            Write-Host "  [DEBUG] Full object:" -ForegroundColor Cyan
                            $restrictedInfo | Format-List | Out-String | ForEach-Object { Write-Host $_ -ForegroundColor Cyan }
                            
                            if ($restrictedInfo) {
                                # Get the Enabled property
                                $restrictionEnabled = $restrictedInfo.Enabled
                                
                                Write-Host "  [DEBUG] Enabled property value: $restrictionEnabled" -ForegroundColor Cyan
                                Write-Host "  [DEBUG] Enabled property type: $($restrictionEnabled.GetType().Name)" -ForegroundColor Cyan
                                
                            } else {
                                Write-Host "  [DEBUG] restrictedInfo is null!" -ForegroundColor Red
                            }
                        } catch {
                            Write-Host "" # New line
                            Write-Host "  [DEBUG] Exception caught: $($_.Exception.Message)" -ForegroundColor Red
                            Write-Host "  [DEBUG] Exception type: $($_.Exception.GetType().Name)" -ForegroundColor Red
                            $restrictionEnabled = $null
                        }
                    } else {
                        Write-Host "" # New line
                        Write-Host "  [DEBUG] Module type is not SPO: $moduleType" -ForegroundColor Yellow
                    }
                    
                    Write-Host "  [SharePoint] Final evaluation: restrictionEnabled = $restrictionEnabled" -ForegroundColor Cyan
                    
                    # Evaluate the result
                    if ($null -eq $restrictionEnabled) {
                        # Could not determine
                        Write-Host "  [SharePoint] Result: " -NoNewline -ForegroundColor Gray
                        Write-Host "Could not verify (cmdlet unavailable)" -ForegroundColor Gray
                        $spConfig.Settings.SiteCreation = "Unknown"
                    } elseif ($restrictionEnabled -eq $false) {
                        # Enabled = False → Users CAN create sites → BAD!
                        Write-Host "  [SharePoint] Result: " -NoNewline -ForegroundColor Gray
                        Write-Host "ENABLED (users can create sites - restriction not active)" -ForegroundColor Yellow
                        $spConfig.Settings.SiteCreation = "Enabled"
                        $spConfig.Errors += "Site creation should be restricted - Get-SPORestrictedSiteCreation Enabled should be True"
                    } else {
                        # Enabled = True → Users CANNOT create sites → GOOD!
                        Write-Host "  [SharePoint] Result: " -NoNewline -ForegroundColor Gray
                        Write-Host "DISABLED (site creation is restricted)" -ForegroundColor Green
                        $spConfig.Settings.SiteCreation = "Disabled"
                    }
                    
                } catch {
                    Write-Host " ⚠ ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $spConfig.Settings.SiteCreation = "Error"
                    $spConfig.Errors += "Error checking site creation: $($_.Exception.Message)"
                }
                               
                # ============================================================
                # CHECK 3: Legacy Browser Auth
                # ============================================================
                try {
                    Write-Host "  [SharePoint] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Legacy Browser Auth blocking..." -NoNewline
                    
                    $legacyAuthBlocked = $null
                    
                    # Try LegacyBrowserAuthProtocolsEnabled property
                    if ($null -ne $tenant.LegacyBrowserAuthProtocolsEnabled) {
                        $legacyAuthBlocked = -not $tenant.LegacyBrowserAuthProtocolsEnabled
                    } elseif ($null -ne $tenant.LegacyAuthProtocolsEnabled) {
                        $legacyAuthBlocked = -not $tenant.LegacyAuthProtocolsEnabled
                    }
                    
                    if ($null -eq $legacyAuthBlocked) {
                        Write-Host " ⓘ Property not available" -ForegroundColor Gray
                        $spConfig.Settings.LegacyAuthBlocked = "Unknown"
                    } elseif ($legacyAuthBlocked) {
                        Write-Host " ✓ BLOCKED" -ForegroundColor Green
                        $spConfig.Settings.LegacyAuthBlocked = $true
                    } else {
                        Write-Host " ⚠ ALLOWED" -ForegroundColor Yellow
                        $spConfig.Settings.LegacyAuthBlocked = $false
                        $spConfig.Errors += "Legacy browser auth should be blocked"
                    }
                    
                } catch {
                    Write-Host " ⚠ ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $spConfig.Settings.LegacyAuthBlocked = "Error"
                    $spConfig.Errors += "Error checking legacy auth: $($_.Exception.Message)"
                }
                
            } else {
                # Not connected
                Write-Host "  ⚠ Not connected to SharePoint" -ForegroundColor Yellow
                $spConfig.Errors += "Not connected to SharePoint Online"
                $spConfig.Settings.SharePointExternalSharing = "Not Connected"
                $spConfig.Settings.OneDriveExternalSharing = "Not Connected"
                $spConfig.Settings.SiteCreation = "Not Connected"
                $spConfig.Settings.LegacyAuthBlocked = "Not Connected"
            }
            
        } else {
            Write-Host "  ⚠ SharePoint PowerShell module not found" -ForegroundColor Yellow
            $spConfig.Errors += "SharePoint PowerShell module not installed"
        }
        
        # Determine overall compliance
        $spConfig.Compliant = ($spConfig.Settings.SharePointExternalSharing -eq "Anyone") -and
                              ($spConfig.Settings.OneDriveExternalSharing -eq "Disabled") -and
                              ($spConfig.Settings.SiteCreation -eq "Disabled") -and
                              ($spConfig.Settings.LegacyAuthBlocked -eq $true) -and
                              ($spConfig.Errors.Count -eq 0)
        
    } catch {
        Write-Host "Error during SharePoint configuration check: $($_.Exception.Message)" -ForegroundColor Red
        $spConfig.Errors += "General error: $($_.Exception.Message)"
    }
    
    # Summary output
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  SHAREPOINT CONFIGURATION SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  SharePoint Ext. Sharing: " -NoNewline -ForegroundColor White
    if ($spConfig.Settings.SharePointExternalSharing -eq "Anyone") {
        Write-Host "Anyone (✓)" -ForegroundColor Green
    } elseif ($spConfig.Settings.SharePointExternalSharing) {
        Write-Host "$($spConfig.Settings.SharePointExternalSharing) (✗)" -ForegroundColor Yellow
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  OneDrive Ext. Sharing:   " -NoNewline -ForegroundColor White
    if ($spConfig.Settings.OneDriveExternalSharing -eq "Disabled") {
        Write-Host "Only Organization (✓)" -ForegroundColor Green
    } elseif ($spConfig.Settings.OneDriveExternalSharing) {
        Write-Host "$($spConfig.Settings.OneDriveExternalSharing) (✗)" -ForegroundColor Yellow
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  Site Creation:           " -NoNewline -ForegroundColor White
    if ($spConfig.Settings.SiteCreation -eq "Disabled") {
        Write-Host "Disabled (✓)" -ForegroundColor Green
    } elseif ($spConfig.Settings.SiteCreation -eq "Enabled") {
        Write-Host "Enabled (✗)" -ForegroundColor Yellow
    } elseif ($spConfig.Settings.SiteCreation) {
        Write-Host "$($spConfig.Settings.SiteCreation) (?)" -ForegroundColor Gray
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  Legacy Auth Blocked:     " -NoNewline -ForegroundColor White
    if ($spConfig.Settings.LegacyAuthBlocked -eq $true) {
        Write-Host "Yes (✓)" -ForegroundColor Green
    } elseif ($spConfig.Settings.LegacyAuthBlocked -eq $false) {
        Write-Host "No (✗)" -ForegroundColor Yellow
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    return @{
        Status = $spConfig
        CheckPerformed = $spConfig.CheckPerformed
    }
}

#============================================================================
# TEAMS CONFIGURATION CHECK
#============================================================================
function Test-TeamsConfiguration {
    param(
        [bool]$CompactView = $false
    )
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  TEAMS CONFIGURATION CHECK" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $teamsConfig = @{
        Settings = @{
            ExternalAccessEnabled = $null
            CloudStorageCitrix = $null
            CloudStorageDropbox = $null
            CloudStorageBox = $null
            CloudStorageGoogleDrive = $null
            CloudStorageEgnyte = $null
            AnonymousUsersCanJoin = $null
            AnonymousUsersCanStartMeeting = $null
            DefaultPresenterRole = $null
        }
        Compliant = $false
        Errors = @()
        CheckPerformed = $false
    }
    
    try {
        Write-Host "Checking Teams configuration..." -ForegroundColor Yellow
        Write-Host ""
        
        # Check if Teams PowerShell module is available
        $teamsModuleAvailable = $false
        
        if (Get-Module -ListAvailable -Name "MicrosoftTeams") {
            $teamsModuleAvailable = $true
        }
        
        if ($teamsModuleAvailable) {
            Write-Host "  [Teams] Using MicrosoftTeams module" -ForegroundColor Gray
            
            # Check if connected to Teams
            $teamsConnected = $false
            
            try {
                $csConfig = Get-CsTeamsClientConfiguration -ErrorAction Stop
                $teamsConnected = $true
            } catch {
                Write-Host "  [Teams] Not connected to Microsoft Teams" -ForegroundColor Yellow
                Write-Host "  [Teams] Please connect first: Connect-MicrosoftTeams" -ForegroundColor Gray
            }
            
            if ($teamsConnected) {
                $teamsConfig.CheckPerformed = $true
                Write-Host ""
                
                # ============================================================
                # CHECK 1: Meetings with unmanaged MS Accounts
                # ============================================================
                try {
                    Write-Host "  [Teams] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Meetings with unmanaged MS Accounts..." -NoNewline
                    
                    $federationConfig = Get-CsTenantFederationConfiguration -ErrorAction Stop
                    $externalAccessEnabled = $federationConfig.AllowTeamsConsumer
                    
                    if ($externalAccessEnabled -eq $false) {
                        Write-Host " ✓ DISABLED (unmanaged Teams blocked)" -ForegroundColor Green
                        $teamsConfig.Settings.ExternalAccessEnabled = $false
                    } else {
                        Write-Host " ⚠ ENABLED (unmanaged Teams allowed)" -ForegroundColor Yellow
                        $teamsConfig.Settings.ExternalAccessEnabled = $true
                        $teamsConfig.Errors += "External access to unmanaged Teams should be disabled"
                    }
                    
                } catch {
                    Write-Host " ⚠ ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $teamsConfig.Settings.ExternalAccessEnabled = "Error"
                    $teamsConfig.Errors += "Error checking external access: $($_.Exception.Message)"
                }
                
                # ============================================================
                # CHECK 2: Cloud Storage Providers
                # ============================================================
                try {
                    Write-Host "  [Teams] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Cloud Storage settings..." -NoNewline
                    
                    $clientConfig = Get-CsTeamsClientConfiguration -ErrorAction Stop
                    
                    # Check each cloud storage provider
                    $allDisabled = $true
                    $enabledProviders = @()
                    
                    # Citrix Files
                    $citrixValue = $clientConfig.AllowCitrixContentSharing
                    if ($citrixValue -eq $true -or $citrixValue -eq "Enabled") {
                        $allDisabled = $false
                        $enabledProviders += "Citrix"
                    }
                    $teamsConfig.Settings.CloudStorageCitrix = if ($citrixValue -eq $false -or $citrixValue -eq "Disabled") { "Disabled" } else { "Enabled" }
                    
                    # Dropbox
                    $dropboxValue = $clientConfig.AllowDropBox
                    if ($dropboxValue -eq $true -or $dropboxValue -eq "Enabled") {
                        $allDisabled = $false
                        $enabledProviders += "Dropbox"
                    }
                    $teamsConfig.Settings.CloudStorageDropbox = if ($dropboxValue -eq $false -or $dropboxValue -eq "Disabled") { "Disabled" } else { "Enabled" }
                    
                    # Box
                    $boxValue = $clientConfig.AllowBox
                    if ($boxValue -eq $true -or $boxValue -eq "Enabled") {
                        $allDisabled = $false
                        $enabledProviders += "Box"
                    }
                    $teamsConfig.Settings.CloudStorageBox = if ($boxValue -eq $false -or $boxValue -eq "Disabled") { "Disabled" } else { "Enabled" }
                    
                    # Google Drive
                    $googleValue = $clientConfig.AllowGoogleDrive
                    if ($googleValue -eq $true -or $googleValue -eq "Enabled") {
                        $allDisabled = $false
                        $enabledProviders += "Google Drive"
                    }
                    $teamsConfig.Settings.CloudStorageGoogleDrive = if ($googleValue -eq $false -or $googleValue -eq "Disabled") { "Disabled" } else { "Enabled" }
                    
                    # Egnyte
                    $egnyteValue = $clientConfig.AllowEgnyte
                    if ($egnyteValue -eq $true -or $egnyteValue -eq "Enabled") {
                        $allDisabled = $false
                        $enabledProviders += "Egnyte"
                    }
                    $teamsConfig.Settings.CloudStorageEgnyte = if ($egnyteValue -eq $false -or $egnyteValue -eq "Disabled") { "Disabled" } else { "Enabled" }
                    
                    if ($allDisabled) {
                        Write-Host " ✓ ALL DISABLED" -ForegroundColor Green
                    } else {
                        Write-Host " ⚠ ENABLED: $($enabledProviders -join ', ')" -ForegroundColor Yellow
                        $teamsConfig.Errors += "All cloud storage providers should be disabled: $($enabledProviders -join ', ')"
                    }
                    
                } catch {
                    Write-Host " ⚠ ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $teamsConfig.Errors += "Error checking cloud storage: $($_.Exception.Message)"
                }
                
                # ============================================================
                # CHECK 3: Meeting & Lobby Settings
                # ============================================================
                try {
                    Write-Host "  [Teams] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Meeting & Lobby settings..." -NoNewline
                    
                    $meetingConfig = Get-CsTeamsMeetingConfiguration -ErrorAction Stop
                    
                    # Anonymous users can join
                    $anonymousCanJoin = $meetingConfig.DisableAnonymousJoin -eq $false
                    $teamsConfig.Settings.AnonymousUsersCanJoin = if ($anonymousCanJoin) { "Enabled" } else { "Disabled" }
                    
                    # Anonymous users can start meeting
                    $anonymousCanStart = -not $meetingConfig.EnabledAnonymousUsersRequireLobby
                    $teamsConfig.Settings.AnonymousUsersCanStartMeeting = if ($anonymousCanStart) { "Enabled" } else { "Disabled" }
                    
                    $meetingIssues = @()
                    
                    if ($anonymousCanJoin) {
                        $meetingIssues += "Anonymous join enabled"
                    }
                    
                    if ($anonymousCanStart) {
                        $meetingIssues += "Anonymous can start meetings"
                    }
                    
                    if ($meetingIssues.Count -eq 0) {
                        Write-Host " ✓ COMPLIANT" -ForegroundColor Green
                    } else {
                        Write-Host " ⚠ ISSUES: $($meetingIssues -join ', ')" -ForegroundColor Yellow
                        $teamsConfig.Errors += "Anonymous users should not be able to join or start meetings"
                    }
                    
                } catch {
                    Write-Host " ⚠ ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $teamsConfig.Settings.AnonymousUsersCanJoin = "Error"
                    $teamsConfig.Settings.AnonymousUsersCanStartMeeting = "Error"
                    $teamsConfig.Errors += "Error checking meeting settings: $($_.Exception.Message)"
                }
                
                # ============================================================
                # CHECK 4: Content Sharing - Who can present
                # ============================================================
                try {
                    Write-Host "  [Teams] " -NoNewline -ForegroundColor Gray
                    Write-Host "Checking Content Sharing settings..." -NoNewline
                    
                    $meetingPolicy = Get-CsTeamsMeetingPolicy -Identity Global -ErrorAction Stop
                    $presenterRole = $meetingPolicy.DesignatedPresenterRoleMode
                    
                    $teamsConfig.Settings.DefaultPresenterRole = $presenterRole
                    
                    # EveryoneUserOverride means "Everyone" can present
                    if ($presenterRole -eq "EveryoneUserOverride") {
                        Write-Host " ✓ EVERYONE (Compliant)" -ForegroundColor Green
                    } else {
                        Write-Host " ⚠ $presenterRole (Non-Compliant)" -ForegroundColor Yellow
                        $teamsConfig.Errors += "Default presenter role should be 'Everyone' (EveryoneUserOverride)"
                    }
                    
                } catch {
                    Write-Host " ⚠ ERROR: $($_.Exception.Message)" -ForegroundColor Yellow
                    $teamsConfig.Settings.DefaultPresenterRole = "Error"
                    $teamsConfig.Errors += "Error checking presenter settings: $($_.Exception.Message)"
                }
                
            } else {
                Write-Host "  ⚠ Not connected to Microsoft Teams" -ForegroundColor Yellow
                $teamsConfig.Errors += "Not connected to Microsoft Teams"
            }
            
        } else {
            Write-Host "  ⚠ MicrosoftTeams PowerShell module not found" -ForegroundColor Yellow
            Write-Host "  Install with: Install-Module -Name MicrosoftTeams" -ForegroundColor Gray
            $teamsConfig.Errors += "MicrosoftTeams PowerShell module not installed"
        }
        
        # Determine overall compliance
        $teamsConfig.Compliant = ($teamsConfig.Settings.ExternalAccessEnabled -eq $false) -and
                                  ($teamsConfig.Settings.CloudStorageCitrix -eq "Disabled") -and
                                  ($teamsConfig.Settings.CloudStorageDropbox -eq "Disabled") -and
                                  ($teamsConfig.Settings.CloudStorageBox -eq "Disabled") -and
                                  ($teamsConfig.Settings.CloudStorageGoogleDrive -eq "Disabled") -and
                                  ($teamsConfig.Settings.CloudStorageEgnyte -eq "Disabled") -and
                                  ($teamsConfig.Settings.AnonymousUsersCanJoin -eq "Disabled") -and
                                  ($teamsConfig.Settings.AnonymousUsersCanStartMeeting -eq "Disabled") -and
                                  ($teamsConfig.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") -and
                                  ($teamsConfig.Errors.Count -eq 0)
        
    } catch {
        Write-Host "Error during Teams configuration check: $($_.Exception.Message)" -ForegroundColor Red
        $teamsConfig.Errors += "General error: $($_.Exception.Message)"
    }
    
    # Summary output
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  TEAMS CONFIGURATION SUMMARY" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Meetings w/ unmanaged MS: " -NoNewline -ForegroundColor White
    if ($teamsConfig.Settings.ExternalAccessEnabled -eq $false) {
        Write-Host "Disabled (✓)" -ForegroundColor Green
    } elseif ($teamsConfig.Settings.ExternalAccessEnabled -eq $true) {
        Write-Host "Enabled (✗)" -ForegroundColor Yellow
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  Cloud Storage:           " -NoNewline -ForegroundColor White
    $allStorageDisabled = ($teamsConfig.Settings.CloudStorageCitrix -eq "Disabled") -and
                          ($teamsConfig.Settings.CloudStorageDropbox -eq "Disabled") -and
                          ($teamsConfig.Settings.CloudStorageBox -eq "Disabled") -and
                          ($teamsConfig.Settings.CloudStorageGoogleDrive -eq "Disabled") -and
                          ($teamsConfig.Settings.CloudStorageEgnyte -eq "Disabled")
    
    if ($allStorageDisabled) {
        Write-Host "All Disabled (✓)" -ForegroundColor Green
    } else {
        # Build list of enabled providers
        $enabledList = @()
        if ($teamsConfig.Settings.CloudStorageCitrix -eq "Enabled") { $enabledList += "Citrix" }
        if ($teamsConfig.Settings.CloudStorageDropbox -eq "Enabled") { $enabledList += "Dropbox" }
        if ($teamsConfig.Settings.CloudStorageBox -eq "Enabled") { $enabledList += "Box" }
        if ($teamsConfig.Settings.CloudStorageGoogleDrive -eq "Enabled") { $enabledList += "Google Drive" }
        if ($teamsConfig.Settings.CloudStorageEgnyte -eq "Enabled") { $enabledList += "Egnyte" }
        
        if ($enabledList.Count -gt 0) {
            Write-Host "Enabled: $($enabledList -join ', ') (✗)" -ForegroundColor Yellow
        } else {
            Write-Host "Unknown" -ForegroundColor Gray
        }
    }
    
    Write-Host "  Anonymous Join:          " -NoNewline -ForegroundColor White
    if ($teamsConfig.Settings.AnonymousUsersCanJoin -eq "Disabled") {
        Write-Host "Disabled (✓)" -ForegroundColor Green
    } elseif ($teamsConfig.Settings.AnonymousUsersCanJoin -eq "Enabled") {
        Write-Host "Enabled (✗)" -ForegroundColor Yellow
    } elseif ($teamsConfig.Settings.AnonymousUsersCanJoin -eq "Error") {
        Write-Host "Error - Could not check" -ForegroundColor Red
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  Anonymous Can Start:     " -NoNewline -ForegroundColor White
    if ($teamsConfig.Settings.AnonymousUsersCanStartMeeting -eq "Disabled") {
        Write-Host "Disabled (✓)" -ForegroundColor Green
    } elseif ($teamsConfig.Settings.AnonymousUsersCanStartMeeting -eq "Enabled") {
        Write-Host "Enabled (✗)" -ForegroundColor Yellow
    } elseif ($teamsConfig.Settings.AnonymousUsersCanStartMeeting -eq "Error") {
        Write-Host "Error - Could not check" -ForegroundColor Red
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "  Who Can Present:         " -NoNewline -ForegroundColor White
    if ($teamsConfig.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") {
        Write-Host "Everyone (✓)" -ForegroundColor Green
    } elseif ($teamsConfig.Settings.DefaultPresenterRole) {
        Write-Host "$($teamsConfig.Settings.DefaultPresenterRole) (✗)" -ForegroundColor Yellow
    } else {
        Write-Host "Not Checked" -ForegroundColor Gray
    }
    
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    return @{
        Status = $teamsConfig
        CheckPerformed = $teamsConfig.CheckPerformed
    }
}

function Export-HTMLReport {
    param(
        [string]$BCID,
        [string]$CustomerName,
        [string]$SubscriptionName,
        [object]$AzureResults,
        [object]$IntuneResults,
        [object]$EntraIDResults,
        [object]$IntuneConnResults,
        [object]$DefenderResults,
        [object]$SoftwareResults,
        [object]$SharePointResults,
        [object]$TeamsResults,
        [bool]$OverallStatus
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $reportDate = Get-Date -Format "yyyyMMdd_HHmmss"
    
    # Include customer name in filename if provided
    if ($CustomerName) {
        $safeCustomerName = $CustomerName -replace '[^\w\s-]', '' -replace '\s+', '_'
        $reportPath = "BWS_Check_Report_${safeCustomerName}_${BCID}_${reportDate}.html"
    } else {
        $reportPath = "BWS_Check_Report_${BCID}_${reportDate}.html"
    }
    
    # Build HTML
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Business Workplace Services Check Report - $(if ($CustomerName) { "$CustomerName - " })BCID $BCID</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #0082C9 0%, #001155 100%);
            padding: 20px;
            color: #333;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #0082C9 0%, #001155 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        
        .header .meta {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .status-badge {
            display: inline-block;
            padding: 10px 30px;
            border-radius: 25px;
            font-weight: bold;
            font-size: 1.2em;
            margin-top: 15px;
            text-transform: uppercase;
        }
        
        .status-pass {
            background: #10b981;
            color: white;
        }
        
        .status-fail {
            background: #ef4444;
            color: white;
        }
        
        .toc {
            background: #f8fafc;
            padding: 30px;
            border-bottom: 3px solid #e2e8f0;
        }
        
        .toc h2 {
            color: #1e293b;
            margin-bottom: 20px;
            font-size: 1.8em;
        }
        
        .toc ul {
            list-style: none;
        }
        
        .toc li {
            margin: 12px 0;
        }
        
        .toc a {
            color: #0082C9;
            text-decoration: none;
            font-size: 1.1em;
            transition: all 0.3s;
            display: inline-block;
        }
        
        .toc a:hover {
            color: #001155;
            transform: translateX(5px);
        }
        
        .content {
            padding: 30px;
        }
        
        .section {
            margin-bottom: 40px;
            padding: 25px;
            background: #f8fafc;
            border-radius: 8px;
            border-left: 5px solid #0082C9;
        }
        
        .section h2 {
            color: #1e293b;
            margin-bottom: 20px;
            font-size: 1.8em;
            display: flex;
            align-items: center;
        }
        
        .section-icon {
            width: 40px;
            height: 40px;
            margin-right: 15px;
            background: #0082C9;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 1.5em;
        }
        
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }
        
        .summary-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
        }
        
        .summary-card h3 {
            color: #64748b;
            font-size: 0.9em;
            margin-bottom: 10px;
            text-transform: uppercase;
        }
        
        .summary-card .value {
            font-size: 2.5em;
            font-weight: bold;
            color: #1e293b;
        }
        
        .summary-card.success .value {
            color: #10b981;
        }
        
        .summary-card.warning .value {
            color: #f59e0b;
        }
        
        .summary-card.error .value {
            color: #ef4444;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        thead {
            background: #0082C9;
            color: white;
        }
        
        th {
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }
        
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #e2e8f0;
        }
        
        tr:last-child td {
            border-bottom: none;
        }
        
        tbody tr:hover {
            background: #f8fafc;
        }
        
        .status-icon {
            font-size: 1.2em;
            font-weight: bold;
        }
        
        .status-found {
            color: #10b981;
        }
        
        .status-missing {
            color: #ef4444;
        }
        
        .status-error {
            color: #f59e0b;
        }
        
        .info-list {
            list-style: none;
            margin: 15px 0;
        }
        
        .info-list li {
            padding: 10px;
            margin: 8px 0;
            background: white;
            border-radius: 5px;
            border-left: 3px solid #0082C9;
        }
        
        .footer {
            background: #1e293b;
            color: white;
            text-align: center;
            padding: 20px;
            font-size: 0.9em;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            
            .container {
                box-shadow: none;
            }
            
            .toc a {
                color: #000;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🛡️ Business Workplace Services Check Report</h1>
"@
    
    if ($CustomerName) {
        $html += @"
            <div style="font-size: 1.8em; font-weight: bold; margin: 15px 0; text-shadow: 1px 1px 2px rgba(0,0,0,0.2);">
                $CustomerName
            </div>
"@
    }
    
    $html += @"
            <div class="meta">
                <strong>BCID:</strong> <span style="font-size: 1.3em; font-weight: bold;">$BCID</span> | 
                <strong>Date:</strong> $timestamp | 
                <strong>Subscription:</strong> $SubscriptionName
            </div>
            <div class="status-badge $(if ($OverallStatus) { 'status-pass' } else { 'status-fail' })">
                $(if ($OverallStatus) { '✓ Passed' } else { '✗ Issues Found' })
            </div>
        </div>
        
        <div class="toc">
            <h2>📋 Table of Contents</h2>
            <ul>
                <li><a href="#summary">→ Executive Summary</a></li>
                <li><a href="#azure">→ Azure Resources</a></li>
                <li><a href="#intune">→ Intune Policies</a></li>
                <li><a href="#software">→ BWS Software Packages</a></li>
                <li><a href="#sharepoint">→ SharePoint Configuration</a></li>
                <li><a href="#entra">→ Entra ID Connect</a></li>
                <li><a href="#hybrid">→ Hybrid Azure AD Join</a></li>
                <li><a href="#defender">→ Defender for Endpoint</a></li>
            </ul>
        </div>
        
        <div class="content">
"@

    # Summary Section
    $html += @"
            <div class="section" id="summary">
                <h2><span class="section-icon">📊</span>Executive Summary</h2>
                <div class="summary-grid">
"@

    if ($AzureResults) {
        $azureClass = if ($AzureResults.Missing.Count -eq 0) { "success" } else { "error" }
        $html += @"
                    <div class="summary-card $azureClass">
                        <h3>Azure Resources</h3>
                        <div class="value">$($AzureResults.Found.Count)/$($AzureResults.Total)</div>
                        <p>Found</p>
                    </div>
"@
    }

    if ($IntuneResults -and $IntuneResults.CheckPerformed) {
        $intuneClass = if ($IntuneResults.Missing.Count -eq 0) { "success" } else { "error" }
        $html += @"
                    <div class="summary-card $intuneClass">
                        <h3>Intune Policies</h3>
                        <div class="value">$($IntuneResults.Found.Count)/$($IntuneResults.Total)</div>
                        <p>Found</p>
                    </div>
"@
    }

    if ($EntraIDResults -and $EntraIDResults.CheckPerformed) {
        $entraClass = if ($EntraIDResults.Status.IsRunning) { "success" } else { "error" }
        $html += @"
                    <div class="summary-card $entraClass">
                        <h3>Entra ID Sync</h3>
                        <div class="value">$(if ($EntraIDResults.Status.IsRunning) { '✓' } else { '✗' })</div>
                        <p>$(if ($EntraIDResults.Status.IsRunning) { 'Active' } else { 'Inactive' })</p>
                    </div>
"@
    }

    if ($DefenderResults -and $DefenderResults.CheckPerformed) {
        $defenderClass = if ($DefenderResults.Status.ConnectorActive -and $DefenderResults.Status.FilesMissing.Count -eq 0) { "success" } elseif ($DefenderResults.Status.ConnectorActive) { "warning" } else { "error" }
        $html += @"
                    <div class="summary-card $defenderClass">
                        <h3>Defender Status</h3>
                        <div class="value">$(if ($DefenderResults.Status.ConnectorActive) { '✓' } else { '✗' })</div>
                        <p>$($DefenderResults.Status.FilesFound.Count)/4 Files</p>
                    </div>
"@
    }

    if ($SoftwareResults -and $SoftwareResults.CheckPerformed) {
        $softwareClass = if ($SoftwareResults.Status.Missing.Count -eq 0) { "success" } else { "error" }
        $html += @"
                    <div class="summary-card $softwareClass">
                        <h3>BWS Software</h3>
                        <div class="value">$($SoftwareResults.Status.Found.Count)/$($SoftwareResults.Status.Total)</div>
                        <p>Packages</p>
                    </div>
"@
    }

    if ($SharePointResults -and $SharePointResults.CheckPerformed) {
        $spClass = if ($SharePointResults.Status.Compliant) { "success" } else { "warning" }
        $html += @"
                    <div class="summary-card $spClass">
                        <h3>SharePoint Config</h3>
                        <div class="value">$(if ($SharePointResults.Status.Compliant) { '✓' } else { '⚠' })</div>
                        <p>$(if ($SharePointResults.Status.Compliant) { 'Compliant' } else { 'Issues' })</p>
                    </div>
"@
    }

    $html += @"
                </div>
            </div>
"@

    # Azure Resources Section
    if ($AzureResults) {
        $html += @"
            <div class="section" id="azure">
                <h2><span class="section-icon">☁️</span>Azure Resources</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Resource Type</th>
                            <th>Resource Name</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($resource in $AzureResults.Found) {
            $html += @"
                        <tr>
                            <td><span class="status-icon status-found">✓</span></td>
                            <td>$($resource.Type)</td>
                            <td>$($resource.Name)</td>
                        </tr>
"@
        }

        foreach ($resource in $AzureResults.Missing) {
            $html += @"
                        <tr>
                            <td><span class="status-icon status-missing">✗</span></td>
                            <td>$($resource.Type)</td>
                            <td>$($resource.Name)</td>
                        </tr>
"@
        }

        $html += @"
                    </tbody>
                </table>
            </div>
"@
    }

    # Intune Policies Section
    if ($IntuneResults -and $IntuneResults.CheckPerformed) {
        $html += @"
            <div class="section" id="intune">
                <h2><span class="section-icon">🔒</span>Intune Policies</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Policy Name</th>
                            <th>Match Type</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($policy in $IntuneResults.Found) {
            $matchType = if ($policy.MatchType) { $policy.MatchType } else { "Exact" }
            $html += @"
                        <tr>
                            <td><span class="status-icon status-found">✓</span></td>
                            <td>$($policy.PolicyName)</td>
                            <td>$matchType</td>
                        </tr>
"@
        }

        foreach ($policy in $IntuneResults.Missing) {
            $html += @"
                        <tr>
                            <td><span class="status-icon status-missing">✗</span></td>
                            <td>$($policy.PolicyName)</td>
                            <td>Not Found</td>
                        </tr>
"@
        }

        $html += @"
                    </tbody>
                </table>
            </div>
"@
    }

    # BWS Software Packages Section
    if ($SoftwareResults -and $SoftwareResults.CheckPerformed) {
        $html += @"
            <div class="section" id="software">
                <h2><span class="section-icon">📦</span>BWS Standard Software Packages</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Status</th>
                            <th>Required Software</th>
                            <th>Actual Name</th>
                            <th>Match Type</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($app in $SoftwareResults.Status.Found) {
            $html += @"
                        <tr>
                            <td><span class="status-icon status-found">✓</span></td>
                            <td>$($app.SoftwareName)</td>
                            <td>$($app.ActualName)</td>
                            <td>$($app.MatchType)</td>
                        </tr>
"@
        }

        foreach ($app in $SoftwareResults.Status.Missing) {
            $html += @"
                        <tr>
                            <td><span class="status-icon status-missing">✗</span></td>
                            <td>$($app.SoftwareName)</td>
                            <td>Not Found</td>
                            <td>-</td>
                        </tr>
"@
        }

        $html += @"
                    </tbody>
                </table>
            </div>
"@
    }

    # SharePoint Configuration Section
    if ($SharePointResults -and $SharePointResults.CheckPerformed) {
        $html += @"
            <div class="section" id="sharepoint">
                <h2><span class="section-icon">🌐</span>SharePoint Configuration</h2>
"@
        
        # Add Tenant URL if available
        if ($SharePointResults.Status.Settings.TenantUrl) {
            $html += @"
                <ul class="info-list">
                    <li><strong>Tenant URL:</strong> $($SharePointResults.Status.Settings.TenantUrl)</li>
                    <li><strong>Connection Method:</strong> $($SharePointResults.Status.Settings.ConnectionMethod)</li>
                </ul>
                <h3>Configuration Settings:</h3>
"@
        }
        
        $html += @"
                <ul class="info-list">
                    <li><strong>SharePoint External Sharing:</strong> $(if ($SharePointResults.Status.Settings.SharePointExternalSharing -eq 'Anyone') { '<span class="status-found">✓ Anyone (Compliant)</span>' } elseif ($SharePointResults.Status.Settings.SharePointExternalSharing -like '*Unknown*' -or $SharePointResults.Status.Settings.SharePointExternalSharing -like '*Not Connected*') { '<span class="status-error">⚠ Could not verify - Not connected</span>' } elseif ($SharePointResults.Status.Settings.SharePointExternalSharing) { "<span class='status-error'>⚠ $($SharePointResults.Status.Settings.SharePointExternalSharing) (Non-Compliant - should be 'Anyone')</span>" } else { '<span class="status-error">⚠ Check not performed</span>' })</li>
                    <li><strong>OneDrive External Sharing:</strong> $(if ($SharePointResults.Status.Settings.OneDriveExternalSharing -eq 'Disabled') { '<span class="status-found">✓ Only people in your organization (Compliant)</span>' } elseif ($SharePointResults.Status.Settings.OneDriveExternalSharing -like '*Unknown*' -or $SharePointResults.Status.Settings.OneDriveExternalSharing -like '*Not Connected*') { '<span class="status-error">⚠ Could not verify - Not connected</span>' } elseif ($SharePointResults.Status.Settings.OneDriveExternalSharing) { "<span class='status-error'>⚠ $($SharePointResults.Status.Settings.OneDriveExternalSharing) (Non-Compliant - should be 'Disabled')</span>" } else { '<span class="status-error">⚠ Check not performed</span>' })</li>
                    <li><strong>Site Creation:</strong> $(if ($SharePointResults.Status.Settings.SiteCreation -eq 'Disabled') { '<span class="status-found">✓ Disabled - Users cannot create sites (Compliant)</span>' } elseif ($SharePointResults.Status.Settings.SiteCreation -eq 'Enabled') { '<span class="status-error">✗ Enabled - Users can create sites (Non-Compliant)</span>' } elseif ($SharePointResults.Status.Settings.SiteCreation -like '*Unknown*') { '<span class="status-error">⚠ Could not verify</span>' } elseif ($SharePointResults.Status.Settings.SiteCreation) { "<span class='status-error'>⚠ $($SharePointResults.Status.Settings.SiteCreation)</span>" } else { '<span class="status-error">⚠ Check not performed</span>' })</li>
                    <li><strong>Legacy Browser Auth Blocked:</strong> $(if ($SharePointResults.Status.Settings.LegacyAuthBlocked -eq $true) { '<span class="status-found">✓ Yes - Legacy browser auth protocols blocked (Compliant)</span>' } elseif ($SharePointResults.Status.Settings.LegacyAuthBlocked -eq $false) { '<span class="status-error">✗ No - Legacy browser auth protocols allowed (Non-Compliant)</span>' } elseif ($SharePointResults.Status.Settings.LegacyAuthBlocked -like '*Property Not Available*') { '<span class="status-error">⚠ Property not available in tenant</span>' } else { '<span class="status-error">⚠ Check not performed</span>' })</li>
                </ul>
"@
        
        $html += "</div>"
    } elseif ($SharePointResults) {
        # SharePoint check was attempted but not performed (no connection)
        $html += @"
            <div class="section" id="sharepoint">
                <h2><span class="section-icon">🌐</span>SharePoint Configuration</h2>
                <ul class="info-list">
                    <li><strong>Status:</strong> <span class="status-error">⚠ Check not performed</span></li>
"@
        if ($SharePointResults.Status.Errors.Count -gt 0) {
            $html += "<li><strong>Reason:</strong></li></ul><ul class='info-list'>"
            foreach ($error in $SharePointResults.Status.Errors) {
                $html += "<li><span class='status-error'>⚠</span> $error</li>"
            }
        }
        $html += @"
                </ul>
                <p style="color: #666; font-style: italic;">
                    Tip: Use -SharePointUrl parameter to connect automatically:<br>
                    <code>-SharePointUrl "https://TENANT-admin.sharepoint.com"</code>
                </p>
            </div>
"@
    }

    # Teams Configuration Section
    if ($TeamsResults -and $TeamsResults.CheckPerformed) {
        $html += @"
            <div class="section" id="teams">
                <h2><span class="section-icon">💬</span>Teams Configuration</h2>
                <h3>Configuration Settings:</h3>
                <ul class="info-list">
                    <li><strong>Meetings with unmanaged MS Accounts:</strong> $(if ($TeamsResults.Status.Settings.ExternalAccessEnabled -eq $false) { '<span class="status-found">✓ Disabled (Compliant)</span>' } elseif ($TeamsResults.Status.Settings.ExternalAccessEnabled -eq $true) { '<span class="status-error">✗ Enabled (Non-Compliant)</span>' } else { '<span class="status-error">⚠ Check not performed</span>' })</li>
                    <li><strong>Cloud Storage Providers:</strong>
                        <ul style="margin-top: 5px;">
                            <li>Citrix Files: $(if ($TeamsResults.Status.Settings.CloudStorageCitrix -eq "Disabled") { '<span class="status-found">✓ Disabled</span>' } else { '<span class="status-error">✗ Enabled</span>' })</li>
                            <li>Dropbox: $(if ($TeamsResults.Status.Settings.CloudStorageDropbox -eq "Disabled") { '<span class="status-found">✓ Disabled</span>' } else { '<span class="status-error">✗ Enabled</span>' })</li>
                            <li>Box: $(if ($TeamsResults.Status.Settings.CloudStorageBox -eq "Disabled") { '<span class="status-found">✓ Disabled</span>' } else { '<span class="status-error">✗ Enabled</span>' })</li>
                            <li>Google Drive: $(if ($TeamsResults.Status.Settings.CloudStorageGoogleDrive -eq "Disabled") { '<span class="status-found">✓ Disabled</span>' } else { '<span class="status-error">✗ Enabled</span>' })</li>
                            <li>Egnyte: $(if ($TeamsResults.Status.Settings.CloudStorageEgnyte -eq "Disabled") { '<span class="status-found">✓ Disabled</span>' } else { '<span class="status-error">✗ Enabled</span>' })</li>
                        </ul>
                    </li>
                    <li><strong>Anonymous Users Can Join:</strong> $(if ($TeamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Disabled") { '<span class="status-found">✓ Disabled (Compliant)</span>' } elseif ($TeamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Enabled") { '<span class="status-error">✗ Enabled (Non-Compliant)</span>' } else { '<span class="status-error">⚠ Check not performed</span>' })</li>
                    <li><strong>Anonymous Users Can Start Meeting:</strong> $(if ($TeamsResults.Status.Settings.AnonymousUsersCanStartMeeting -eq "Disabled") { '<span class="status-found">✓ Disabled (Compliant)</span>' } elseif ($TeamsResults.Status.Settings.AnonymousUsersCanStartMeeting -eq "Enabled") { '<span class="status-error">✗ Enabled (Non-Compliant)</span>' } else { '<span class="status-error">⚠ Check not performed</span>' })</li>
                    <li><strong>Who Can Present:</strong> $(if ($TeamsResults.Status.Settings.DefaultPresenterRole -eq 'EveryoneUserOverride') { '<span class="status-found">✓ Everyone (Compliant)</span>' } elseif ($TeamsResults.Status.Settings.DefaultPresenterRole) { "<span class='status-error'>✗ $($TeamsResults.Status.Settings.DefaultPresenterRole) (Non-Compliant)</span>" } else { '<span class="status-error">⚠ Check not performed</span>' })</li>
                </ul>
            </div>
"@
    } elseif ($TeamsResults) {
        # Teams check was attempted but not performed
        $html += @"
            <div class="section" id="teams">
                <h2><span class="section-icon">💬</span>Teams Configuration</h2>
                <p class="status-error">⚠ Check not performed</p>
                <p style="color: #666; font-style: italic;">
                    Reason: Not connected to Microsoft Teams<br>
                    Tip: Connect first with: <code>Connect-MicrosoftTeams</code>
                </p>
            </div>
"@
    }

    # Entra ID Connect Section
    if ($EntraIDResults -and $EntraIDResults.CheckPerformed) {
        $html += @"
            <div class="section" id="entra">
                <h2><span class="section-icon">🔗</span>Entra ID Connect</h2>
                <ul class="info-list">
                    <li><strong>Sync Enabled:</strong> $(if ($EntraIDResults.Status.IsInstalled) { '<span class="status-found">✓ Yes</span>' } else { '<span class="status-missing">✗ No</span>' })</li>
                    <li><strong>Sync Active:</strong> $(if ($EntraIDResults.Status.IsRunning) { '<span class="status-found">✓ Yes</span>' } else { '<span class="status-missing">✗ No</span>' })</li>
"@
        if ($EntraIDResults.Status.LastSyncTime) {
            $html += @"
                    <li><strong>Last Sync:</strong> $($EntraIDResults.Status.LastSyncTime)</li>
"@
        }
        if ($EntraIDResults.Status.SyncErrors.Count -gt 0) {
            $html += @"
                    <li><strong>Errors:</strong> <span class="status-error">$($EntraIDResults.Status.SyncErrors.Count)</span></li>
"@
        }
        $html += @"
                </ul>
            </div>
"@
    }

    # Hybrid Join Section
    if ($IntuneConnResults -and $IntuneConnResults.CheckPerformed) {
        $html += @"
            <div class="section" id="hybrid">
                <h2><span class="section-icon">🔐</span>Hybrid Azure AD Join</h2>
                <ul class="info-list">
                    <li><strong>Check Performed:</strong> <span class="status-found">✓ Yes</span></li>
                    <li><strong>Errors:</strong> $(if ($IntuneConnResults.Status.Errors.Count -eq 0) { '<span class="status-found">0</span>' } else { "<span class='status-error'>$($IntuneConnResults.Status.Errors.Count)</span>" })</li>
                </ul>
            </div>
"@
    }

    # Defender Section
    if ($DefenderResults -and $DefenderResults.CheckPerformed) {
        $html += @"
            <div class="section" id="defender">
                <h2><span class="section-icon">🛡️</span>Microsoft Defender for Endpoint</h2>
                <ul class="info-list">
                    <li><strong>Policies Configured:</strong> $($DefenderResults.Status.ConfiguredPolicies)</li>
                    <li><strong>Compatible Devices:</strong> $($DefenderResults.Status.OnboardedDevices)</li>
                    <li><strong>Onboarding Files:</strong> $($DefenderResults.Status.FilesFound.Count)/4</li>
                    <li><strong>Status:</strong> $(if ($DefenderResults.Status.ConnectorActive) { '<span class="status-found">✓ Active</span>' } else { '<span class="status-missing">✗ Not Configured</span>' })</li>
                </ul>
"@
        
        if ($DefenderResults.Status.FilesFound.Count -gt 0) {
            $html += @"
                <h3>Found Onboarding Files:</h3>
                <ul class="info-list">
"@
            foreach ($file in $DefenderResults.Status.FilesFound) {
                $html += @"
                    <li><span class="status-found">✓</span> $file</li>
"@
            }
            $html += "</ul>"
        }
        
        if ($DefenderResults.Status.FilesMissing.Count -gt 0) {
            $html += @"
                <h3>Missing Onboarding Files:</h3>
                <ul class="info-list">
"@
            foreach ($file in $DefenderResults.Status.FilesMissing) {
                $html += @"
                    <li><span class="status-missing">✗</span> $file</li>
"@
            }
            $html += "</ul>"
        }
        
        $html += "</div>"
    }

    $html += @"
        </div>
        
        <div class="footer">
            Generated by BWS Checking Script | $timestamp
        </div>
    </div>
</body>
</html>
"@

    # Write HTML file
    $html | Out-File -FilePath $reportPath -Encoding UTF8
    
    return $reportPath
}

function Export-PDFReport {
    param(
        [string]$HTMLPath
    )
    
    # Validate input
    if ([string]::IsNullOrWhiteSpace($HTMLPath)) {
        Write-Host "  ✗ Error: HTML path is empty" -ForegroundColor Red
        return $null
    }
    
    if (-not (Test-Path $HTMLPath)) {
        Write-Host "  ✗ Error: HTML file not found: $HTMLPath" -ForegroundColor Red
        return $null
    }
    
    Write-Host "Converting HTML to PDF..." -ForegroundColor Yellow
    
    # Get absolute paths
    $htmlItem = Get-Item $HTMLPath
    $htmlFullPath = $htmlItem.FullName
    $pdfPath = $htmlFullPath -replace '\.html$', '.pdf'
    $conversionSuccess = $false
    
    # Method 1: wkhtmltopdf
    $wkhtmltopdf = Get-Command "wkhtmltopdf" -ErrorAction SilentlyContinue
    if ($wkhtmltopdf) {
        Write-Host "  Using wkhtmltopdf..." -ForegroundColor Gray
        try {
            $process = Start-Process -FilePath $wkhtmltopdf.Source `
                -ArgumentList "--enable-local-file-access", "--no-stop-slow-scripts", "--javascript-delay", "1000", "`"$htmlFullPath`"", "`"$pdfPath`"" `
                -Wait -PassThru -NoNewWindow -ErrorAction Stop
            
            if ($process.ExitCode -eq 0 -and (Test-Path $pdfPath)) {
                $conversionSuccess = $true
                Write-Host "  ✓ PDF created successfully with wkhtmltopdf" -ForegroundColor Green
            }
        } catch {
            Write-Host "  ✗ wkhtmltopdf error: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # Method 2: Chrome/Edge Headless
    if (-not $conversionSuccess) {
        $chromePaths = @(
            "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
            "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
            "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
            "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
        )
        
        $chromePath = $null
        foreach ($path in $chromePaths) {
            if (Test-Path $path) {
                $chromePath = $path
                break
            }
        }
        
        if ($chromePath) {
            Write-Host "  Using Chrome/Edge Headless..." -ForegroundColor Gray
            
            try {
                # Chrome needs file:/// URL
                $htmlFileUrl = "file:///$($htmlFullPath.Replace('\', '/'))"
                
                $chromeArgs = @(
                    "--headless=new"
                    "--disable-gpu"
                    "--no-sandbox"
                    "--print-to-pdf=`"$pdfPath`""
                    "`"$htmlFileUrl`""
                )
                
                $process = Start-Process -FilePath $chromePath -ArgumentList $chromeArgs `
                    -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
                
                # Wait for PDF
                $waitCount = 0
                while (-not (Test-Path $pdfPath) -and $waitCount -lt 10) {
                    Start-Sleep -Milliseconds 500
                    $waitCount++
                }
                
                if (Test-Path $pdfPath) {
                    $conversionSuccess = $true
                    Write-Host "  ✓ PDF created successfully with Chrome/Edge" -ForegroundColor Green
                }
            } catch {
                Write-Host "  ✗ Chrome/Edge error: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
    
    # Method 3: Microsoft Word
    if (-not $conversionSuccess) {
        Write-Host "  Trying Microsoft Word..." -ForegroundColor Gray
        try {
            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            $word.Visible = $false
            $doc = $word.Documents.Open($htmlFullPath)
            $doc.SaveAs([ref]$pdfPath, [ref]17)
            $doc.Close()
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            [System.GC]::Collect()
            
            if (Test-Path $pdfPath) {
                $conversionSuccess = $true
                Write-Host "  ✓ PDF created successfully with Word" -ForegroundColor Green
            }
        } catch {
            Write-Host "  ✗ Word error: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # Failed
    if (-not $conversionSuccess) {
        Write-Host ""
        Write-Host "  ⚠ PDF conversion failed" -ForegroundColor Yellow
        Write-Host "  Install wkhtmltopdf: https://wkhtmltopdf.org/downloads.html" -ForegroundColor Gray
        Write-Host "  Or use: winget install wkhtmltopdf" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  HTML report: $htmlFullPath" -ForegroundColor Cyan
        return $null
    }
    
    return $pdfPath
}


#============================================================================
# GUI Mode
#============================================================================

if ($GUI) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "BWS Checking Tool v$script:Version - GUI"
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
    
    # Customer Name Input
    $labelCustomer = New-Object System.Windows.Forms.Label
    $labelCustomer.Location = New-Object System.Drawing.Point(400, 20)
    $labelCustomer.Size = New-Object System.Drawing.Size(100, 20)
    $labelCustomer.Text = "Kunde (optional):"
    $form.Controls.Add($labelCustomer)
    
    $textCustomer = New-Object System.Windows.Forms.TextBox
    $textCustomer.Location = New-Object System.Drawing.Point(510, 18)
    $textCustomer.Size = New-Object System.Drawing.Size(200, 20)
    $textCustomer.Text = $CustomerName
    $form.Controls.Add($textCustomer)
    
    # Subscription ID Input
    $labelSubID = New-Object System.Windows.Forms.Label
    $labelSubID.Location = New-Object System.Drawing.Point(20, 50)
    $labelSubID.Size = New-Object System.Drawing.Size(150, 20)
    $labelSubID.Text = "Subscription ID (optional):"
    $form.Controls.Add($labelSubID)
    
    $textSubID = New-Object System.Windows.Forms.TextBox
    $textSubID.Location = New-Object System.Drawing.Point(170, 48)
    $textSubID.Size = New-Object System.Drawing.Size(540, 20)
    $textSubID.Text = $SubscriptionId
    $form.Controls.Add($textSubID)
    
    # SharePoint URL Input
    $labelSPUrl = New-Object System.Windows.Forms.Label
    $labelSPUrl.Location = New-Object System.Drawing.Point(20, 78)
    $labelSPUrl.Size = New-Object System.Drawing.Size(150, 20)
    $labelSPUrl.Text = "SharePoint URL (optional):"
    $form.Controls.Add($labelSPUrl)
    
    $textSPUrl = New-Object System.Windows.Forms.TextBox
    $textSPUrl.Location = New-Object System.Drawing.Point(170, 76)
    $textSPUrl.Size = New-Object System.Drawing.Size(540, 20)
    $textSPUrl.Text = $SharePointUrl
    $textSPUrl.PlaceholderText = "https://TENANT-admin.sharepoint.com"
    $form.Controls.Add($textSPUrl)
    
    # GroupBox for Check Selection
    $groupBoxChecks = New-Object System.Windows.Forms.GroupBox
    $groupBoxChecks.Location = New-Object System.Drawing.Point(20, 110)
    $groupBoxChecks.Size = New-Object System.Drawing.Size(300, 250)
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
    
    # Software Packages Check Checkbox
    $chkSoftware = New-Object System.Windows.Forms.CheckBox
    $chkSoftware.Location = New-Object System.Drawing.Point(15, 150)
    $chkSoftware.Size = New-Object System.Drawing.Size(280, 20)
    $chkSoftware.Text = "BWS Software Packages Check"
    $chkSoftware.Checked = $true
    $groupBoxChecks.Controls.Add($chkSoftware)
    
    # SharePoint Configuration Check Checkbox
    $chkSharePoint = New-Object System.Windows.Forms.CheckBox
    $chkSharePoint.Location = New-Object System.Drawing.Point(15, 175)
    $chkSharePoint.Size = New-Object System.Drawing.Size(280, 20)
    $chkSharePoint.Text = "SharePoint Configuration Check"
    $chkSharePoint.Checked = $true
    $groupBoxChecks.Controls.Add($chkSharePoint)
    
    # Teams Configuration Check Checkbox
    $chkTeams = New-Object System.Windows.Forms.CheckBox
    $chkTeams.Location = New-Object System.Drawing.Point(15, 200)
    $chkTeams.Size = New-Object System.Drawing.Size(280, 20)
    $chkTeams.Text = "Teams Configuration Check"
    $chkTeams.Checked = $true
    $groupBoxChecks.Controls.Add($chkTeams)
    
    # Options GroupBox
    $groupBoxOptions = New-Object System.Windows.Forms.GroupBox
    $groupBoxOptions.Location = New-Object System.Drawing.Point(340, 110)
    $groupBoxOptions.Size = New-Object System.Drawing.Size(300, 250)
    $groupBoxOptions.Text = "Options"
    $form.Controls.Add($groupBoxOptions)
    
    # Compact View Checkbox
    $chkCompact = New-Object System.Windows.Forms.CheckBox
    $chkCompact.Location = New-Object System.Drawing.Point(15, 25)
    $chkCompact.Size = New-Object System.Drawing.Size(250, 20)
    $chkCompact.Text = "Compact View"
    $chkCompact.Checked = $true
    $groupBoxOptions.Controls.Add($chkCompact)
    
    # Verbose Checkbox
    $chkShowAll = New-Object System.Windows.Forms.CheckBox
    $chkShowAll.Location = New-Object System.Drawing.Point(15, 50)
    $chkShowAll.Size = New-Object System.Drawing.Size(250, 20)
    $chkShowAll.Text = "Verbose"
    $chkShowAll.Checked = $false
    $groupBoxOptions.Controls.Add($chkShowAll)
    
    # Export Report Checkbox
    $chkExport = New-Object System.Windows.Forms.CheckBox
    $chkExport.Location = New-Object System.Drawing.Point(15, 75)
    $chkExport.Size = New-Object System.Drawing.Size(250, 20)
    $chkExport.Text = "Export Report"
    $chkExport.Checked = $false
    $groupBoxOptions.Controls.Add($chkExport)
    
    # Export Format Label
    $lblExportFormat = New-Object System.Windows.Forms.Label
    $lblExportFormat.Location = New-Object System.Drawing.Point(15, 100)
    $lblExportFormat.Size = New-Object System.Drawing.Size(100, 20)
    $lblExportFormat.Text = "Export Format:"
    $groupBoxOptions.Controls.Add($lblExportFormat)
    
    # HTML Radio Button
    $radioHTML = New-Object System.Windows.Forms.RadioButton
    $radioHTML.Location = New-Object System.Drawing.Point(30, 120)
    $radioHTML.Size = New-Object System.Drawing.Size(70, 20)
    $radioHTML.Text = "HTML"
    $radioHTML.Checked = $true
    $groupBoxOptions.Controls.Add($radioHTML)
    
    # PDF Radio Button
    $radioPDF = New-Object System.Windows.Forms.RadioButton
    $radioPDF.Location = New-Object System.Drawing.Point(110, 120)
    $radioPDF.Size = New-Object System.Drawing.Size(60, 20)
    $radioPDF.Text = "PDF"
    $radioPDF.Checked = $false
    $groupBoxOptions.Controls.Add($radioPDF)
    
    # Both Radio Button
    $radioBoth = New-Object System.Windows.Forms.RadioButton
    $radioBoth.Location = New-Object System.Drawing.Point(180, 120)
    $radioBoth.Size = New-Object System.Drawing.Size(60, 20)
    $radioBoth.Text = "Both"
    $radioBoth.Checked = $false
    $groupBoxOptions.Controls.Add($radioBoth)
    
    # Run Button
    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Location = New-Object System.Drawing.Point(660, 110)
    $btnRun.Size = New-Object System.Drawing.Size(150, 60)
    $btnRun.Text = "Run Check"
    $btnRun.BackColor = [System.Drawing.Color]::LightGreen
    $btnRun.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnRun)
    
    # Clear Button
    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Location = New-Object System.Drawing.Point(660, 230)
    $btnClear.Size = New-Object System.Drawing.Size(150, 30)
    $btnClear.Text = "Clear Output"
    $form.Controls.Add($btnClear)
    
    # Status Label
    $labelStatus = New-Object System.Windows.Forms.Label
    $labelStatus.Location = New-Object System.Drawing.Point(20, 345)
    $labelStatus.Size = New-Object System.Drawing.Size(800, 20)
    $labelStatus.Text = "Ready - Please select checks and click 'Run Check'"
    $labelStatus.ForeColor = [System.Drawing.Color]::Blue
    $form.Controls.Add($labelStatus)
    
    # Progress Bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 370)
    $progressBar.Size = New-Object System.Drawing.Size(950, 20)
    $progressBar.Style = "Continuous"
    $form.Controls.Add($progressBar)
    
    # Output TextBox
    $textOutput = New-Object System.Windows.Forms.TextBox
    $textOutput.Location = New-Object System.Drawing.Point(20, 400)
    $textOutput.Size = New-Object System.Drawing.Size(950, 290)
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
        $customerName = $textCustomer.Text
        $subId = $textSubID.Text
        $SharePointUrl = $textSPUrl.Text
        $runAzure = $chkAzure.Checked
        $runIntune = $chkIntune.Checked
        $runEntraID = $chkEntraID.Checked
        $runIntuneConn = $chkIntuneConn.Checked
        $runDefender = $chkDefender.Checked
        $runSoftware = $chkSoftware.Checked
        $runSharePoint = $chkSharePoint.Checked
        $runTeams = $chkTeams.Checked
        $compact = $chkCompact.Checked
        $showAll = $chkShowAll.Checked
        $export = $chkExport.Checked
        
        # Determine export format
        $exportFormat = "HTML"
        if ($radioPDF.Checked) { $exportFormat = "PDF" }
        if ($radioBoth.Checked) { $exportFormat = "Both" }
        
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
            $softwareResults = $null
            $sharePointResults = $null
            
            $totalChecks = ($runAzure -as [int]) + ($runIntune -as [int]) + ($runEntraID -as [int]) + ($runIntuneConn -as [int]) + ($runDefender -as [int]) + ($runSoftware -as [int]) + ($runSharePoint -as [int]) + ($runTeams -as [int])
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
            
            # Run BWS Software Packages Check
            if ($runSoftware) {
                $labelStatus.Text = "Running BWS Software Packages Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $softwareResults = Test-BWSSoftwarePackages -CompactView $compact
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run SharePoint Configuration Check
            if ($runSharePoint) {
                $labelStatus.Text = "Running SharePoint Configuration Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $sharePointResults = Test-SharePointConfiguration -CompactView $compact -SharePointUrl $SharePointUrl
                $currentCheck++
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
            }
            
            # Run Teams Configuration Check
            if ($runTeams) {
                $labelStatus.Text = "Running Teams Configuration Check..."
                $progressBar.Value = 10 + ($currentCheck * $progressIncrement)
                $form.Refresh()
                
                $teamsResults = Test-TeamsConfiguration -CompactView $compact
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
            
            if ($runSoftware -and $softwareResults -and $softwareResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  BWS Software Packages:" -ForegroundColor White
                Write-Host "    Total:        $($softwareResults.Status.Total)" -ForegroundColor White
                Write-Host "    Found:        $($softwareResults.Status.Found.Count)" -ForegroundColor $(if ($softwareResults.Status.Found.Count -eq $softwareResults.Status.Total) { "Green" } else { "Yellow" })
                Write-Host "    Missing:      $($softwareResults.Status.Missing.Count)" -ForegroundColor $(if ($softwareResults.Status.Missing.Count -eq 0) { "Green" } else { "Red" })
            }
            
            if ($runSharePoint -and $sharePointResults -and $sharePointResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  SharePoint Configuration:" -ForegroundColor White
                Write-Host "    SP Ext. Sharing:   " -NoNewline -ForegroundColor White
                Write-Host $(if ($sharePointResults.Status.Settings.SharePointExternalSharing -eq "Anyone") { "Anyone (✓)" } else { "$($sharePointResults.Status.Settings.SharePointExternalSharing) (✗)" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.SharePointExternalSharing -eq "Anyone") { "Green" } else { "Yellow" })
                Write-Host "    OD Ext. Sharing:   " -NoNewline -ForegroundColor White
                Write-Host $(if ($sharePointResults.Status.Settings.OneDriveExternalSharing -eq "Disabled") { "Only Organization (✓)" } else { "$($sharePointResults.Status.Settings.OneDriveExternalSharing) (✗)" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.OneDriveExternalSharing -eq "Disabled") { "Green" } else { "Yellow" })
                Write-Host "    Site Creation:     " -NoNewline -ForegroundColor White  
                Write-Host $(if ($sharePointResults.Status.Settings.SiteCreation -eq "Disabled") { "Disabled (✓)" } else { "$($sharePointResults.Status.Settings.SiteCreation) (✗)" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.SiteCreation -eq "Disabled") { "Green" } else { "Yellow" })
                Write-Host "    Legacy Auth Block: " -NoNewline -ForegroundColor White
                Write-Host $(if ($sharePointResults.Status.Settings.LegacyAuthBlocked -eq $true) { "Yes (✓)" } else { "No (✗)" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.LegacyAuthBlocked) { "Green" } else { "Yellow" })
            }
            
            if ($runTeams -and $teamsResults -and $teamsResults.CheckPerformed) {
                Write-Host ""
                Write-Host "  Teams Configuration:" -ForegroundColor White
                Write-Host "    Meetings w/ unmanaged MS: " -NoNewline -ForegroundColor White
                Write-Host $(if ($teamsResults.Status.Settings.ExternalAccessEnabled -eq $false) { "Disabled (✓)" } else { "Enabled (✗)" }) -ForegroundColor $(if ($teamsResults.Status.Settings.ExternalAccessEnabled -eq $false) { "Green" } else { "Yellow" })
                
                $allStorageDisabled = ($teamsResults.Status.Settings.CloudStorageCitrix -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageDropbox -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageBox -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageGoogleDrive -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageEgnyte -eq "Disabled")
                Write-Host "    Cloud Storage:     " -NoNewline -ForegroundColor White
                if ($allStorageDisabled) {
                    Write-Host "All Disabled (✓)" -ForegroundColor Green
                } else {
                    $enabledList = @()
                    if ($teamsResults.Status.Settings.CloudStorageCitrix -eq "Enabled") { $enabledList += "Citrix" }
                    if ($teamsResults.Status.Settings.CloudStorageDropbox -eq "Enabled") { $enabledList += "Dropbox" }
                    if ($teamsResults.Status.Settings.CloudStorageBox -eq "Enabled") { $enabledList += "Box" }
                    if ($teamsResults.Status.Settings.CloudStorageGoogleDrive -eq "Enabled") { $enabledList += "Google Drive" }
                    if ($teamsResults.Status.Settings.CloudStorageEgnyte -eq "Enabled") { $enabledList += "Egnyte" }
                    Write-Host "Enabled: $($enabledList -join ', ') (✗)" -ForegroundColor Yellow
                }
                
                Write-Host "    Anonymous Join:    " -NoNewline -ForegroundColor White
                Write-Host $(if ($teamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Disabled") { "Disabled (✓)" } else { "Enabled (✗)" }) -ForegroundColor $(if ($teamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Disabled") { "Green" } else { "Yellow" })
                
                Write-Host "    Who Can Present:   " -NoNewline -ForegroundColor White
                Write-Host $(if ($teamsResults.Status.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") { "Everyone (✓)" } else { "$($teamsResults.Status.Settings.DefaultPresenterRole) (✗)" }) -ForegroundColor $(if ($teamsResults.Status.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") { "Green" } else { "Yellow" })
            }
            
            Write-Host "======================================================" -ForegroundColor Cyan
            
            if ($compact) {
                Write-Host ""
                Write-Host "Note: Compact View enabled" -ForegroundColor Gray
            }
            
            # Export report if requested
            if ($export) {
                Write-Host ""
                
                $currentContext = Get-AzContext
                $subName = if ($currentContext) { $currentContext.Subscription.Name } else { "Unknown" }
                
                $overallStatus = ($azureResults.Missing.Count -eq 0 -and $azureResults.Errors.Count -eq 0) -and 
                                 (-not $intuneResults -or ($intuneResults.Missing.Count -eq 0 -and $intuneResults.Errors.Count -eq 0)) -and
                                 (-not $entraIDResults -or ($entraIDResults.Status.IsRunning)) -and
                                 (-not $intuneConnResults -or ($intuneConnResults.Status.Errors.Count -eq 0)) -and
                                 (-not $defenderResults -or ($defenderResults.Status.ConnectorActive -and $defenderResults.Status.FilesMissing.Count -eq 0)) -and
                                 (-not $softwareResults -or ($softwareResults.Status.Missing.Count -eq 0)) -and
                                 (-not $sharePointResults -or ($sharePointResults.Status.Compliant)) -and
                                 (-not $teamsResults -or ($teamsResults.Status.Compliant))
                
                # Generate HTML report
                if ($exportFormat -eq "HTML" -or $exportFormat -eq "Both") {
                    Write-Host "Generating HTML Report..." -ForegroundColor Yellow
                    $htmlPath = Export-HTMLReport -BCID $bcid -CustomerName $customerName -SubscriptionName $subName `
                        -AzureResults $azureResults -IntuneResults $intuneResults `
                        -EntraIDResults $entraIDResults -IntuneConnResults $intuneConnResults `
                        -DefenderResults $defenderResults -SoftwareResults $softwareResults `
                        -SharePointResults $sharePointResults -TeamsResults $teamsResults -OverallStatus $overallStatus
                    
                    Write-Host "HTML Report exported to: $htmlPath" -ForegroundColor Green
                }
                
                # Generate PDF report
                if ($exportFormat -eq "PDF" -or $exportFormat -eq "Both") {
                    if (-not $htmlPath) {
                        # Need HTML first for PDF conversion
                        $htmlPath = Export-HTMLReport -BCID $bcid -CustomerName $customerName -SubscriptionName $subName `
                            -AzureResults $azureResults -IntuneResults $intuneResults `
                            -EntraIDResults $entraIDResults -IntuneConnResults $intuneConnResults `
                            -DefenderResults $defenderResults -SoftwareResults $softwareResults `
                            -SharePointResults $sharePointResults -TeamsResults $teamsResults -OverallStatus $overallStatus
                    }
                    
                    $pdfPath = Export-PDFReport -HTMLPath $htmlPath
                    if ($pdfPath) {
                        Write-Host "PDF Report exported to: $pdfPath" -ForegroundColor Green
                    }
                    
                    # Clean up temp HTML if only PDF was requested
                    if ($exportFormat -eq "PDF" -and $htmlPath -and (Test-Path $htmlPath)) {
                        Remove-Item $htmlPath -Force -ErrorAction SilentlyContinue
                    }
                }
                
                Write-Host ""
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

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  BWS-Checking-Script v$script:Version" -ForegroundColor Cyan
Write-Host "  Command Line Mode" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

# Set CompactView as default if not explicitly overridden
if (-not $PSBoundParameters.ContainsKey('CompactView')) {
    $CompactView = $true
}

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

# Run BWS Software Packages Check
$softwareResults = $null
if (-not $SkipSoftware) {
    $softwareResults = Test-BWSSoftwarePackages -CompactView $CompactView
}

# Run SharePoint Configuration Check
$sharePointResults = $null
if (-not $SkipSharePoint) {
    $sharePointResults = Test-SharePointConfiguration -CompactView $CompactView -SharePointUrl $SharePointUrl
}

# Run Teams Configuration Check
$teamsResults = $null
if (-not $SkipTeams) {
    $teamsResults = Test-TeamsConfiguration -CompactView $CompactView
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

if ($softwareResults -and $softwareResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  BWS Software Packages:" -ForegroundColor White
    Write-Host "    Total:        $($softwareResults.Status.Total)" -ForegroundColor White
    Write-Host "    Found:        $($softwareResults.Status.Found.Count)" -ForegroundColor $(if ($softwareResults.Status.Found.Count -eq $softwareResults.Status.Total) { "Green" } else { "Yellow" })
    Write-Host "    Missing:      $($softwareResults.Status.Missing.Count)" -ForegroundColor $(if ($softwareResults.Status.Missing.Count -eq 0) { "Green" } else { "Red" })
}

if ($sharePointResults -and $sharePointResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  SharePoint Configuration:" -ForegroundColor White
    Write-Host "    SP Ext. Sharing:   " -NoNewline -ForegroundColor White
    Write-Host $(if ($sharePointResults.Status.Settings.SharePointExternalSharing -eq "Anyone") { "Anyone (✓)" } else { "$($sharePointResults.Status.Settings.SharePointExternalSharing) (✗)" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.SharePointExternalSharing -eq "Anyone") { "Green" } else { "Yellow" })
    Write-Host "    OD Ext. Sharing:   " -NoNewline -ForegroundColor White
    Write-Host $(if ($sharePointResults.Status.Settings.OneDriveExternalSharing -eq "Disabled") { "Only Organization (✓)" } else { "$($sharePointResults.Status.Settings.OneDriveExternalSharing) (✗)" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.OneDriveExternalSharing -eq "Disabled") { "Green" } else { "Yellow" })
    Write-Host "    Site Creation:     " -NoNewline -ForegroundColor White  
    Write-Host $(if ($sharePointResults.Status.Settings.SiteCreation -eq "Disabled") { "Disabled (✓)" } else { "$($sharePointResults.Status.Settings.SiteCreation) (✗)" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.SiteCreation -eq "Disabled") { "Green" } else { "Yellow" })
    Write-Host "    Legacy Auth Block: " -NoNewline -ForegroundColor White
    Write-Host $(if ($sharePointResults.Status.Settings.LegacyAuthBlocked -eq $true) { "Yes (✓)" } else { "No (✗)" }) -ForegroundColor $(if ($sharePointResults.Status.Settings.LegacyAuthBlocked) { "Green" } else { "Yellow" })
}

if ($teamsResults -and $teamsResults.CheckPerformed) {
    Write-Host ""
    Write-Host "  Teams Configuration:" -ForegroundColor White
    Write-Host "    Meetings w/ unmanaged MS: " -NoNewline -ForegroundColor White
    Write-Host $(if ($teamsResults.Status.Settings.ExternalAccessEnabled -eq $false) { "Disabled (✓)" } else { "Enabled (✗)" }) -ForegroundColor $(if ($teamsResults.Status.Settings.ExternalAccessEnabled -eq $false) { "Green" } else { "Yellow" })
    
    $allStorageDisabled = ($teamsResults.Status.Settings.CloudStorageCitrix -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageDropbox -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageBox -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageGoogleDrive -eq "Disabled") -and ($teamsResults.Status.Settings.CloudStorageEgnyte -eq "Disabled")
    Write-Host "    Cloud Storage:     " -NoNewline -ForegroundColor White
    if ($allStorageDisabled) {
        Write-Host "All Disabled (✓)" -ForegroundColor Green
    } else {
        $enabledList = @()
        if ($teamsResults.Status.Settings.CloudStorageCitrix -eq "Enabled") { $enabledList += "Citrix" }
        if ($teamsResults.Status.Settings.CloudStorageDropbox -eq "Enabled") { $enabledList += "Dropbox" }
        if ($teamsResults.Status.Settings.CloudStorageBox -eq "Enabled") { $enabledList += "Box" }
        if ($teamsResults.Status.Settings.CloudStorageGoogleDrive -eq "Enabled") { $enabledList += "Google Drive" }
        if ($teamsResults.Status.Settings.CloudStorageEgnyte -eq "Enabled") { $enabledList += "Egnyte" }
        Write-Host "Enabled: $($enabledList -join ', ') (✗)" -ForegroundColor Yellow
    }
    
    Write-Host "    Anonymous Join:    " -NoNewline -ForegroundColor White
    Write-Host $(if ($teamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Disabled") { "Disabled (✓)" } else { "Enabled (✗)" }) -ForegroundColor $(if ($teamsResults.Status.Settings.AnonymousUsersCanJoin -eq "Disabled") { "Green" } else { "Yellow" })
    
    Write-Host "    Who Can Present:   " -NoNewline -ForegroundColor White
    Write-Host $(if ($teamsResults.Status.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") { "Everyone (✓)" } else { "$($teamsResults.Status.Settings.DefaultPresenterRole) (✗)" }) -ForegroundColor $(if ($teamsResults.Status.Settings.DefaultPresenterRole -eq "EveryoneUserOverride") { "Green" } else { "Yellow" })
}

Write-Host ""
$overallStatus = ($azureResults.Missing.Count -eq 0 -and $azureResults.Errors.Count -eq 0) -and 
                 (-not $intuneResults -or ($intuneResults.Missing.Count -eq 0 -and $intuneResults.Errors.Count -eq 0)) -and
                 (-not $entraIDResults -or ($entraIDResults.Status.IsRunning)) -and
                 (-not $intuneConnResults -or ($intuneConnResults.Status.Errors.Count -eq 0)) -and
                 (-not $defenderResults -or ($defenderResults.Status.ConnectorActive -and $defenderResults.Status.FilesMissing.Count -eq 0)) -and
                 (-not $softwareResults -or ($softwareResults.Status.Missing.Count -eq 0)) -and
                 (-not $sharePointResults -or ($sharePointResults.Status.Compliant)) -and
                 (-not $teamsResults -or ($teamsResults.Status.Compliant))

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
    
    # Generate HTML report
    if ($ExportFormat -eq "HTML" -or $ExportFormat -eq "Both") {
        Write-Host "Generating HTML Report..." -ForegroundColor Yellow
        
        $htmlPath = Export-HTMLReport -BCID $BCID -CustomerName $CustomerName -SubscriptionName $currentContext.Subscription.Name `
            -AzureResults $azureResults -IntuneResults $intuneResults `
            -EntraIDResults $entraIDResults -IntuneConnResults $intuneConnResults `
            -DefenderResults $defenderResults -SoftwareResults $softwareResults `
            -SharePointResults $sharePointResults -TeamsResults $teamsResults -OverallStatus $overallStatus
        
        Write-Host "HTML Report exported to: $htmlPath" -ForegroundColor Green
    }
    
    # Generate PDF report
    if ($ExportFormat -eq "PDF" -or $ExportFormat -eq "Both") {
        if (-not $htmlPath) {
            # Need HTML first for PDF conversion
            $htmlPath = Export-HTMLReport -BCID $BCID -CustomerName $CustomerName -SubscriptionName $currentContext.Subscription.Name `
                -AzureResults $azureResults -IntuneResults $intuneResults `
                -EntraIDResults $entraIDResults -IntuneConnResults $intuneConnResults `
                -DefenderResults $defenderResults -SoftwareResults $softwareResults `
                -SharePointResults $sharePointResults -TeamsResults $teamsResults -OverallStatus $overallStatus
        }
        
        $pdfPath = Export-PDFReport -HTMLPath $htmlPath
        if ($pdfPath) {
            Write-Host "PDF Report exported to: $pdfPath" -ForegroundColor Green
        }
        
        # Clean up temp HTML if only PDF was requested
        if ($ExportFormat -eq "PDF" -and $htmlPath -and (Test-Path $htmlPath)) {
            Remove-Item $htmlPath -Force -ErrorAction SilentlyContinue
        }
    }
    
    Write-Host "HTML Report exported to: $reportPath" -ForegroundColor Green
    Write-Host ""
}