#Parameters

#BCID - Business Continuity ID
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$BCID = "0000"
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





#BWS Base Check



## Azure
### Storage Accounts
### VM
### Azure Vaults
### vNets
### Azure Gateways
### vNics
### NSGs
### Public IPs
### BNS/EC Connection Points



## M365
### Intune

### 

#BWS TA Check

#BWS Report
