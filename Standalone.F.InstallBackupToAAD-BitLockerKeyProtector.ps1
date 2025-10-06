$AppName = "BackupToAAD-BitLockerKeyProtector"
$RegPath = "HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$AppName"

# Create BWS registry key
If ( $Null -eq ( Get-Item -Path "HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$($AppName)" -ErrorAction SilentlyContinue ) )
{
    New-Item -Path $RegPath -Name $AppName -Force
}
# Get the BitLocker volume
$BLV = Get-BitLockerVolume -MountPoint "C:"
# Find the recovery password key protector
$KeyProtector = $BLV.KeyProtector | Where-Object { $_.KeyProtectorType -eq "RecoveryPassword" }
# Backup the key protector to AAD 
  BackupToAAD-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $KeyProtector.KeyProtectorId

### Detection script ###
### Look for Bitlocker Recovery Key Backup events of Systemdrive
$BLSysVolume = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
$BLRecoveryProtector = $BLSysVolume.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } -ErrorAction Stop
$BLprotectorguid = $BLRecoveryProtector.KeyProtectorId


### obtain backup event for Systemdrive
$BLBackupEvent = Get-WinEvent -ProviderName Microsoft-Windows-BitLocker-API -FilterXPath "*[System[(EventID=845)] and EventData[Data[@Name='ProtectorGUID'] and (Data='$BLprotectorguid')]]" -MaxEvents 1 -ErrorAction Stop

# Check for returned values of events
if ($null -ne $BLBackupEvent) 
{
# If the command succeeds and event present, set the status to Success
   Set-ItemProperty -Path $RegPath -Name Status -Value "Success" -Force
}
else 
{
# If the command fails and no event, set the status to Failed
Set-ItemProperty -Path $RegPath -Name Status -Value "Failed" -Force}