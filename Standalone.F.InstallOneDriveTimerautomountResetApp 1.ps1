#Region script parameters
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true)]    
    [string]$AppVersion
)

#Region auto trim string parameters
($MyInvocation.MyCommand.Parameters).Keys | ForEach-Object {
    $val = (Get-Variable -Name $_ -ErrorAction SilentlyContinue).Value
    if( $val.length -gt 0 -and $val.GetType().Name -ieq 'String' )
    {
        (Get-Variable -Name $_ -ErrorAction SilentlyContinue).Value = $val.Trim()
    }
}
#EndRegion
#EndRegion script parameters

#Region Constants for the script
$TemplateScope = 'Frontend'
$TemplateType = 'Standalone'
$TemplateVersion = [Version]'2.0.0.0'
#EndRegion Constants for the script

#Region OperatingSystem detection
If ( $Null -eq $env:APPDATA )
{
	# Script is not running on Windows
    $BasePath = '/opt/swisscom'
    $CredentialFile = "$($BasePath)/Credentials/root.json"
} Elseif ( $Null -ne $env:OperationMode -and $env:OperationMode -ieq 'Development' )
{
    $BasePath = "C:\Users\$($env:USERNAME)\Documents\SmartICTDev"
    $CredentialFile = "D:\Swisscom\Scripts\Credentials\$env:USERNAME.json"
} Else
{
    $BasePath = 'C:\Swisscom'
    $CredentialFile = "D:\Swisscom\Scripts\Credentials\$env:USERNAME.json"
}
#EndRegion OperatingSystem detection

#Region Loading modules
if(!($env:PSModulePath -like "*$BasePath*"))
{
    $env:PSModulePath = "$BasePath$([System.IO.Path]::DirectorySeparatorChar)Modules$([System.IO.Path]::PathSeparator)" + $env:PSModulePath
}
If ( $Null -ne $env:DevEnvironment )
{
    $ModulePath = $env:DevEnvironment + '$([System.IO.Path]::DirectorySeparatorChar)Modules'
    If (  Test-Path $ModulePath -ErrorAction SilentlyContinue )
    {
        Write-Host "Development mode active. Loading modules with priority to path $ModulePath"
        $env:PSModulePath = "${ModulePath}$([System.IO.Path]::PathSeparator)${env:PSModulePath}"
    }
}
#EndRegion Loading modules

#Region Variables that need to be adapted for every script
$ProcedureAction = "InstallOneDriveTimerautomountResetApp"
#endRegion Variables that need to be adapted for every script

#Region variables auto-initialized
$ScriptName = "$TemplateType.$($TemplateScope[0]).$ProcedureAction"
$ProcedurePath = "$TemplateType$([System.IO.Path]::DirectorySeparatorChar)$TemplateScope"
$LogPath = "$BasePath$([System.IO.Path]::DirectorySeparatorChar)Logs$([System.IO.Path]::DirectorySeparatorChar)$ProcedurePath$([System.IO.Path]::DirectorySeparatorChar)"
$OutPath = "$BasePath$([System.IO.Path]::DirectorySeparatorChar)Output$([System.IO.Path]::DirectorySeparatorChar)$ProcedurePath$([System.IO.Path]::DirectorySeparatorChar)"
# Path to config file
$ConfigFilesPath = "$BasePath$([System.IO.Path]::DirectorySeparatorChar)$ProcedurePath$([System.IO.Path]::DirectorySeparatorChar)ConfigFiles$([System.IO.Path]::DirectorySeparatorChar)"
# Default Failure message
$FailureMessage = "$ScriptName has failed"
#EndRegion variables auto-initialized

#Region Initialization
Try
{    
    #Region Debug and Verbose management
    If ($PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent)
    {
        Write-Debug "Debug mode: on"
        Set-Variable -Name 'DebugPreference' -Value 'Continue' -Scope Global -Force -Confirm:$false -WhatIf:$false
    }else
    {
        Set-Variable -Name 'DebugPreference' -Value 'SilentlyContinue' -Scope Global -Force -Confirm:$false -WhatIf:$false
    }

    If ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)
    {
        Write-Verbose "Verbose mode: on"
        Set-Variable -Name 'VerbosePreference' -Value 'Continue' -Scope Global -Force -Confirm:$false -WhatIf:$false
    } Else
    {
        Set-Variable -Name 'VerbosePreference' -Value 'SilentlyContinue' -Scope Global -Force -Confirm:$false -WhatIf:$false
    }
    #EndRegion Debug and Verbose management 

    #Region Logging
    If ( -not ( Test-Path $ConfigFilesPath ) )
    {
        New-Item $ConfigFilesPath -ItemType Directory | Out-Null
    }
    If ( -not ( Test-Path $LogPath ) )
    {
        New-Item $LogPath -ItemType Directory | Out-Null
    }
    If ( -not ( Test-Path $OutPath ) )
    {
        New-Item $OutPath -ItemType Directory | Out-Null
    }

	$LogFileFullPath = ( $LogPath + $( $ScriptName ) + ".log" )

    Write-Verbose "Starting transcript to file $LogFileFullPath"
    Start-Transcript -path $LogFileFullPath -append | out-null
    Write-Host "$(Get-Date -format g) : Start of script $ScriptName ..."
    Write-Host "$(Get-Date -format g) : Base template $TemplateType.$($TemplateScope[0]) version $TemplateVersion"
    #EndRegion Logging

    #Region Custom objects

    #EndRegion Custom objects

    #Region Personalized initialization

        $AppName = "OneDriveTimerautomountReset"
        $AppPath = ""
        $RegPath = "HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$AppName"

        #Region Create/Update SmartICT registry keys and values
        Write-Host "$(Get-Date -Format g) : Create 'HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$AppName' key"
        If ( $Null -eq ( Get-Item -Path "HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$($AppName)" -ErrorAction SilentlyContinue ) )
        {
            New-Item -Path HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps -Name $AppName -Force
        }
        If ( $Null -eq ( Get-Item -Path HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$AppName -ErrorAction SilentlyContinue ) )
        {
            Throw "$(Get-Date -Format g) : Unable to create 'HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$AppName' key"
        }

        Write-Host "$(Get-Date -Format g) : Add package Status to $AppName key"
        If ( $Null -eq ( Get-ItemProperty -Path HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$AppName -Name Status -ErrorAction SilentlyContinue ) )
        {
            Set-ItemProperty -Path HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$AppName -Name Status -Value NotInstalled -Force
        }
        If ( $Null -eq ( Get-ItemProperty -Path HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$AppName -Name Status -ErrorAction SilentlyContinue ) )
        {
            Throw "Unable to set package status value to $AppName key"
        }
    
        Write-Host "$(Get-Date -Format g) : Add package Version to $AppName key"
        If ( $Null -eq ( Get-ItemProperty -Path HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$AppName -Name Version -ErrorAction SilentlyContinue ) )
        {
            Set-ItemProperty -Path HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$AppName -Name Version -Value '' -Force
        }
        If ( $Null -eq ( Get-ItemProperty -Path HKLM:\SOFTWARE\Swisscom\BWS\Intune\Apps\$AppName -Name Version -ErrorAction SilentlyContinue ) )
        {
            Throw "Unable to set package version value to $AppName key"
        }
        #EndRegion Create/Update SmartICT registry keys
    #EndRegion Personalized initialization
}
catch
{
    write-host "$(Get-Date -format g) : $FailureMessage" -ForegroundColor Magenta
    write-host "ErrorType: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Error Message: $($_.Exception.Message)" -ForegroundColor Red
    Try { Stop-Transcript | Out-Null } Catch [System.InvalidOperationException]{ Write-Host "Session is not transcripting" -ForegroundColor Yellow }

    Exit 1
}
#EndRegion Initialization

#Region WorkingSet
Try
{
    #Region Working part
    Write-Host "$(Get-Date -format g) : Working part starting"
    
    #Region Install App
    Write-Host "$(Get-Date -format g) : Create scheduled task for Chocolatey packages update"
    $schtaskName = 'OneDriveTimerautomountReset'
	$schtaskPath = '\Swisscom\BWS\'
    $schtaskDescription = 'Reset HKCU\Software\Microsoft\OneDrive\Accounts\Business1\Timerautomount registry key at each login.'
    If ( $Null -ine (Get-ScheduledTask -TaskName $schtaskName -ErrorAction SilentlyContinue).State )
    {
        Write-Host "$(Get-Date -format g) : Deleting existing scheduled task for Chocolatey packages update"
        Unregister-ScheduledTask -TaskName $schtaskName -Confirm:$false
    }
    Write-Host "$(Get-Date -format g) : Generating scheduled task for OneDriveTimerautomountReset"
	$action = New-ScheduledTaskAction -Execute "reg.exe" -Argument 'add HKCU\Software\Microsoft\OneDrive\Accounts\Business1 /v Timerautomount /t REG_QWORD /d 1 /f'
	$trigger = New-ScheduledTaskTrigger -AtLogOn
	$principal= New-ScheduledTaskPrincipal -UserId 'SYSTEM'
	$settings= New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -minutes 5)
    Write-Host "$(Get-Date -format g) : Registering scheduled task for OneDriveTimerautomountReset"
    Register-ScheduledTask -TaskName $schtaskName -Trigger $trigger -Action $action -Settings $settings -Description $schtaskDescription -Force

    Write-Host "$(Get-Date -format g) : Working part done"
    #EndRegion Working part
    
    #Region Validation
    Write-Host "$(Get-Date -format g) : Validation starting"
    If ( $Null -ine (Get-ScheduledTask -TaskName $schtaskName -ErrorAction SilentlyContinue).State )
    {
        If ( (Get-ScheduledTask -TaskName $schtaskName -ErrorAction SilentlyContinue).State -ieq "Ready")
        {
            Set-ItemProperty -Path $RegPath -Name Status -Value "Installed" -Force | Out-Null
            Set-ItemProperty -Path $RegPath -Name Version -Value $AppVersion -Force | Out-Null
            Write-Host "$(Get-Date -format g) : $($AppName) has been installed."
        }
        Else {
            Write-Host "$(Get-Date -format g) : $($AppName) is in incorrect state."
            Set-ItemProperty -Path $RegPath -Name Status -Value "NotInstalled" -Force | Out-Null
            Set-ItemProperty -Path $RegPath -Name Version -Value $AppVersion -Force | Out-Null
        }
    }
    Else {
        Write-Host "$(Get-Date -format g) : $($AppName) has not been installed."
        Set-ItemProperty -Path $RegPath -Name Status -Value "NotInstalled" -Force | Out-Null
        Set-ItemProperty -Path $RegPath -Name Version -Value $AppVersion -Force | Out-Null
    }
    Write-Host "$(Get-Date -format g) : Validation done"
    #EndRegion Validation
} Catch
{
    write-host "$(Get-Date -format g) : $FailureMessage" -ForegroundColor Magenta
    write-host "ErrorType: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Error Message: $($_.Exception.Message)" -ForegroundColor Red
    Try { Stop-Transcript | Out-Null } Catch [System.InvalidOperationException]{ Write-Host "Session is not transcripting" -ForegroundColor Yellow }

    Exit 1
}
#EndRegion WorkingSet

#Region Cleanup and EndOf script
Write-Host "$(Get-Date -format g) : End of run"
Try { Stop-Transcript | Out-Null } Catch [System.InvalidOperationException]{ Write-Host "Session is not transcripting" -ForegroundColor Yellow }
Remove-Item -Path $LogFileFullPath -Force -Confirm:$False -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

Exit 0
#EndRegion Cleanup and EndOf script