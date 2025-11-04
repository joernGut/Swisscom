Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# Globale Variable für Berechtigungen
$global:PermissionsExport = @()

function Get-NTFSPermissions {
    param ($Path)

    $acl = Get-Acl -Path $Path -ErrorAction SilentlyContinue
    $acl.Access | ForEach-Object {
        if ($_.IdentityReference -like "*\*") {
            [PSCustomObject]@{
                Path        = $Path
                Identity    = $_.IdentityReference
                Rights      = $_.FileSystemRights
                AccessType  = $_.AccessControlType
                Inherited   = $_.IsInherited
            }
        }
    }
}

function Export-Permissions {
    param ($RootPath)

    $global:PermissionsExport = @()

    Get-ChildItem -Path $RootPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
        $perms = Get-NTFSPermissions -Path $_.FullName
        if ($perms) {
            $global:PermissionsExport += $perms
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Berechtigungen erfolgreich erfasst.","Info")
}

function Migrate-GroupsToNewDomain {
    param (
        [string]$NewDomain,
        [string]$OUPath
    )

    $groups = $global:PermissionsExport.Identity | Select-Object -Unique

    foreach ($group in $groups) {
        $oldGroupName = ($group -split '\\')[1]
        $newGroupName = "GRP_$oldGroupName"

        if (-not (Get-ADGroup -Filter { Name -eq $newGroupName } -Server $NewDomain -ErrorAction SilentlyContinue)) {
            try {
                New-ADGroup -Name $newGroupName -SamAccountName $newGroupName -GroupScope Global -Path $OUPath -Server $NewDomain
            } catch {
                Write-Warning "Fehler beim Erstellen der Gruppe $newGroupName"
                continue
            }
        }

        try {
            $members = Get-ADGroupMember -Identity $oldGroupName -ErrorAction Stop
        } catch {
            continue
        }

        foreach ($user in $members) {
            $userFound = Get-ADUser -Filter { SamAccountName -eq $user.SamAccountName } -Server $NewDomain -ErrorAction SilentlyContinue
            if ($userFound) {
                try {
                    Add-ADGroupMember -Identity $newGroupName -Members $userFound -Server $NewDomain
                } catch {
                    Write-Warning "Fehler beim Hinzufügen von $($userFound.SamAccountName)"
                }
            }
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Gruppen in neuer Domäne erstellt und befüllt.","Fertig")
}

function Migrate-DataAndApplyPermissions {
    param (
        [string]$SourcePath,
        [string]$TargetPath
    )

    # Kopieren der Daten
    robocopy $SourcePath $TargetPath /E /COPYALL /R:1 /W:1 | Out-Null

    foreach ($entry in $global:PermissionsExport) {
        $relative = $entry.Path.Replace($SourcePath, "").TrimStart('\')
        $targetFile = Join-Path $TargetPath $relative
        if (-not (Test-Path $targetFile)) { continue }

        $acl = Get-Acl -Path $targetFile
        $newGroup = "GRP_" + ($entry.Identity -split '\\')[1]
        $id = "$env:USERDOMAIN\$newGroup"

        try {
            $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($id, $entry.Rights, "ContainerInherit,ObjectInherit", "None", $entry.AccessType)
            $acl.SetAccessRule($accessRule)
            Set-Acl -Path $targetFile -AclObject $acl
        } catch {
            Write-Warning "Berechtigung konnte nicht gesetzt werden: $targetFile"
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Migration abgeschlossen.","Fertig")
}

# GUI-Start
function Start-GUI {
    $form = New-Object Windows.Forms.Form
    $form.Text = "Fileserver Migrations-Tool"
    $form.Size = '700,400'
    $form.StartPosition = 'CenterScreen'

    # Pfade
    $lblSource = New-Object Windows.Forms.Label
    $lblSource.Text = "Quellpfad:"
    $lblSource.Location = "10,20"
    $form.Controls.Add($lblSource)

    $txtSource = New-Object Windows.Forms.TextBox
    $txtSource.Size = '500,20'
    $txtSource.Location = "100,20"
    $form.Controls.Add($txtSource)

    $btnBrowseSource = New-Object Windows.Forms.Button
    $btnBrowseSource.Text = "..."
    $btnBrowseSource.Location = "610,18"
    $btnBrowseSource.Width = 30
    $btnBrowseSource.Add_Click({
        $dlg = New-Object Windows.Forms.FolderBrowserDialog
        if ($dlg.ShowDialog() -eq "OK") {
            $txtSource.Text = $dlg.SelectedPath
        }
    })
    $form.Controls.Add($btnBrowseSource)

    # Zielpfad
    $lblTarget = New-Object Windows.Forms.Label
    $lblTarget.Text = "Zielpfad:"
    $lblTarget.Location = "10,60"
    $form.Controls.Add($lblTarget)

    $txtTarget = New-Object Windows.Forms.TextBox
    $txtTarget.Size = '500,20'
    $txtTarget.Location = "100,60"
    $form.Controls.Add($txtTarget)

    $btnBrowseTarget = New-Object Windows.Forms.Button
    $btnBrowseTarget.Text = "..."
    $btnBrowseTarget.Location = "610,58"
    $btnBrowseTarget.Width = 30
    $btnBrowseTarget.Add_Click({
        $dlg = New-Object Windows.Forms.FolderBrowserDialog
        if ($dlg.ShowDialog() -eq "OK") {
            $txtTarget.Text = $dlg.SelectedPath
        }
    })
    $form.Controls.Add($btnBrowseTarget)

    # Neue Domäne
    $lblDomain = New-Object Windows.Forms.Label
    $lblDomain.Text = "Neue Domäne:"
    $lblDomain.Location = "10,100"
    $form.Controls.Add($lblDomain)

    $txtDomain = New-Object Windows.Forms.TextBox
    $txtDomain.Size = '250,20'
    $txtDomain.Location = "100,100"
    $form.Controls.Add($txtDomain)

    # OU Pfad
    $lblOU = New-Object Windows.Forms.Label
    $lblOU.Text = "Ziel-OU:"
    $lblOU.Location = "370,100"
    $form.Controls.Add($lblOU)

    $txtOU = New-Object Windows.Forms.TextBox
    $txtOU.Size = '270,20'
    $txtOU.Location = "430,100"
    $form.Controls.Add($txtOU)

    # Buttons
    $btnExport = New-Object Windows.Forms.Button
    $btnExport.Text = "1. Analysiere Berechtigungen"
    $btnExport.Location = "10,150"
    $btnExport.Size = '200,30'
    $btnExport.Add_Click({
        Export-Permissions -RootPath $txtSource.Text
    })
    $form.Controls.Add($btnExport)

    $btnMigrateGroups = New-Object Windows.Forms.Button
    $btnMigrateGroups.Text = "2. Gruppen migrieren"
    $btnMigrateGroups.Location = "230,150"
    $btnMigrateGroups.Size = '200,30'
    $btnMigrateGroups.Add_Click({
        Migrate-GroupsToNewDomain -NewDomain $txtDomain.Text -OUPath $txtOU.Text
    })
    $form.Controls.Add($btnMigrateGroups)

    $btnMigrateData = New-Object Windows.Forms.Button
    $btnMigrateData.Text = "3. Daten + Berechtigungen migrieren"
    $btnMigrateData.Location = "450,150"
    $btnMigrateData.Size = '220,30'
    $btnMigrateData.Add_Click({
        Migrate-DataAndApplyPermissions -SourcePath $txtSource.Text -TargetPath $txtTarget.Text
    })
    $form.Controls.Add($btnMigrateData)

    $form.Topmost = $true
    [void]$form.ShowDialog()
}

Start-GUI
