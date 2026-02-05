<#
.BWSCheckRules.ps1
Zweck:
- Enthält alle Prüfregeln (Checks) als Objekte.
- Jede Regel beschreibt:
  - Id, Name, Product, Description, Severity
  - Requires (Graph/Exchange/Teams/SharePoint)
  - MinimumScopes (Graph Scopes für Interactive Auth)
  - ScriptBlock: die eigentliche Prüfung (nimmt $Context entgegen)

WIE ERSTELLE ICH EINE NEUE REGEL?
1) Kopiere eine bestehende Regel als Vorlage.
2) Vergib eine eindeutige Id (z.B. "ENTRA-XYZ-001").
3) Setze Product (EntraID/Intune/SPO/Teams/OneDrive/Exchange) und Severity.
4) Definiere Requires.* und MinimumScopes passend zur API.
5) Implementiere den ScriptBlock:
   - param($ctx)
   - Nutze $ctx.Helper.InvokeGraph(...) oder Exchange/Teams/SPO Cmdlets.
   - Gib IMMER ein Result-Objekt via New-BwsRuleResult zurück.
6) Füge die Regel am Ende zu $script:BwsCheckRules hinzu.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function New-BwsRuleResult {
    param(
        [Parameter(Mandatory)][string]$RuleId,
        [Parameter(Mandatory)][string]$RuleName,
        [Parameter(Mandatory)][string]$Product,
        [Parameter(Mandatory)][ValidateSet('Pass','Fail','Warn','Info','Error','Skipped')]
        [string]$Status,
        [string]$Severity = 'Info',
        [string]$Summary,
        $Details,
        [string]$Remediation,
        [string[]]$Evidence
    )

    [pscustomobject]@{
        Timestamp = (Get-Date).ToString("s")
        RuleId    = $RuleId
        RuleName  = $RuleName
        Product   = $Product
        Severity  = $Severity
        Status    = $Status
        Summary   = $Summary
        Details   = $Details
        Remediation = $Remediation
        Evidence  = $Evidence
    }
}

function New-BwsCheckRule {
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Product,
        [Parameter(Mandatory)][string]$Description,
        [ValidateSet('High','Medium','Low','Info')]
        [string]$Severity = 'Info',

        # Requires: welche Services müssen verbunden sein?
        [hashtable]$Requires = @{ Graph=$true; Exchange=$false; Teams=$false; SharePoint=$false },

        # Für Interactive Auth: welche Scopes braucht diese Regel minimal?
        [string[]]$MinimumScopes = @(),

        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )

    [pscustomobject]@{
        Id            = $Id
        Name          = $Name
        Product       = $Product
        Description   = $Description
        Severity      = $Severity
        Requires      = $Requires
        MinimumScopes = $MinimumScopes
        ScriptBlock   = $ScriptBlock
    }
}

# -----------------------------------------
# Beispiel-Regeln (anpassbar/erweiterbar)
# -----------------------------------------
$script:BwsCheckRules = @()

# ENTRA: Mind. 2 Global Admins (Break-Glass/Redundanz)
$script:BwsCheckRules += New-BwsCheckRule `
    -Id 'ENTRA-GA-COUNT' `
    -Name 'Global Admins: Mindestanzahl prüfen' `
    -Product 'EntraID' `
    -Severity 'High' `
    -Description 'Prüft, ob mindestens 2 Benutzer Mitglied der Rolle "Global Administrator" sind (Redundanz / Break-Glass).' `
    -Requires @{ Graph=$true; Exchange=$false; Teams=$false; SharePoint=$false } `
    -MinimumScopes @('RoleManagement.Read.Directory','Directory.Read.All') `
    -ScriptBlock {
        param($ctx)

        try {
            $role = $ctx.Helper.InvokeGraph("GET","/v1.0/directoryRoles?`$filter=displayName eq 'Global Administrator'")
            $roleId = $role.value[0].id
            if (-not $roleId) {
                return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                    -Status 'Skipped' -Summary 'Rolle "Global Administrator" ist evtl. nicht aktiviert (directoryRoles leer).' `
                    -Remediation 'Stelle sicher, dass die Rolle in directoryRoles sichtbar ist (Rollen werden erst nach Aktivierung gelistet).' `
                    -Evidence @('GET /directoryRoles?$filter=displayName eq ''Global Administrator''')
            }

            $members = $ctx.Helper.InvokeGraph("GET","/v1.0/directoryRoles/$roleId/members?`$select=id,displayName,userPrincipalName")
            $users = @($members.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' })
            $count = $users.Count

            $status = if ($count -ge 2) { 'Pass' } else { 'Fail' }
            $summary = "Gefunden: $count Global Admin(s)."

            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status $status -Summary $summary `
                -Details ($users | Select-Object displayName,userPrincipalName,id) `
                -Remediation 'Empfehlung: Mindestens 2 (besser 3) Global Admins, davon 1-2 Break-Glass Accounts, und starke MFA/CA Absicherung.' `
                -Evidence @("GET /directoryRoles/$roleId/members")
        }
        catch {
            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status 'Error' -Summary $_.Exception.Message `
                -Remediation 'Prüfe Graph-Berechtigungen (RoleManagement.Read.Directory / Directory.Read.All) und ob du korrekt authentifiziert bist.'
        }
    }

# ENTRA: Gäste-Benutzer Anzahl (Info/Warn)
$script:BwsCheckRules += New-BwsCheckRule `
    -Id 'ENTRA-GUEST-COUNT' `
    -Name 'Guests: Anzahl & Trend (Info)' `
    -Product 'EntraID' `
    -Severity 'Low' `
    -Description 'Zählt Guest-User (userType=Guest). Warnung ab Schwellwert.' `
    -Requires @{ Graph=$true; Exchange=$false; Teams=$false; SharePoint=$false } `
    -MinimumScopes @('Directory.Read.All') `
    -ScriptBlock {
        param($ctx)

        $threshold = 50
        try {
            $res = $ctx.Helper.InvokeGraph("GET","/v1.0/users?`$filter=userType eq 'Guest'&`$select=id,displayName,userPrincipalName,createdDateTime&`$top=999")
            $count = @($res.value).Count

            $status = if ($count -ge $threshold) { 'Warn' } else { 'Info' }
            $summary = "Guest-User: $count (Warn ab $threshold)."

            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status $status -Summary $summary `
                -Details (@($res.value) | Select-Object displayName,userPrincipalName,createdDateTime | Sort-Object createdDateTime -Descending | Select-Object -First 50) `
                -Remediation 'Empfehlung: Guest-Lifecycle (Access Reviews, Expiration), klare Owner-Verantwortung, externe Sharing Policies prüfen.' `
                -Evidence @("GET /users?filter=userType eq 'Guest'")
        }
        catch {
            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status 'Error' -Summary $_.Exception.Message `
                -Remediation 'Prüfe Graph-Berechtigungen (Directory.Read.All) und Auth.'
        }
    }

# INTUNE: Non-compliant Devices
$script:BwsCheckRules += New-BwsCheckRule `
    -Id 'INTUNE-NONCOMPLIANT' `
    -Name 'Intune: Non-compliant Devices (Top 50)' `
    -Product 'Intune' `
    -Severity 'Medium' `
    -Description 'Listet nicht-konforme Geräte (complianceState != compliant) und gibt Warnung bei >0.' `
    -Requires @{ Graph=$true; Exchange=$false; Teams=$false; SharePoint=$false } `
    -MinimumScopes @('DeviceManagementManagedDevices.Read.All') `
    -ScriptBlock {
        param($ctx)

        try {
            # Hinweis: Filter-Syntax je nach Tenant/API; wir nutzen eine robuste Variante ohne zu harte Filter:
            $uri = "/v1.0/deviceManagement/managedDevices?`$select=id,deviceName,operatingSystem,osVersion,complianceState,lastSyncDateTime,userPrincipalName&`$top=999"
            $res = $ctx.Helper.InvokeGraph("GET",$uri)

            $non = @($res.value | Where-Object { $_.complianceState -and $_.complianceState -ne 'compliant' })
            $count = $non.Count

            $status = if ($count -gt 0) { 'Warn' } else { 'Pass' }
            $summary = "Non-compliant Devices: $count"

            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status $status -Summary $summary `
                -Details ($non | Sort-Object lastSyncDateTime -Descending | Select-Object -First 50 deviceName,operatingSystem,osVersion,complianceState,lastSyncDateTime,userPrincipalName) `
                -Remediation 'Empfehlung: Compliance Policies, Conditional Access (Require compliant device), Ausnahmen prüfen und Remediation Workflows etablieren.' `
                -Evidence @("GET $uri")
        }
        catch {
            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status 'Error' -Summary $_.Exception.Message `
                -Remediation 'Prüfe Graph-Berechtigungen (DeviceManagementManagedDevices.Read.All) und Intune/Graph Zugriff.'
        }
    }

# TEAMS: Teams ohne Owner (Graph-basiert über M365 Groups)
$script:BwsCheckRules += New-BwsCheckRule `
    -Id 'TEAMS-OWNER-HYGIENE' `
    -Name 'Teams: Owner-Hygiene (0/1 Owner)' `
    -Product 'Teams' `
    -Severity 'Medium' `
    -Description 'Findet Teams (M365 Groups mit Team) und prüft Owner-Anzahl. Warnung bei 0 oder 1 Owner.' `
    -Requires @{ Graph=$true; Exchange=$false; Teams=$false; SharePoint=$false } `
    -MinimumScopes @('Group.Read.All') `
    -ScriptBlock {
        param($ctx)

        try {
            $groups = $ctx.Helper.InvokeGraph("GET","/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,visibility&`$top=200")

            $findings = New-Object System.Collections.Generic.List[object]

            foreach ($g in @($groups.value)) {
                $owners = $ctx.Helper.InvokeGraph("GET","/v1.0/groups/$($g.id)/owners?`$select=id,displayName,userPrincipalName&`$top=999")
                $ownerUsers = @($owners.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' })
                $oc = $ownerUsers.Count

                if ($oc -le 1) {
                    $findings.Add([pscustomobject]@{
                        TeamName   = $g.displayName
                        Visibility = $g.visibility
                        OwnerCount = $oc
                        Owners     = ($ownerUsers | ForEach-Object { $_.userPrincipalName } ) -join '; '
                        GroupId    = $g.id
                    })
                }
            }

            $count = $findings.Count
            $status = if ($count -gt 0) { 'Warn' } else { 'Pass' }
            $summary = "Teams mit OwnerCount <= 1: $count"

            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status $status -Summary $summary `
                -Details ($findings | Sort-Object OwnerCount,TeamName | Select-Object -First 200) `
                -Remediation 'Empfehlung: pro Team mind. 2 Owner, Lifecycle/Owner-Prozess definieren, ggf. Owner-Automation (z.B. via Governance/Workflows).' `
                -Evidence @('GET /groups?filter=resourceProvisioningOptions/Any(x:x eq ''Team'')','GET /groups/{id}/owners')
        }
        catch {
            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status 'Error' -Summary $_.Exception.Message `
                -Remediation 'Prüfe Graph-Berechtigungen (Group.Read.All) und Auth.'
        }
    }

# SHAREPOINT ONLINE: Tenant Sharing Capability (SPO Mgmt Shell)
$script:BwsCheckRules += New-BwsCheckRule `
    -Id 'SPO-SHARING-CAPABILITY' `
    -Name 'SharePoint: External Sharing Capability (Tenant)' `
    -Product 'SharePoint Online' `
    -Severity 'High' `
    -Description 'Liest SharingCapability des SPO Tenants (Get-SPOTenant). Bewertung je nach Setting.' `
    -Requires @{ Graph=$false; Exchange=$false; Teams=$false; SharePoint=$true } `
    -MinimumScopes @() `
    -ScriptBlock {
        param($ctx)

        try {
            if (-not (Get-Command Get-SPOTenant -ErrorAction SilentlyContinue)) {
                return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                    -Status 'Skipped' -Summary 'Get-SPOTenant nicht verfügbar (SPO Modul fehlt oder keine SPO-Connection).' `
                    -Remediation 'Installiere Microsoft.Online.SharePoint.PowerShell und verbinde mit Connect-SPOService.'
            }

            $t = Get-SPOTenant
            $cap = $t.SharingCapability

            # Grobe Bewertung (deine Policy kann abweichen)
            $status = switch ($cap) {
                'Disabled' { 'Pass' }
                'ExternalUserSharingOnly' { 'Warn' }
                'ExternalUserAndGuestSharing' { 'Fail' }
                default { 'Info' }
            }

            $summary = "SharingCapability = $cap"
            $details = [pscustomobject]@{
                SharingCapability = $cap
                DefaultSharingLinkType = $t.DefaultSharingLinkType
                PreventExternalUsersFromResharing = $t.PreventExternalUsersFromResharing
                FileAnonymousLinkType = $t.FileAnonymousLinkType
                FolderAnonymousLinkType = $t.FolderAnonymousLinkType
            }

            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status $status -Summary $summary -Details $details `
                -Remediation 'Empfehlung: Externes Sharing minimieren, "Anyone"-Links vermeiden/abschalten, Resharing kontrollieren, Sensitivity Labels & DLP ergänzen.' `
                -Evidence @('Get-SPOTenant')
        }
        catch {
            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status 'Error' -Summary $_.Exception.Message `
                -Remediation 'Prüfe SPO Admin Rechte, SharePointAdminUrl und Verbindung (Connect-SPOService).'
        }
    }

# EXCHANGE ONLINE: Admin Audit Log Enabled
$script:BwsCheckRules += New-BwsCheckRule `
    -Id 'EXO-ADMIN-AUDIT' `
    -Name 'Exchange: Admin Audit Log aktiviert' `
    -Product 'Exchange Online' `
    -Severity 'High' `
    -Description 'Prüft, ob AdminAuditLogEnabled aktiv ist (Get-AdminAuditLogConfig).' `
    -Requires @{ Graph=$false; Exchange=$true; Teams=$false; SharePoint=$false } `
    -MinimumScopes @() `
    -ScriptBlock {
        param($ctx)

        try {
            if (-not (Get-Command Get-AdminAuditLogConfig -ErrorAction SilentlyContinue)) {
                return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                    -Status 'Skipped' -Summary 'Get-AdminAuditLogConfig nicht verfügbar (EXO Modul fehlt oder keine EXO-Connection).' `
                    -Remediation 'Installiere ExchangeOnlineManagement und verbinde mit Connect-ExchangeOnline.'
            }

            $cfg = Get-AdminAuditLogConfig
            $enabled = [bool]$cfg.AdminAuditLogEnabled

            $status = if ($enabled) { 'Pass' } else { 'Fail' }
            $summary = "AdminAuditLogEnabled = $enabled"

            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status $status -Summary $summary -Details $cfg `
                -Remediation 'Empfehlung: Admin Audit Log aktivieren und Aufbewahrung/Export/Alerting definieren.' `
                -Evidence @('Get-AdminAuditLogConfig')
        }
        catch {
            return New-BwsRuleResult -RuleId $ctx.Rule.Id -RuleName $ctx.Rule.Name -Product $ctx.Rule.Product -Severity $ctx.Rule.Severity `
                -Status 'Error' -Summary $_.Exception.Message `
                -Remediation 'Prüfe EXO Rechte, Auth und Modulversion.'
        }
    }
