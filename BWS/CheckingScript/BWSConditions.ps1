<#
BWSConditions.ps1
This file defines the checks (“conditions”). It is loaded by BWSCheckScript.ps1.

FORMAT
- This file MUST return an array of PSCustomObjects.
- Each condition should be a PSCustomObject with these (recommended) properties:

  Id              : unique ID (String), e.g. "ENTRA-001"
  Product         : 'EntraID' | 'Azure' | 'AVD' | 'Intune' | 'AD'
  Title           : short headline
  Description     : (optional) longer explanation
  Severity        : 'Low' | 'Medium' | 'High' | 'Critical'
  Tags            : String[] (optional), e.g. @('MFA','Baseline')
  Expected        : (optional) text/object describing the desired state (used in report)
  RequiresGraph   : [bool] requires Microsoft Graph connectivity
  GraphScopes     : String[] delegated scopes needed (merged across conditions)
  RequiresAzure   : [bool] requires Az connectivity
  RequiresAD      : [bool] requires ActiveDirectory module (RSAT)
  Remediation     : remediation text (shown in report)
  Test            : Scriptblock with signature:
                    param([hashtable]$Context, [pscustomobject]$Condition)
                    return PSCustomObject:
                       IsCompliant : [bool]
                       Actual      : any
                       Expected    : (optional)
                       Evidence    : (optional)
                       Message     : (optional)

BEST PRACTICES
- Keep tests idempotent and read-only.
- Put concrete identifiers in Evidence (PolicyId, SubscriptionId, etc.).
- Keep Graph scopes minimal (least privilege).
#>

return @(
    # ---------------- Entra ID (Graph) ----------------
    [pscustomobject]@{
        Id            = "ENTRA-001"
        Product       = "EntraID"
        Title         = "Security Defaults are enabled"
        Description   = "Checks whether Entra Security Defaults are enabled (quick baseline when no CA policies exist)."
        Severity      = "High"
        Tags          = @("Baseline","MFA")
        Expected      = "identitySecurityDefaultsEnforcementPolicy.isEnabled = true"
        RequiresGraph = $true
        GraphScopes   = @("Policy.Read.All")
        RequiresAzure = $false
        RequiresAD    = $false
        Remediation   = @"
If you do NOT use Conditional Access:
- Enable Security Defaults in the Entra admin center.
If you DO use Conditional Access:
- Ensure MFA / block legacy auth / privileged access controls are covered by CA policies.
"@
        Test          = {
            param([hashtable]$Context, [pscustomobject]$Condition)

            # Microsoft Graph SDK:
            # Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy
            $p = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy

            [pscustomobject]@{
                IsCompliant = [bool]$p.IsEnabled
                Actual      = "IsEnabled=$($p.IsEnabled)"
                Evidence    = ($p | ConvertTo-Json -Depth 5)
                Message     = $null
            }
        }
    }

    [pscustomobject]@{
        Id            = "ENTRA-002"
        Product       = "EntraID"
        Title         = "At least one Conditional Access policy exists (signal check)"
        Description   = "Basic existence check. This does NOT confirm correct MFA / legacy auth coverage."
        Severity      = "Medium"
        Tags          = @("CA")
        Expected      = ">= 1 Conditional Access policy"
        RequiresGraph = $true
        GraphScopes   = @("Policy.Read.All")
        RequiresAzure = $false
        RequiresAD    = $false
        Remediation   = "Create CA policies (e.g., MFA for all users, block legacy auth, harden admin roles) based on your target design."
        Test          = {
            param([hashtable]$Context, [pscustomobject]$Condition)

            $pol = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
            $count = @($pol).Count

            [pscustomobject]@{
                IsCompliant = ($count -ge 1)
                Actual      = $count
                Evidence    = ($pol | Select-Object -First 10 -Property Id,DisplayName,State | Format-Table | Out-String)
                Message     = if ($count -eq 0) { "No Conditional Access policies found." } else { $null }
            }
        }
    }

    # ---------------- Azure (Az) ----------------
    [pscustomobject]@{
        Id            = "AZ-001"
        Product       = "Azure"
        Title         = "Subscriptions have tags 'Owner' and 'CostCenter'"
        Description   = "Governance baseline example: enforce tagging."
        Severity      = "Low"
        Tags          = @("Governance","Tags")
        Expected      = "All subscriptions contain tags Owner and CostCenter"
        RequiresGraph = $false
        GraphScopes   = @()
        RequiresAzure = $true
        RequiresAD    = $false
        Remediation   = @"
Add tags at subscription level or enforce via Azure Policy (Append/Modify/Audit).
"@
        Test          = {
            param([hashtable]$Context, [pscustomobject]$Condition)

            Import-Module Az.Resources -ErrorAction Stop

            $subs = Get-AzSubscription -ErrorAction Stop
            $missing = foreach ($s in $subs) {
                $t = $s.Tags
                $hasOwner = $t.ContainsKey("Owner")
                $hasCC    = $t.ContainsKey("CostCenter")
                if (-not ($hasOwner -and $hasCC)) {
                    [pscustomobject]@{
                        SubscriptionId = $s.Id
                        Name           = $s.Name
                        Missing        = (@(
                            if (-not $hasOwner) { "Owner" }
                            if (-not $hasCC)    { "CostCenter" }
                        ) -join ",")
                    }
                }
            }

            [pscustomobject]@{
                IsCompliant = (@($missing).Count -eq 0)
                Actual      = if (@($missing).Count -eq 0) { "OK" } else { "$(@($missing).Count) subscription(s) missing tags" }
                Evidence    = ($missing | Format-Table | Out-String)
                Message     = $null
            }
        }
    }

    # ---------------- Azure Virtual Desktop (Az.DesktopVirtualization) ----------------
    [pscustomobject]@{
        Id            = "AVD-001"
        Product       = "AVD"
        Title         = "Host pools: StartVMOnConnect is enabled"
        Description   = "Cost/UX baseline example."
        Severity      = "Medium"
        Tags          = @("Cost","AVD")
        Expected      = "HostPool.StartVMOnConnect = true"
        RequiresGraph = $false
        GraphScopes   = @()
        RequiresAzure = $true
        RequiresAD    = $false
        Remediation   = "Enable Start VM on Connect per host pool (portal or Update-AzWvdHostPool)."
        Test          = {
            param([hashtable]$Context, [pscustomobject]$Condition)

            if (-not (Get-Module -ListAvailable -Name Az.DesktopVirtualization)) {
                throw "Az.DesktopVirtualization is missing. Install-Module Az.DesktopVirtualization"
            }
            Import-Module Az.DesktopVirtualization -ErrorAction Stop

            $hps = Get-AzWvdHostPool -ErrorAction Stop
            $bad = $hps | Where-Object { -not $_.StartVMOnConnect } | Select-Object Name, ResourceGroupName, Location, StartVMOnConnect

            [pscustomobject]@{
                IsCompliant = (@($bad).Count -eq 0)
                Actual      = if (@($bad).Count -eq 0) { "OK" } else { "Non-compliant: $(@($bad).Count) host pool(s)" }
                Evidence    = ($bad | Format-Table | Out-String)
                Message     = $null
            }
        }
    }

    # ---------------- Intune (Graph DeviceManagement) ----------------
    [pscustomobject]@{
        Id            = "INTUNE-001"
        Product       = "Intune"
        Title         = "At least one Compliance Policy exists"
        Description   = "Signal check: without compliance policies, compliance-based Conditional Access is often ineffective."
        Severity      = "High"
        Tags          = @("Intune","Baseline")
        Expected      = ">= 1 Device Compliance Policy"
        RequiresGraph = $true
        GraphScopes   = @("DeviceManagementConfiguration.Read.All")
        RequiresAzure = $false
        RequiresAD    = $false
        Remediation   = "Create at least one compliance policy (e.g., BitLocker, PIN, OS version, jailbreak/root, etc.)."
        Test          = {
            param([hashtable]$Context, [pscustomobject]$Condition)

            $pol = Get-MgDeviceManagementDeviceCompliancePolicy -All -ErrorAction Stop
            $count = @($pol).Count

            [pscustomobject]@{
                IsCompliant = ($count -ge 1)
                Actual      = $count
                Evidence    = ($pol | Select-Object -First 10 -Property Id,DisplayName,CreatedDateTime | Format-Table | Out-String)
                Message     = $null
            }
        }
    }

    # ---------------- Active Directory (RSAT) ----------------
    [pscustomobject]@{
        Id            = "AD-001"
        Product       = "AD"
        Title         = "Default domain password policy: MinLength >= 12"
        Description   = "Checks the Default Domain Password Policy (example)."
        Severity      = "High"
        Tags          = @("AD","Password")
        Expected      = "MinPasswordLength >= 12"
        RequiresGraph = $false
        GraphScopes   = @()
        RequiresAzure = $false
        RequiresAD    = $true
        Remediation   = "Adjust the default domain password policy (GPO) or use FGPP depending on your design."
        Test          = {
            param([hashtable]$Context, [pscustomobject]$Condition)

            if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
                throw "ActiveDirectory module is missing (RSAT)."
            }
            Import-Module ActiveDirectory -ErrorAction Stop

            $p = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop

            [pscustomobject]@{
                IsCompliant = ($p.MinPasswordLength -ge 12)
                Actual      = "MinPasswordLength=$($p.MinPasswordLength)"
                Evidence    = ($p | Select-Object MinPasswordLength,PasswordHistoryCount,MaxPasswordAge,MinPasswordAge,ComplexityEnabled | Format-List | Out-String)
                Message     = $null
            }
        }
    }
)
