<#
.SYNOPSIS
    Manage the Microsoft Entra Passkey Dynamic Migration setting associated
    with the Microsoft-provided SMS/Voice retirement transition.

.DESCRIPTION
    Provides three options:

        1 - Verify current configuration
        2 - Opt OUT of automatic Passkey Dynamic Migration
        3 - Re-enable automatic Passkey Dynamic Migration

    Microsoft Graph property:

        optOutSettings.passkeyDynamicMigration

    IMPORTANT SEMANTICS:

        true  = OPT OUT of Microsoft's automatic Passkey Dynamic Migration
        false = Remove the opt-out / allow automatic Passkey Dynamic Migration

    IMPORTANT:
    - This controls the temporary automatic migration behavior associated
      with the September 2026 -> February 2027 transition.
    - It does NOT opt the tenant out of the February 1, 2027 retirement
      of Microsoft-provided SMS/Voice authentication.
    - It does NOT disable passkeys generally.
    - It does NOT necessarily disable a Registration Campaign explicitly
      configured by an administrator.
    - Option 3 removes the dynamic-migration opt-out. It should not be
      interpreted as directly enabling an administrator-created
      Registration Campaign.
    - The API used for this setting is Microsoft Graph BETA.

.REQUIREMENTS
    Microsoft Graph delegated permission:
        Policy.ReadWrite.AuthenticationMethod

    Recommended Microsoft Entra role:
        Authentication Policy Administrator

.NOTES
    Microsoft documentation:
    https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sms-voice-retirement
#>


# ============================================================
# Configuration
# ============================================================

$GraphUri = "https://graph.microsoft.com/beta/policies/authenticationmethodspolicy"


# ============================================================
# Function: Connect to Microsoft Graph
# ============================================================

function Connect-EntraGraph {

    Write-Host ""
    Write-Host "Checking Microsoft Graph PowerShell prerequisites..." `
        -ForegroundColor Cyan

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {

        Write-Host ""
        Write-Host "Microsoft.Graph.Authentication module not found." `
            -ForegroundColor Yellow

        Write-Host "Installing Microsoft.Graph.Authentication..." `
            -ForegroundColor Yellow

        try {

            Install-Module Microsoft.Graph.Authentication `
                -Scope CurrentUser `
                -Force `
                -AllowClobber `
                -ErrorAction Stop

        }
        catch {

            Write-Host ""
            Write-Host "ERROR installing Microsoft Graph PowerShell module." `
                -ForegroundColor Red

            Write-Host $_.Exception.Message `
                -ForegroundColor Red

            throw
        }
    }

    Import-Module Microsoft.Graph.Authentication `
        -ErrorAction Stop

    Write-Host ""
    Write-Host "Connecting to Microsoft Graph..." `
        -ForegroundColor Cyan

    Write-Host ""
    Write-Host "Requested delegated permission:"
    Write-Host "  Policy.ReadWrite.AuthenticationMethod"
    Write-Host ""

    try {

        Connect-MgGraph `
            -Scopes "Policy.ReadWrite.AuthenticationMethod" `
            -NoWelcome `
            -ErrorAction Stop

    }
    catch {

        Write-Host ""
        Write-Host "ERROR connecting to Microsoft Graph." `
            -ForegroundColor Red

        Write-Host $_.Exception.Message `
            -ForegroundColor Red

        throw
    }

    $context = Get-MgContext

    Write-Host ""
    Write-Host "Connected successfully." `
        -ForegroundColor Green

    Write-Host ""
    Write-Host "Tenant ID : $($context.TenantId)"
    Write-Host "Account   : $($context.Account)"
    Write-Host ""
}


# ============================================================
# Function: Read current configuration
# ============================================================

function Get-PasskeyDynamicMigrationStatus {

    param (
        [switch]$Quiet
    )

    if (-not $Quiet) {

        Write-Host ""
        Write-Host "Reading Authentication Methods Policy..." `
            -ForegroundColor Cyan
    }

    try {

        $policy = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $GraphUri `
            -ErrorAction Stop

    }
    catch {

        Write-Host ""
        Write-Host "ERROR reading Authentication Methods Policy." `
            -ForegroundColor Red

        Write-Host $_.Exception.Message `
            -ForegroundColor Red

        throw
    }


    $value = $null

    if ($null -ne $policy.optOutSettings) {

        $value = $policy.optOutSettings.passkeyDynamicMigration
    }


    if (-not $Quiet) {

        Write-Host ""
        Write-Host "======================================================" `
            -ForegroundColor White

        Write-Host " Passkey Dynamic Migration Configuration" `
            -ForegroundColor White

        Write-Host "======================================================" `
            -ForegroundColor White

        Write-Host ""
        Write-Host "Graph endpoint:"
        Write-Host "  $GraphUri"

        Write-Host ""
        Write-Host "Graph property:"
        Write-Host "  optOutSettings.passkeyDynamicMigration = $value"

        Write-Host ""

        if ($value -eq $true) {

            Write-Host "STATUS: OPTED OUT" `
                -ForegroundColor Yellow

            Write-Host ""
            Write-Host "Microsoft automatic Passkey Dynamic Migration"
            Write-Host "is currently opted out for this tenant."

            Write-Host ""
            Write-Host "Interpretation:"
            Write-Host "  passkeyDynamicMigration = TRUE"
            Write-Host "  -> Automatic dynamic migration OPT-OUT is active."

        }
        elseif ($value -eq $false) {

            Write-Host "STATUS: OPT-OUT NOT ACTIVE" `
                -ForegroundColor Green

            Write-Host ""
            Write-Host "Microsoft automatic Passkey Dynamic Migration"
            Write-Host "is allowed to apply to this tenant."

            Write-Host ""
            Write-Host "Interpretation:"
            Write-Host "  passkeyDynamicMigration = FALSE"
            Write-Host "  -> Automatic dynamic migration OPT-OUT is not active."

        }
        else {

            Write-Host "STATUS: UNKNOWN / PROPERTY NOT RETURNED" `
                -ForegroundColor Yellow

            Write-Host ""
            Write-Host "The passkeyDynamicMigration property was not"
            Write-Host "returned with an explicit True/False value."

            Write-Host ""
            Write-Host "Review the complete Authentication Methods Policy"
            Write-Host "or verify the current Microsoft Graph API behavior."
        }

        Write-Host ""
        Write-Host "IMPORTANT:" `
            -ForegroundColor Yellow

        Write-Host "This setting does NOT opt the tenant out of the"
        Write-Host "February 1, 2027 Microsoft-provided SMS/Voice retirement."

        Write-Host ""
        Write-Host "======================================================"
    }

    return $value
}


# ============================================================
# Function: Set configuration
# ============================================================

function Set-PasskeyDynamicMigration {

    param (

        [Parameter(Mandatory = $true)]
        [bool]$OptOut
    )


    $body = @{

        optOutSettings = @{

            passkeyDynamicMigration = $OptOut
        }

    } | ConvertTo-Json -Depth 5


    Write-Host ""
    Write-Host "Submitting configuration to Microsoft Graph..." `
        -ForegroundColor Cyan


    try {

        Invoke-MgGraphRequest `
            -Method PATCH `
            -Uri $GraphUri `
            -Body $body `
            -ContentType "application/json" `
            -ErrorAction Stop

        Write-Host ""
        Write-Host "Configuration submitted successfully." `
            -ForegroundColor Green

    }
    catch {

        Write-Host ""
        Write-Host "ERROR applying configuration." `
            -ForegroundColor Red

        Write-Host $_.Exception.Message `
            -ForegroundColor Red

        throw
    }
}


# ============================================================
# Function: Opt out
# ============================================================

function Enable-PasskeyMigrationOptOut {

    Write-Host ""
    Write-Host "======================================================" `
        -ForegroundColor Yellow

    Write-Host " OPT OUT OF AUTOMATIC PASSKEY DYNAMIC MIGRATION" `
        -ForegroundColor Yellow

    Write-Host "======================================================" `
        -ForegroundColor Yellow

    Write-Host ""
    Write-Host "This operation will set:"
    Write-Host ""

    Write-Host "  optOutSettings.passkeyDynamicMigration = TRUE" `
        -ForegroundColor Yellow

    Write-Host ""
    Write-Host "Meaning:"
    Write-Host ""
    Write-Host "  The tenant will OPT OUT of Microsoft's automatic"
    Write-Host "  Passkey Dynamic Migration behavior associated with"
    Write-Host "  the SMS/Voice retirement transition."

    Write-Host ""
    Write-Host "IMPORTANT:" `
        -ForegroundColor Yellow

    Write-Host ""
    Write-Host "This does NOT:"
    Write-Host ""
    Write-Host "  - Disable passkeys."
    Write-Host "  - Disable an administrator-created Registration Campaign."
    Write-Host "  - Preserve Microsoft-provided SMS/Voice after retirement."
    Write-Host "  - Opt out of the February 1, 2027 enforcement."

    Write-Host ""
    Write-Host "Checking current configuration..." `
        -ForegroundColor Cyan


    $currentValue = Get-PasskeyDynamicMigrationStatus -Quiet


    if ($currentValue -eq $true) {

        Write-Host ""
        Write-Host "No change required." `
            -ForegroundColor Green

        Write-Host ""
        Write-Host "The tenant is already opted out:"
        Write-Host ""

        Write-Host "  passkeyDynamicMigration = TRUE" `
            -ForegroundColor Green

        return
    }


    Write-Host ""
    Write-Host "Current value:"
    Write-Host ""

    Write-Host "  passkeyDynamicMigration = $currentValue"

    Write-Host ""


    $confirmation = Read-Host `
        "Do you want to OPT OUT of automatic Passkey Dynamic Migration? (Y/N)"


    if ($confirmation -notmatch '^[Yy]$') {

        Write-Host ""
        Write-Host "Operation cancelled." `
            -ForegroundColor Yellow

        Write-Host "No changes were made."

        return
    }


    Write-Host ""
    Write-Host "Applying opt-out..." `
        -ForegroundColor Cyan


    Set-PasskeyDynamicMigration -OptOut $true


    Write-Host ""
    Write-Host "Verifying configuration..." `
        -ForegroundColor Cyan


    Start-Sleep -Seconds 2


    $verifiedValue = Get-PasskeyDynamicMigrationStatus -Quiet


    if ($verifiedValue -eq $true) {

        Write-Host ""
        Write-Host "======================================================" `
            -ForegroundColor Green

        Write-Host " SUCCESS" `
            -ForegroundColor Green

        Write-Host "======================================================" `
            -ForegroundColor Green

        Write-Host ""
        Write-Host "The tenant is now opted out of Microsoft's automatic"
        Write-Host "Passkey Dynamic Migration."

        Write-Host ""
        Write-Host "Verified Graph value:"
        Write-Host ""

        Write-Host "  passkeyDynamicMigration = TRUE" `
            -ForegroundColor Green

        Write-Host ""
    }
    else {

        Write-Host ""
        Write-Host "WARNING: Verification failed." `
            -ForegroundColor Yellow

        Write-Host ""
        Write-Host "Expected:"
        Write-Host "  passkeyDynamicMigration = TRUE"

        Write-Host ""
        Write-Host "Returned:"
        Write-Host "  passkeyDynamicMigration = $verifiedValue"

        Write-Host ""
        Write-Host "Review the Authentication Methods Policy manually."
    }
}


# ============================================================
# Function: Remove opt-out / re-enable dynamic migration
# ============================================================

function Disable-PasskeyMigrationOptOut {

    Write-Host ""
    Write-Host "======================================================" `
        -ForegroundColor Green

    Write-Host " REMOVE OPT-OUT / RE-ENABLE DYNAMIC MIGRATION" `
        -ForegroundColor Green

    Write-Host "======================================================" `
        -ForegroundColor Green

    Write-Host ""
    Write-Host "This operation will set:"
    Write-Host ""

    Write-Host "  optOutSettings.passkeyDynamicMigration = FALSE" `
        -ForegroundColor Green

    Write-Host ""
    Write-Host "Meaning:"
    Write-Host ""
    Write-Host "  The temporary opt-out will be removed."
    Write-Host ""
    Write-Host "  Microsoft's automatic Passkey Dynamic Migration"
    Write-Host "  behavior will be allowed to apply again."

    Write-Host ""
    Write-Host "IMPORTANT:" `
        -ForegroundColor Yellow

    Write-Host ""
    Write-Host "This option removes the dynamic-migration opt-out."
    Write-Host ""
    Write-Host "It should NOT be interpreted as directly enabling"
    Write-Host "an administrator-created Registration Campaign."

    Write-Host ""
    Write-Host "Checking current configuration..." `
        -ForegroundColor Cyan


    $currentValue = Get-PasskeyDynamicMigrationStatus -Quiet


    if ($currentValue -eq $false) {

        Write-Host ""
        Write-Host "No change required." `
            -ForegroundColor Green

        Write-Host ""
        Write-Host "The dynamic-migration opt-out is already disabled:"
        Write-Host ""

        Write-Host "  passkeyDynamicMigration = FALSE" `
            -ForegroundColor Green

        return
    }


    Write-Host ""
    Write-Host "Current value:"
    Write-Host ""

    Write-Host "  passkeyDynamicMigration = $currentValue"

    Write-Host ""

    Write-Host "WARNING" `
        -ForegroundColor Yellow

    Write-Host ""
    Write-Host "Removing the opt-out can cause eligible users to"
    Write-Host "become subject again to Microsoft's automatic"
    Write-Host "Passkey Dynamic Migration / Microsoft-managed"
    Write-Host "registration experience."

    Write-Host ""


    $confirmation = Read-Host `
        "Do you want to REMOVE the opt-out? (Y/N)"


    if ($confirmation -notmatch '^[Yy]$') {

        Write-Host ""
        Write-Host "Operation cancelled." `
            -ForegroundColor Yellow

        Write-Host "No changes were made."

        return
    }


    Write-Host ""
    Write-Host "Removing opt-out..." `
        -ForegroundColor Cyan


    Set-PasskeyDynamicMigration -OptOut $false


    Write-Host ""
    Write-Host "Verifying configuration..." `
        -ForegroundColor Cyan


    Start-Sleep -Seconds 2


    $verifiedValue = Get-PasskeyDynamicMigrationStatus -Quiet


    if ($verifiedValue -eq $false) {

        Write-Host ""
        Write-Host "======================================================" `
            -ForegroundColor Green

        Write-Host " SUCCESS" `
            -ForegroundColor Green

        Write-Host "======================================================" `
            -ForegroundColor Green

        Write-Host ""
        Write-Host "The Passkey Dynamic Migration opt-out has been removed."

        Write-Host ""
        Write-Host "Verified Graph value:"
        Write-Host ""

        Write-Host "  passkeyDynamicMigration = FALSE" `
            -ForegroundColor Green

        Write-Host ""
        Write-Host "Microsoft's automatic migration behavior is now"
        Write-Host "allowed to apply again."

        Write-Host ""
    }
    else {

        Write-Host ""
        Write-Host "WARNING: Verification failed." `
            -ForegroundColor Yellow

        Write-Host ""
        Write-Host "Expected:"
        Write-Host "  passkeyDynamicMigration = FALSE"

        Write-Host ""
        Write-Host "Returned:"
        Write-Host "  passkeyDynamicMigration = $verifiedValue"

        Write-Host ""
        Write-Host "Review the Authentication Methods Policy manually."
    }
}


# ============================================================
# Main
# ============================================================

Clear-Host


Write-Host ""
Write-Host "======================================================" `
    -ForegroundColor Cyan

Write-Host " Microsoft Entra Passkey Dynamic Migration Manager" `
    -ForegroundColor Cyan

Write-Host "======================================================" `
    -ForegroundColor Cyan

Write-Host ""
Write-Host "Microsoft-provided SMS / Voice Retirement Transition"
Write-Host ""

Write-Host "Graph API:"
Write-Host "  Microsoft Graph BETA"

Write-Host ""
Write-Host "Required permission:"
Write-Host "  Policy.ReadWrite.AuthenticationMethod"

Write-Host ""
Write-Host "Recommended Entra role:"
Write-Host "  Authentication Policy Administrator"

Write-Host ""


# ============================================================
# Connect
# ============================================================

try {

    Connect-EntraGraph

}
catch {

    Write-Host ""
    Write-Host "Unable to continue." `
        -ForegroundColor Red

    Write-Host ""
    exit 1
}


# ============================================================
# Menu
# ============================================================

do {

    Write-Host ""
    Write-Host "------------------------------------------------------"
    Write-Host "Select an option:"
    Write-Host "------------------------------------------------------"
    Write-Host ""

    Write-Host "  1 - Verify current configuration" `
        -ForegroundColor Cyan

    Write-Host ""

    Write-Host "  2 - OPT OUT of automatic Passkey Dynamic Migration" `
        -ForegroundColor Yellow

    Write-Host ""

    Write-Host "  3 - REMOVE OPT-OUT / re-enable Dynamic Migration" `
        -ForegroundColor Green

    Write-Host ""

    Write-Host "  Q - Quit"

    Write-Host ""


    $choice = Read-Host "Selection"


    switch ($choice) {

        "1" {

            try {

                Get-PasskeyDynamicMigrationStatus | Out-Null

            }
            catch {

                Write-Host ""
                Write-Host "Unable to retrieve configuration." `
                    -ForegroundColor Red
            }
        }


        "2" {

            try {

                Enable-PasskeyMigrationOptOut

            }
            catch {

                Write-Host ""
                Write-Host "Unable to configure opt-out." `
                    -ForegroundColor Red
            }
        }


        "3" {

            try {

                Disable-PasskeyMigrationOptOut

            }
            catch {

                Write-Host ""
                Write-Host "Unable to remove opt-out." `
                    -ForegroundColor Red
            }
        }


        { $_ -match '^[Qq]$' } {

            Write-Host ""
            Write-Host "Exiting..." `
                -ForegroundColor Cyan
        }


        default {

            Write-Host ""
            Write-Host "Invalid selection." `
                -ForegroundColor Red

            Write-Host ""
            Write-Host "Please select 1, 2, 3, or Q."
        }
    }

}
until ($choice -match '^[Qq]$')


# ============================================================
# Disconnect
# ============================================================

try {

    Disconnect-MgGraph `
        -ErrorAction SilentlyContinue

}
catch {

    # Ignore disconnect errors
}


Write-Host ""
Write-Host "Disconnected from Microsoft Graph." `
    -ForegroundColor Cyan

Write-Host ""
Write-Host "Done."
Write-Host ""