<#
.SYNOPSIS
    Synchronizes the 'company' attribute from JumpCloud users to the 'division' attribute within the 'employeeOrgData' object for Microsoft 365 (Entra ID) users.

.DESCRIPTION
    This script connects to both the JumpCloud and Microsoft Graph APIs. It fetches the 'company' attribute from JumpCloud users
    and updates the 'division' property within the 'employeeOrgData' object for the corresponding M365 user.
    
    The script also preserves the existing 'costCenter' value by reading it before writing, as the API requires both
    properties to be included in any update to the 'employeeOrgData' object.
    This version uses delegated user permissions and requires an interactive login from an administrator.

.NOTES
    Prerequisites:
        - PowerShell 7+
        - JumpCloud PowerShell Module (Install-Module JumpCloud)
        - Microsoft Graph PowerShell SDK (Install-Module Microsoft.Graph)
        - A JumpCloud API Key.
        - An M365 user with permissions to read and write user attributes (e.g., User Administrator or Global Administrator role).
#>

#requires -Module JumpCloud
#requires -Module Microsoft.Graph

# --- SCRIPT CONFIGURATION ---
# Set to $true to run the script in a "what if" mode. No changes will be made to M365.
$DryRun = $false

# --- FUNCTIONS ---

Function Connect-ToJumpCloud {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ApiKey
    )
    try {
        Write-Host "Connecting to JumpCloud..." -ForegroundColor Cyan
        Connect-JCOnline -JumpCloudApiKey $ApiKey
        Write-Host "Successfully connected to JumpCloud." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to JumpCloud. The specific error was: $($_.Exception.Message)"
        exit 1
    }
}

Function Connect-ToMicrosoftGraph {
    # Define the required permissions (scopes) for the Microsoft Graph API
    $requiredScopes = @("User.Read.All", "User.ReadWrite.All", "RoleManagement.Read.Directory")

    try {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        # Check if already connected with the required permissions
        $token = Get-MgContext
        if ($null -eq $token -or !($requiredScopes | ForEach-Object { $token.Scopes -contains $_ } | Where-Object { $_ })) {
            Write-Host "No valid session found or missing permissions. Authenticating..." -ForegroundColor Yellow
            Connect-MgGraph -Scopes $requiredScopes
        }
        Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
        # Verify connection by getting the current user
        Get-MgContext
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph. The specific error was: $($_.Exception.Message)"
        exit 1
    }
}


# --- SCRIPT START ---

Write-Host "Starting JumpCloud Company to M365 Division Sync" -ForegroundColor Green

# 1. Authenticate to Services
$jcApiKey = Read-Host -Prompt "Please enter your JumpCloud API Key" -AsSecureString
$jcApiKeyPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($jcApiKey))

Connect-ToJumpCloud -ApiKey $jcApiKeyPlainText
Connect-ToMicrosoftGraph

# 2. Get All JumpCloud Users
Write-Host "Fetching all users from JumpCloud..." -ForegroundColor Cyan
$jcUsers = Get-JCUser -ErrorAction Stop | Select-Object email, company
$jcUsersHashTable = @{}
foreach ($user in $jcUsers) { if (-not [string]::IsNullOrEmpty($user.email)) { $jcUsersHashTable[$user.email.ToLower()] = $user.company } }
Write-Host "Found $($jcUsersHashTable.Count) users in JumpCloud with an email address." -ForegroundColor Green

# 3. Get All Microsoft 365 Users
Write-Host "Fetching all users from Microsoft 365..." -ForegroundColor Cyan
# Fetch the employeeOrgData object which contains division and costCenter
$m365Users = Get-MgUser -All -Property 'id', 'userPrincipalName', 'employeeOrgData' -ErrorAction Stop
Write-Host "Found $($m365Users.Count) users in Microsoft 365." -ForegroundColor Green

# 4. Compare and Synchronize
Write-Host "Comparing users and preparing for synchronization..." -ForegroundColor Cyan
$updatesCounter = 0
$skipsCounter = 0

foreach ($m365User in $m365Users) {
    $upnLower = $m365User.UserPrincipalName.ToLower()

    if ($jcUsersHashTable.ContainsKey($upnLower)) {
        $jcCompany = $jcUsersHashTable[$upnLower]
        $m365Division = $m365User.EmployeeOrgData.Division

        if (-not [string]::IsNullOrEmpty($jcCompany) -and $jcCompany -ne $m365Division) {
            Write-Host ("UPDATE REQUIRED for $($m365User.UserPrincipalName):") -ForegroundColor Yellow
            Write-Host ("  - JumpCloud Company: '$jcCompany'")
            Write-Host ("  - M365 Division:     '$m365Division'")

            if (-not $DryRun) {
                try {
                    $userId = $m365User.Id
                    $uri = "https://graph.microsoft.com/v1.0/users/$userId"
                    
                    # Per documentation, we must provide the entire employeeOrgData object.
                    # First, get the current costCenter to avoid overwriting it.
                    $currentCostCenter = $m365User.EmployeeOrgData.CostCenter
                    if ($null -eq $currentCostCenter) { $currentCostCenter = "" }

                    $orgDataPayload = @{
                        "division"   = $jcCompany
                        "costCenter" = $currentCostCenter
                    }

                    $bodyPayload = @{ "employeeOrgData" = $orgDataPayload }
                    
                    Invoke-MgGraphRequest -Uri $uri -Method 'PATCH' -Body ($bodyPayload | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop

                    Write-Host "  - SUCCESS: Updated M365 employeeOrgData." -ForegroundColor Green
                    $updatesCounter++
                }
                catch {
                    $errorMessage = $_.Exception.Message
                    if ($_.ErrorDetails) { $errorMessage += " | Details: $($_.ErrorDetails.Message)" }
                    Write-Error "  - Failed to update user $($m365User.UserPrincipalName). Error: $errorMessage"
                }
            } else {
                Write-Host "  - DRY RUN: No changes were made." -ForegroundColor Magenta
                $updatesCounter++
            }
        } else { $skipsCounter++ }
    } else { $skipsCounter++ }
}

Write-Host "`n--- Sync Summary ---" -ForegroundColor Green
if ($DryRun) {
    Write-Host "DRY RUN MODE ENABLED. No actual changes were made." -ForegroundColor Magenta
    Write-Host "$updatesCounter users would have been updated."
} else {
    Write-Host "$updatesCounter users were successfully updated in Microsoft 365."
}
Write-Host "$skipsCounter users were skipped (already in sync, no source data, or not found in JumpCloud)."
Write-Host "--------------------`n"

# 5. Disconnect Sessions
Write-Host "Disconnecting Microsoft Graph session..." -ForegroundColor Cyan
Disconnect-MgGraph
Write-Host "Script finished." -ForegroundColor Green

