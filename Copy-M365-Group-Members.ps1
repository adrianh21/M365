<#
.SYNOPSIS
Copy members between M365 groups
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[\w\.\-]+@[\w\.\-]+\.[a-zA-Z]{2,}$')]
    [string]$SourceGroup,

    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[\w\.\-]+@[\w\.\-]+\.[a-zA-Z]{2,}$')]
    [string]$DestinationGroup,

    [string]$LogPath
)


# --- Interactive prompts if not provided ---
if (-not $SourceGroup)      { $SourceGroup      = Read-Host "Enter Source Group email address" }
if (-not $DestinationGroup) { $DestinationGroup = Read-Host "Enter Destination Group email address" }

# --- Fail-fast group existence checks ---
try { $null = Get-UnifiedGroup -Identity $SourceGroup -ErrorAction Stop }
catch { throw "Source group not found: $SourceGroup" }

try { $null = Get-UnifiedGroup -Identity $DestinationGroup -ErrorAction Stop }
catch { throw "Destination group not found: $DestinationGroup" }

# --- Optional logging ---
if ($LogPath) {
    $folder = Split-Path -Path $LogPath
    if ($folder -and -not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    New-Item -Path $LogPath -ItemType File -Force | Out-Null
    "Action,Member,Group,Timestamp,Message" | Out-File -FilePath $LogPath -Encoding utf8
}

# --- Get members ---
$srcMembers = (Get-UnifiedGroupLinks -Identity $SourceGroup -LinkType Members -ResultSize Unlimited -ErrorAction Stop).PrimarySmtpAddress
$dstSet     = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$dstMembers = (Get-UnifiedGroupLinks -Identity $DestinationGroup -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue).PrimarySmtpAddress

foreach ($m in @($dstMembers)) {
    if ($m) { $null = $dstSet.Add($m.Trim()) }
}

$added = 0; $skipped = 0; $errors = 0

foreach ($m in @($srcMembers)) {
    $addr = ($m | ForEach-Object { $_.ToString().Trim() })
    if (-not $addr) { $skipped++; continue }
    if ($dstSet.Contains($addr)) { $skipped++; continue }

    if ($PSCmdlet.ShouldProcess($DestinationGroup, "Add member $addr")) {
        try {
            Add-UnifiedGroupLinks -Identity $DestinationGroup -LinkType Members -Links $addr -Confirm:$false -ErrorAction Stop
            $added++
            if ($LogPath) { "Added,$addr,$DestinationGroup,$((Get-Date).ToString('s'))," | Out-File -FilePath $LogPath -Append -Encoding utf8 }
        } catch {
            $errors++
            if ($LogPath) { "Error,$addr,$DestinationGroup,$((Get-Date).ToString('s')),$($_.Exception.Message)" | Out-File -FilePath $LogPath -Append -Encoding utf8 }
        }
    }
}

Write-Host "Processed: $(@($srcMembers).Count)"
Write-Host "Added:    $added"
Write-Host "Skipped:  $skipped"
Write-Host "Errors:   $errors"