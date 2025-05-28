<#
.SYNOPSIS
Checks users from CSV for W365 group membership and disabled account status.

.DESCRIPTION
This script reads user principal names from a CSV file, checks if they are members 
of groups containing 'w365' in the name, and identifies disabled accounts. Results 
are displayed in the terminal and optionally exported to CSV.

.PARAMETER CsvPath
Path to the input CSV file containing user principal names.

.PARAMETER OutputCsvPath
Optional path to export results to CSV file.

.PARAMETER UPNColumnName
Name of the column containing user principal names (default: "UserPrincipalName").

.EXAMPLE
.\Get-GroupDisabledMembers.ps1 -CsvPath "C:\users.csv"

.EXAMPLE
.\Get-GroupDisabledMembers.ps1 -CsvPath "C:\users.csv" -OutputCsvPath "C:\results.csv" -UPNColumnName "Email"

.NOTES
Requires Microsoft.Graph PowerShell SDK and appropriate permissions:
- User.Read.All
- Group.Read.All  
- GroupMember.Read.All
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputCsvPath = $null,
    
    [Parameter(Mandatory=$false)]
    [string]$UPNColumnName = "UserPrincipalName"
)

# Import required modules
try {
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}
catch {
    Write-Error "Failed to import Microsoft Graph modules. Please install Microsoft.Graph PowerShell SDK."
    exit 1
}

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
try {
    Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "GroupMember.Read.All" -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

# Read CSV file
Write-Host "Reading CSV file: $CsvPath" -ForegroundColor Yellow
try {
    $users = Import-Csv -Path $CsvPath -ErrorAction Stop
    Write-Host "Found $($users.Count) users in CSV" -ForegroundColor Green
}
catch {
    Write-Error "Failed to read CSV file: $($_.Exception.Message)"
    exit 1
}

# Initialize results array
$results = @()

# Process each user
Write-Host "Processing $($users.Count) users..." -ForegroundColor Yellow
$counter = 0
$errorCount = 0

foreach ($user in $users) {
    $counter++
    $upn = $user.$UPNColumnName
    
    # Show progress every 50 users
    if ($counter % 50 -eq 0 -or $counter -eq $users.Count) {
        Write-Progress -Activity "Processing Users" -Status "Processed $counter of $($users.Count)" -PercentComplete (($counter / $users.Count) * 100)
    }
    
    if ([string]::IsNullOrEmpty($upn)) {
        $errorCount++
        continue
    }
    
    try {
        # Get user details
        $mgUser = Get-MgUser -UserId $upn -Property "Id,UserPrincipalName,AccountEnabled" -ErrorAction Stop
        
        # Get user's group memberships
        $userGroups = Get-MgUserMemberOf -UserId $mgUser.Id -All -ErrorAction Stop
        
        # Filter groups that contain 'w365' in the name
        $w365Groups = $userGroups | Where-Object { 
            $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group" -and 
            $_.AdditionalProperties["displayName"] -like "*w365*" 
        }
        
        # Check if user is disabled and member of w365 groups
        if (-not $mgUser.AccountEnabled -and $w365Groups.Count -gt 0) {
            foreach ($group in $w365Groups) {
                $result = [PSCustomObject]@{
                    UserPrincipalName = $mgUser.UserPrincipalName
                    UserId = $mgUser.Id
                    GroupName = $group.AdditionalProperties["displayName"]
                    GroupId = $group.Id
                    AccountEnabled = $mgUser.AccountEnabled
                }
                $results += $result
            }
        }
    }
    catch {
        $errorCount++
        # Silently continue processing
    }
}

# Clear progress bar
Write-Progress -Activity "Processing Users" -Completed

# Display results
Write-Host "`n=== PROCESSING SUMMARY ===" -ForegroundColor Yellow
Write-Host "Total users processed: $counter" -ForegroundColor White
Write-Host "Errors encountered: $errorCount" -ForegroundColor $(if($errorCount -gt 0) { "Yellow" } else { "Green" })
Write-Host "Disabled users in W365 groups: $($results.Count)" -ForegroundColor $(if($results.Count -gt 0) { "Red" } else { "Green" })

Write-Host "`n=== RESULTS ===" -ForegroundColor Yellow

if ($results.Count -gt 0) {
    $results | Format-Table -AutoSize
    
    # Export to CSV if specified
    if ($OutputCsvPath) {
        try {
            $results | Export-Csv -Path $OutputCsvPath -NoTypeInformation -ErrorAction Stop
            Write-Host "Results exported to: $OutputCsvPath" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to export results to CSV: $($_.Exception.Message)"
        }
    }
}
else {
    Write-Host "No disabled users found in W365 groups." -ForegroundColor Green
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph
Write-Host "`nScript completed successfully!" -ForegroundColor Green