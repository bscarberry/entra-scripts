# Add devices using Computer Name from a CSV file to a Microsoft 365 group
# Author: Brandon Scarberry
# Date: 2025-05-20

# Connect to Microsoft Graph (if not already connected)
Connect-MgGraph -Scopes "Directory.ReadWrite.All", "GroupMember.ReadWrite.All"

# Set the Group ID
$groupId = "<Your-Group-Id-Here>"

# Import the CSV file with a 'ComputerName' column
$computers = Import-Csv -Path "<Path-To-CSV-File>"

$addedCount = 0

foreach ($computer in $computers) {
    # Find the device object by display name (computer name)
    $device = Get-MgDevice -Filter "displayName eq '$($computer.ComputerName)'" -Property Id,DisplayName

    if ($device) {
        try {
            Add-MgGroupMember -GroupId $groupId -DirectoryObjectId $device.Id
            Write-Host "Added $($computer.ComputerName) to group."
            $addedCount++
        } catch {
            Write-Warning "Failed to add $($computer.ComputerName): $_"
        }
    } else {
        Write-Warning "Device $($computer.ComputerName) not found."
    }
}

Write-Host "Total devices added: $addedCount"