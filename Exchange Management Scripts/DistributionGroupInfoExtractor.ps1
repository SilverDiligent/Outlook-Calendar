<#
.SYNOPSIS
This script extracts information about distribution groups from a CSV file and retrieves the manager details for each group.

.DESCRIPTION
The script reads the content of a CSV file, performs some modifications on the content, and saves it as a text file. It then reads the group names from the text file and retrieves the manager details for each group using the Get-DistributionGroup cmdlet. If a group has a manager set, it outputs the group name and manager details. If a group does not have a manager set, it outputs a message indicating that. If a group is not found, it outputs a message indicating that as well.

.PARAMETER FilePath
The path to the CSV file containing the distribution group names.

.OUTPUTS
The script outputs the group names and manager details for each group found.

.EXAMPLE
.\DistributionGroupInfoExtractor.ps1

This example runs the script using the default file path "./task0871534.csv" and outputs the group names and manager details.

.NOTES
Author: [Alexis Crawford]
Date: [12/23/2023]
Version: [Version 1.0]
#>
$filePath = "./task0871534.csv"
$fileContent = get-content -Path $filePath

$modifiedContent = $fileContent -replace '"','' -replace ',',"`r`n" -replace '-','' -replace 'Platforms ', "Platforms`r`n"

$lines = $modifiedContent -split "`r`n"
$trimmedLines = $lines | foreach-object {
  $_.Trim()
}

# $finalString = $trimmedLines -join "`r`n"
# Skip the first line and join the rest
$finalString = ($trimmedLines | select-object -skip 1) -join "`r`n"

$textFilePath = "./task0871534.txt"

 $finalString | set-content -Path $textFilePath

write-output "Text file saved at: $textFilePath"

$groupNames = get-content -Path $textFilePath

foreach ($groupName in $groupNames) {
  if ([string]::IsNullOrWhiteSpace($groupName)) {
    continue
  }

  $group = get-distributiongroup -Identity $groupName -ErrorAction SilentlyContinue

  if ($group) {
    if ($group.Managedby) {
      # Output group name and manager details
      Write-Output "Group: $groupName is managed by: $($group.ManagedBy)"
    } else {  
      Write-Output "Group: $groupName does not have a manager set. "
    } else {
      Write-Output "Group: '$groupName' not found."
    }
  }
}


