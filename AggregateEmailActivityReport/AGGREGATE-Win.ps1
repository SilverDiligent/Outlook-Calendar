<#
.SYNOPSIS
This script retrieves email activity counts from Microsoft Graph API and adds them as list items in SharePoint.

Author: Alexis Crawford
Created: 2024-01-01

.DESCRIPTION
The script performs the following steps:
1. Checks if the MSAL.PS module is installed and installs it if not.
2. Loads the configuration data from a JSON file.
3. Sets the configuration parameters.
4. Acquires an access token using MSAL.PS module.
5. Retrieves email activity counts from Microsoft Graph API.
6. Parses the response and converts it to CSV format.
7. Filters the data to include only the previous month's activity.
8. Prepares a list item for each row of data.
9. Posts the list item to SharePoint using Microsoft Graph API.

.PARAMETER None

.EXAMPLE
.\AGGREGATE-Win.ps1

.NOTES
- This script requires the MSAL.PS module to be installed.
- The configuration data should be provided in a JSON file named 'config.json' in the same directory as the script.
- The script assumes that the SharePoint site, list, and required permissions are already set up.
- The script uses the 'getEmailActivityCounts' endpoint of Microsoft Graph API to retrieve email activity counts.
- The script requires an Azure AD application with the necessary permissions to access Microsoft Graph API.
- The script logs any errors that occur during the API calls or data processing.
#>


# Check if MSAL.PS is installed, if not, install it
if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
  Install-Module MSAL.PS -Force -Scope CurrentUser
}
# Loading the configuration data from the JSON file
$configData = get-content -path './config.json' | ConvertFrom-Json

# Set the configuration parameters
$clientId = $configData.clientId
$clientSecret = $configData.clientSecret
$clientSecretString = $configData.clientSecretString
$tenantId = $configData.tenantId
$siteId = $configData.siteId
$listId = $configData.listId
$scope = $configData.scope
$clientSecret = convertTo-SecureString -String $clientSecretString -asPlainText -force

# Acquire the Access Token
$token = Get-MsalToken -ClientId $clientId -TenantId $tenantId -ClientSecret $clientSecret -Scopes $scope

# Get Email Activity Counts
$activityCountsUrl = "https://graph.microsoft.com/v1.0/reports/getEmailActivityCounts(period='D90')"

# Delay for 3 second
Start-Sleep -Seconds 3
try {
  $activityCountsResponse = Invoke-RestMethod -Uri $activityCountsUrl -Headers @{Authorization = "Bearer $($token.AccessToken)" } -ContentType "application/json"
  $csvData = ConvertFrom-csv $activityCountsResponse
}
catch {
  write-host "Error occurred while making the API call: $_"
  return
}

$lastMonth = (Get-Date).Addmonths(-1)
# Initialize the variables
$receivedItems = 0
$sentItems = 0
$reportDate = ""


foreach ($row in $csvData) {
  try {
    # Get the report date as a DateTime object
    $reportDateObject = [dateTime]::ParseExact($row.'Report Date', "yyyy-MM-dd", $null)

    # Skip this row if it's not from the previous month
    if ($reportDateObject.Year -ne $lastMonth.Year -or $reportDateObject.Month -ne $lastMonth.Month) {
      continue
    }
    $receivedItems = [int]$row.Receive
    $sentItems = [int]$row.Send
    $reportDate = $row.'Report Date'

    # Prepare list item for SharePoint
    $listItem = @{
      fields = @{
        Title         = "Email activity for " + $reportDate
        ReceivedItems = $receivedItems
        SentItems     = $sentItems
        ReportDate    = $reportDate
      }
    } | ConvertTo-Json
  }
  catch {
    write-host "Error occurred while processing the row: $_"
    continue
  }

  try {
    $addToListUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items"
    $response = Invoke-RestMethod -Uri $addToListUrl -Method Post -Body $listItem -ContentType "application/json" -Headers @{Authorization = "Bearer $($token.AccessToken)" }
    write-host "List item for date $reportDate created successfully."
  }
  catch {
    Write-Host "Error occurred while posting to SharePoint: $($_.Exception.Response)"
    Write-Host "Status Code: $($_.Exception.Response.StatusCode.Value__)"
    Write-Host "Status Description: $($_.Exception.Response.StatusDescription)"
     
    # Read the response content
    $responseContent = $_.Exception.Response.Content.ReadAsStringAsync().Result
    Write-Host "Response Body: $responseContent"
  }
}