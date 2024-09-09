<#
.SYNOPSIS
   Delete-Folders-Win.ps1 - Script to delete specified folders in multiple mailboxes using Microsoft Graph API.

   Author: Alexis Crawford
   Date: 2024-07-24

.DESCRIPTION
   This script deletes specified folders in multiple mailboxes using Microsoft Graph API. It retrieves access token using MSAL.PS module and performs the deletion operation for each mailbox and folder combination provided in the CSV file.

.PARAMETER None

.EXAMPLE
   .\Delete-Folders-Win.ps1
   - This example runs the script without any parameters. It will read the configuration data from the "config.json" file and the mailbox and folder information from the "accounts.csv" file.

.NOTES
   - This script requires the MSAL.PS module to be installed. You can install it using the following command:
       Install-Module -Name MSAL.PS -Scope CurrentUser

   - The script expects the following files to be present in the specified paths:
       - config.json: Contains the configuration data including tenant ID, client ID, and thumbprint.
       - accounts.csv: Contains the mailbox and folder information in CSV format.

   - Make sure to provide the correct paths to the configuration and CSV files in the script.

   - The CSV file should have at least 2 columns: Email and FolderPath. Here is an example of the format:
       Email,FolderPath
       user1@domain.com,Inbox/Subfolder1

.LINK
   Microsoft Graph API: https://docs.microsoft.com/en-us/graph/overview

#>


$configData = Get-Content -Path "../config.json" | ConvertFrom-Json

# Get access token
$tenantId = $configData.TENANT_ID
$clientId = $configData.CLIENT_ID
$thumbprint = $configData.Thumbprint
$authority = "https://login.microsoftonline.com/$tenantId"
$resourceUrl = "https://graph.microsoft.com"

# Load the certificate from the local machine store
try {
  $certStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "LocalMachine")
  $certStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
  $pfxCert = $certStore.Certificates | Where-Object { $_.Thumbprint -eq $thumbprint }
  $certStore.Close()

  if ($pfxCert -eq $null) {
      Write-Host "Certificate with thumbprint $thumbprint not found in the local machine store."
      return
  } else {
      Write-Host "Certificate loaded successfully."
  }
} catch {
  Write-Host "Failed to load certificate: $_"
  return
}

# Request access token using MSAL.PS module
try {
  $tokenRequest = Get-MsalToken -ClientId $clientId -TenantId $tenantId -ClientCertificate $pfxCert -Authority $authority -Scopes "$resourceUrl/.default"
  $accessToken = $tokenRequest.AccessToken
  Write-Host "Successfully retrieved access token."
} catch {
  Write-Host "Failed to retrieve access token: $_"
  return
}

if (-not $accessToken) {
  Write-Host "Access token is empty. Please check your credentials and network connection."
  return
}

$headers = @{
  Authorization = "Bearer $accessToken"
  ContentType   = "application/json"
}

# Import holiday data from CSV
$holidayData = Import-Csv -Path "E:\ProductionScripts\Outlook-Calendar\resultCalendar-NOW.csv"

# List of user emails to delete events from their calendars
$userEmails = Import-csv -Path "E:\ProductionScripts\Outlook-Calendar\accounts.csv"

# Log file path
$logFilePath = "E:\ProductionScripts\Outlook-Calendar\Remove-Events-From-Calendar\Remove-Log.txt"

# Loop through each holiday
foreach ($holiday in $holidayData) {
  # Prepare the holiday details from CSV data
  $holidaySubject = $holiday.Subject

  # Loop through each user and delete the event from their calendar
  foreach ($user in $userEmails) {
      try {
          $email = $user.Email  # Access the Email property of the user object

          # Prepare the encoded holiday subject
          $encodedSubject = [System.Uri]::EscapeDataString($holidaySubject)

          # Debugging output
          Write-Host "Searching for event with subject: $holidaySubject"
          Write-Host "Encoded Subject: $encodedSubject"

          # Search for existing events with the encoded subject
          $searchUri = "https://graph.microsoft.com/v1.0/users/$($email)/calendar/events?`$filter=startswith(subject,'$encodedSubject')"
          Write-Host "Search URI: $searchUri"

          # Perform the search
          $searchResponse = Invoke-RestMethod -Uri $searchUri -Headers $headers -Method Get

          # Check if existing event found
          $existingEvent = $searchResponse.value | Where-Object { $_.subject -eq $holidaySubject }

          if ($existingEvent) {
              # Event exists, delete the event
              $deleteUri = "https://graph.microsoft.com/v1.0/users/$($email)/calendar/events/$($existingEvent.id)"
              Invoke-RestMethod -Uri $deleteUri -Headers $headers -Method Delete

              # Log information to the file
              $logMessage = "User: $($email), Holiday: $($holidaySubject), Time Deleted: $(Get-Date)"
              Add-Content -Path $logFilePath -Value $logMessage
              Write-Host "Deleted event '$holidaySubject' for $($email)"
          } else {
              Write-Host "Event with subject '$holidaySubject' not found for $($email), skipping deletion."
          }
      }
      catch {
          Write-Error "Error processing calendar event for $($email): $_"
      }
  }
}