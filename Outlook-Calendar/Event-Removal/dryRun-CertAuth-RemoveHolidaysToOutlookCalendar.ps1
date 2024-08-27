<#
.SYNOPSIS
This script removes events from Outlook calendars based on specified holidays.

.DESCRIPTION
The script retrieves access tokens, loads configuration data, and imports holiday data from a CSV file. It then loops through each holiday, searches for events containing any of the holiday words in the subject, and deletes the events from the calendars of specified users. The script logs information about the deleted events to a log file.

.PARAMETER DryRun
A switch parameter that, when specified, enables dry run mode. In dry run mode, the script only lists the events that would be deleted without actually deleting them.

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
None. The script does not generate any output.

.EXAMPLE
.\dryRun.ps1 -DryRun
Runs the script in dry run mode, listing the events that would be deleted without actually deleting them.

.NOTES
- This script requires the MSAL.PS module to be installed.
- The script requires a valid configuration file (config.json) and a CSV file containing holiday data (resultcalendar-now.csv).
- The script requires the accounts.csv file to contain a list of user emails to delete events from their calendars.
- The script requires the specified certificate (identified by thumbprint) to be present in the local machine store.


#>
param (
   [switch]$DryRun
)

# Load configuration data
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
$holidayData = Import-Csv -Path "e:\ProductionScripts\Calendar-Management\resultcalendar-now.csv"

# List of user emails to delete events from their calendars
$userEmails = Import-csv -Path "E:\ProductionScripts\Calendar-Management\accounts.csv"

# Log file path
$logFilePath = "E:\ProductionScripts\Calendar-Management\Event-Removal\Remove-Log.txt"

# Loop through each holiday
foreach ($holiday in $holidayData) {
   # Prepare the holiday details from CSV data
   $holidaySubject = $holiday.Subject
   $dateString = $holiday.'Start Date'

   # Parse the date from the CSV file
   try {
       Write-Host "Parsing date for holiday '$holidaySubject': $dateString"
       $holidayDate = [datetime]::ParseExact($dateString, "M/d/yyyy", $null).ToString('yyyy-MM-dd')
   } catch {
       Write-Error "Failed to parse date for holiday '$holidaySubject': $_"
       continue
   }

   # Split the holiday subject into individual words
   $holidayWords = $holidaySubject -split ' '

   # Loop through each user and delete the event from their calendar
   foreach ($user in $userEmails) {
       try {
           $email = $user.Email  # Access the Email property of the user object

           # Debugging output
           Write-Host "Searching for events containing any of the words in '$holidaySubject' for user: $email on date: $holidayDate"

           # Construct the search query to check for each word in the holiday subject and the date
           $searchQuery = ($holidayWords | ForEach-Object { "contains(subject,'$_')" }) -join ' or '
           $searchQuery += " and start/dateTime ge '$holidayDate' and start/dateTime lt '"
           try {
               Write-Host "Parsing date: $dateString"
               $trimmedDateString = $dateString.Trim()
               Write-Host "Trimmed date string: $trimmedDateString"
               $parsedDate = [datetime]::ParseExact($trimmedDateString, 'M/d/yyyy', $null)
               $searchQuery += $parsedDate.AddDays(1).ToString('yyyy-MM-dd') + "'"
           } catch {
               Write-Error "Failed to parse date: $dateString - Exception: $_"
               throw
           }

           $searchUri = "https://graph.microsoft.com/v1.0/users/$($email)/calendar/events?`$filter=$searchQuery"
           Write-Host "Search URI: $searchUri"

           # Perform the search
           $searchResponse = Invoke-RestMethod -Uri $searchUri -Headers $headers -Method Get

           # Check if existing event found
           $existingEvents = $searchResponse.value | Where-Object { $event = $_; $holidayWords | ForEach-Object { $event.subject -match $_ } }

           foreach ($existingEvent in $existingEvents) {
               if ($DryRun) {
                   # Dry run mode: just list the event
                   Write-Host "Dry Run: Found event '$($existingEvent.subject)' for $($email) that would be deleted."
               } else {
                   # Event exists, delete the event
                   $deleteUri = "https://graph.microsoft.com/v1.0/users/$($email)/calendar/events/$($existingEvent.id)"
                   Invoke-RestMethod -Uri $deleteUri -Headers $headers -Method Delete

                   # Log information to the file
                   $logMessage = "User: $($email), Holiday: $($existingEvent.subject), Time Deleted: $(Get-Date)"
                   Add-Content -Path $logFilePath -Value $logMessage
                   Write-Host "Deleted event '$($existingEvent.subject)' for $($email)"
               }
           }
       } catch {
           Write-Error "Error processing calendar event for $($email): $_"
       }
   }
}