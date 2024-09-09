<#
.SYNOPSIS
Script to add holiday events to Outlook calendars.

.DESCRIPTION
This script retrieves holiday data from a CSV file and adds corresponding events to the Outlook calendars of specified users. It uses the Microsoft Graph API to interact with the calendars.

.PARAMETER DryRun
Switch parameter to perform a dry run without actually creating events.

.INPUTS
None. You must provide a valid CSV file containing holiday data.

.OUTPUTS
The script generates a report of the actions performed, including the created events and skipped events. The report is saved as a CSV file.

.EXAMPLE
.\DryRun-ADDOutlookHolidayEventsWithReport-Cert.ps1 -DryRun
Runs the script in dry run mode, where events are not actually created.

.NOTES
- This script requires the MSAL.PS module to retrieve an access token.
- The script assumes that the necessary configuration data is available in a JSON file.
- The script requires a valid certificate with the specified thumbprint to authenticate with the Microsoft Graph API.
- The script uses the primary SMTP addresses of user mailboxes to add events to their calendars.
- The script assumes that the CSV file containing holiday data has the required columns: Subject and Start Date.
- The script uses the US Mountain Standard Time zone for event start and end times.
- The generated report includes the email address, action (created or skipped), holiday subject, and date for each event.
- The report is saved in the specified directory with a timestamped file name.

.LINK
Microsoft Graph API documentation: https://docs.microsoft.com/graph/api/overview?view=graph-rest-1.0

#>
param (
   [switch]$DryRun
)

# Initialize an empty list to store the report objects
$Report = [System.Collections.Generic.List[Object]]::new()

# Load configuration data
$configData = Get-Content -Path "../config.json" | ConvertFrom-Json

# Log file path
#$logFilePath = "E:\ProductionScripts\Outlook-Calendar\Add-Events-To-Calendar\Add-Log.txt"

# Function to log messages

# Get access token
$tenantId = $configData.TENANT_ID
$clientId = $configData.CLIENT_ID
$thumbprint = $configData.Thumbprint
$tenantName = $configData.TENANT_NAME
$authority = "https://login.microsoftonline.com/$tenantId"
$resourceUrl = "https://graph.microsoft.com"

# Load the certificate from the local machine store
try {
   $certStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "LocalMachine")
   $certStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
   $pfxCert = $certStore.Certificates | Where-Object { $_.Thumbprint -eq $thumbprint }
   $certStore.Close()

   if ($pfxCert -eq $null) {
       $message = "Certificate with thumbprint $thumbprint not found in the local machine store."
       Write-Host $message
       
       return
   } else {
       Write-Host "Certificate loaded successfully."
   }
} catch {
   $message = "Failed to load certificate: $_"
   Write-Host $message
   
   return
}

# Request access token using MSAL.PS module
try {
   $tokenRequest = Get-MsalToken -ClientId $clientId -TenantId $tenantId -ClientCertificate $pfxCert -Authority $authority -Scopes "$resourceUrl/.default"
   $accessToken = $tokenRequest.AccessToken
   Write-Host "Successfully retrieved access token."
} catch {
   $message = "Failed to retrieve access token: $_"
   Write-Host $message
   
   return
}

if (-not $accessToken) {
   $message = "Access token is empty. Please check your credentials and network connection."
   Write-Host $message
   
   return
}

$headers = @{
   Authorization = "Bearer $accessToken"
   ContentType   = "application/json"
}

# Retrieve the list of user mailboxes dynamically
# $usermailboxes = @("alexis.crawford@srpnet.com", "Lina.Smith@srpnet.com")

 $userMailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Select-Object -ExpandProperty PrimarySmtpAddress

# Import holiday data from CSV
$holidayData = Import-Csv -Path "E:\ProductionScripts\Calendar-Management\resultCalendar-NOW.csv"

# Loop through each holiday
foreach ($holiday in $holidayData) {
   $holidaySubject = $holiday.Subject
   $dateString = $holiday.'Start Date'

   # Parse the date from the CSV file
   try {
       Write-Host "Parsing date for holiday '$holidaySubject': $dateString"
       $holidayDate = [datetime]::ParseExact($dateString.Trim(), "M/d/yyyy", $null).ToString('yyyy-MM-dd')
   } catch {
       Write-Error "Failed to parse date for holiday '$holidaySubject': $_"
       continue
   }

   # Split the holiday subject into individual words
   $holidayWords = $holidaySubject -split ' '

   # Loop through each mailbox retrieved dynamically
   foreach ($email in $userMailboxes) {
       try {
           Write-Host "Searching for events containing any of the words in '$holidaySubject' for user: $email on date: $holidayDate"

           # Construct the search query to check for each word in the holiday subject and the date
           $searchQuery = ($holidayWords | ForEach-Object { "contains(subject,'$_')" }) -join ' and '
           $searchQuery += " and start/dateTime ge '$holidayDate' and start/dateTime lt '"
           try {
               $trimmedDateString = $dateString.Trim()
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

           if ($existingEvents.Count -eq 0) {
               if ($DryRun) {
                   Write-Host "Dry Run: Would create new event: $($holidaySubject) for $($email) on $holidayDate"
                   
               } else {
                   # Event doesn't exist, create new event
                   $holidayDetails = @{
                       subject  = $holidaySubject
                       start    = @{
                           dateTime = $holidayDate
                           timeZone = "US Mountain Standard Time"
                       }
                       end      = @{
                           dateTime = $parsedDate.AddDays(1).ToString('yyyy-MM-dd')
                           timeZone = "US Mountain Standard Time"
                       }
                       isAllDay = $true
                   }

                   $jsonBody = $holidayDetails | ConvertTo-Json

                   $createUri = "https://graph.microsoft.com/v1.0/users/$email/calendar/events"
                   $createResponse = Invoke-RestMethod -Uri $createUri -Headers $headers -Method Post -ContentType "application/json" -Body $jsonBody

                   Write-Host "Created new event: $($holidaySubject) for $($email)"
                   
               }
           } else {
               $message = "Event with subject $($holidaySubject) already exists for $($email), skipping creation."
               Write-Host $message
               
           }

           # Add the mailbox and action to the report
           $reportItem = [PSCustomObject]@{
               EmailAddress = $email
               Action       = if ($existingEvents.Count -eq 0) { "Created" } else { "Skipped" }
               Holiday      = $holidaySubject
               Date         = $holidayDate
           }
           $Report.Add($reportItem)
       }
       catch {
           $message = "Error processing calendar event for ${email}: $_"
           Write-Error $message
           
       }
   }
}



# Output the report to the console
$Report | Format-Table -AutoSize

# Get the current date and time to create a unique file name
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportFileName = "EventAdd-Report_$timestamp.csv"
$reportFilePath = "E:\ProductionScripts\Calendar-Management\Report\$reportFileName"

# Export the report to a CSV file with the timestamped name
$Report | Export-Csv -Path $reportFilePath -NoTypeInformation

# Output the path of the generated report
Write-Host "Report generated and saved to: $reportFilePath"


