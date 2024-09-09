param (
    [switch]$DryRun
)

# Load configuration data
$configData = Get-Content -Path "../config.json" | ConvertFrom-Json

# Log file path
$logFilePath = "E:\ProductionScripts\Outlook-Calendar\Add-Events-To-Calendar\Add-Log.txt"

# Function to log messages
function Log-Message {
    param (
        [string]$message
    )
    Add-Content -Path $logFilePath -Value $message
}

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
        $message = "Certificate with thumbprint $thumbprint not found in the local machine store."
        Write-Host $message
        Log-Message $message
        return
    } else {
        Write-Host "Certificate loaded successfully."
    }
} catch {
    $message = "Failed to load certificate: $_"
    Write-Host $message
    Log-Message $message
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
    Log-Message $message
    return
}

if (-not $accessToken) {
    $message = "Access token is empty. Please check your credentials and network connection."
    Write-Host $message
    Log-Message $message
    return
}

$headers = @{
    Authorization = "Bearer $accessToken"
    ContentType   = "application/json"
}

# Import holiday data from CSV
$holidayData = Import-Csv -Path "E:\ProductionScripts\Outlook-Calendar\resultCalendar-NOW.csv"

# List of user emails to add the holidays to their calendars
$userEmails = Import-csv -Path "E:\ProductionScripts\Outlook-Calendar\accounts.csv" | foreach-object { $_.Email }

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

    # Loop through each user and search for the event in their calendar
    foreach ($email in $userEmails) {
        try {
            # Debugging output
            Write-Host "Searching for events containing any of the words in '$holidaySubject' for user: $email on date: $holidayDate"

            # Construct the search query to check for each word in the holiday subject and the date
            $searchQuery = ($holidayWords | ForEach-Object { "contains(subject,'$_')" }) -join ' or '
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
                    Log-Message "DryRun - User: $email, Holiday: $($holidaySubject), Date: $holidayDate"
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
                    Log-Message "User: $email, Holiday: $($holidaySubject), Time Created: $(Get-Date)"
                }
            } else {
                $message = "Event with subject $($holidaySubject) already exists for $($email), skipping creation."
                Write-Host $message
                Log-Message $message
            }
        }
        catch {
            $message = "Error processing calendar event for ${email}: $_"
            Write-Error $message
            Log-Message $message
        }
    }
}