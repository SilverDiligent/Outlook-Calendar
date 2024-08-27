$configData = Get-Content -Path './mock-Config.json' | ConvertFrom-Json

# Set configuration parameters
$clientId = $configData.clientId
$p12Path = $configData.p12Path
$tenantId = $configData.tenantId
$p12Password = $configData.p12Password
$logFilePath = './emaillog.txt'

if (!(Test-Path -Path $logFilePath)) {
  New-Item -ItemType File -Path $logFilePath -Force
}

# Load the certificate from the P12 file
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($p12Path, $p12Password)

# Acquire an Access Token (Assuming this function is defined in your environment)
$token = Get-MsalToken -ClientId $clientId -TenantId $tenantId -ClientCertificate $cert

# Load the report entries from a JSON file and ensure it's an array
$reportEntries = Get-Content -Path "./Mock-report.json" -Raw | ConvertFrom-Json
$reportEntries = @($reportEntries)


# Read the email log and create a hashtable for notifications
$notificationLog = @{}
Get-Content -Path "./emaillog.txt" | ForEach-Object {
  if ($_ -match "Would send email to: (.+) on (.+)") {
    $email = $matches[1].Trim()
    $date = $matches[2].Trim()
    # Store both the email and the date in an array
    $notificationLog[$email.ToLower()] = @($email, $date)
  }
}

# Debug: Output the content of the notification log to verify correct data
$notificationLog.Keys | ForEach-Object { Write-Host "$_ : $($notificationLog[$_])" }

# Function to convert the report entry to an HTML table row
function ConvertTo-HTMLRow {
  param (
    [Parameter(Mandatory = $true)]
    [PSCustomObject]$entry
  )

  # Extract the email and date from the NotificationSentTo property
  $notifiedEmail = $entry.NotificationSentTo[0]
  $notifiedDate = $entry.NotificationSentTo[1]

  $htmlRow = @"
    <tr>
        <td>$($entry.Mailbox)</td>
        <td>$($entry.GrantedTo)</td>
        <td>$($entry.AccessRights -join ', ')</td>
        <td>$notifiedEmail</td>
        <td>$notifiedDate</td>
    </tr>
"@
  return $htmlRow
}

# Initialize the HTML report with the table header
$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
<style>
    table { width: 100%; border-collapse: collapse; }
    th, td { border: 1px solid black; padding: 5px; text-align: left; }
    th { background-color: #f2f2f2; }
</style>
</head>
<body>
<table>
    <tr>
        <th>Mailbox</th>
        <th>Granted To</th>
        <th>Access Rights</th>
        <th>Email Notified</th>
        <th>Date Notified</th>
    </tr>
"@

# Process each report entry
foreach ($entry in $reportEntries) {
  $userMailbox = $entry.Mailbox.ToLower().Trim()

  # Skip the entry if the mailbox is null or empty
  if ([string]::IsNullOrWhiteSpace($userMailbox)) {
    Write-Warning "No mailbox found in the entry. Skipping."
    continue
  }

  # Determine the notification info based on the presence of the mailbox in the notification log
  $notificationInfo = if ($notificationLog.ContainsKey($userMailbox)) {
    $notificationLog[$userMailbox]
  }
  else {
    @("N/A", "N/A")
  }

  # Add the 'NotificationSentTo' property to the entry
  $entry | Add-Member -NotePropertyName "NotificationSentTo" -NotePropertyValue $notificationInfo -Force

  # Append the HTML row to the report
  $htmlReport += ConvertTo-HTMLRow -entry $entry
}

# Finalize the HTML report
$htmlReport += @"
</table>
</body>
</html>
"@

