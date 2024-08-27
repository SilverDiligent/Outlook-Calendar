# Read the JSON report and convert it to PowerShell objects
$reportEntries = Get-Content -Path './mock-report.json' -Raw | ConvertFrom-Json

# Read the email log and create a hashtable for notifications
$notificationLog = @{}
Get-Content -Path "./emaillog.txt" | ForEach-Object {
  if ($_ -match "Would send email to: (.+) on (.+)") {
    $mailbox = $matches[1].Trim()
    $date = $matches[2].Trim()
    $notificationLog[$mailbox] = $date
  }
}

# Function to convert each report entry to an HTML row
function ConvertTo-HTMLRow {
  param (
    [Parameter(Mandatory = $true)]
    [PSCustomObject]$entry
  )
   
  # Assuming that the 'NotificationSentTo' property is already added to each $entry object
  $htmlRow = @"
    <tr>
        <td>$($entry.Mailbox)</td>
        <td>$($entry.GrantedTo)</td>
        <td>$($entry.AccessRights -join ', ')</td>
        <td>$($entry.NotificationSentTo)</td>
    </tr>
"@

  return $htmlRow
}

# Initialize the HTML report with headers
$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
<style>
    table {
        width: 100%;
        border-collapse: collapse;
    }
    table, th, td {
        border: 1px solid black;
        padding: 5px;
        text-align: left;
    }
    th {
        background-color: #f2f2f2;
    }
</style>
</head>
<body>
<table>
    <tr>
        <th>Mailbox</th>
        <th>Granted To</th>
        <th>Access Rights</th>
        <th>Notification Sent To</th>
    </tr>
"@

# Generate the rows of the table
foreach ($entry in $reportEntries) {
  # Add the NotificationSentTo property to each entry
  if ($notificationLog.ContainsKey($entry.Mailbox)) {
    $entry | Add-Member -NotePropertyName "NotificationSentTo" -NotePropertyValue $notificationLog[$entry.Mailbox]
  }
  else {
    $entry | Add-Member -NotePropertyName "NotificationSentTo" -NotePropertyValue (Get-Date).ToString()
  }

  # Append each row to the HTML report
  $htmlReport += ConvertTo-HTMLRow -entry $entry
}

# Finish the HTML report
$htmlReport += @"
</table>
</body>
</html>
"@

# Save the HTML report to a file
Set-Content -Path "./HTMLPermissionsReport.html" -Value $htmlReport
 