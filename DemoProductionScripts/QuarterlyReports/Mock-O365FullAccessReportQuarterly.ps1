# Define admin credentials
$adminUsername = 'acrawford@alexislab.com'
$adminPassword = ConvertTo-SecureString "#i^1_SvMNg&6^GD/UK69" -AsPlainText -Force # Replace with the admin's password
$credential = New-Object System.Management.Automation.PSCredential($adminUsername, $adminPassword)

# Function to connect to Exchange Online
function Connect-ExchangeOnlineWithCredential {
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
}

# Function to get mailbox permissions report
function Get-MailboxPermissionsReport {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $MailboxIdentity
    )

    # Get permissions for a single mailbox
    Get-MailboxPermission -Identity $MailboxIdentity | Where-Object {
        ($_.AccessRights -contains "FullAccess") -and 
        (-not $_.IsInherited) -and 
        ($_.User -notlike "NT AUTHORITY\SELF") -and
        ($_.User -ne $MailboxIdentity)  # Exclude explicit self full access
    } | Where-Object {
        $recipientDetails = Get-recipient -Identity $_.User.toString() -ErrorAction SilentlyContinue
        -not ($recipientDetails -ne $null -and $recipientDetails.RecipientType -eq "MailUniversalSecurityGroup")
    } | ForEach-Object {
        [PSCustomObject]@{
            Mailbox      = $MailboxIdentity
            GrantedTo    = $_.User
            AccessRights = $_.AccessRights
        }
    }
   
}

# Connect to Exchange Online
Connect-ExchangeOnlineWithCredential -Credential $credential

# Initialize an array to hold all mailbox reports
$allReports = @()

# Fetch mailboxes and generate reports
$mailboxes = Get-Mailbox -ResultSize Unlimited
# $mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Filter { ((CustomAttribute2 -eq 'empl') -or (CustomAttribute2 -eq 'cont')) -and (Name -notlike 'pending.delete.*') }

foreach ($mailbox in $mailboxes) {
    # Get the permissions report for this mailbox
    $report = Get-MailboxPermissionsReport -MailboxIdentity $mailbox.PrimarySmtpAddress

    # Add the report to the array
    $allReports += $report

    # Extract the first name for email
    $FirstName = $mailbox.PrimarySmtpAddress.Split("@")[0].Split(".")[0]

    # Call the Outlook-sendas.ps1 script with parameters
    ./Mock-Outlook-SendAs -mailbox $mailbox.PrimarySmtpAddress -firstName $FirstName
}
   
# Convert report to JSON and save
$allReports | ConvertTo-Json | Set-Content -Path "./Mock-report.json"

# Disconnect from Exchange Online at the end of the script
Disconnect-ExchangeOnline -Confirm:$false