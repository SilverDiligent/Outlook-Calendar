<#
.SYNOPSIS
This script removes the "assistant" attribute from a user in Active Directory and updates the assistant's name in Exchange.

.DESCRIPTION
The script prompts the user to enter the email address of the user and the assistant's name. It then searches for the user in Active Directory using the email address and retrieves the user's distinguished name (DN). The script uses ADSI to bind to the user object and deletes the "assistant" attribute in Active Directory. It then sets the assistant's name in Exchange using the Set-Mailbox cmdlet.

.PARAMETER userEmail
The email address of the user.

.PARAMETER assistantName
The name of the assistant.

.EXAMPLE
RemoveAndUpdateAssistantAttribute.ps1
Prompts the user to enter the email address of the user and the assistant's name. Removes the "assistant" attribute from the user in Active Directory and updates the assistant's name in Exchange.

.NOTES
- This script requires the ActiveDirectory and Exchange modules to be imported.
- The script assumes that the user exists in Active Directory and the assistant's name can be set in Exchange.
- Modify the Set-Mailbox cmdlet based on your Exchange environment.
#>

# Import the ActiveDirectory and Exchange modules (if not already imported)
Import-Module ActiveDirectory
Import-Module Exchange

# Prompt the user for the email address of the user
$userEmail = Read-Host "Enter the email address of the user"

# Prompt the user for the assistant's display name
$assistantName = Read-Host "Enter the assistant's display name"

# Search for the user in Active Directory and retrieve the DN
$user = Get-ADUser -Filter { EmailAddress -eq $userEmail }

if ($user) {
  $userDN = $user.DistinguishedName

  # Use ADSI to bind to the user object
  $adsiUser = [ADSI]"LDAP://$userDN"

  # Delete the "assistant" attribute in Active Directory
  $adsiUser.PSBase.Properties["assistant"].Clear()
  $adsiUser.CommitChanges()

  Write-Host "The 'assistant' attribute has been deleted for user: $userEmail"

  # Set the assistant's name in Exchange (you may need to modify this based on your Exchange environment)
  Set-Mailbox -Identity $userEmail -AssistantName $assistantName

  Write-Host "The 'assistant' name has been updated to '$assistantName' for user: $userEmail"
}
else {
  Write-Host "User with email address '$userEmail' not found in Active Directory."
}
