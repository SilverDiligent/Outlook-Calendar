<#
.SYNOPSIS
 This script creates folders in mailboxes using Microsoft Graph API.

 Author: Alexis Crawford
 Created: July 31 2024

.DESCRIPTION
 The script imports the required module MSAL.PS and retrieves an access token using the client ID, tenant ID, and client certificate specified in the config.json file.
 It then imports a CSV file containing mailbox information and folder paths to be created.
 The script uses the Get-FolderId function to get the ID of the parent folder based on the provided folder path.
 If the parent folder ID is not found, the script creates the folder under 'msgfolderroot'.
 The Create-Folder function is used to create a new folder if it does not already exist.
 The script loops through each mailbox in the CSV file and creates the specified folder.

.PARAMETER config.json
 The config.json file contains the following information:
 - tenant_Id: The ID of the Azure AD tenant.
 - client_Id: The ID of the registered application in Azure AD.
 - thumbprint: The thumbprint of the client certificate used for authentication.
 - pfxFilePath: The file path of the client certificate in PFX format.

.PARAMETER accountsToCreate.csv
 The accountsToCreate.csv file contains the following information for each mailbox:
 - Email: The email address of the mailbox.
 - FolderPath: The path of the parent folder where the new folder will be created.
 - NewFolderName: The name of the new folder to be created.
   Email,FolderPath,NewFolderName
   alexis.crawford@srpnet.com,root,FolderToCreate
   alexis.crawford@srpnet.com,dev-FoldertoCreateNow,TEST
   alexis.crawford@srpnet.com,dev-FolderToCreateNow/TEST,2024

.EXAMPLE
 .\Create-Folders-Win.ps1

.NOTES
 - This script requires the MSAL.PS module to be installed.
 - The script requires appropriate permissions to access and modify mail folders in Microsoft 365.


#>


# Import Required Module
Import-Module MSAL.PS
$configData = Get-Content -Path "../config.json" | ConvertFrom-Json

# Get access token
$tenantId = $configData.tenant_Id
$clientId = $configData.client_Id
$thumbprint = $configData.thumbprint
$pfxFilePath = $configData.pfxFilePath
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

# Import mailboxes from CSV file
$mailboxes = Import-Csv -Path "../accountsToCreate.csv"

# Helper function to get folder ID by path
function Get-FolderId {
   param (
       [string]$mailbox,
       [string]$folderPath
   )
   $parentFolderId = "msgfolderroot"
   $folders = $folderPath -split "/"
   foreach ($folder in $folders) {
       $encodedMailbox = [System.Web.HttpUtility]::UrlEncode($mailbox)
       $encodedParentFolderId = [System.Web.HttpUtility]::UrlEncode($parentFolderId)
       $uri = "$resourceUrl/v1.0/users/$encodedMailbox/mailFolders/$encodedParentFolderId/childFolders"
       
       $allFolders = @()
       $skip = 0
       $moreResults = $true
       
       while ($moreResults) {
           $pagedUri = "$uri`?`$skip=$skip"
           Write-Host "Constructed URI: $pagedUri"  # Log the constructed URI
           try {
               $response = Invoke-RestMethod -Uri $pagedUri -Method Get -Headers $headers
               $allFolders += $response.value
               if ($response.'@odata.nextLink') {
                   $skip += 10
               } else {
                   $moreResults = $false
               }
           } catch {
               Write-Host "Error retrieving child folders: $_"
               return $null
           }
       }
       if ($allFolders) {
           $parentFolder = $allFolders | Where-Object { $_.displayName -eq $folder }
           if ($parentFolder) {
               Write-Host "Found folder '$folder' with ID '$($parentFolder.id)' in mailbox '$mailbox'"
               $parentFolderId = $parentFolder.id
           } else {
               Write-Host "Could not find a folder named '$folder' in the mailbox of $mailbox"
               return $null
           }
       } else {
           Write-Host "No child folders found under parent folder ID '$parentFolderId' in mailbox '$mailbox'"
           return $null
       }
   }
   return $parentFolderId
}

# Function to create a new folder if it does not exist
function Create-Folder {
   param (
       [string]$mailbox,
       [string]$parentFolderId,
       [string]$newFolderName
   )
   try {
       $encodedMailbox = [System.Web.HttpUtility]::UrlEncode($mailbox)
       $encodedParentFolderId = [System.Web.HttpUtility]::UrlEncode($parentFolderId)
       $uri = "$resourceUrl/v1.0/users/$encodedMailbox/mailFolders/$encodedParentFolderId/childFolders"
       
       # Check if the folder already exists
       $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
       $existingFolder = $response.value | Where-Object { $_.displayName -eq $newFolderName }
       
       if ($existingFolder) {
           Write-Host "Folder '$newFolderName' already exists in $mailbox's mailbox"
           return
       }
       
       $body = @{
           displayName = $newFolderName
       } | ConvertTo-Json -Depth 10
       Write-Host "Constructed URI for folder creation: $uri"
       Write-Host "Request Body: $body"
       Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body -ContentType "application/json"
       Write-Host "Successfully created folder '$newFolderName' in $mailbox's mailbox"
   } catch {
       Write-Host ("Failed to create folder in " + $mailbox + ": " + $_)
   }
}

# Loop through each mailbox and create the specified folder
foreach ($mailbox in $mailboxes) {
   $folderPath = $mailbox.FolderPath
   $newFolderName = $mailbox.NewFolderName
   # Ensure the new folder name is not empty
   if (-not $newFolderName) {
       Write-Host "New folder name is empty for mailbox '$($mailbox.Email)'. Skipping..."
       continue
   }
   # Get the parent folder ID
   $parentFolderId = Get-FolderId -mailbox $mailbox.Email -folderPath $folderPath
   if (-not $parentFolderId) {
       Write-Host "Parent folder ID could not be found, creating folder '$newFolderName' under 'msgfolderroot'."
       $parentFolderId = "msgfolderroot"
   }
   # Make the API request to create the folder if it does not exist
   Create-Folder -mailbox $mailbox.Email -parentFolderId $parentFolderId -newFolderName $newFolderName
}

