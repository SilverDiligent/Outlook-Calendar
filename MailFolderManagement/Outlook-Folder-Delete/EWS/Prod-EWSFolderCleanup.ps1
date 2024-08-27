#Provide the Folders to be deleted
$csvFile = "././accounts.csv"
$configData = Get-Content -Path '././config.json' | ConvertFrom-Json

$AppID = $configData.clientId
$TenantId = $configData.tenantId
$ClientSecretString = $configData.clientSecretString

# Import the EWS API dll
Import-Module "C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.Exchange.WebServices.2.2\lib\40\Microsoft.Exchange.WebServices.dll"

$Scopes = "https://outlook.office365.com/.default"
$AuthResult = Get-MsalToken -TenantId $TenantId -ClientId $AppID -ClientSecret ($ClientSecretString | ConvertTo-SecureString -AsPlainText -Force) -Scopes $Scopes

#Create the Exchange Service object
$Service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
$Service.Url = "https://outlook.office365.com/EWS/Exchange.asmx"

#Set the recommended Instrumentation Headers
$Service.ClientRequestId = [guid]::NewGuid().ToString()
$Service.ReturnClientRequestId = $True
$Service.UserAgent = "OAuth_SampleScriptAppOnly"

#Using OAuth authentication
$Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$AuthResult.AccessToken

Import-CSV $csvFile | ForEach-Object {
  try {
    $account = $_.account
    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $account)
    $Service.HttpHeaders.Add("X-AnchorMailbox", $account)

    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $account)
    $targetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderId)

    $fldArray = $_.folder.Split("\")
    for ($lint = 0; $lint -lt $fldArray.Length; $lint++) {
      $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10)
      $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $fldArray[$lint])

      $findFolderResults = $service.FindFolders($targetFolder.Id, $searchFilter, $folderView)
      if ($findFolderResults.TotalCount -gt 0) {
        $targetFolder = $findFolderResults.Folders[0]
      }
      else {
        $targetFolder = $null  
        break  
      }    
    }  

    if ($targetFolder -ne $null) {
      $targetFolder.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
      Write-Host "Folder $.Folder has been deleted for $account" -ForeGroundColor Green
    }  
  }
  catch {
    Write-Host "Error encountered: $_" -ForeGroundColor Red
  }
}
