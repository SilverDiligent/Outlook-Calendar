$global:configPath = '/Users/alexiscrawford/Library/CloudStorage/OneDrive-Alexislabs/_Coding/ProductionScripts/Config.Json'
$global:configData = Get-Content -Path $global:configPath | ConvertFrom-Json
$global:password = $global:configData.password | ConvertTo-SecureString -AsPlainText -Force
$global:username = $global:configData.username

function Connect-ExchangeSRP {

  $credsSRPnet = New-Object System.Management.Automation.PSCredential -ArgumentList $global:username, $global:password
  $S = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://SRPEXM221.srp.gov/PowerShell/ -Authentication Kerberos -Credential $credsSRPnet
  Import-PSSession -Session $S
}

function ConnectToExchangeOnline {
  $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $global:userName, $global:password
  Connect-ExchangeOnline -Credential $credential
}

function PerformMessageTrace {
  param(
    [Parameter(Mandatory = $false)]
    [string]$senderAddress,
    [Parameter(Mandatory = $true)]
    [string]$recipientAddress,
    [Parameter(Mandatory = $true)]
    [datetime]$startDate,
    [Parameter(Mandatory = $false)]
    [datetime]$endDate = (Get-Date)  
  )
  
  Get-MessageTrace -SenderAddress $senderAddress -RecipientAddress $recipientAddress -StartDate $startDate -EndDate $endDate
}
