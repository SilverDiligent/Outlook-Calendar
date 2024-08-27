# Read the configuration files.
$configData = Get-Content -Path '.\config.json' | ConvertFrom-Json
$emailConfigData = Get-Content -Path '.\emailConfig.json' | ConvertFrom-Json


# Set the configuration parameters
$clientId = $configData.appId
$clientSecret = $configData.clientSecret
$clientSecretString = $configData.clientSecretString
$tenantId = $configData.tenantId
$scope = $configData.scope
$fromEmail = $emailConfigData.fromEmail
$recipientEmail = "alexis.crawford@srpnet.com"
$clientSecret = convertTo-SecureString -String $clientSecretString -asPlainText -force

# Define the send-GraphEmail function.
function send-GraphEmail {
    param (
        [Parameter(Mandatory = $true)]
        [String]$accessToken,

        [Parameter(Mandatory = $true)]
        [String]$recipientEmail,

        [Parameter(Mandatory = $true)]
        [String]$subject,

        [Parameter(Mandatory = $true)]
        [String]$content,

        [Parameter(Mandatory = $true)]
        [String]$fromEmail = $Global:emailConfig.fromEmail
    
    )

    
    $uri = "https://graph.microsoft.com/v1.0/users/$fromEmail/sendMail"

    $emailContent = @{
        message         = @{
            subject      = $subject
            from         = @{
                emailAddress = @{
                    address = $fromEmail
                }
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $recipientEmail
                    }
                }
            )
            
        }
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 4


    $headers = @{
        Authorization  = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }
    

    # Add this line before sending the email
    Write-Host "About to send email to $recipientEmail from $fromEmail."
    Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -Body $emailContent -ContentType "application/json"
}   

# Get the access token
$tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $clientId
    scope         = $scope
    client_secret = $clientSecretString
    grant_type    = 'client_credentials'
}
$response = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
$accessToken = $response.access_token

# Check if the $accessToken is populated
if (-not $accessToken) {
    Write-Error "Failed to obtain the access token. Terminating script."
    return
}


#send-GraphEmail -accessToken $accessToken -recipientEmail "alexis.crawford@srpnet.com" -subject "Test" -content "this is a test email" -fromEmail "noreply@srpnet.com"