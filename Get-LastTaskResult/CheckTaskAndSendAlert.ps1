# ------------------------
# 1. Initialization
# ------------------------

$logFilePath = ".\errorTasks.log"
Set-location -Path 'E:\scheduledTasks\Monitor'

#Dot source the send-GraphEmail Function
. .\send-GraphEmail.ps1

# ------------------------
# 2. Function Declarations
# ------------------------

function isTaskLogged($taskName) {
  if (Test-Path $logFilePath) {
    $loggedTasks = Get-Content $logFilePath
    return $taskName -in $loggedTasks
  }
  return $false
}

function logTask($taskName) {
  Add-Content -Path $logFilePath -Value $taskName
}


function Get-LastTaskResult { 
  param (
    [string]$taskName
  )

  try {

    $results = schtasks /query /tn $taskName /v /fo LIST 2>$null | Select-String "Last Result" -SimpleMatch

    if ($results) {
      return $results[0].ToString().Split(":")[1].Trim()
      # return $results Note: The next line return $results is unreachable due to the previous return statement. This seems to be a mistake in the code.
    }
    else {
      return $null
    }
  }
  catch {
    Write-Warning "Error processing ${taskname}: $_"
    Write-Output "schtasks output for ${taskname}:"
    schtasks /query /tn $taskName /v /fo LIST 2>$null

    return $null
  }

}        
 

# ------------------------
# 3. Main Processing
# ------------------------

$taskNames = Get-ScheduledTask | Where-Object { $_.TaskPath -eq "\" -and $_.State -ne "Disabled" -and $_.State -ne "Running" } | 
ForEach-Object { $_.TaskName }

$results = @{}

foreach ($taskName in $taskNames) {
  $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
  $taskInfo = $null

  if ($task) {
    $taskInfo = Get-ScheduledTaskInfo -TaskName $taskName
  }

  $lastResult = Get-LastTaskResult -taskName ("\\" + $taskName)

  Write-Host "Debug for $taskname:Last Result = $lastresult"
  if ($lastResult -ne "0" -and $lastResult -ne "0x0") {

    # Check if task is logged and if not, send the email and log it
    if (-not (isTaskLogged $taskName)) {
      # Construct email content
      $serverName = $env:Computername
      $subject = "Error in Scheduled Task: $taskName on server: $serverName"
      $content = "There was an error in the scheduled task $taskname on server $serverName. The error code is $lastResult."
        
      # Send the email
      send-GraphEmail -accessToken $accessToken -recipientEmail $recipientEmail -subject $subject -Content $content -fromEmail $fromEmail

      # Log the task
      logTask $taskname
    }
    
    # Add the task details to results
    $results[$taskName] = @{
      'TaskName'      = $taskName
      'LastRunTime'   = $taskInfo.LastRunTime
      'LastRunResult' = $lastResult
    }
  }

  #Send email notification
  $serverName = $env:Computername
  $subject = "Error in Scheduled Task: $taskName on server: $serverName"
  $content = "There was an error in the scheduled task $taskname on server env:COMPUTERNAME. The error code is $lastResult."
    
}

Start-Sleep -Milliseconds 100 #Adding the sleep here


# ------------------------
# 4. Reporting
# ------------------------

$results.GetEnumerator() | ForEach-Object {
  New-Object PSObject -Property @{
    TaskName      = $_.Name
    LastRunTime   = if ($_.Value.LastRunTime -eq $null) { 'Never Executed' } else { $_.Value.LastRunTime }
    LastRunResult = switch ($_.Value.LastRunResult) {
      '0x0' { 'Successful' }
      0 { 'Successful' }
      $null { 'No Reported Result' }
      default { "Error Code: $($_.Value.LastRunResult)" }
    }
  }
} | Format-Table -AutoSize