Clear-Host

# 1 - Define application variables

$AppID = "AppId"
$TenantID = "TenantID"
$ClientSecret = "Secret"

# 2 - Connect to Graph API as the application

$ReqToken = @{
    Grant_Type = "Client_Credentials"
    client_id = $AppId
    client_secret = $ClientSecret
    Scope = "https://graph.microsoft.com/.default"
}

$Uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
$tokenResponse = Invoke-RestMethod -Uri $Uri -Method Post -Body $ReqToken -ContentType "application/x-www-form-urlencoded"

# 3 - Get Applications values, from which to retrieve secret expiration date

Function Get-MSGraphRequest {
    param (
        [system.string]$Uri,
        [system.string]$AccessToken
    )
    begin {
        [System.Array]$allPages = @()
        $ReqTokenBody = @{
            Headers = @{
                "Content-Type"  = "application/json"
                "Authorization" = "Bearer $($AccessToken)"
            }
            Method  = "Get"
            Uri     = $Uri
        }
    }
    process {
        write-verbose "GET request at endpoint: $Uri"
        $data = Invoke-RestMethod @ReqTokenBody
        while ($data.'@odata.nextLink') {
            $allPages += $data.value
            $ReqTokenBody.Uri = $data.'@odata.nextLink'
            $Data = Invoke-RestMethod @ReqTokenBody
            # to avoid throttling, the loop will sleep for 3 seconds
            Start-Sleep -Seconds 3
        }
        $allPages += $data.value
    }
    end {
        Write-Verbose "Returning all results"
        $allPages
    }
}

$Applications = Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/applications/"

# 4 - Finding expired secret

$TimeZone = (Get-TimeZone).Id
$Results = @()

$Applications | Sort-Object displayName | Where-Object {$_.passwordCredentials -match "key"} | Foreach-Object {
    #If there are more than one password credentials, we need to get the expiration of each one
    if ($_.passwordCredentials.endDateTime.count -gt 1) {
        $endDates = $_.passwordCredentials.endDateTime
        [int[]]$daysUntilExpiration = @()
        foreach ($Date in $endDates) {
            $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($Date, $TimeZone)
            $daysUntilExpiration += (New-TimeSpan -Start ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, $TimeZone)) -End $Date).Days
        }
    }
    Elseif ($_.passwordCredentials.endDateTime.count -eq 1) {
        $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($_.passwordCredentials.endDateTime, $TimeZone)
        $daysUntilExpiration = (New-TimeSpan -Start ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, $TimeZone)) -End $Date).Days 
    }
    $hash = [ordered]@{
        ID = $_.id
        DisplayName = $_.DisplayName
        DaysUntilExpiration = $daysUntilExpiration
    }
    $Item = New-Object PSOBject -Property $hash
    $Results = $Results + $Item
}

$ResultTabs = $Results | Select-Object ID, DisplayName, @{L = "DaysUntilExpiration"; E = { ($_.DaysUntilExpiration) -join "|"}}
$FinalResults = @()

Foreach($ResultTab in $ResultTabs){
    if($ResultTab.DaysUntilExpiration -match "-"){
        $Status = "The application contains at least one expired secret"
    }
    if($ResultTab.DaysUntilExpiration -notmatch "-"){
        $Status = "The application secret is still valid"
    }
    $hash = [ordered]@{
        ID = $ResultTab.id
        DisplayName = $ResultTab.DisplayName
        DaysUntilExpiration = $ResultTab.daysUntilExpiration
        Status = $Status
    }
    $Item = New-Object PSOBject -Property $hash
    $FinalResults = $FinalResults + $Item
}

#Count how much secret are expired or will expire within 30 days

$Array = ($FinalResults.DaysUntilExpiration).split("|")
$j = 0

foreach ($s in $Array){
    [int]$s = $s
    if ($s -lt 30){
        $j ++
    }
}

$PathFile = ".\AppRegistrationSecretExpiration.csv"
$FinalResults | Export-csv -path $PathFile -NoTypeInformation -Encoding UTF8

#OPTIONAL 5 - SendMail to a specified recipient (the sender must have a valid Exchange license!)

$SenderMail = "sender@contoso.com"
$Recipient = "recipient@contoso.com"
$URI = "https://graph.microsoft.com/v1.0/users/$SenderMail/sendMail"
$base64string = [Convert]::ToBase64String([IO.File]::ReadAllBytes($PathFile))
$FileName = (Get-Item -Path $PathFile).name

Function Send-MSGraphEmail {
    param (
        [system.string]$Uri,
        [system.string]$AccessToken,
        [system.string]$To,
        [system.string]$Subject = "App Secret Expiration Notification",
        [system.string]$Body
    )
    begin {
        $headers = @{
            "Authorization" = "Bearer $($AccessToken)"
            "Content-type"  = "application/json"
        }

        $BodyJsonsend = @"
{
   "message": {
   "subject": "$Subject",
   "body": {
      "contentType": "HTML",
      "content": "Hi workplace team, <br>
      <br>
      Please find as attachment the report about the days remaining for an application secret to expire.
      Negative values indicate that the secret is already expired and you should replace it as soon as possible. <br>
      <br>
      <b>
      Note that $($j) secret are already expired or will expired within 30 days! 
      <b>
      "
   },
   "toRecipients": [
      {
      "emailAddress": {
      "address": "$to"
          }
      }
   ]
   ,"attachments": [
      {
      "@odata.type": "#microsoft.graph.fileAttachment",
      "name": "$FileName",
      "contentType": "text/plain",
      "contentBytes": "$base64string"
      }
   ]
   },
   "saveToSentItems": "true"
}
"@
    }
    process {
        $data = Invoke-RestMethod -Method POST -Uri $Uri -Headers $headers -Body $BodyJsonsend
    }
    end {
        $data
    }
}

Send-MSGraphEmail -Uri $URI -AccessToken $tokenResponse.access_token -To $Recipient -Body $Body

Remove-Item -path $PathFile -Force