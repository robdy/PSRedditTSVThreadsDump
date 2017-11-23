# Variables

$clientID = ""
$secretID = ""
$redirectURI = "https://not-an-aardvark.github.io/reddit-oauth-helper/"
$refresh_token = ""
$TSVFile = "$env:USERPROFILE\Desktop\tsv.csv"
$TSVFileProcessed = "$env:USERPROFILE\Desktop\tsv-processed.csv"
$TSVFileG6 = "$env:USERPROFILE\Desktop\tsv-6.csv"
$TSVFileProcessedG6 = "$env:USERPROFILE\Desktop\tsv-processed-6.csv"

$Credentials = New-Object System.Management.Automation.PSCredential ($clientID, ($secretID|ConvertTo-SecureString -AsPlainText -Force))


<#
.Synopsis
   Refreshes access token for Reddit OAuth
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Refresh-RedditAccessToken
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Refresh token
        [string][Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $refreshToken,

                [System.Management.Automation.PSCredential][Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $Credentials
    )

    Begin
    {
        # Request-specific values,
        $uri_refreshing = "https://www.reddit.com/api/v1/access_token"
        $body_refreshing = "grant_type=refresh_token&refresh_token=$refreshToken"

    }
    Process
    {
        $res = Invoke-webrequest -Uri $uri_refreshing -Method Post -Credential $Credentials -ContentType "application/x-www-form-urlencoded"  -UserAgent "SVHelper/0.1 by robdy" -Body $body_refreshing #-Credential $creds
        $access_token = ($res.Content | ConvertFrom-Json).access_token
        $expirationTime = (get-date).AddSeconds(($res.Content|ConvertFrom-Json).expires_in)
    }
    End
    {
        $refreshedTokenData = New-Object PSObject
        Add-Member -InputObject $refreshedTokenData -MemberType NoteProperty -Name accessToken -Value $access_token
        Add-Member -InputObject $refreshedTokenData -MemberType NoteProperty -Name expirationTime -Value $expirationTime
        $refreshedTokenData

    }
}

$linkList = @()

$StartTime = Get-Date

$TSVMaxValue = 4095

# Removing old data
if (Test-Path $TSVFile) {Remove-Item $TSVFile}
if (Test-Path $TSVFileProcessed) {Remove-Item $TSVFileProcessed}
if (Test-Path $TSVFileG6) {Remove-Item $TSVFileG6}
if (Test-Path $TSVFileProcessedG6) {Remove-Item $TSVFileProcessedG6}

$i = 0
    do {

        $SecondsElapsed = ((Get-Date) - $StartTime).TotalSeconds
        $SecondsRemaining = ($SecondsElapsed / ($($i+1) / $TSVMaxValue)) - $SecondsElapsed
        Write-Progress -Activity "Processing Record $($i+1) of $($TSVMaxValue)" -PercentComplete (($($i+1)/$($TSVMaxValue)) * 100) -CurrentOperation "$("{0:N2}" -f ((($($i+1)/$($TSVMaxValue)) * 100),2))% Complete" -SecondsRemaining $SecondsRemaining

        $TSVToBeChecked = $i.ToString("0000")
        $requestURI = "https://oauth.reddit.com/r/SVExchange/search?q=title:$TSVToBeChecked and flair%3A(TSV+(Gen+7))&sort=new&show=all&type=link&restrict_sr=true&t=all&limit=100"
        do {
        Start-Sleep -Seconds 1
        $timeToExpire = New-TimeSpan -Start (get-date) -End $expirationTime

        if ($timeToExpire.Minutes -le 3 -or !($timeToExpire)) {

            $res = Refresh-RedditAccessToken -refreshToken $refresh_token -Credentials $Credentials

            $access_token = $res.accessToken
            $expirationTime = $res.expirationTime
            
            $headers = @{ "Authorization"="bearer $access_token";
            "User-Agent" = "SVHelper/0.1 by robdy"}
            Write-Host -ForegroundColor Green "Token refreshed"
            $timeToExpire = New-TimeSpan -Start (get-date) -End $expirationTime

        }

        $res = Invoke-WebRequest -Method Get -Uri $requestURI -Headers $headers
        } while ($? -eq $false)
        $list = ($res | ConvertFrom-Json).data.children
        $list_processed = $list|select -ExpandProperty data |select  -Property *  -ExcludeProperty selftext,selftext_html,data
        $list_processed |Export-Csv -Path $TSVFile -NoClobber -NoTypeInformation -Append
        $linkList += $list
        $list[-1].data.url
        Start-Sleep -Milliseconds 200
        $i++

    }
    while ($i -lt 4096)
    $EndTime = Get-Date



###### CONVERTING DATA BEFORE IMPORT TO GOOGLE SHEET #######

$dataFromCSV = Import-Csv -Path $TSVFile |Where-Object over_18 -eq "FALSE"
$processedData = $dataFromCSV | Select-Object title,id,url,author,author_flair_css_class,link_flair_css_class,over_18
$processedData | Export-Csv -Path $TSVFileProcessed -NoClobber -NoTypeInformation

$linkList = @()

$StartTime = Get-Date

$TSVMaxValue = 4095

$i = 0
    do {

        $SecondsElapsed = ((Get-Date) - $StartTime).TotalSeconds
        $SecondsRemaining = ($SecondsElapsed / ($($i+1) / $TSVMaxValue)) - $SecondsElapsed
        Write-Progress -Activity "Processing Record $($i+1) of $($TSVMaxValue)" -PercentComplete (($($i+1)/$($TSVMaxValue)) * 100) -CurrentOperation "$("{0:N2}" -f ((($($i+1)/$($TSVMaxValue)) * 100),2))% Complete" -SecondsRemaining $SecondsRemaining

        $TSVToBeChecked = $i.ToString("0000")
        $requestURI = "https://oauth.reddit.com/r/SVExchange/search?q=title:$TSVToBeChecked and flair%3A(TSV+(Gen+6))&sort=new&show=all&type=link&restrict_sr=true&t=all&limit=100"
        do {
        Start-Sleep -Seconds 1
        $timeToExpire = New-TimeSpan -Start (get-date) -End $expirationTime

        if ($timeToExpire.Minutes -le 3 -or !($timeToExpire)) {

            $res = Refresh-RedditAccessToken -refreshToken $refresh_token -Credentials $Credentials

            $access_token = $res.accessToken
            $expirationTime = $res.expirationTime
            
            $headers = @{ "Authorization"="bearer $access_token";
            "User-Agent" = "SVHelper/0.1 by robdy"}
            Write-Host -ForegroundColor Green "Token refreshed"
            $timeToExpire = New-TimeSpan -Start (get-date) -End $expirationTime

        }

        $res = Invoke-WebRequest -Method Get -Uri $requestURI -Headers $headers
        } while ($? -eq $false)
        $list = ($res | ConvertFrom-Json).data.children
        $list_processed = $list|select -ExpandProperty data |select  -Property *  -ExcludeProperty selftext,selftext_html,data
        $list_processed |Export-Csv -Path $TSVFileG6 -NoClobber -NoTypeInformation -Append
        $linkList += $list
        $list[-1].data.url
        Start-Sleep -Milliseconds 200
        $i++

    }
    while ($i -lt 4096)
    $EndTime = Get-Date


###### CONVERTING DATA BEFORE IMPORT TO GOOGLE SHEET #######

$dataFromCSV = Import-Csv -Path $TSVFileG6 |Where-Object over_18 -eq "FALSE"
$processedData = $dataFromCSV | Select-Object title,id,url,author,author_flair_css_class,link_flair_css_class,over_18
$processedData | Export-Csv -Path $TSVFileProcessedG6 -NoClobber -NoTypeInformation
