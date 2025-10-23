<#
.SYNOPSIS
  Gathers and downloads files from Midwest Management Summit conference sessions
.DESCRIPTION
  This script gathers and downloads files from Midwest Management Summit conference sessions. You must
  have a valid login to Sched for the year you're attempting to download.
.INPUTS
  None
.OUTPUTS
  All session content from the specified years.
.NOTES
  Version:        1.7.3
  Author:         Andrew Johnson
  Modified Date:  10/22/2025
  Purpose/Change: Fixes credential prompt issue when launching PowerShell in certain ways on Windows

  Original author (2015 script): Duncan Russell - http://www.sysadmintechnotes.com
  Edits made by:
    Evan Yeung - https://www.forevanyeung.com
    Chris Kibble - https://www.christopherkibble.com
    Jon Warnken - https://www.mrbodean.net
    Oliver Baddeley - Edited for Desert Edition
    Benjamin Reynolds - https://sqlbenjamin.wordpress.com/
    Jorge Suarez - https://github.com/jorgeasaurus
    Nathan Ziehnert - https://z-nerd.com
    Piotr Gardy - https://garit.pro


  TODO:
  [ ] Create a version history in these notes? Something like this:
  Version History/Notes:
    Date          Version    Author                    Notes
    ??/??/2015    1.0        Duncan Russell            Initial Creation?
    11/13/2019    1.1        Andrew Johnson            Added logic to only authenticate if content for the specified sessions has not been made public
    11/02/2021    1.2        Benjamin Reynolds         Added SingleEvent, MultipleEvent, and AllEvent parameters/logic; simplified logic; added a Session Info
                                                       text file containing details of the event
    04/05/2023    1.3        Jorge Suarez              Modified login body string for downloading session content
    11/06/2023    1.4        Nathan Ziehnert           Adds support for PowerShell 7.x, revamps the webscraping bit to be cross platform (no html parser in core). 
                                                       Sets default directory for non-Microsoft OS to be $HOME\Downloads\MMSContent. Ugly basic HTML parser for the
                                                       session info file, but it should suffice for now.
    04/28/2024    1.5        Andrew Johnson            Updated and tested to include 2024 at MOA
    10/20/2024    1.6        Andrew Johnson            Updated and tested to include MMS Flamingo Edition
    10/26/2024    1.6.1      Piotr Gardy               Adds functionality to re-download and check if file was updated on server
    5/1/2025      1.7        Andrew Johnson            Updated and tested to include 2025 at MOA
    5/12/2025     1.7.1      Nathan Ziehnert           Fixes a bug where the script hangs on Windows PowerShell on logon for some users
                                                       Fixes the regex for the session descriptions and speakers (unknown when this broke)
                                                       Adds throttling to avoid 429 Too Many Requests errors (if request fails due to 429, script waits 20 seconds and retries)
                                                       Could be improved with exponential backoff, but this is a start - also 20 seconds seemed to work best (15 almost worked)
    10/13/2025    1.7.2      Andrew Johnson            Updated and tested to include 2025 Music City Edition
    10/22/2025    1.7.3      Nathan Ziehnert           Fixes credential prompt issue when launching PowerShell in certain ways on Windows                                                       

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -ConferenceList @('2025atmoa','2025music');

  Downloads all MMS session content from 2025 at MOA and 2025 Music City Edition on to C:\Conferences\MMS\

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -DownloadLocation "C:\Temp\MMS" -ConferenceId 2025music

  Downloads all MMS session content from 2025 at MOA to C:\Temp\MMS\

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -All

  Downloads all MMS session content from all years to C:\Conferences\MMS\

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -All -ExcludeSessionDetails;

  Downloads all MMS session content from all years to C:\Conferences\MMS\ BUT does not include a "Session Info.txt" file for each session containing the session details

.LINK
  Project URL - https://github.com/AndrewJNet/CopyMMSFiles
#>
[cmdletbinding(PositionalBinding = $false)]
Param(
  [Parameter(Mandatory = $false)][string]$DownloadLocation = "C:\Conferences\MMS", # could validate this: [ValidateScript({(Test-Path -Path (Split-Path $PSItem))})]
  [Parameter(Mandatory = $true, ParameterSetName = 'SingleEvent')]
  [ValidateSet("2015", "2016", "2017", "2018", "de2018", "2019", "jazz", "miami", "2022atmoa", "2023atmoa", "2023miami", "2024atmoa", "2024fll", "2025atmoa", "2025music")]
  [string]$ConferenceId,
  [Parameter(Mandatory = $true, ParameterSetName = 'MultipleEvents', HelpMessage = "This needs to bwe a list or array of conference ids/years!")]
  [System.Collections.Generic.List[string]]$ConferenceList,
  [Parameter(Mandatory = $true, ParameterSetName = 'AllEvents')][switch]$All,
  [Parameter(Mandatory = $false)][switch]$ExcludeSessionDetails,
  [Parameter(Mandatory = $false)][switch]$ReDownloadIsHashIsDifferent
)

function Invoke-BasicHTMLParser ($html) {
  $html = $html.Replace("<br>", "`r`n").Replace("<br/>", "`r`n").Replace("<br />", "`r`n") # replace <br> with new line

  # Speaker Spacing
  $html = $html.Replace("<div class=`"sched-person-session`">", "`r`n`r`n")

  # Link parsing
  $linkregex = '(?<texttoreplace><a.*?href="(?<link>.*?)".*?>(?<content>.*?)<\/a>)'
  $links = [regex]::Matches($html, $linkregex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  foreach ($l in $links) {
    if (-not $l.Groups['link'].Value.StartsWith("http")) { $link = "$SchedBaseURL/$($l.Groups['link'].Value)" }else { $link = $l.Groups['link'].Value }
    $html = $html.Replace($l.Groups['texttoreplace'].Value, " [$($l.Groups['content'].Value)]($link)")
  }

  # List Parsing
  $listRegex = '(?<texttoreplace><ul[^>]?>(?<content>.*?)<\/ul>)'
  $lists = [regex]::Matches($html, $listRegex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  foreach ($l in $lists) {
    $content = $l.Groups['content'].Value.Replace("<li>", "`r`n* ").Replace("</li>", "")
    $html = $html.Replace($l.Groups['texttoreplace'].Value, $content)
  }

  # General Cleanup
  $html = $html.replace("&rarr;", "")
  $html = $html -replace '<div[^>]+>', "`r`n"
  $html = $html -replace '<[^>]+>', '' # Strip all HTML tags

  ## Future revisions
  # do something about <b> / <i> / <strong> / etc...
  # maybe a converter to markdown
  
  return $html
}
## Hide Invoke-WebRequest progress bar. There's a bug that doesn't clear the bar after a request is finished. 
$ProgressPreference = "SilentlyContinue"
## Determine OS... sorta
if ($PSEdition -eq "Desktop" -or $isWindows) { $win = $true }
else { 
  $win = $false
  if ($DownloadLocation -eq "C:\Conferences\MMS") { $DownloadLocation = "$HOME\Downloads\MMSContent" }
}

## Make sure there aren't any trailing backslashes:
$DownloadLocation = $DownloadLocation.Trim('\')

## Setup
$PublicContentYears = @('2015', '2016', '2017', '2019', 'jazz', 'miami', '2022atmoa', '2023atmoa','2023miami', '2024atmoa', '2024fll')
$PrivateContentYears = @('2018', 'de2018', '2025atmoa', '2025music')
$ConferenceYears = New-Object -TypeName System.Collections.Generic.List[string]
[int]$PublicYearsCount = $PublicContentYears.Count
[int]$PrivateYearsCount = $PrivateContentYears.Count

if ($All) {
  for ($i = 0; $i -lt $PublicYearsCount; $i++) {
    $ConferenceYears.Add($PublicContentYears[$i])
  }
  Remove-Variable -Name i -ErrorAction SilentlyContinue
  for ($i = 0; $i -lt $PrivateYearsCount; $i++) {
    $ConferenceYears.Add($PrivateContentYears[$i])
  }
  Remove-Variable -Name i -ErrorAction SilentlyContinue
}
elseif ($PsCmdlet.ParameterSetName -eq 'SingleEvent') {
  $ConferenceYears.Add($ConferenceId)
}
else {
  $ConfListCount = $ConferenceList.Count
  for ($i = 0; $i -lt $ConfListCount; $i++) {
    if ($ConferenceList[$i] -in ($PublicContentYears + $PrivateContentYears)) {
      $ConferenceYears.Add($ConferenceList[$i])
    }
    else {
      Write-Output "The Conference Id '$($ConferenceList[$i])' is not valid. Item will be skipped."
    }
  }
  Remove-Variable -Name i -ErrorAction SilentlyContinue
}

Write-Output "Base Download URL is $DownloadLocation"
Write-Output "Searching for content from these sessions: $([String]::Join(',',$ConferenceYears))"

##
$ConferenceYears | ForEach-Object -Process {
  [string]$Year = $_

  if ($Year -in $PrivateContentYears) {
    ## We're going to generate the credential prompt manually to avoid issues
    ## with the credential prompt not working in some scenarios. Specifically
    ## when launching PowerShell from "run" or directly double-clicking on
    ## powershell.exe in Windows versions that then launch the Windows Terminal.
    Write-Host "Credentials required for $Year content."
    $un = Read-Host "Enter Username for $Year Sched"
    $pw = Read-Host "Enter Password for $Year Sched" -AsSecureString
    $creds = [System.Management.Automation.PSCredential]::new($un, $pw)
    $un, $pw = $null
  }

  $SchedBaseURL = "https://mms" + $Year + ".sched.com"
  $SchedLoginURL = $SchedBaseURL + "/login"
  Add-Type -AssemblyName System.Web
  $web = Invoke-WebRequest $SchedLoginURL -SessionVariable mms -UseBasicParsing
  ## Connect to Sched

  if ($creds) {
    #$form = $web.Forms[1]
    #$form.fields['username'] = $creds.UserName;
    #$form.fields['password'] = $creds.GetNetworkCredential().Password;

    $username = $creds.UserName
    $password = $creds.GetNetworkCredential().Password

    # Updated POST body
    $body = "landing_conf=" + [System.Uri]::EscapeDataString($SchedBaseURL) + "&username=" + [System.Uri]::EscapeDataString($username) + "&password=" + [System.Uri]::EscapeDataString($password) + "&login="

    # SEND IT
    $web = Invoke-WebRequest $SchedLoginURL -SessionVariable mms -Method POST -Body $body -UseBasicParsing

  }
  else {
    $web = Invoke-WebRequest $SchedLoginURL -SessionVariable mms -UseBasicParsing
  }

  $SessionDownloadPath = $DownloadLocation + '\mms' + $Year
  Write-Output "Logging in to $SchedBaseURL"

  ## Check if we connected (if required):
  if ((-Not ($web.InputFields.FindByName("login")) -and ($Year -in $PrivateContentYears)) -or ($Year -in $PublicContentYears)) {
    ##
    Write-Output "Downloaded content can be found in $SessionDownloadPath"

    $sched = Invoke-WebRequest -Uri $($SchedBaseURL + "/list/descriptions") -WebSession $mms -UseBasicParsing
    $links = $sched.Links
    # For indexing available downloads later
    $eventsList = New-Object -TypeName System.Collections.Generic.List[int]
    $links | ForEach-Object -Process {
      if ($_.href -like "event/*") {
        [void]$eventsList.Add($links.IndexOf($_))
      }
    }
    $eventCount = $eventsList.Count

    for ($i = 0; $i -lt $eventCount; $i++) {
      [int]$linkIndex = $eventsList[$i]
      [int]$nextLinkIndex = $eventsList[$i + 1]
      $eventobj = $links[($eventsList[$i])]

      # Get/Fix the Session Title:
      $titleRegex = '<a.*?href="(?<url>.*?)".*?>(?<title>.*?)(<span|<\/a>)'
      $titleMatches = [regex]::Matches($eventobj.outerHTML.Replace("`r", "").Replace("`n", ""), $titleRegex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
      [string]$eventTitle = $titleMatches.Groups[0].Groups['title'].Value.Trim()
      [string]$eventUrl = $titleMatches.Groups[0].Groups['url'].Value.Trim()

      # Generate session info string
      [string]$sessionInfoText = ""
      $sessionInfoText += "Session Title: `r`n$eventTitle`r`n`r`n"
      $downloadTitle = $eventTitle -replace "[^A-Za-z0-9-_. ]", ""
      $downloadTitle = $downloadTitle.Trim()
      $downloadTitle = $downloadTitle -replace "\W+", "_"

      ## Set the download destination:
      $downloadPath = $SessionDownloadPath + "\" + $downloadTitle

      ## Get session info if required:
      if (-not $ExcludeSessionDetails) {
        try{
          #Wait-Debugger
          $sessionLinkInfo = (Invoke-WebRequest -Uri $($SchedBaseURL + "/" + $eventUrl) -WebSession $mms -UseBasicParsing).Content.Replace("`r", "").Replace("`n", "")
        }
        catch {
          if (($_.Exception.GetType().FullName -eq "System.Net.WebException" -or $_.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") `
              -and $_.Exception.Response.StatusCode -eq 429) {
            Write-Warning "Received 429 Too Many Requests error. Waiting 20 seconds and retrying..."
            Start-Sleep -Seconds 20
            $sessionLinkInfo = (Invoke-WebRequest -Uri $($SchedBaseURL + "/" + $eventUrl) -WebSession $mms -UseBasicParsing).Content.Replace("`r", "").Replace("`n", "")
          }
        }

        $descriptionPattern = '<div class="tip-description">(?<description>.*?)(<div class="tip-roles">|<div class="sched-event-details-timeandplace">)'
        $description = [regex]::Matches($sessionLinkInfo, $descriptionPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($description.Count -gt 0) { $sessionInfoText += "$(Invoke-BasicHTMLParser -html $description.Groups[0].Groups['description'].Value)`r`n`r`n" }

        $rolesPattern = '<div class="tip-roles">(?<roles>.*?)<div class="sched-file">|<div class="sched-event-details-timeandplace">'
        $roles = [regex]::Matches($sessionLinkInfo, $rolesPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($roles.Count -gt 0) { $sessionInfoText += "$(Invoke-BasicHTMLParser -html $roles.Groups[0].Groups['roles'].Value)`r`n`r`n" }

        if ((Test-Path -Path $($downloadPath)) -eq $false) { New-Item -ItemType Directory -Force -Path $downloadPath | Out-Null }
        Out-File -FilePath "$downloadPath\Session Info.txt" -InputObject $sessionInfoText -Force -Encoding default
      }

      $downloads = $links[($linkIndex + 1)..($nextLinkIndex - 1)] | Where-Object { $_.href -like "*hosted_files*" } #prefilter
      foreach ($download in $downloads) {
        $filename = Split-Path $download.href -Leaf
        # Replace HTTP Encoding Characters (e.g. %20) with the proper equivalent.
        $filename = [System.Web.HttpUtility]::UrlDecode($filename)
        # Replace non-standard characters
        $filename = $filename -replace "[^A-Za-z0-9\.\-_ ]", ""

        $outputFilePath = $downloadPath + '\' + $filename

        # Reduce Total Path to 255 characters.
        $outputFilePathLen = $outputFilePath.Length
        if ($outputFilePathLen -ge 255) {
          $fileExt = [System.IO.Path]::GetExtension($outputFilePath)
          $newFileName = $outputFilePath.Substring(0, $($outputFilePathLen - $fileExt.Length))
          $newFileName = $newFileName.Substring(0, $(255 - $fileExt.Length)).trim()
          $newFileName = "$newFileName$fileExt"
          $outputFilePath = $newFileName
        }

        # Download the file
        if ((Test-Path -Path $($downloadPath)) -eq $false) { New-Item -ItemType Directory -Force -Path $downloadPath | Out-Null }
        if ((Test-Path -Path $outputFilePath) -eq $false) {
          Write-host -ForegroundColor Green "...attempting to download '$filename' because it doesn't exist"
          try {
            try{
              Invoke-WebRequest -Uri $download.href -OutFile $outputfilepath -WebSession $mms -UseBasicParsing
            }
            catch {
              if (($_.Exception.GetType().FullName -eq "System.Net.WebException" -or $_.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") `
                  -and $_.Exception.Response.StatusCode -eq 429) {
                Write-Warning "Received 429 Too Many Requests error. Waiting 20 seconds and retrying..."
                Start-Sleep -Seconds 20
                Invoke-WebRequest -Uri $download.href -OutFile $outputfilepath -WebSession $mms -UseBasicParsing
              }
            }
            if ($win) { Unblock-File $outputFilePath }
          }
          catch {
            Write-Output ".................$($PSItem.Exception) for '$($download.href)'...moving to next file..."
          }
        }
        else {
          if ($ReDownloadIsHashIsDifferent) {
            Write-Output "...attempting to download '$filename'"
            $oldHash = (Get-FileHash $outputFilePath).Hash
            try {
              try{
                Invoke-WebRequest -Uri $download.href -OutFile "$($outputfilepath).new" -WebSession $mms -UseBasicParsing
              }
              catch {
                if (($_.Exception.GetType().FullName -eq "System.Net.WebException" -or $_.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") `
                    -and $_.Exception.Response.StatusCode -eq 429) {
                  Write-Warning "Received 429 Too Many Requests error. Waiting 20 seconds and retrying..."
                  Start-Sleep -Seconds 20
                  Invoke-WebRequest -Uri $download.href -OutFile "$($outputfilepath).new" -WebSession $mms -UseBasicParsing
                }
              }
              if ($win) { Unblock-File "$($outputfilepath).new" }
              $NewHash = (Get-FileHash "$($outputfilepath).new").Hash
              if ($NewHash -ne $oldHash) {
                Write-Host -ForegroundColor Green " => HASH is different. Keeping new file"
                Move-Item "$($outputfilepath).new" $outputfilepath -Force
              }
              else {
                Write-Output " => Hash is the same. "
                Remove-item "$($outputfilepath).new" -Force
              }
            }
            catch {
              Write-Output ".................$($PSItem.Exception) for '$($download.href)'...moving to next file..."
            }
          }
        }
      } # end procesing downloads
    } # end processing session
  } # end connectivity/login check
  else {
    Write-Output "Login to $SchedBaseUrl failed."
  }
}
