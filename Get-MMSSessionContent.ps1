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
  Version:        1.7.4
  Author:         Andrew Johnson
  Modified Date:  5/5/2025
  Purpose/Change: Updated and tested to include 2026 at MOA 

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
    10/22/2025    1.7.4      Andrew Johnson            Updated and tested to include 2026 at MOA                                                  

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -ConferenceList @('2026atmoa','2025music');

  Downloads all MMS session content from 2026 at MOA and 2025 Music City Edition on to C:\Conferences\MMS\

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -DownloadLocation "C:\Temp\MMS" -ConferenceId 2026atmoa

  Downloads all MMS session content from 2026 at MOA to C:\Temp\MMS\

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
  [ValidateSet("2015", "2016", "2017", "2018", "de2018", "2019", "jazz", "miami", "2022atmoa", "2023atmoa", "2023miami", "2024atmoa", "2024fll", "2025atmoa", "2025music","2026atmoa")]
  [string]$ConferenceId,
  [Parameter(Mandatory = $true, ParameterSetName = 'MultipleEvents', HelpMessage = "This needs to be a list or array of conference IDs/years!")]
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

function Get-EventLinkPairs {
  param(
    [Parameter(Mandatory = $true)]
    [object]$SchedulePage
  )

  $eventPairs = New-Object -TypeName System.Collections.Generic.List[object]
  $eventTitleByUrl = @{}

  if ($SchedulePage.Content) {
    $eventRegex = '<a[^>]+href="(?<url>event\/[^"#?]+)"[^>]*>(?<title>.*?)<\/a>'
    $eventMatches = [regex]::Matches($SchedulePage.Content.Replace("`r", "").Replace("`n", ""), $eventRegex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($m in $eventMatches) {
      $u = [string]$m.Groups['url'].Value.Trim()
      $t = [regex]::Replace($m.Groups['title'].Value, '<[^>]+>', '').Trim()
      if (-not [string]::IsNullOrWhiteSpace($u) -and -not [string]::IsNullOrWhiteSpace($t) -and -not $eventTitleByUrl.ContainsKey($u)) {
        $eventTitleByUrl[$u] = $t
      }
    }
  }

  # First attempt: use parsed Links collection
  if ($SchedulePage.Links) {
    foreach ($lnk in $SchedulePage.Links) {
      if ($lnk.href -like 'event/*') {
        [string]$url = [string]$lnk.href
        [string]$title = [string]$lnk.innerText

        if ([string]::IsNullOrWhiteSpace($title) -and $lnk.outerHTML) {
          $title = [regex]::Replace([string]$lnk.outerHTML, '<[^>]+>', '').Trim()
        }

        if ([string]::IsNullOrWhiteSpace($title) -and $eventTitleByUrl.ContainsKey($url)) {
          $title = [string]$eventTitleByUrl[$url]
        }

        if ([string]::IsNullOrWhiteSpace($title) -and $url -match '^event\/[^\/]+\/(?<slug>[^\/?#]+)') {
          $title = [System.Uri]::UnescapeDataString($matches['slug']) -replace '-', ' '
        }

        $eventPairs.Add([PSCustomObject]@{
            Url = $url
            Title = $title
          })
      }
    }
  }

  # Fallback: parse anchor tags from raw HTML when Links parsing is incomplete
  if ($eventPairs.Count -eq 0 -and $SchedulePage.Content) {
    foreach ($entry in $eventTitleByUrl.GetEnumerator()) {
      $eventPairs.Add([PSCustomObject]@{
          Url = [string]$entry.Key
          Title = [string]$entry.Value
        })
    }
  }

  # De-duplicate by URL while preserving order
  $seen = @{}
  $uniquePairs = New-Object -TypeName System.Collections.Generic.List[object]
  foreach ($p in $eventPairs) {
    if ([string]::IsNullOrWhiteSpace($p.Url)) { continue }
    if (-not $seen.ContainsKey($p.Url)) {
      $seen[$p.Url] = $true
      $uniquePairs.Add($p)
    }
  }

  return $uniquePairs
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
$PublicContentYears = @('2015', '2016', '2017', '2019', 'jazz', 'miami', '2022atmoa', '2023atmoa','2023miami', '2024atmoa', '2024fll','2025atmoa', '2025music')
$PrivateContentYears = @('2018', 'de2018', '2026atmoa')
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

  if ($creds) {
    #$form = $web.Forms[1]
    #$form.fields['username'] = $creds.UserName;
    #$form.fields['password'] = $creds.GetNetworkCredential().Password;

    $username = $creds.UserName
    $password = $creds.GetNetworkCredential().Password

    # Updated POST body
    $body = "landing_conf=" + [System.Uri]::EscapeDataString($SchedBaseURL) + "&username=" + [System.Uri]::EscapeDataString($username) + "&password=" + [System.Uri]::EscapeDataString($password) + "&login="

    # SEND IT
    $web = Invoke-WebRequest $SchedLoginURL -WebSession $mms -Method POST -Body $body -UseBasicParsing -Headers @{ Referer = $SchedLoginURL; Origin = $SchedBaseURL }
  }
  else {
    $web = Invoke-WebRequest $SchedLoginURL -SessionVariable mms -UseBasicParsing
  }

  $SessionDownloadPath = $DownloadLocation + '\mms' + $Year

  ## Check if we connected (if required):
  if ((-Not ($web.InputFields.FindByName("login")) -and ($Year -in $PrivateContentYears)) -or ($Year -in $PublicContentYears)) {
    try {
      $sched = Invoke-WebRequest -Uri $($SchedBaseURL + "/list/descriptions") -WebSession $mms -UseBasicParsing -ErrorAction Stop
    }
    catch {
      Write-Output "FAILED: unable to retrieve session list"
      continue
    }

    $events = Get-EventLinkPairs -SchedulePage $sched
    $eventCount = $events.Count

    for ($i = 0; $i -lt $eventCount; $i++) {
      $event = $events[$i]

      [string]$eventTitle = [string]$event.Title
      [string]$eventUrl = [string]$event.Url
      if ([string]::IsNullOrWhiteSpace($eventTitle)) { $eventTitle = $eventUrl }
      if ([string]::IsNullOrWhiteSpace($eventUrl)) { continue }

      # Strip capacity labels (FULL, LIMITED, FILLING, etc.) that Sched appends to the anchor text
      $eventTitle = $eventTitle -replace '\s*(FULL|LIMITED|SOLD OUT|WAITLIST|FILLING)\s*$', ''
      $eventTitle = $eventTitle.Trim()

      # Skip social/networking/activity events that never have downloadable content
      if ($eventTitle -match '^Fishing with' -or $eventTitle -match '^Career Connections' -or
          $eventTitle -match '^Great Big Game Show' -or $eventTitle -match 'Escape Game') { continue }

      # Generate session info string
      [string]$sessionInfoText = ""
      $sessionInfoText += "Session Title: `r`n$eventTitle`r`n`r`n"
      $downloadTitle = $eventTitle -replace "[^A-Za-z0-9-_. ]", ""
      $downloadTitle = $downloadTitle.Trim()
      $downloadTitle = $downloadTitle -replace "\W+", "_"

      ## Set the download destination:
      $downloadPath = $SessionDownloadPath + "\" + $downloadTitle

      # Fetch the individual event page. Hosted files are now reliably discovered there.
      $sessionPage = $null
      try {
        $sessionPage = Invoke-WebRequest -Uri $($SchedBaseURL + "/" + $eventUrl) -WebSession $mms -UseBasicParsing -ErrorAction Stop
      }
      catch {
        if (($_.Exception.GetType().FullName -eq "System.Net.WebException" -or $_.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") `
            -and $_.Exception.Response.StatusCode -eq 429) {
          Start-Sleep -Seconds 20
          try {
            $sessionPage = Invoke-WebRequest -Uri $($SchedBaseURL + "/" + $eventUrl) -WebSession $mms -UseBasicParsing -ErrorAction Stop
          }
          catch {
            $sessionPage = $null
          }
        }
      }

      if (-not $sessionPage) {
        Write-Output "FAILED: $eventTitle"
        continue
      }

      $sessionLinkInfo = $sessionPage.Content.Replace("`r", "").Replace("`n", "")

      ## Get session info if required:
      if (-not $ExcludeSessionDetails) {
        $descriptionPattern = '<div class="tip-description">(?<description>.*?)(<div class="tip-roles">|<div class="sched-event-details-timeandplace">)'
        $description = [regex]::Matches($sessionLinkInfo, $descriptionPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($description.Count -gt 0) { $sessionInfoText += "$(Invoke-BasicHTMLParser -html $description.Groups[0].Groups['description'].Value)`r`n`r`n" }

        $rolesPattern = '<div class="tip-roles">(?<roles>.*?)<div class="sched-file">|<div class="sched-event-details-timeandplace">'
        $roles = [regex]::Matches($sessionLinkInfo, $rolesPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($roles.Count -gt 0) { $sessionInfoText += "$(Invoke-BasicHTMLParser -html $roles.Groups[0].Groups['roles'].Value)`r`n`r`n" }

        if ((Test-Path -Path $($downloadPath)) -eq $false) { New-Item -ItemType Directory -Force -Path $downloadPath | Out-Null }
        Out-File -FilePath "$downloadPath\Session Info.txt" -InputObject $sessionInfoText -Force -Encoding default
      }

      $downloads = $sessionPage.Links | Where-Object {
        $_.href -like "*hosted-files*" -or
        $_.href -like "*hosted-files.sched.co*" -or
        $_.href -like "//hosted-files.sched.co/*"
      } | Select-Object -Unique

      # Fallback: parse hosted-files URLs directly from HTML when Links collection is sparse
      if ($downloads.Count -eq 0 -and $sessionPage.Content) {
        $hostedFilesRegex = 'href="((?:https?:)?//hosted-files\.sched\.co/[^"]+)"'
        $hostedMatches = [regex]::Matches($sessionPage.Content, $hostedFilesRegex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        if ($hostedMatches.Count -gt 0) {
          $downloads = @()
          foreach ($m in $hostedMatches) {
            $url = $m.Groups[1].Value
            if ($url -notin ($downloads | Select-Object -ExpandProperty href)) {
              $downloads += [PSCustomObject]@{ href = $url }
            }
          }
        }
      }

      if ($downloads.Count -eq 0) {
        Write-Output "No files found for: $eventTitle"
      }
      

      foreach ($download in $downloads) {
        $downloadUrl = $download.href
        if ($downloadUrl -like "//hosted-files.sched.co/*") {
          $downloadUrl = "https:$downloadUrl"
        }
        elseif ($downloadUrl -notmatch '^https?://') {
          $downloadUrl = "$SchedBaseURL/$downloadUrl"
        }

        $filename = Split-Path $downloadUrl -Leaf
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
          Write-Output "DOWNLOADING: $filename"
          try {
            try{
              Invoke-WebRequest -Uri $downloadUrl -OutFile $outputfilepath -WebSession $mms -UseBasicParsing -ErrorAction Stop
            }
            catch {
              if (($_.Exception.GetType().FullName -eq "System.Net.WebException" -or $_.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") `
                  -and $_.Exception.Response.StatusCode -eq 429) {
                Start-Sleep -Seconds 20
                Invoke-WebRequest -Uri $downloadUrl -OutFile $outputfilepath -WebSession $mms -UseBasicParsing -ErrorAction Stop
              }
              else { throw }
            }
            if ($win) { Unblock-File $outputFilePath }
          }
          catch {
            Write-Output "FAILED: $filename"
          }
        }
        else {
          if ($ReDownloadIsHashIsDifferent) {
            Write-Output "DOWNLOADING: $filename"
            $oldHash = (Get-FileHash $outputFilePath).Hash
            try {
              try{
                Invoke-WebRequest -Uri $downloadUrl -OutFile "$($outputfilepath).new" -WebSession $mms -UseBasicParsing -ErrorAction Stop
              }
              catch {
                if (($_.Exception.GetType().FullName -eq "System.Net.WebException" -or $_.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") `
                    -and $_.Exception.Response.StatusCode -eq 429) {
                  Start-Sleep -Seconds 20
                  Invoke-WebRequest -Uri $downloadUrl -OutFile "$($outputfilepath).new" -WebSession $mms -UseBasicParsing -ErrorAction Stop
                }
                else { throw }
              }
              if ($win) { Unblock-File "$($outputfilepath).new" }
              $NewHash = (Get-FileHash "$($outputfilepath).new").Hash
              if ($NewHash -ne $oldHash) {
                Move-Item "$($outputfilepath).new" $outputfilepath -Force
              }
              else {
                Write-Output "SKIPPING: $filename"
                Remove-item "$($outputfilepath).new" -Force
              }
            }
            catch {
              Write-Output "FAILED: $filename"
            }
          }
          else {
            Write-Output "SKIPPING: $filename"
          }
        }
      } # end procesing downloads
    } # end processing session
    } # end connectivity/login check
  else {
    Write-Output "FAILED: login to $SchedBaseUrl"
  }
}

