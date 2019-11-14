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
  Version:        1.1
  Author:         Andrew Johnson
  Modified Date:  11/13/2019
  Purpose/Change: Added logic to only authenticate if content for the specified sessions has not been made public

  Original author (2015 script): Duncan Russell - http://www.sysadmintechnotes.com
  Edits made by:
    Evan Yeung - https://www.forevanyeung.com
    Chris Kibble - https://www.christopherkibble.com
    Jon Warnken - https://www.mrbodean.net
    Oliver Baddeley - Edited for Desert Edition



.EXAMPLE
  .\Get-MMSSessionContent.ps1

  Downloads all MMS session content from the current year to C:\Conferences\MMS\

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -DownloadLocation "C:\Temp\MMS" -ConferenceYears 2015

  Downloads all MMS session content from 2015 to C:\Temp\MMS\

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -All $True

  Downloads all MMS session content from all years to C:\Conferences\MMS\

.LINK
  Project URL - https://github.com/AndrewJNet/CopyMMSFiles
#>


Param(
  $DownloadLocation = "C:\Conferences\MMS\",
  $ConferenceYears = (Get-Date).year,
  $All = $false
)

$PublicContentYears = @('2015', '2016', '2017')
$PrivateContentYears = @('2018', 'de2018', '2019', 'jazz')

if ($All -eq $true) {
  $ConferenceYears = $PublicContentYears + $PrivateContentYears
  $ConferenceYearsPublic = $PublicContentYears
  $ConferenceYearsPrivate = $PrivateContentYears
}

Write-Output "Base Download URL is $DownloadLocation"
Write-Output "Searching for content from these sessions: $([String]::Join(',',$ConferenceYears))"

$ConferenceYearsPublic = $ConferenceYears | Where-Object { $_ -notin $PrivateContentYears }
$ConferenceYearsPrivate = $ConferenceYears | Where-Object { $_ -notin $PublicContentYears }

if ($ConferenceYearsPrivate.count -gt 0) {
  $c = $host.UI.PromptForCredential('Sched Credentials', 'Enter Credentials', '', '')
  foreach ($year in $ConferenceYearsPrivate) {
    $SchedBaseURL = "https://mms" + $Year + ".sched.com"
    $SchedLoginURL = $SchedBaseURL + "/login"
    Add-Type -AssemblyName System.Web
    $web = Invoke-WebRequest $SchedLoginURL -SessionVariable mms
    $form = $web.Forms[1]
    $form.fields['username'] = $c.UserName
    $form.fields['password'] = $c.GetNetworkCredential().Password
    $SessionDownloadPath = $DownloadLocation + '\mms' + $Year

    "Logging in to $SchedBaseURL"


    $mmsHome = Invoke-WebRequest $SchedBaseURL -WebSession $mms

    $htmlDate = $mmsHome.ParsedHtml.IHTMLDocument3_GetElementById('sched-sidebar-filters-dates')
    $htmlPopoverBody = $htmlDate.getElementsByClassName('popover')
    $htmlDateList = $htmlPopoverBody[0].getElementsByTagName('li')

    $MMSDates = @()
    $htmlDateList | ForEach-Object {
      out-null -InputObject  $($_.innerHTML -match "\d{4}-\d{2}-\d{2}")
      $MMSDates += $matches[0]
    }

    $web = Invoke-WebRequest $SchedLoginURL -WebSession $mms -Method POST -Body $form.Fields
    if (-Not ($web.InputFields.FindByName("login"))) {
      Write-Output "Downloaded content can be found in $SessionDownloadPath"
      ForEach ($Date in $MMSDates) {
        "Checking day '{0}' for downloads" -f $Date

        $sched = Invoke-WebRequest -Uri $($SchedBaseURL + "/" + $Date + "/list/descriptions") -WebSession $mms

        $links = $sched.Links
        $eventsIndex = @()
        $links | ForEach-Object { if (($_.href -like "*/event/*") -and ($_.innerText -notlike "here")) {
            $eventsIndex += (, ($links.IndexOf($_), $_.innerText))
          } }
        $i = 0
        While ($i -lt $eventsIndex.Count) {
          $eventTitle = $eventsIndex[$i][1]
          $eventTitle = $eventTitle -replace "[^A-Za-z0-9-_. ]", ""
          $eventTitle = $eventTitle.Trim()
          $eventTitle = $eventTitle -replace "\W+", "_"
          $links[$eventsIndex[$i][0]..$(if ($i -eq $eventsIndex.Count - 1) { $links.Count - 1 } else { $eventsIndex[$i + 1][0] })] | ForEach-Object {

            if ($_.href -like "*hosted_files*") {
              $downloadPath = $DownloadLocation + '\mms' + $Year + '\' + $Date + '\' + $eventTitle
              $filename = $_.href
              $filename = $filename.substring(40)
              # Replace HTTP Encoding Characters (e.g. %20) with the proper equivalent.
              $filename = [System.Web.HttpUtility]::UrlDecode($filename)
              # Replace non-standard characters
              $filename = $filename -replace "[^A-Za-z0-9\.\-_ ]", ""
              # Remove 7 characters added to filenames in 2019
              $filename = $filename.Substring(7)
              $outputFilePath = $downloadPath + '\' + $filename
              # Reduce Total Path to 255 characters.
              $outputFilePathLen = $outputFilePath.Length

              If ($outputFilePathLen -ge 255) {
                $fileExt = [System.IO.Path]::GetExtension($outputFilePath)
                $newFileName = $outputFilePath.Substring(0, $($outputFilePathLen - $fileExt.Length))
                $newFileName = $newFileName.Substring(0, $(255 - $fileExt.Length)).trim()
                $newFileName = "$newFileName$fileExt"
                $outputFilePath = $newFileName
              }

              if ((Test-Path -Path $($downloadPath)) -eq $false) { New-Item -ItemType Directory -Force -Path $downloadPath | Out-Null }
              if ((Test-Path -Path $outputFilePath) -eq $false) {
                "...attempting to download '{0}'" -f $filename
                Invoke-WebRequest -Uri $_.href -OutFile $outputfilepath -WebSession $mms
                Unblock-File $outputFilePath
              }
            }
          }

          $i++
        }
      }
    }
    else {
      "Login to $SchedBaseUrl failed."
    }
  }
}

foreach ($year in $ConferenceYearsPublic) {
  $SchedBaseURL = "https://mms" + $Year + ".sched.com"
  $SchedLoginURL = $SchedBaseURL
  Add-Type -AssemblyName System.Web
  $web = Invoke-WebRequest $SchedLoginURL -SessionVariable mms
  $SessionDownloadPath = $DownloadLocation + '\mms' + $Year

  "Logging in to $SchedBaseURL"


  $mmsHome = Invoke-WebRequest $SchedBaseURL -WebSession $mms

  $htmlDate = $mmsHome.ParsedHtml.IHTMLDocument3_GetElementById('sched-sidebar-filters-dates')
  $htmlPopoverBody = $htmlDate.getElementsByClassName('popover')
  $htmlDateList = $htmlPopoverBody[0].getElementsByTagName('li')

  $MMSDates = @()
  $htmlDateList | ForEach-Object {
    out-null -InputObject  $($_.innerHTML -match "\d{4}-\d{2}-\d{2}")
    $MMSDates += $matches[0]
  }

  $web = Invoke-WebRequest $SchedLoginURL -WebSession $mms
  Write-Output "Downloaded content can be found in $SessionDownloadPath"
  ForEach ($Date in $MMSDates) {
    "Checking day '{0}' for downloads" -f $Date

    $sched = Invoke-WebRequest -Uri $($SchedBaseURL + "/" + $Date + "/list/descriptions") -WebSession $mms

    $links = $sched.Links
    $eventsIndex = @()
    $links | ForEach-Object { if (($_.href -like "*/event/*") -and ($_.innerText -notlike "here")) {
        $eventsIndex += (, ($links.IndexOf($_), $_.innerText))
      } }
    $i = 0
    While ($i -lt $eventsIndex.Count) {
      $eventTitle = $eventsIndex[$i][1]
      $eventTitle = $eventTitle -replace "[^A-Za-z0-9-_. ]", ""
      $eventTitle = $eventTitle.Trim()
      $eventTitle = $eventTitle -replace "\W+", "_"
      $links[$eventsIndex[$i][0]..$(if ($i -eq $eventsIndex.Count - 1) { $links.Count - 1 } else { $eventsIndex[$i + 1][0] })] | ForEach-Object {

        if ($_.href -like "*hosted_files*") {
          $downloadPath = $DownloadLocation + '\mms' + $Year + '\' + $Date + '\' + $eventTitle
          $filename = $_.href
          $filename = $filename.substring(40)
          # Replace HTTP Encoding Characters (e.g. %20) with the proper equivalent.
          $filename = [System.Web.HttpUtility]::UrlDecode($filename)
          # Replace non-standard characters
          $filename = $filename -replace "[^A-Za-z0-9\.\-_ ]", ""
          # Remove 7 characters added to filenames in 2019
          $filename = $filename.Substring(7)
          $outputFilePath = $downloadPath + '\' + $filename
          # Reduce Total Path to 255 characters.
          $outputFilePathLen = $outputFilePath.Length

          If ($outputFilePathLen -ge 255) {
            $fileExt = [System.IO.Path]::GetExtension($outputFilePath)
            $newFileName = $outputFilePath.Substring(0, $($outputFilePathLen - $fileExt.Length))
            $newFileName = $newFileName.Substring(0, $(255 - $fileExt.Length)).trim()
            $newFileName = "$newFileName$fileExt"
            $outputFilePath = $newFileName
          }

          if ((Test-Path -Path $($downloadPath)) -eq $false) { New-Item -ItemType Directory -Force -Path $downloadPath | Out-Null }
          if ((Test-Path -Path $outputFilePath) -eq $false) {
            "...attempting to download '{0}'" -f $filename
            Invoke-WebRequest -Uri $_.href -OutFile $outputfilepath -WebSession $mms
            Unblock-File $outputFilePath
          }
        }
      }

      $i++
    }
  }
}