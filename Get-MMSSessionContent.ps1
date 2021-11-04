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
    Benjamin Reynolds - https://sqlbenjamin.wordpress.com/

  TODO:
  [ ] Create a version history in these notes? Something like this:
  Version History/Notes:
    Date          Version    Author                    Notes
    ??/??/2015    1.0        Duncan Russell            Initial Creation? 
    11/13/2019    1.1        Andrew Johnson            Added logic to only authenticate if content for the specified sessions has not been made public
    11/02/2021    1.2        Benjamin Reynolds         Added SingleEvent, MultipleEvent, and AllEvent parameters/logic; simplified logic; added a Session Info
                                                       text file containing details of the event

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -ConferenceList @('2015','2018');

  Downloads all MMS session content from 2015 and 2018 to C:\Conferences\MMS\

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -DownloadLocation "C:\Temp\MMS" -ConferenceId 2015

  Downloads all MMS session content from 2015 to C:\Temp\MMS\

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -All

  Downloads all MMS session content from all years to C:\Conferences\MMS\

.EXAMPLE
  .\Get-MMSSessionContent.ps1 -All -ExcludeSessionDetails;

  Downloads all MMS session content from all years to C:\Conferences\MMS\ BUT does not include a "Session Info.txt" file for each session containing the session details

.LINK
  Project URL - https://github.com/AndrewJNet/CopyMMSFiles
#>

[cmdletbinding(PositionalBinding=$false)]
Param(
  [Parameter(Mandatory=$false)][string]$DownloadLocation = "C:\Conferences\MMS", # could validate this: [ValidateScript({(Test-Path -Path (Split-Path $PSItem))})]
  [Parameter(Mandatory=$true,ParameterSetName='SingleEvent')]
  [ValidateSet("2015","2016","2017","2018","de2018","2019","jazz","miami")]
  [string]$ConferenceId,
  [Parameter(Mandatory=$true,ParameterSetName='MultipleEvents',HelpMessage="This needs to be a list or array of conference ids/years!")]
  [System.Collections.Generic.List[string]]$ConferenceList,
  [Parameter(Mandatory=$true,ParameterSetName='AllEvents')][switch]$All,
  [Parameter(Mandatory=$false)][switch]$ExcludeSessionDetails
)

## Make sure there aren't any trailing backslashes:
$DownloadLocation = $DownloadLocation.Trim('\');

## Setup
$PublicContentYears = @('2015', '2016', '2017');
$PrivateContentYears = @('2018', 'de2018', '2019', 'jazz', 'miami');
$ConferenceYears = New-Object -TypeName System.Collections.Generic.List[string];
[int]$PublicYearsCount = $PublicContentYears.Count;
[int]$PrivateYearsCount = $PrivateContentYears.Count;

if ($All) {
  for ($i = 0;$i -lt $PublicYearsCount;$i++)
  {
    $ConferenceYears.Add($PublicContentYears[$i]);
  }
  Remove-Variable -Name i -ErrorAction SilentlyContinue;
  for ($i = 0;$i -lt $PrivateYearsCount;$i++)
  {
    $ConferenceYears.Add($PrivateContentYears[$i]);
  }
  Remove-Variable -Name i -ErrorAction SilentlyContinue;
}
elseif ($PsCmdlet.ParameterSetName -eq 'SingleEvent')
{
  $ConferenceYears.Add($ConferenceId);
}
else
{
  $ConfListCount = $ConferenceList.Count;
  for ($i = 0;$i -lt $ConfListCount;$i++)
  {
    if ($ConferenceList[$i] -in ($PublicContentYears + $PrivateContentYears))
    {
      $ConferenceYears.Add($ConferenceList[$i]);
    }
    else
    {
      Write-Output "The Conference Id '$($ConferenceList[$i])' is not valid. Item will be skipped.";
    }
  }
  Remove-Variable -Name i -ErrorAction SilentlyContinue;
}

Write-Output "Base Download URL is $DownloadLocation"
Write-Output "Searching for content from these sessions: $([String]::Join(',',$ConferenceYears))"

## 
$ConferenceYears | ForEach-Object -Process {
  [string]$Year = $_;
  
  if ($Year -in $PrivateContentYears)
  {
    $creds = $host.UI.PromptForCredential('Sched Credentials', "Enter Credentials for the MMS Event: $Year", '', '');
  }

  $SchedBaseURL = "https://mms" + $Year + ".sched.com";
  $SchedLoginURL = $SchedBaseURL + "/login";
  Add-Type -AssemblyName System.Web;
  $web = Invoke-WebRequest $SchedLoginURL -SessionVariable mms;
  if ($creds)
  {
    $form = $web.Forms[1]
    $form.fields['username'] = $creds.UserName;
    $form.fields['password'] = $creds.GetNetworkCredential().Password;
  }

  $SessionDownloadPath = $DownloadLocation + '\mms' + $Year;
  Write-Output "Logging in to $SchedBaseURL";

  ## Connect to Sched
  if ($form)
  {
    $web = Invoke-WebRequest $SchedLoginURL -WebSession $mms -Method POST -Body $form.Fields;
  }
  else
  {
    $web = Invoke-WebRequest $SchedLoginURL -WebSession $mms;
  }
  ## Check if we connected (if required):
  if ((-Not ($web.InputFields.FindByName("login")) -and ($Year -in $PrivateContentYears)) -or ($Year -in $PublicContentYears)) {
    ##
    Write-Output "Downloaded content can be found in $SessionDownloadPath";

    $sched = Invoke-WebRequest -Uri $($SchedBaseURL + "/list/descriptions") -WebSession $mms;
    $links = $sched.Links;

    $eventsList = New-Object -TypeName System.Collections.Generic.List[int];

    $links | ForEach-Object -Process {
      if ($_.href -like "event/*")
      {
          [void]$eventsList.Add($links.IndexOf($_));
      }
    }

    [int]$eventCount = $eventsList.Count;

    for ($i = 0; $i -lt $eventCount; $i++)
    {
      [int]$linkIndex = $eventsList[$i];
      [int]$nextLinkIndex = $eventsList[$i+1];
      [string]$sessionInfoText = "";
      
      ## Get/Fix the Session Title:
      [string]$eventTitle = $links[$linkIndex].innerText;
      ## testing Write-output $eventTitle
      $sessionInfoText += "Session Title:`r`n$eventTitle`r`n`r`n";
      $eventTitle = $eventTitle -replace "[^A-Za-z0-9-_. ]", "";
      $eventTitle = $eventTitle.Trim();
      $eventTitle = $eventTitle -replace "\W+", "_";

      ## Set the download destination:
      $downloadPath = $SessionDownloadPath + '\' + $eventTitle;

      ## Get Session information if required:
      if (-Not $ExcludeSessionDetails)
      {
        $sessionLinkInfo = Invoke-WebRequest -Uri $($SchedBaseURL + "/" + $links[$linkIndex].href) -WebSession $mms;
        
        $sessionLinkInfo.ParsedHtml.getElementsByTagName("div") | ForEach-Object {
          if ($_.className -eq 'tip-description')
          {
            $sessionInfoText += "$($_.outerText)`r`n`r`n";
          }
          if ($_.className -eq 'tip-roles')
          {
            $sessionInfoText += "$($_.outerText)`r`n`r`n";
          }
        } # end of Foreach-Object processing the div tags to get the classes

        if ((Test-Path -Path $($downloadPath)) -eq $false) { New-Item -ItemType Directory -Force -Path $downloadPath | Out-Null }
        Out-File -FilePath "$downloadPath\Session Info.txt" -InputObject $sessionInfoText -Force -Encoding default;
      }
      
      ## Get/Fix/Download any files uploaded for the current session:
      $links[($linkIndex+1)..($nextLinkIndex-1)] | ForEach-Object -Process {
        if ($_.href -like "*hosted_files*")
        {
          $filename = Split-Path $_.href -Leaf;
          # Replace HTTP Encoding Characters (e.g. %20) with the proper equivalent.
          $filename = [System.Web.HttpUtility]::UrlDecode($filename);
          # Replace non-standard characters
          $filename = $filename -replace "[^A-Za-z0-9\.\-_ ]", "";
          
          $outputFilePath = $downloadPath + '\' + $filename;
          
          # Reduce Total Path to 255 characters.
          $outputFilePathLen = $outputFilePath.Length;
          if ($outputFilePathLen -ge 255) {
            $fileExt = [System.IO.Path]::GetExtension($outputFilePath)
            $newFileName = $outputFilePath.Substring(0, $($outputFilePathLen - $fileExt.Length))
            $newFileName = $newFileName.Substring(0, $(255 - $fileExt.Length)).trim()
            $newFileName = "$newFileName$fileExt"
            $outputFilePath = $newFileName
          }
          
          if ((Test-Path -Path $($downloadPath)) -eq $false) { New-Item -ItemType Directory -Force -Path $downloadPath | Out-Null }
          if ((Test-Path -Path $outputFilePath) -eq $false) {
            Write-Output "...attempting to download '$filename'";
            try
            {
              Invoke-WebRequest -Uri $_.href -OutFile $outputfilepath -WebSession $mms;
              Unblock-File $outputFilePath;
            }
            catch
            {
              Write-Output ".................$($PSItem.Exception) for '$($_.href)'...moving to next file...";
            }
          }
        }
      } # end of Foreach-Object: processing download files
    } # end of for loop processing every session
  } # end of if we successfully logged in if required
  else
  {
    Write-Output "Login to $SchedBaseUrl failed.";
  }
}
