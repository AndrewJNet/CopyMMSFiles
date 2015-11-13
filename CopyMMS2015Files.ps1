##############################################
#                                            #
# File:     CopyMMS2015Files.ps1             #
# Author:   Duncan Russell                   #
#           http://www.sysadmintechnotes.com #
# Edited:   Andrew Johnson                   #
#           http://www.andrewj.net           #
#                                            #
##############################################

$baseLocation = 'C:\Conferences\MMS'
Clear-Host
$MMSDates='http://mms2015.sched.org/2015-11-08','http://mms2015.sched.org/2015-11-09','http://mms2015.sched.org/2015-11-10','http://mms2015.sched.org/2015-11-11'
ForEach ($Date in $MMSDates)
{
$uri = $Date
$sched = Invoke-WebRequest -Uri $uri -WebSession $mms
$links = $sched.Links
$links | ForEach-Object {
    if(($PSItem.href -like '*event/*') -and ($PSItem.innerText -notlike '*birds*'))
    {
        $eventUrl = $PSItem.href
        $eventTitle = $($PSItem.innerText -replace "full$", "") -replace "filling$", "" 
        "Checking session '{0}' for downloads" -f $eventTitle
        $eventTitle = $eventTitle -replace "\W+", "_"
        
        $uri = 'http://mms2015.sched.org'
        $event = Invoke-WebRequest -Uri $($uri + $eventUrl)
        $eventLinks = $event.Links

        $eventLinks | ForEach-Object { 
            $eventFileUrl = $PSItem.href;$filename = $PSItem.href;if($eventFileUrl -like '*hosted_files*')
            {
                $downloadPath = $baseLocation + '\mms2015\' + $eventTitle
                $filename = $filename.substring(39)
                $outputFilePath = $downloadPath + '\' + $filename
                if((Test-Path -Path $($downloadPath)) -eq $false){New-Item -ItemType Directory -Force -Path $downloadPath | Out-Null}
                if((Test-Path -Path $outputFilePath) -eq $false)
                {
                    "...attempting to download '{0}'" -f $filename
                    Invoke-WebRequest -Uri $eventFileUrl -OutFile $outputfilepath -WebSession $mms;$doDownload=$false;
                    Unblock-File $outputFilePath
                    $stopit = $true
                }
            } 
        }
    }
}
}