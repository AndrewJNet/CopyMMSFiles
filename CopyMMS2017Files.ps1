##############################################
#                                            #
# File:     Copymms2017Files.ps1             #
# Author:   Duncan Russell                   #
#           http://www.sysadmintechnotes.com #
# Edited:   Andrew Johnson                   #
#           http://www.andrewj.net           #
#           Evan Yeung                       #
#           http://www.forevanyeung.com      #
#                                            #
##############################################

$baseLocation = 'C:\Conferences\MMS'
Clear-Host
$MMSDates='2017-05-14', '2017-05-15', '2017-05-16','2017-05-17','2017-05-18'

$web = Invoke-WebRequest 'https://mms2017.sched.com/login' -SessionVariable mms
$c = $host.UI.PromptForCredential('Sched Credentials', 'Enter Credentials', '', '')
$form = $web.Forms[1]
$form.fields['username'] = $c.UserName
$form.fields['password'] = $c.GetNetworkCredential().Password
"Logging In..."
$web = Invoke-WebRequest -Uri 'https://mms2017.sched.com/login' -WebSession $mms -Method POST -Body $form.Fields

ForEach ($Date in $MMSDates) {
    "Checking day '{0}' for downloads" -f $Date
    
    $sched = Invoke-WebRequest -Uri $("https://mms2017.sched.org/" + $Date + "/list/descriptions") -WebSession $mms
    $links = $sched.Links

    $eventsIndex = @()
    $links | ForEach-Object { if(($_.href -like "*/event/*") -and ($_.innerText -notlike "here")) { 
        $eventsIndex += (, ($links.IndexOf($_), $_.innerText))
    } }

    $i = 0
    While($i -lt $eventsIndex.Count) {
        $eventTitle = $eventsIndex[$i][1]
        $eventTitle = $eventTitle -replace "\W+", "_"

        $links[$eventsIndex[$i][0]..$(if($i -eq $eventsIndex.Count - 1) {$links.Count-1} else {$eventsIndex[$i+1][0]})] | ForEach-Object { 
            if($_.href -like "*hosted_files*") { 
                $downloadPath = $baseLocation + '\mms2017\' + $Date + '\' + $eventTitle
                $filename = $_.href
                $filename = $filename.substring(39)
                $filename = $filename.Replace('%20',' ')
                $outputFilePath = $downloadPath + '\' + $filename
                if((Test-Path -Path $($downloadPath)) -eq $false) { New-Item -ItemType Directory -Force -Path $downloadPath | Out-Null }
                if((Test-Path -Path $outputFilePath) -eq $false)
                {
                    "...attempting to download '{0}'" -f $filename
                    Invoke-WebRequest -Uri $_.href -OutFile $outputfilepath -WebSession $mms
                    Unblock-File $outputFilePath
                    $stopit = $true
                }
            } 
        }
        $i++
    }
}
