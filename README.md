# CopyMMSFiles

These scripts are used to download the session files that were made available during [Midwest Management Summit 2015-2022](http://mmsmoa.com). If you are involved in configuration/infrastructure/identity/mobile device management primarily in the Microsoft space, consider attending this event!

## Usage

For content just from 2022 in a custom directory (default is C:\Conferences\MMS\$conferenceyear), use the following:

``` .\Get-MMSSessionContent.ps1 -DownloadLocation "C:\Temp\MMS" -ConferenceId 2022atmoa```

For multiple years:

``` .\Get-MMSSessionContent.ps1 -ConferenceList @('2015','2022atmoa')```

To exclude session details:

``` .\Get-MMSSessionContent.ps1 -All -ExcludeSessionDetails```

## Acknowledgements

Thank you to [Duncan Russell](http://www.sysadmintechnotes.com/) for providing the initial script for MMS 2014 and helping me test the changes I made for it to work with the more recent conferences.

Thank you to [Evan Yeung](https://github.com/forevanyeung) for cleaning up processing and file naming.

Thank you to [Chris Kibble](https://www.christopherkibble.com) for continued testing and improvements made to the script.

Thank you to [Benjamin Reynolds](https://sqlbenjamin.wordpress.com) for loads of great changes and additional testing. 

This script is provided as-is with no guarantees. It was tested with PowerShell 5.1 on Windows 10. It does not currently function in PowerShell 6/7.
