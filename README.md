# Get-MMSSessionContent

These scripts are used to download the session files that were made available during [Midwest Management Summit 2015-2024](http://mmsmoa.com). If you are involved in configuration/infrastructure/identity/mobile device management primarily in the Microsoft space, consider attending this event!

## Usage

For content just from 2024 in a custom directory (default is C:\Conferences\MMS\$conferenceyear), use the following:

``` .\Get-MMSSessionContent.ps1 -DownloadLocation "C:\Temp\MMS" -ConferenceId 2024atmoa```

For multiple years:

``` .\Get-MMSSessionContent.ps1 -ConferenceList @('2015','2024atmoa')```

To exclude session details:

``` .\Get-MMSSessionContent.ps1 -All -ExcludeSessionDetails```

## Acknowledgements

Thank you to:
- [Duncan Russell](http://www.sysadmintechnotes.com/) for providing the initial script for MMS 2014 and helping me test the changes I made for it to work with the more recent conferences.
- [Evan Yeung](https://github.com/forevanyeung) for cleaning up processing and file naming.
- [Chris Kibble](https://www.christopherkibble.com) for continued testing and improvements made to the script.
- [Benjamin Reynolds](https://sqlbenjamin.wordpress.com) for loads of great changes and additional testing.
- [Nathan Ziehnert](https://z-nerd.com/) for adding PowerShell 7 support
- As well as edits by [Jon Warnken](https://github.com/mrbodean), [Oliver Baddeley](https://github.com/BaddMann), and [Jorge Suarez](https://github.com/jorgeasaurus)

This script is provided as-is with no guarantees. As of April 28, 2024, version 1.5 was tested with no errors using the following configurations:

- Windows 11 23H2 using Windows PowerShell 5.1
- Windows 11 23H2 using PowerShell 7.4.2
- Ubuntu using PowerShell 7.4.2
