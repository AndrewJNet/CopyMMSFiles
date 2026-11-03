# Get-MMSSessionContent_GUI

A Windows PowerShell graphical interface for discovering, filtering, and downloading Midwest Management Summit (MMS) session content from Sched.

Built on top of the excellent CopyMMSFiles project created by Andrew Johnson and contributors, this GUI edition simplifies MMS content discovery and download operations through an intuitive Windows interface.

---

## Features

- Automatic MMS conference discovery
- Conference accessibility validation
- Sched credential authentication
- Public-content access
- Conference filtering
- Session filtering
- Session selection with checkboxes
- Download all sessions for selected conferences
- Download individual sessions
- Download logging
- Excel export support
- CSV fallback reporting
- PowerShell 5.1 support
- PowerShell 7 support
- Automatic STA relaunch when required

---

## Installation

1. Download the latest release.
2. Open PowerShell.
3. Run:

```powershell
.\Get-MMSSessionContent_GUI.ps1
```

---

## Requirements

- Windows
- PowerShell 5.1 or later
- Internet access
- Valid Sched account for protected conference content

---

## Credits

This project is based on the CopyMMSFiles project.

Original Project:
https://github.com/AndrewJNet/CopyMMSFiles

Original Author:
Andrew Johnson

Additional credit goes to all contributors of the original project.

---

## Author

Keith Nail

My 02D Cents

https://www.my02dcents.com

---

## License

This script is provided strictly AS-IS, with no guarantees, warranties, or support commitments. 
The author assumes no responsibility for any issues, data loss, service interruptions, or compatibility changes resulting from its use. 
Users should thoroughly test the script in their own environment before relying on it. As of May 6, 2026, the script was tested successfully using the following configurations:

Windows 11 26H2 using PowerShell 7.6.6
