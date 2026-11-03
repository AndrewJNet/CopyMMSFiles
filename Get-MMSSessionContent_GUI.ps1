<#
.SYNOPSIS
    GUI for discovering, filtering, and downloading Midwest Management Summit (MMS)
    session content from Sched.

.DESCRIPTION
    This script provides a graphical user interface for discovering,
    selecting, filtering, and downloading Midwest Management Summit (MMS)
    session content from Sched.

    ## Startup and Conference Selection

    The main window opens without a credential popup or automatic network query.
    All 17 preloaded conferences are initially visible with an access status of
    Not queried. Conference checkboxes select query targets and filter the loaded
    session list. Select All includes visible conferences that are not yet queried.

    Query Selected refreshes only checked conferences, including checked past
    years. It retains the catalog and sessions previously loaded for other
    conferences. Refreshing a conference replaces its old sessions and selections.

    ## Authentication

    Public schedules may be queried without signing in, but all conference-file
    downloads require a completed Sign In. Enter a username and password and click
    Sign In. Authentication is verified against
    the first checked conference, or 2026atmoa when none is checked. A successful
    sign-in does not guarantee access to every conference. Before downloading, the
    application verifies authenticated access to every distinct target conference.
    If any check fails or falls back to public mode, the job stops before writing
    session metadata or attachments. Cached anonymous sessions cannot bypass this gate.

    Editing either credential field clears the active credential and web-session
    caches and resets the button to Sign In. Passwords are not written to logs
    or results; query messages may include the username. The masked password
    textbox retains its value in process memory.
    Both download buttons remain disabled until Sign In succeeds and are disabled
    again when either credential field changes. Schedule browsing remains available.

    ## Discovery and Manual Queries

    Discover IDs probes configured year/suffix combinations and appends valid
    conferences without loading their sessions or clearing existing results.
    Discovery uses the Show past years setting to choose its probing year range.

    Add and Query accepts the ID suffix only, such as 2027atmoa or midway, rather
    than a full URL. IDs must contain only letters and numbers. A new ID is appended
    to the catalog, checked, and immediately queried. Existing IDs are reused.
    Manually queried entries remain visible despite year/access visibility filters.

    ## Session Filtering and Downloads

    The session list shows loaded sessions from checked conferences, or all loaded
    sessions not hidden by the year setting when no conferences are checked.
    Text filtering matches session titles or conference IDs without regard to case.
    Filtering changes the display, not the underlying checked-session selections.

    Download SELECTED Sessions uses all checked session objects, including checked
    sessions currently outside the text filter. Download ALL Sessions of Checked
    Conferences uses all loaded sessions from checked, accessible, year-visible
    conferences and requests confirmation before starting.

    Files are grouped by conference and sanitized session title. Session Info.txt
    is written unless excluded. Existing files are skipped unless hash comparison
    is enabled; hash comparison downloads a temporary copy and replaces changed
    content. Only Sched hosted-file links are downloaded, not arbitrary linked sites.

    ## Logging, Reporting, and Cancellation

    The GUI displays operation status and a timestamped log. Reports use ImportExcel
    when available and fall back to UTF-8 CSV when Excel export is unavailable or
    fails. Reports and logs remain under C:\Scripts even if the download folder changes.

    Cancel stops the current operation, including an in-flight HTTP request, while
    leaving the window available for a new operation. Closing the window hides it
    immediately if busy, cancels pending requests, and closes it after cleanup.
    Incomplete new files and hash-check temporary files are removed; completed files
    are retained. Cancelled jobs do not begin result export. A report already being
    written finishes before cleanup. Window-close tracking is native WinForms code
    and does not invoke a PowerShell FormClosing callback on a stopped pipeline.

    ## Attribution

    This project is based on Get-MMSSessionContent.ps1 Version 1.7.4 from
    the CopyMMSFiles project and would not exist without the work of the
    original authors and contributors listed below.

    Original Project:
        Andrew Johnson
        https://github.com/AndrewJNet/CopyMMSFiles

    Original Contributors:

        Duncan Russell
        http://www.sysadmintechnotes.com

        Andrew Johnson
        https://github.com/AndrewJNet

        Evan Yeung
        https://www.forevanyeung.com

        Chris Kibble
        https://www.christopherkibble.com

        Jon Warnken
        https://www.mrbodean.net

        Oliver Baddeley
        Desert Edition Contributor

        Benjamin Reynolds
        https://sqlbenjamin.wordpress.com

        Jorge Suarez
        https://github.com/jorgeasaurus

        Nathan Ziehnert
        https://z-nerd.com

        Piotr Gardy
        https://garit.pro

    GUI Edition:

        Keith Nail
        Marion, North Carolina
        https://www.my02dcents.com

    GUI Enhancements Include:

        - Windows Forms graphical interface
        - On-demand MMS conference discovery
        - Conference accessibility validation
        - Conference and session filtering
        - Session selection with checkbox support
        - Bulk conference downloads
        - Individual session downloads
        - In-window sign-in with optional anonymous mode
        - Download logging
        - Results reporting
        - Excel export support with CSV fallback
        - Preloaded conference catalog with on-demand session inventory
        - Automatic STA relaunch when required

    Run with NO parameters:

        1. The GUI opens directly with the preloaded conference catalog.
        2. Query public schedules without signing in; complete Sign In before downloading files.
        3. Select one or more conferences and use Query Selected.
        4. Discover IDs finds additional conferences without querying their sessions.
        5. Manually queried conferences are appended without removing the catalog.
        6. Users can:
             - Download all sessions from selected conferences
             - Download selected sessions only
             - Review session metadata before downloading
        7. Logs are written during operations; download results are exported after a job.

    Requirements:

        - Windows
        - PowerShell 5.1 or PowerShell 7.x
        - Internet connectivity
        - Valid Sched credentials and conference permissions for protected content
        - Write permission to C:\Scripts and the selected download folder
        - Optional ImportExcel module for Excel reports; otherwise CSV is used

.INPUTS
    None. The script has no command-line parameters and does not accept pipeline input.
    Conference IDs, credentials, selections, and download options are entered in the GUI.

.OUTPUTS

    Session Content:
        <DownloadLocation>\mms<ConferenceId>\<SanitizedSessionTitle>\

    Optional Session Metadata:
        <DownloadLocation>\mms<ConferenceId>\<SanitizedSessionTitle>\Session Info.txt

    Default Download Location:
        C:\Scripts\MMSContent

    Log File:
        C:\Scripts\yyyy-MM-dd_HHmm-MMS-Download-Log.txt

    Results Export:
        C:\Scripts\yyyy-MM-dd_HHmm-MMS-Download-Results.xlsx

    CSV Fallback:
        C:\Scripts\yyyy-MM-dd_HHmm-MMS-Download-Results.csv

    Result Fields:
        Timestamp, Conference, Session, File, Status, Path

    Result Statuses:
        Downloaded, Updated, Skipped, Failed, Failed (session page),
        Failed (login), No files

    No result report is written when the download job has no result records.

.NOTES

    Script Name:           Get-MMSSessionContent_GUI.ps1
    Version:               2.1.2-GUI
    Author:                Keith Nail
    Website:               https://www.my02dcents.com
    Created:               2026-10-23
    Documentation Updated: 2026-11-02

    Based On:
        Get-MMSSessionContent.ps1 Version 1.7.4

    Original Project:
        https://github.com/AndrewJNet/CopyMMSFiles

    Purpose:
        Provide a graphical front-end for discovering,
        browsing, filtering, and downloading MMS session
        content without requiring command-line parameters.

        ## Section Map

        - Prerequisites and Assemblies: Windows/STA checks, assemblies, TLS, progress.
        - Global Settings and State: output locations, caches, selections, operation flags.
        - Conference Catalog: known IDs, conference metadata, year/access visibility.
        - Logging Functions: disk/GUI logging, status updates, result records.
        - Helper Functions: responsive waits, HTTP retry, HTML parsing, credentials.
        - Credential Prompt: retained legacy dialog; not used by the current launch path.
        - Authentication and Conference Query Functions: sign-in, schedules, discovery.
        - Download Functions: per-file handling, session metadata, download jobs.
        - Export Functions: timestamped Excel reports with automatic CSV fallback.
        - GUI State Functions: selection synchronization, filtering, busy-state controls.
        - Build GUI: credentials, conference/session lists, downloads, status, log.
        - Event Handlers: user actions, visibility changes, cancellation, startup display.
        - Launch GUI: modal application loop and form disposal.

        ## Operational Notes

        - Windows Forms requires Windows and a single-threaded apartment (STA).
        - Queries and downloads run synchronously while selected helpers process UI events.
        - HTTP 429 responses receive one delayed retry; other request errors propagate.
        - Sched HTML changes can affect login detection, schedule parsing, and hosted links.
        - Known Private conferences require credentials; public/discovered sites may fall
            back to anonymous access if their login attempt is not accepted.
        - The retained legacy dialog is not invoked by startup or by the Sign In button.

    Version History

    Date          Version       Author          Notes
    ----------    -----------   -------------   -----------------------------------
    2026-10-23    2.1.1-GUI     Keith Nail      Initial GUI release based on
                                                Get-MMSSessionContent.ps1 v1.7.4

    2026-10-23    2.1.2-GUI     Keith Nail      Updated GUI release with improved functionality
                                                and bug fixes. Not to initially query past years of MMS conferences, but to allow the user to select the year of interest.

    2026-11-02    2.1.2-GUI     Keith Nail      Current workflow documentation: direct GUI
                                                launch, selected queries, in-window sign-in,
                                                manual append, and on-demand discovery.

.EXAMPLE

    .\Get-MMSSessionContent_GUI.ps1

    Launches the GUI with the conference catalog and in-window sign-in controls.

.EXAMPLE

    powershell.exe -STA -ExecutionPolicy Bypass `
        -File .\Get-MMSSessionContent_GUI.ps1

    Launches the GUI directly in Windows PowerShell with an STA thread.

.EXAMPLE

    pwsh.exe -STA -NoProfile -File .\Get-MMSSessionContent_GUI.ps1

    Launches the GUI using PowerShell 7 on Windows. Select conferences in the
    main window, then click Query Selected to retrieve their sessions.

.LINK

    GUI Project Author
    https://www.my02dcents.com

.LINK

    Original CopyMMSFiles Project
    https://github.com/AndrewJNet/CopyMMSFiles

.LINK

    MMS Conference Website
    https://mmsmoa.sched.com

#>


# ==================================================
# MARK: Prerequisites and Assemblies
# ==================================================
<#
## Prerequisites and Assemblies
### Purpose
Prepare a Windows Forms-compatible execution environment before creating controls.
### Execution Flow
- Reject an explicitly non-Windows platform before loading Windows-only assemblies.
- Relaunch the saved script with `-STA` when the current thread is not STA, then return.
- Load Windows Forms, Drawing, and Web utilities; enable native visual styles.
- Compile the native MmsContentForm lifecycle type once, using the host's assembly paths.
- Attempt to add TLS 1.2 support and suppress console web-request progress output.
### Dependencies and Limitations
Requires Windows and a saved script path for automatic relaunch. The relaunch preserves
the host executable; it does not authenticate, discover conferences, or show a sign-in popup.
MmsContentForm records ShutdownRequested without a PowerShell event callback. When busy,
it hides and defers closing until CompleteOperation; otherwise it closes immediately.
#>
if ($PSVersionTable.PSObject.Properties.Name -contains 'Platform' -and $PSVersionTable.Platform -and $PSVersionTable.Platform -ne 'Win32NT') {
    Write-Host 'This GUI requires Windows (Windows Forms).' -ForegroundColor Red
    return
}

# Windows Forms requires STA; relaunch the saved script in the same host executable if needed.
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    if ($PSCommandPath) {
        $HostExe = (Get-Process -Id $PID).Path
        Start-Process -FilePath $HostExe -ArgumentList @('-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    }
    else {
        Write-Host 'Please run this script from a saved .ps1 file or start PowerShell with the -STA switch.' -ForegroundColor Red
    }
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web
if (-not ('MmsContentForm' -as [type])) {
    $FormReferences = @(
        [System.Windows.Forms.Form].Assembly.Location
        [System.ComponentModel.Component].Assembly.Location
        [System.ComponentModel.CancelEventArgs].Assembly.Location
    ) | Select-Object -Unique
    Add-Type -ReferencedAssemblies $FormReferences -TypeDefinition @'
public class MmsContentForm : System.Windows.Forms.Form
{
    public bool ShutdownRequested { get; private set; }
    public bool OperationActive { get; set; }

    protected override void OnFormClosing(System.Windows.Forms.FormClosingEventArgs CloseArgs)
    {
        ShutdownRequested = true;
        if (OperationActive)
        {
            CloseArgs.Cancel = true;
            Hide();
        }
        base.OnFormClosing(CloseArgs);
    }

    public void CompleteOperation()
    {
        OperationActive = false;
        if (ShutdownRequested && !IsDisposed)
        {
            Close();
        }
    }
}
'@
}
[System.Windows.Forms.Application]::EnableVisualStyles()

try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12 } catch { }

# Suppress console HTTP progress so the GUI's progress bar is the primary progress display.
$ProgressPreference = 'SilentlyContinue'
# ==================================================
# ^^^^^ Prerequisites and Assemblies ^^^^^
# ==================================================

# ==================================================
# MARK: Global Settings and State
# ==================================================
<#
## Global Settings and State
### Purpose
Initialize output defaults and the script-scoped data shared by functions and GUI handlers.
### Output Locations
- `OutputRoot`: `C:\Scripts`, created if missing; receives logs and download-result reports.
- `DefaultDownloadLocation`: `C:\Scripts\MMSContent`; may be overridden in the GUI.
- `RunStamp` and `LogFile`: one timestamped log filename for this application run.
### Data and Operation State
- `Sessions` stores per-conference HTTP sessions; `EventCache` stores parsed schedule entries.
- `SessionList` stores loaded session objects and their download checkbox states.
- `Results` stores the current download job's report records; each new job clears it.
- `Credential`, `CancelRequested`, `IsBusy`, and `SuppressEvents` control authentication,
  cooperative cancellation, overlapping operations, and programmatic checkbox updates.
All state is in memory for this run; previously exported reports are not imported at launch.
#>
$Script:OutputRoot = 'C:\Scripts'
$Script:DefaultDownloadLocation = Join-Path -Path $Script:OutputRoot -ChildPath 'MMSContent'
$Script:RunStamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$Script:LogFile = Join-Path -Path $Script:OutputRoot -ChildPath ($Script:RunStamp + '-MMS-Download-Log.txt')

if (-not (Test-Path -Path $Script:OutputRoot)) {
    New-Item -ItemType Directory -Force -Path $Script:OutputRoot | Out-Null
}

$Script:RegexIgnoreCase = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
$Script:NewLine = [System.Environment]::NewLine

$Script:Sessions = @{}          # Conference ID -> authenticated or anonymous HTTP session.
$Script:EventCache = @{}        # Conference ID -> cached parsed schedule/session objects.
$Script:SessionList = New-Object -TypeName System.Collections.Generic.List[object]   # Loaded sessions and persistent checkbox states.
$Script:Results = New-Object -TypeName System.Collections.Generic.List[object]       # Result records for the current download job.
$Script:Credential = $null
$Script:CancelRequested = $false
$Script:IsBusy = $false
$Script:SuppressEvents = $false
# ==================================================
# ^^^^^ Global Settings and State ^^^^^
# ==================================================

# ==================================================
# MARK: Conference Catalog
# ==================================================
<#
## Conference Catalog
### Purpose
Maintain an ordered conference inventory independently of whether its schedule has been queried.
### Entry Fields
| Field | Meaning |
| --- | --- |
| `Id`, `Name` | Sched ID suffix and display name; the host is `mms<Id>.sched.com`. |
| `Type`, `Year` | Public/Private/Discovered classification and extracted four-digit year. |
| `Access`, `SessionCount` | Last query outcome and number of loaded sessions. |
| `Checked`, `Queried` | Current selection and whether an access query was attempted. |
| `Manual` | Visibility override for an entry queried through Add and Query. |
### Ordering and Visibility
The 17 known entries are added in source order. New discoveries and manual IDs append to
the same list; lookup is case-insensitive. Legacy IDs without a year return year zero and
are normally considered past. Manual entries bypass the past-year visibility restriction.
Discovery combines configured suffixes with years beginning at `ProbeStartYear` when enabled
by the chosen discovery scope; preloaded entries are not proof that a site is still available.
#>
# Extract the first four-digit 20xx year; return zero for IDs with no embedded year.
function Get-ConferenceYear {
    param([Parameter(Mandatory = $true)][string]$ConfId)

    $YearMatch = [System.Text.RegularExpressions.Regex]::Match($ConfId, '20\d{2}')
    if ($YearMatch.Success) { return [int]$YearMatch.Value }
    return 0
}

# Create a consistent catalog record with no loaded sessions and no initial selection.
function New-ConferenceEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Type
    )

    return [PSCustomObject]@{
        Id           = $Id
        Name         = $Name
        Type         = $Type          # Public / Private / Discovered; controls anonymous fallback policy.
        Year         = Get-ConferenceYear -ConfId $Id
        Access       = 'Not queried'
        SessionCount = 0
        Checked      = $false
        Queried      = $false
        Manual       = $false
    }
}

# Preload known conference IDs and access classifications from the original public/private lists.
$Script:ConferenceCatalog = New-Object -TypeName System.Collections.Generic.List[object]
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2015'      -Name 'MMS 2015'                        -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2016'      -Name 'MMS 2016'                        -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2017'      -Name 'MMS 2017'                        -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2018'      -Name 'MMS 2018'                        -Type 'Private'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id 'de2018'    -Name 'MMS 2018 Desert Edition'         -Type 'Private'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2019'      -Name 'MMS 2019'                        -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id 'jazz'      -Name 'MMS Jazz Edition'                -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id 'miami'     -Name 'MMS Miami Edition'               -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2022atmoa' -Name 'MMS 2022 at MOA'                 -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2023atmoa' -Name 'MMS 2023 at MOA'                 -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2023miami' -Name 'MMS 2023 Miami Edition'          -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2024atmoa' -Name 'MMS 2024 at MOA'                 -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2024fll'   -Name 'MMS 2024 Flamingo Edition (FLL)' -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2025atmoa' -Name 'MMS 2025 at MOA'                 -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2025music' -Name 'MMS 2025 Music City Edition'     -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2026atmoa' -Name 'MMS 2026 at MOA'                 -Type 'Private'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id 'midway'    -Name 'MMS 2026 Midway Edition'         -Type 'Public'))


# Discovery candidates combine these suffixes with the requested probing year range.
$Script:ProbeSuffixes = @('atmoa', 'miami', 'fll', 'music')
$Script:ProbeStartYear = 2022

# Return the existing catalog object by case-insensitive ID, or null when it is unknown.
function Get-ConferenceInfo {
    param([Parameter(Mandatory = $true)][string]$ConfId)

    foreach ($Conf in $Script:ConferenceCatalog) {
        if ($Conf.Id -ieq $ConfId) { return $Conf }
    }
    return $null
}

# Compare the stored conference year with the current local year; year-zero legacy entries are past.
function Test-IsPastConference {
    param([Parameter(Mandatory = $true)][object]$Conf)

    return ([int]$Conf.Year -lt (Get-Date).Year)
}

# Apply Show past years; manually queried entries remain visible regardless of their year.
function Test-IsHiddenConference {
    param([Parameter(Mandatory = $true)][object]$Conf)

    if ($Conf.Manual) { return $false }
    if ($Script:ChkPast -and $Script:ChkPast.Checked) { return $false }
    return (Test-IsPastConference -Conf $Conf)
}
# ==================================================
# ^^^^^ Conference Catalog ^^^^^
# ==================================================

# ==================================================
# MARK: Logging Functions
# ==================================================
<#
## Logging Functions
### Purpose
Keep disk logging, the GUI activity log, the status label, and result records consistent.
### Function Responsibilities
- `Write-Log`: prefix messages with `yyyy-MM-dd_HHmmss` and a level, append to the log
    file, and mirror messages to the GUI log when that control exists.
- `Set-Status`: replace the current status-label text without creating a result record.
- `Add-DownloadResult`: append Timestamp/Conference/Session/File/Status/Path fields
    for a file outcome or a session-level failure/no-files outcome.
### UI and Failure Behavior
Logging and status updates process pending UI events. Disk log-write failures are swallowed
so an unavailable log destination does not itself abort the requested operation. Result
records are exported later by the download job, not by these helpers.
#>
# Append a leveled timestamped message to disk and the GUI, then process pending UI events.
function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [string]$Level = 'INFO'
    )

    if ((Test-CancellationRequested) -and $Level -eq 'ERROR') { return }
    $Line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'), $Level, $Message
    try { Add-Content -Path $Script:LogFile -Value $Line -ErrorAction Stop } catch { }

    if ($Script:TxtLog -and -not $Script:TxtLog.IsDisposed) {
        $Script:TxtLog.AppendText($Line + $Script:NewLine)
    }
    Update-UiEvents
}

# Update the single current-status label when available; this does not append to the log.
function Set-Status {
    param([string]$Text)

    if ($Script:Form -and ($Script:Form.IsDisposed -or $Script:Form.ShutdownRequested)) { return }
    if ($Script:LblStatus -and -not $Script:LblStatus.IsDisposed) {
        $Script:LblStatus.Text = $Text
    }
    Update-UiEvents
}

# Record one file/session outcome using the fixed report schema consumed by the export function.
function Add-DownloadResult {
    param(
        [string]$ConfId,
        [string]$SessionTitle,
        [string]$FileName,
        [string]$Status,
        [string]$Path
    )

    $Script:Results.Add([PSCustomObject]@{
            Timestamp  = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            Conference = $ConfId
            Session    = $SessionTitle
            File       = $FileName
            Status     = $Status
            Path       = $Path
        })
}
# ==================================================
# ^^^^^ Logging Functions ^^^^^
# ==================================================

# ==================================================
# MARK: Helper Functions
# ==================================================
<#
## Helper Functions
### Purpose
Provide shared request handling, HTML normalization, session-link extraction, and credential creation.
### HTTP and Responsiveness
`Wait-WithDoEvents` waits in short intervals and observes cooperative cancellation.
`Get-ErrorStatusCode` extracts an HTTP response status when available, otherwise zero.
`Update-UiEvents` processes messages only while the form is open and maps native shutdown
state to the cancellation flag. `Test-CancellationRequested` checks both operation and
native form state. `Invoke-CancellableWebRequest` runs HTTP work in an isolated PowerShell
runspace, processes GUI events while waiting, stops pending work on cancellation, and
disposes its pipeline in finally. Request parameters and cookie sessions stay in process.
`Invoke-SchedRequest` forwards options with a 30-second timeout and retries HTTP 429 once
after a 20-second wait; cancellation prevents that retry. Discovery probes use 15 seconds.
### Parsing
`Invoke-BasicHTMLParser` converts selected line breaks, links, speaker blocks, and lists
to readable metadata text before stripping remaining tags. `Get-EventLinkPairs` combines
parsed anchors with raw HTML/title fallbacks and deduplicates session URLs. These parsers
target Sched markup and are not general-purpose HTML parsers.
### Credentials
`Get-CurrentCredential` returns null for incomplete GUI credentials or constructs a
PSCredential from the current fields. Construction alone does not authenticate to Sched.
#>
# Process pending UI messages without invoking controls after native shutdown begins.
function Update-UiEvents {
    if ($Script:Form -and -not $Script:Form.IsDisposed -and -not $Script:Form.ShutdownRequested) {
        [System.Windows.Forms.Application]::DoEvents()
    }
    if ($Script:Form -and ($Script:Form.IsDisposed -or $Script:Form.ShutdownRequested)) {
        $Script:CancelRequested = $true
    }
}

# Treat window shutdown/disposal as cancellation even before the script flag is synchronized.
function Test-CancellationRequested {
    return ($Script:CancelRequested -or ($Script:Form -and ($Script:Form.IsDisposed -or $Script:Form.ShutdownRequested)))
}

# Keep the UI responsive during a bounded wait; stop waiting when cancellation is requested.
function Wait-WithDoEvents {
    param([int]$Seconds = 1)

    $EndTime = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $EndTime -and -not (Test-CancellationRequested)) {
        Start-Sleep -Milliseconds 200
        Update-UiEvents
    }
}

# Read an HTTP status from the exception response when supported by the current host.
function Get-ErrorStatusCode {
    param($ErrorRecord)

    try {
        if ($ErrorRecord.Exception.Response) { return [int]$ErrorRecord.Exception.Response.StatusCode }
    }
    catch { }
    return 0
}

# Treat null, empty strings, and whitespace-only values as blank.
function Test-IsBlank {
    param($Value)
    return ([System.String]::IsNullOrWhiteSpace([string]$Value))
}

# Run HTTP work off the UI thread and stop/dispose the worker when cancellation is requested.
function Invoke-CancellableWebRequest {
    param([Parameter(Mandatory = $true)][hashtable]$RequestParameters)

    if (Test-CancellationRequested) { throw [System.OperationCanceledException]::new('Operation cancelled.') }
    $RequestPipeline = [System.Management.Automation.PowerShell]::Create()
    $PendingRequest = $null
    try {
        [void]$RequestPipeline.AddScript('param($RequestParameters) $ProgressPreference = "SilentlyContinue"; Invoke-WebRequest @RequestParameters').AddArgument($RequestParameters)
        $PendingRequest = $RequestPipeline.BeginInvoke()
        while (-not $PendingRequest.IsCompleted) {
            Update-UiEvents
            if (Test-CancellationRequested) {
                $RequestPipeline.Stop()
                throw [System.OperationCanceledException]::new('Operation cancelled.')
            }
            [void]$PendingRequest.AsyncWaitHandle.WaitOne(50)
        }
        if (Test-CancellationRequested) { throw [System.OperationCanceledException]::new('Operation cancelled.') }
        try {
            return $RequestPipeline.EndInvoke($PendingRequest)
        }
        catch {
            if ($RequestPipeline.Streams.Error.Count -gt 0) { throw $RequestPipeline.Streams.Error[0] }
            throw
        }
    }
    finally {
        if ($PendingRequest -and -not $PendingRequest.IsCompleted) { $RequestPipeline.Stop() }
        $RequestPipeline.Dispose()
    }
}

# Centralize Sched options and the single delayed HTTP 429 retry, respecting cancellation.
function Invoke-SchedRequest {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)]$WebSession,
        [string]$Method = 'GET',
        [string]$Body,
        [hashtable]$Headers,
        [string]$OutFile
    )

    $Params = @{
        Uri             = $Uri
        WebSession      = $WebSession
        Method          = $Method
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
        TimeoutSec      = 30
    }
    if ($Body) { $Params['Body'] = $Body }
    if ($Headers) { $Params['Headers'] = $Headers }
    if ($OutFile) { $Params['OutFile'] = $OutFile }

    try {
        return (Invoke-CancellableWebRequest -RequestParameters $Params)
    }
    catch {
        $Err = $_
        if (Test-CancellationRequested) { throw [System.OperationCanceledException]::new('Operation cancelled.') }
        if ((Get-ErrorStatusCode -ErrorRecord $Err) -eq 429) {
            Write-Log -Message 'HTTP 429 (Too Many Requests) received. Waiting 20 seconds, then retrying once.' -Level 'WARN'
            Wait-WithDoEvents -Seconds 20
            if (Test-CancellationRequested) { throw [System.OperationCanceledException]::new('Operation cancelled.') }
            return (Invoke-CancellableWebRequest -RequestParameters $Params)
        }
        throw
    }
}

# Convert supported Sched description markup into readable text with Markdown-style links/lists.
function Invoke-BasicHTMLParser {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Html,
        [Parameter(Mandatory = $true)][string]$SchedBaseURL
    )

    $Html = $Html.Replace('<br>', "`r`n").Replace('<br/>', "`r`n").Replace('<br />', "`r`n") # Normalize common HTML line-break variants.

    # Preserve separation between speaker/session blocks in the exported metadata text.
    $Html = $Html.Replace('<div class="sched-person-session">', "`r`n`r`n")

    # Render anchors as [label](URL), resolving relative links against the conference site.
    $LinkRegex = '(?<texttoreplace><a.*?href="(?<link>.*?)".*?>(?<content>.*?)<\/a>)'
    $LinkMatches = [System.Text.RegularExpressions.Regex]::Matches($Html, $LinkRegex, $Script:RegexIgnoreCase)
    foreach ($LinkMatch in $LinkMatches) {
        $LinkTarget = $LinkMatch.Groups['link'].Value
        if (-not $LinkTarget.StartsWith('http')) { $LinkValue = $SchedBaseURL + '/' + $LinkTarget } else { $LinkValue = $LinkTarget }
        $OpenSquare = [System.Char]91
        $CloseSquare = [System.Char]93
        $ReplaceWith = ' ' + $OpenSquare + $LinkMatch.Groups['content'].Value + $CloseSquare + '(' + $LinkValue + ')'
        $Html = $Html.Replace($LinkMatch.Groups['texttoreplace'].Value, $ReplaceWith)
    }

    # Convert supported unordered-list markup to line-separated bullet text.
    $ListRegex = '(?<texttoreplace><ul[^>]?>(?<content>.*?)<\/ul>)'
    $ListMatches = [System.Text.RegularExpressions.Regex]::Matches($Html, $ListRegex, $Script:RegexIgnoreCase)
    foreach ($ListMatch in $ListMatches) {
        $ListContent = $ListMatch.Groups['content'].Value.Replace('<li>', "`r`n* ").Replace('</li>', '')
        $Html = $Html.Replace($ListMatch.Groups['texttoreplace'].Value, $ListContent)
    }

    # Remove the arrow entity, preserve div breaks, and strip remaining HTML tags.
    $Html = $Html.Replace('&rarr;', '')
    $Html = $Html -replace '<div[^>]+>', "`r`n"
    $Html = $Html -replace '<[^>]+>', '' # Leave text after the supported transformations.

    return $Html
}

# Build unique event URL/title pairs, recovering titles from HTML or URL slugs when needed.
function Get-EventLinkPairs {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SchedulePage
    )

    $EventPairs = New-Object -TypeName System.Collections.Generic.List[object]
    $EventTitleByUrl = @{}

    if ($SchedulePage.Content) {
        $EventRegex = '<a[^>]+href="(?<url>event\/[^"#?]+)"[^>]*>(?<title>.*?)<\/a>'
        $FlatContent = $SchedulePage.Content.Replace("`r", '').Replace("`n", '')
        $EventMatches = [System.Text.RegularExpressions.Regex]::Matches($FlatContent, $EventRegex, $Script:RegexIgnoreCase)
        foreach ($EventMatch in $EventMatches) {
            $EventUrlValue = ([string]$EventMatch.Groups['url'].Value).Trim()
            $EventTitleValue = ($EventMatch.Groups['title'].Value -replace '<[^>]+>', '').Trim()
            if (-not (Test-IsBlank $EventUrlValue) -and -not (Test-IsBlank $EventTitleValue) -and -not $EventTitleByUrl.ContainsKey($EventUrlValue)) {
                $EventTitleByUrl[$EventUrlValue] = $EventTitleValue
            }
        }
    }

    # Prefer the response's parsed Links collection, supplementing missing titles from raw HTML.
    if ($SchedulePage.Links) {
        foreach ($Lnk in $SchedulePage.Links) {
            if ($Lnk.href -like 'event/*') {
                $LnkUrl = [string]$Lnk.href
                $LnkTitle = [string]$Lnk.innerText

                if ((Test-IsBlank $LnkTitle) -and $Lnk.outerHTML) {
                    $LnkTitle = ([string]$Lnk.outerHTML -replace '<[^>]+>', '').Trim()
                }

                if ((Test-IsBlank $LnkTitle) -and $EventTitleByUrl.ContainsKey($LnkUrl)) {
                    $LnkTitle = [string]$EventTitleByUrl[$LnkUrl]
                }

                if ((Test-IsBlank $LnkTitle) -and $LnkUrl -match '^event\/[^\/]+\/(?<slug>[^\/?#]+)') {
                    $LnkTitle = [System.Uri]::UnescapeDataString($Matches['slug']) -replace '-', ' '
                }

                $EventPairs.Add([PSCustomObject]@{
                        Url   = $LnkUrl
                        Title = $LnkTitle
                    })
            }
        }
    }

    # Use raw HTML event anchors when the parsed Links collection yielded no event pairs.
    if ($EventPairs.Count -eq 0 -and $SchedulePage.Content) {
        foreach ($Entry in $EventTitleByUrl.GetEnumerator()) {
            $EventPairs.Add([PSCustomObject]@{
                    Url   = [string]$Entry.Key
                    Title = [string]$Entry.Value
                })
        }
    }

    # Deduplicate URLs in their collected order; the HTML-map fallback has no guaranteed source order.
    $Seen = @{}
    $UniquePairs = New-Object -TypeName System.Collections.Generic.List[object]
    foreach ($Pair in $EventPairs) {
        if (Test-IsBlank $Pair.Url) { continue }
        if (-not $Seen.ContainsKey($Pair.Url)) {
            $Seen[$Pair.Url] = $true
            $UniquePairs.Add($Pair)
        }
    }

    return $UniquePairs
}

# Read the main-window fields into a PSCredential without performing a sign-in request.
function Get-CurrentCredential {
    if ((Test-IsBlank $Script:TxtUser.Text) -or [System.String]::IsNullOrEmpty($Script:TxtPass.Text)) { return $null }
    $SecurePass = ConvertTo-SecureString -String $Script:TxtPass.Text -AsPlainText -Force
    return (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Script:TxtUser.Text.Trim(), $SecurePass)
}
# ==================================================
# ^^^^^ Helper Functions ^^^^^
# ==================================================

# ==================================================
# MARK: Credential Prompt
# ==================================================
<#
## Credential Prompt
### Purpose and Current Status
Retain the legacy modal credential dialog without changing its implementation. The current
application launch and Sign In button do not call `Show-CredentialPrompt`; authentication
now occurs from the username/password controls in the main window.
### Legacy Contract
If explicitly called, the dialog returns a PSCredential when its OK action is selected
with complete fields, or null when skipped/closed or fields are incomplete. It disposes
its own form before returning and does not itself authenticate or query conferences.
### Maintenance Note
The dialog's historical captions describe the older workflow. They are retained as
legacy UI text, not documentation of the current application startup behavior.
#>
# Retained legacy dialog: returns entered credentials or null; not invoked at startup.
function Show-CredentialPrompt {
    $Dlg = New-Object -TypeName System.Windows.Forms.Form
    $Dlg.Text = 'MMS Sched Sign-in'
    $Dlg.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 430, 250
    $Dlg.StartPosition = 'CenterScreen'
    $Dlg.FormBorderStyle = 'FixedDialog'
    $Dlg.MaximizeBox = $false
    $Dlg.MinimizeBox = $false
    $Dlg.TopMost = $true
    $Dlg.Font = New-Object -TypeName System.Drawing.Font -ArgumentList 'Segoe UI', 9

    $LblInfo = New-Object -TypeName System.Windows.Forms.Label
    $LblInfo.Text = 'Enter your Sched credentials. The script will query current and upcoming MMS conferences and list the ones you can access. Choose Skip for public content only.'
    $LblInfo.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 15, 12
    $LblInfo.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 390, 50
    $Dlg.Controls.Add($LblInfo)

    $LblDlgUser = New-Object -TypeName System.Windows.Forms.Label
    $LblDlgUser.Text = 'Username:'
    $LblDlgUser.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 15, 75
    $LblDlgUser.AutoSize = $true
    $Dlg.Controls.Add($LblDlgUser)

    $TxtDlgUser = New-Object -TypeName System.Windows.Forms.TextBox
    $TxtDlgUser.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 95, 72
    $TxtDlgUser.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 300, 24
    $Dlg.Controls.Add($TxtDlgUser)

    $LblDlgPass = New-Object -TypeName System.Windows.Forms.Label
    $LblDlgPass.Text = 'Password:'
    $LblDlgPass.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 15, 110
    $LblDlgPass.AutoSize = $true
    $Dlg.Controls.Add($LblDlgPass)

    $TxtDlgPass = New-Object -TypeName System.Windows.Forms.TextBox
    $TxtDlgPass.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 95, 107
    $TxtDlgPass.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 300, 24
    $TxtDlgPass.UseSystemPasswordChar = $true
    $Dlg.Controls.Add($TxtDlgPass)

    $BtnDlgOk = New-Object -TypeName System.Windows.Forms.Button
    $BtnDlgOk.Text = 'Sign in and Query'
    $BtnDlgOk.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 95, 150
    $BtnDlgOk.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 140, 32
    $BtnDlgOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Dlg.Controls.Add($BtnDlgOk)

    $BtnDlgSkip = New-Object -TypeName System.Windows.Forms.Button
    $BtnDlgSkip.Text = 'Skip (Public Only)'
    $BtnDlgSkip.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 245, 150
    $BtnDlgSkip.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 150, 32
    $BtnDlgSkip.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
    $Dlg.Controls.Add($BtnDlgSkip)

    $Dlg.AcceptButton = $BtnDlgOk
    $Dlg.CancelButton = $BtnDlgSkip
    $Dlg.Add_Shown({ $TxtDlgUser.Focus() | Out-Null })

    $DlgResult = $Dlg.ShowDialog()
    $ReturnCred = $null
    if ($DlgResult -eq [System.Windows.Forms.DialogResult]::OK -and -not (Test-IsBlank $TxtDlgUser.Text) -and -not [System.String]::IsNullOrEmpty($TxtDlgPass.Text)) {
        $SecurePass = ConvertTo-SecureString -String $TxtDlgPass.Text -AsPlainText -Force
        $ReturnCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $TxtDlgUser.Text.Trim(), $SecurePass
    }
    $Dlg.Dispose()
    return $ReturnCred
}
# ==================================================
# ^^^^^ Credential Prompt ^^^^^
# ==================================================

# ==================================================
# MARK: Authentication and Conference Query Functions
# ==================================================
<#
## Authentication and Conference Query Functions
### Purpose
Manage conference-specific HTTP sessions, access checks, schedule caches, and discovery scopes.
### Authentication and Schedule Loading
`Connect-SchedConference` returns Success/Mode/Message and stores successful HTTP sessions.
Known Private entries reject anonymous access. A rejected login on a non-private entry may
fall back to public mode; the Sign In handler accepts only Authenticated mode as verification.
`Get-ConferenceEvents` reads `/list/descriptions`, normalizes titles, omits configured
activity-title patterns, and caches session objects. `Test-ConferenceAccess` refreshes one
conference's HTTP/cache/session data and updates its access status and session count.
### Discovery and Query Modes
| Mode | Targets and State Handling |
| --- | --- |
| `SelectedOnly` | Query checked entries across years; retain other results; do not probe. |
| `DiscoverOnly` | Probe unknown year/suffix IDs and append valid entries; do not load sessions. |
| `PastOnly` | Probe older candidates and query older entries not yet queried; retain loaded results. |
| No switch | Reset loaded sessions/query statuses, probe, then query the year scope. |
The GUI uses SelectedOnly and DiscoverOnly; the remaining modes are retained function paths.
`Test-ConferenceExists` verifies the final response host with a 15-second probe timeout.
`Confirm-ConferenceSession` requires an active verified credential and a fresh Authenticated
connection for download access; cached sessions and public fallback are not sufficient.
### Failure and Cancellation
Access outcomes are stored per conference; request failures are logged. Discovery/query loops
observe cancellation between candidates/conferences and restore busy-state controls in finally.
#>
# Open a conference HTTP session and return its success, access mode, and diagnostic message.
# Supplied credentials trigger login; non-private sites may fall back to anonymous access.
function Connect-SchedConference {
    param(
        [Parameter(Mandatory = $true)][string]$ConfId,
        [Parameter(Mandatory = $true)][string]$Type,
        [System.Management.Automation.PSCredential]$Credential
    )

    $SchedBaseURL = 'https://mms' + $ConfId + '.sched.com'
    $SchedLoginURL = $SchedBaseURL + '/login'
    $LoginFormPattern = '<input[^>]+name=[''"]login[''"]'

    try {
        $WebSess = New-Object -TypeName Microsoft.PowerShell.Commands.WebRequestSession
        $Web = Invoke-SchedRequest -Uri $SchedLoginURL -WebSession $WebSess

        if (-not $Credential) {
            if ($Type -eq 'Private') {
                return [PSCustomObject]@{ Success = $false; Mode = 'Failed'; Message = 'Login required' }
            }
            $Script:Sessions[$ConfId] = $WebSess
            return [PSCustomObject]@{ Success = $true; Mode = 'Public'; Message = 'Anonymous session opened' }
        }

        $UserName = $Credential.UserName
        $Password = $Credential.GetNetworkCredential().Password

        # URL-encode the original script's Sched login fields and release temporary password text.
        $Body = 'landing_conf=' + [System.Uri]::EscapeDataString($SchedBaseURL) + '&username=' + [System.Uri]::EscapeDataString($UserName) + '&password=' + [System.Uri]::EscapeDataString($Password) + '&login='
        $Password = $null

        $Web = Invoke-SchedRequest -Uri $SchedLoginURL -WebSession $WebSess -Method 'POST' -Body $Body -Headers @{ Referer = $SchedLoginURL; Origin = $SchedBaseURL }
        $Body = $null

        # A returned login form indicates rejection; only non-private sites allow anonymous fallback.
        if ($Web.Content -match $LoginFormPattern) {
            if ($Type -ne 'Private') {
                $WebSess = New-Object -TypeName Microsoft.PowerShell.Commands.WebRequestSession
                $Web = Invoke-SchedRequest -Uri $SchedLoginURL -WebSession $WebSess
                $Script:Sessions[$ConfId] = $WebSess
                return [PSCustomObject]@{ Success = $true; Mode = 'Public'; Message = 'Login not accepted, using anonymous session' }
            }
            return [PSCustomObject]@{ Success = $false; Mode = 'Failed'; Message = 'Login failed / no access' }
        }

        $Script:Sessions[$ConfId] = $WebSess
        return [PSCustomObject]@{ Success = $true; Mode = 'Authenticated'; Message = 'Authenticated' }
    }
    catch {
        $Err = $_
        return [PSCustomObject]@{ Success = $false; Mode = 'Failed'; Message = 'Unavailable: ' + $Err.Exception.Message }
    }
}

# Return cached schedule objects unless Refresh is requested; otherwise use the existing HTTP session.
function Get-ConferenceEvents {
    param(
        [Parameter(Mandatory = $true)][string]$ConfId,
        [switch]$Refresh
    )

    if (-not $Refresh -and $Script:EventCache.ContainsKey($ConfId)) { return @($Script:EventCache[$ConfId]) }

    $SchedBaseURL = 'https://mms' + $ConfId + '.sched.com'
    $WebSess = $Script:Sessions[$ConfId]

    $Sched = Invoke-SchedRequest -Uri ($SchedBaseURL + '/list/descriptions') -WebSession $WebSess

    $EventList = New-Object -TypeName System.Collections.Generic.List[object]
    $Pairs = Get-EventLinkPairs -SchedulePage $Sched

    foreach ($Pair in $Pairs) {
        $EventTitle = [string]$Pair.Title
        $EventUrl = [string]$Pair.Url
        if (Test-IsBlank $EventUrl) { continue }
        if (Test-IsBlank $EventTitle) { $EventTitle = $EventUrl }

        # Remove trailing capacity labels so schedule availability text does not become part of the title.
        $EventTitle = $EventTitle -replace '\s*(FULL|LIMITED|SOLD OUT|WAITLIST|FILLING)\s*$', ''
        $EventTitle = $EventTitle.Trim()

        # Omit configured social/activity title patterns intended to exclude non-content sessions.
        if ($EventTitle -match '^Fishing with' -or $EventTitle -match '^Career Connections' -or
            $EventTitle -match '^Great Big Game Show' -or $EventTitle -match 'Escape Game') { continue }

        $EventList.Add([PSCustomObject]@{
                ConfId  = $ConfId
                Title   = $EventTitle
                Url     = $EventUrl
                Checked = $false
            })
    }

    $Script:EventCache[$ConfId] = $EventList.ToArray()
    return @($Script:EventCache[$ConfId])
}

# Refresh one conference's connection, schedule, Access, and SessionCount without resetting other entries.
function Test-ConferenceAccess {
    param(
        [Parameter(Mandatory = $true)][object]$Conf,
        [System.Management.Automation.PSCredential]$Credential
    )

    $Conf.Access = 'Checking...'
    $Conf.SessionCount = 0
    $Conf.Queried = $true

    # Replace this conference's loaded sessions, including their old session checkbox states.
    for ($SessIndex = $Script:SessionList.Count - 1; $SessIndex -ge 0; $SessIndex--) {
        if ($Script:SessionList[$SessIndex].ConfId -eq $Conf.Id) { $Script:SessionList.RemoveAt($SessIndex) }
    }

    $Script:Sessions.Remove($Conf.Id)
    $Script:EventCache.Remove($Conf.Id)

    $Connect = Connect-SchedConference -ConfId $Conf.Id -Type $Conf.Type -Credential $Credential
    if (Test-CancellationRequested) { $Conf.Access = 'Cancelled'; return }
    if (-not $Connect.Success) {
        $Conf.Access = $Connect.Message
        Write-Log -Message ($Conf.Id + ': ' + $Connect.Message) -Level 'WARN'
        return
    }

    try {
        $Events = @(Get-ConferenceEvents -ConfId $Conf.Id -Refresh)
        $Conf.SessionCount = $Events.Count
        if ($Events.Count -gt 0) {
            if ($Connect.Mode -eq 'Authenticated') { $Conf.Access = 'Yes (Logged in)' } else { $Conf.Access = 'Yes (Public)' }
            foreach ($EventItem in $Events) { $Script:SessionList.Add($EventItem) }
        }
        else {
            $Conf.Access = 'No sessions found'
        }
        Write-Log -Message ('{0}: {1} - {2} session(s)' -f $Conf.Id, $Conf.Access, $Events.Count)
    }
    catch {
        $Err = $_
        if (Test-CancellationRequested) { $Conf.Access = 'Cancelled'; return }
        $Conf.Access = 'Unable to read schedule'
        Write-Log -Message ($Conf.Id + ': unable to retrieve session list - ' + $Err.Exception.Message) -Level 'WARN'
    }
}

# Accept a discovery candidate only when the final response host matches mms<ConfId>.sched.com.
function Test-ConferenceExists {
    param([Parameter(Mandatory = $true)][string]$ConfId)

    $ExpectedHost = 'mms' + $ConfId + '.sched.com'
    $Url = 'https://' + $ExpectedHost + '/list/descriptions'
    try {
        $Resp = Invoke-CancellableWebRequest -RequestParameters @{ Uri = $Url; UseBasicParsing = $true; TimeoutSec = 15; ErrorAction = 'Stop' }
        $FinalHost = ''
        if ($Resp.BaseResponse.ResponseUri) { $FinalHost = [string]$Resp.BaseResponse.ResponseUri.Host }
        elseif ($Resp.BaseResponse.RequestMessage.RequestUri) { $FinalHost = [string]$Resp.BaseResponse.RequestMessage.RequestUri.Host }
        return ($FinalHost -ieq $ExpectedHost)
    }
    catch {
        return $false
    }
}

# Probe unknown year/suffix candidates within the supplied range and append matching Sched sites.
function Find-NewConferences {
    param(
        [Parameter(Mandatory = $true)][int]$StartYear,
        [Parameter(Mandatory = $true)][int]$EndYear
    )

    for ($Year = $StartYear; $Year -le $EndYear; $Year++) {
        foreach ($Suffix in $Script:ProbeSuffixes) {
            if (Test-CancellationRequested) { return }
            $ProbeId = [string]$Year + $Suffix
            if (Get-ConferenceInfo -ConfId $ProbeId) { continue }

            Set-Status -Text ('Probing Sched for conference ID ' + $ProbeId + ' ...')
            if (Test-ConferenceExists -ConfId $ProbeId) {
                $Script:ConferenceCatalog.Add((New-ConferenceEntry -Id $ProbeId -Name ('MMS ' + $ProbeId + ' (discovered)') -Type 'Discovered'))
                Write-Log -Message ('Discovered conference ID: ' + $ProbeId)
            }
        }
    }
}

# Coordinate query/discovery scopes; the GUI invokes one mode switch at a time.
#   SelectedOnly: refresh checked conferences across years without probing or clearing other results.
#   DiscoverOnly: append discovered IDs without querying sessions or resetting existing results.
#   PastOnly: probe past years and query past entries not already queried, retaining loaded results.
#   Default: reset query results, probe IDs, and query all years or current/future according to Show past years.
function Invoke-ConferenceDiscovery {
    param([switch]$PastOnly, [switch]$SelectedOnly, [switch]$DiscoverOnly)

    if ($Script:IsBusy -or ($Script:Form -and ($Script:Form.IsDisposed -or $Script:Form.ShutdownRequested))) { return }

    Sync-ConferenceChecks
    $SelectedConferences = @($Script:ConferenceCatalog | Where-Object { $_.Checked })
    if ($SelectedOnly -and $SelectedConferences.Count -eq 0) {
        Set-Status -Text 'No conferences selected.'
        return
    }
    $Script:CancelRequested = $false
    $CurrentYear = (Get-Date).Year
    $IncludePast = [bool]$Script:ChkPast.Checked
    Set-UiBusy -Busy $true

    if (-not $PastOnly -and -not $SelectedOnly -and -not $DiscoverOnly) {
        # Only the retained no-switch path resets loaded sessions and catalog query statuses.
        $Script:SessionList.Clear()
        foreach ($Conf in $Script:ConferenceCatalog) {
            $Conf.Queried = $false
            $Conf.Access = 'Not queried'
            $Conf.SessionCount = 0
        }
        Update-ConferenceListView
        Update-SessionListView
    }

    try {
        if ($SelectedOnly) { $ScopeText = 'selected conferences' }
        elseif ($PastOnly) { $ScopeText = 'past years' }
        elseif ($IncludePast) { $ScopeText = 'all years' }
        else { $ScopeText = ('{0} and later' -f $CurrentYear) }

        if ($Script:Credential) { Write-Log -Message ('Querying MMS conferences (' + $ScopeText + ') as ' + $Script:Credential.UserName + ' ...') }
        else { Write-Log -Message ('Querying MMS conferences (' + $ScopeText + ', public content only) ...') }

        # Selected queries never probe; other modes discover unknown IDs in their configured year scope.
        if (-not $SelectedOnly) {
            if ($PastOnly) { Find-NewConferences -StartYear $Script:ProbeStartYear -EndYear ($CurrentYear - 1) }
            elseif ($IncludePast) { Find-NewConferences -StartYear $Script:ProbeStartYear -EndYear ($CurrentYear + 1) }
            else { Find-NewConferences -StartYear $CurrentYear -EndYear ($CurrentYear + 1) }
        }

        if (Test-CancellationRequested) { Set-Status -Text 'Conference query cancelled.'; return }
        if ($DiscoverOnly) {
            Update-ConferenceListView
            Set-Status -Text 'Conference discovery complete.'
            return
        }

        # Build targets from checked entries, unqueried past entries, or the retained default year scope.
        $Targets = New-Object -TypeName System.Collections.Generic.List[object]
        foreach ($Conf in $Script:ConferenceCatalog) {
            $IsPast = Test-IsPastConference -Conf $Conf
            if ($SelectedOnly) {
                if ($SelectedConferences -contains $Conf) { $Targets.Add($Conf) }
            }
            elseif ($PastOnly) {
                if ($IsPast -and -not $Conf.Queried) { $Targets.Add($Conf) }
            }
            elseif ($IncludePast -or -not $IsPast) {
                $Targets.Add($Conf)
            }
        }

        $Total = $Targets.Count
        $Script:PbMain.Minimum = 0
        $Script:PbMain.Maximum = [System.Math]::Max($Total, 1)
        $Script:PbMain.Value = 0

        for ($Index = 0; $Index -lt $Total; $Index++) {
            if ($Script:CancelRequested) { Write-Log -Message 'Query cancelled by user.' -Level 'WARN'; break }
            $Conf = $Targets[$Index]
            Set-Status -Text ('Checking {0} of {1}: {2} ({3})' -f ($Index + 1), $Total, $Conf.Name, $Conf.Id)
            Test-ConferenceAccess -Conf $Conf -Credential $Script:Credential
            if (Test-CancellationRequested) { Set-Status -Text 'Conference query cancelled.'; return }
            Update-ConferenceListView
            $Script:PbMain.Value = $Index + 1
        }

        $Accessible = @($Script:ConferenceCatalog | Where-Object { $_.Access -like 'Yes*' -and -not (Test-IsHiddenConference -Conf $_) }).Count
        Update-ConferenceListView
        Update-SessionListView
        $StatusText = 'Query complete ({0}). You can access {1} conference(s) with {2} loaded session(s).' -f $ScopeText, $Accessible, $Script:SessionList.Count
        if (-not $IncludePast -and -not $SelectedOnly) { $StatusText += ' Check "Show past years" to display older conferences.' }
        Set-Status -Text $StatusText
        Write-Log -Message ('Query complete ({0}). Conferences queried: {1}. Accessible: {2}. Sessions loaded: {3}' -f $ScopeText, $Total, $Accessible, $Script:SessionList.Count)
    }
    catch {
        $Err = $_
        Write-Log -Message ('Query error: ' + $Err.Exception.Message) -Level 'ERROR'
    }
    finally {
        Set-UiBusy -Busy $false
    }
}

# Verify authenticated conference access before downloading; cached anonymous sessions are insufficient.
function Confirm-ConferenceSession {
    param([Parameter(Mandatory = $true)][string]$ConfId)

    if (-not $Script:Credential) {
        Set-Status -Text 'Sign in before downloading conference files.'
        Write-Log -Message ('Download blocked for ' + $ConfId + ': sign-in required.') -Level 'WARN'
        return $false
    }

    $Conf = Get-ConferenceInfo -ConfId $ConfId
    $ConfType = 'Discovered'
    if ($Conf) { $ConfType = $Conf.Type }

    $Script:Sessions.Remove($ConfId)
    Write-Log -Message ('Verifying authenticated download access for ' + $ConfId + ' ...')
    $Connect = Connect-SchedConference -ConfId $ConfId -Type $ConfType -Credential $Script:Credential
    if (-not $Connect.Success -or $Connect.Mode -ne 'Authenticated') {
        $Script:Sessions.Remove($ConfId)
        Set-Status -Text ('Download blocked for ' + $ConfId + ': sign-in/access not verified.')
        Write-Log -Message ('Download blocked for https://mms' + $ConfId + '.sched.com: ' + $Connect.Message) -Level 'WARN'
        return $false
    }
    return $true
}
# ==================================================
# ^^^^^ Authentication and Conference Query Functions ^^^^^
# ==================================================

# ==================================================
# MARK: Download Functions
# ==================================================
<#
## Download Functions
### Purpose
Download Sched-hosted attachments and optional session metadata, then summarize the job.
### File and Folder Handling
`Save-SessionFile` returns Downloaded, Updated, Skipped, or Failed. Existing files are
skipped by default. Optional hash comparison downloads to `.new`, replaces differing
content, and removes matching/failed temporary copies. Failed new downloads are removed.
`Save-ConferenceSession` builds `<DownloadLocation>\mms<Id>\<SanitizedSessionTitle>`,
retrieves the event page, optionally writes Session Info.txt, and deduplicates hosted links.
File names are decoded/sanitized; long output paths are shortened toward 255 characters
while retaining the extension. This is inherited path handling, not a general long-path solution.
### Job Lifecycle and Reporting
`Start-SessionDownload` reads the GUI options, resets the current job's result list,
ensures each conference has an HTTP session, and advances progress per session.
Before any session content is written, authenticate each distinct target conference once.
If any conference fails that preflight, stop the entire job without writing session metadata
or attachments. The shared job entry point also rejects attempts without completed Sign In.
Outcomes include file statuses, session-page/login failures, and sessions without files.
The finally block logs totals, exports any result records, and restores operation controls.
### Cancellation and Scope
Cancellation is checked between sessions/files and while HTTP requests are pending.
Partial new downloads and hash-check temporary files are removed; completed files remain.
Cancelled jobs skip result export, and shutdown avoids further UI updates.
Only discovered hosted-file URLs are downloaded. Writing metadata and discovering attachments
depend on Sched HTML structure, and metadata generation can occur even when no files exist.
#>
# Download or compare one hosted file; clean up incomplete output on failure or cancellation.
function Save-SessionFile {
    param(
        [Parameter(Mandatory = $true)][string]$DownloadUrl,
        [Parameter(Mandatory = $true)][string]$OutputFilePath,
        [Parameter(Mandatory = $true)]$WebSession,
        [bool]$ReDownloadIfHashIsDifferent = $false
    )

    if (Test-CancellationRequested) { return 'Cancelled' }
    $FileLabel = Split-Path -Path $OutputFilePath -Leaf

    if ((Test-Path -LiteralPath $OutputFilePath) -eq $false) {
        Write-Log -Message ('DOWNLOADING: ' + $FileLabel)
        try {
            Invoke-SchedRequest -Uri $DownloadUrl -WebSession $WebSession -OutFile $OutputFilePath | Out-Null
            Unblock-File -LiteralPath $OutputFilePath -ErrorAction SilentlyContinue
            return 'Downloaded'
        }
        catch {
            if (Test-Path -LiteralPath $OutputFilePath) { Remove-Item -LiteralPath $OutputFilePath -Force -ErrorAction SilentlyContinue }
            if (Test-CancellationRequested) { return 'Cancelled' }
            Write-Log -Message ('FAILED: ' + $FileLabel) -Level 'ERROR'
            return 'Failed'
        }
    }

    if ($ReDownloadIfHashIsDifferent) {
        Write-Log -Message ('DOWNLOADING: ' + $FileLabel + ' (hash check)')
        $OldHash = (Get-FileHash -LiteralPath $OutputFilePath).Hash
        $TempPath = $OutputFilePath + '.new'
        try {
            Invoke-SchedRequest -Uri $DownloadUrl -WebSession $WebSession -OutFile $TempPath | Out-Null
            Unblock-File -LiteralPath $TempPath -ErrorAction SilentlyContinue
            $NewHash = (Get-FileHash -LiteralPath $TempPath).Hash
            if ($NewHash -ne $OldHash) {
                Move-Item -LiteralPath $TempPath -Destination $OutputFilePath -Force
                return 'Updated'
            }
            Write-Log -Message ('SKIPPING: ' + $FileLabel)
            Remove-Item -LiteralPath $TempPath -Force
            return 'Skipped'
        }
        catch {
            if (Test-Path -LiteralPath $TempPath) { Remove-Item -LiteralPath $TempPath -Force -ErrorAction SilentlyContinue }
            if (Test-CancellationRequested) { return 'Cancelled' }
            Write-Log -Message ('FAILED: ' + $FileLabel) -Level 'ERROR'
            return 'Failed'
        }
    }

    Write-Log -Message ('SKIPPING: ' + $FileLabel)
    return 'Skipped'
}

# Download a session's hosted attachments and optionally write its description/speaker metadata.
function Save-ConferenceSession {
    param(
        [Parameter(Mandatory = $true)][string]$ConfId,
        [Parameter(Mandatory = $true)][object]$EventItem,
        [Parameter(Mandatory = $true)][string]$DownloadLocation,
        [bool]$ExcludeSessionDetails = $false,
        [bool]$ReDownloadIfHashIsDifferent = $false
    )

    if (Test-CancellationRequested) { return }
    $WebSess = $Script:Sessions[$ConfId]
    $SchedBaseURL = 'https://mms' + $ConfId + '.sched.com'
    $SessionDownloadPath = $DownloadLocation + '\mms' + $ConfId

    $EventTitle = [string]$EventItem.Title
    $EventUrl = [string]$EventItem.Url

    # Seed the metadata text and sanitize the title used for this session's destination folder.
    $SessionInfoText = 'Session Title: ' + "`r`n" + $EventTitle + "`r`n`r`n"
    $DownloadTitle = $EventTitle -replace '[^A-Za-z0-9-_. ]', ''
    $DownloadTitle = $DownloadTitle.Trim()
    $DownloadTitle = $DownloadTitle -replace '\W+', '_'
    if (Test-IsBlank $DownloadTitle) {
        $DownloadTitle = [System.Uri]::UnescapeDataString($EventUrl) -replace '\W+', '_'
    }

    # Nest the sanitized session folder under the conference's mms<Id> folder.
    $DownloadPath = $SessionDownloadPath + '\' + $DownloadTitle

    # Read the individual event page containing metadata and hosted attachment links.
    $SessionPage = $null
    try {
        $SessionPage = Invoke-SchedRequest -Uri ($SchedBaseURL + '/' + $EventUrl) -WebSession $WebSess
    }
    catch {
        $SessionPage = $null
    }

    if (Test-CancellationRequested) { return }
    if (-not $SessionPage) {
        Write-Log -Message ('FAILED: ' + $EventTitle) -Level 'ERROR'
        Add-DownloadResult -ConfId $ConfId -SessionTitle $EventTitle -FileName '' -Status 'Failed (session page)' -Path ''
        return
    }

    $SessionLinkInfo = $SessionPage.Content.Replace("`r", '').Replace("`n", '')

    # Extract supported description/speaker blocks and write Session Info.txt unless excluded.
    if (-not $ExcludeSessionDetails) {
        $DescriptionPattern = '<div class="tip-description">(?<description>.*?)(<div class="tip-roles">|<div class="sched-event-details-timeandplace">)'
        $DescriptionMatches = [System.Text.RegularExpressions.Regex]::Matches($SessionLinkInfo, $DescriptionPattern, $Script:RegexIgnoreCase)
        if ($DescriptionMatches.Count -gt 0) {
            $SessionInfoText += (Invoke-BasicHTMLParser -Html $DescriptionMatches[0].Groups['description'].Value -SchedBaseURL $SchedBaseURL) + "`r`n`r`n"
        }

        $RolesPattern = '<div class="tip-roles">(?<roles>.*?)<div class="sched-file">|<div class="sched-event-details-timeandplace">'
        $RolesMatches = [System.Text.RegularExpressions.Regex]::Matches($SessionLinkInfo, $RolesPattern, $Script:RegexIgnoreCase)
        if ($RolesMatches.Count -gt 0) {
            $SessionInfoText += (Invoke-BasicHTMLParser -Html $RolesMatches[0].Groups['roles'].Value -SchedBaseURL $SchedBaseURL) + "`r`n`r`n"
        }

        if ((Test-Path -LiteralPath $DownloadPath) -eq $false) { New-Item -ItemType Directory -Force -Path $DownloadPath | Out-Null }
        Out-File -LiteralPath ($DownloadPath + '\Session Info.txt') -InputObject $SessionInfoText -Force -Encoding UTF8
    }

    # Collect hosted-file links from parsed anchors, deduplicating by their href values.
    $Downloads = New-Object -TypeName System.Collections.Generic.List[object]
    $SeenHrefs = @{}
    if ($SessionPage.Links) {
        foreach ($Lnk in $SessionPage.Links) {
            $Href = [string]$Lnk.href
            if ($Href -like '*hosted-files*') {
                if (-not $SeenHrefs.ContainsKey($Href)) {
                    $SeenHrefs[$Href] = $true
                    $Downloads.Add([PSCustomObject]@{ href = $Href })
                }
            }
        }
    }

    # Fall back to raw HTML hosted-file URL extraction when parsed anchors found no downloads.
    if ($Downloads.Count -eq 0 -and $SessionPage.Content) {
        $HostedFilesRegex = 'href="((?:https?:)?//hosted-files\.sched\.co/[^"]+)"'
        $HostedMatches = [System.Text.RegularExpressions.Regex]::Matches($SessionPage.Content, $HostedFilesRegex, $Script:RegexIgnoreCase)
        foreach ($HostedMatch in $HostedMatches) {
            $HostedUrl = $HostedMatch.Groups[1].Value
            if (-not $SeenHrefs.ContainsKey($HostedUrl)) {
                $SeenHrefs[$HostedUrl] = $true
                $Downloads.Add([PSCustomObject]@{ href = $HostedUrl })
            }
        }
    }

    if ($Downloads.Count -eq 0) {
        Write-Log -Message ('No files found for: ' + $EventTitle)
        Add-DownloadResult -ConfId $ConfId -SessionTitle $EventTitle -FileName '' -Status 'No files' -Path ''
        return
    }

    foreach ($Download in $Downloads) {
        if (Test-CancellationRequested) { return }

        $DownloadUrl = [string]$Download.href
        if ($DownloadUrl -like '//hosted-files.sched.co/*') {
            $DownloadUrl = 'https:' + $DownloadUrl
        }
        elseif ($DownloadUrl -notmatch '^https?://') {
            $DownloadUrl = $SchedBaseURL + '/' + $DownloadUrl
        }

        $FileName = Split-Path -Path $DownloadUrl -Leaf
        # Decode URL-encoded filename text, such as %20 for a space.
        $FileName = [System.Web.HttpUtility]::UrlDecode($FileName)
        # Keep the inherited filename-safe character set before constructing the output path.
        $FileName = $FileName -replace '[^A-Za-z0-9\.\-_ ]', ''

        $OutputFilePath = $DownloadPath + '\' + $FileName

        # Shorten paths at or above 255 characters while preserving the file extension.
        $OutputFilePathLen = $OutputFilePath.Length
        if ($OutputFilePathLen -ge 255) {
            $FileExt = [System.IO.Path]::GetExtension($OutputFilePath)
            $NewFileName = $OutputFilePath.Substring(0, ($OutputFilePathLen - $FileExt.Length))
            $NewFileName = $NewFileName.Substring(0, (255 - $FileExt.Length)).Trim()
            $OutputFilePath = $NewFileName + $FileExt
        }

        if ((Test-Path -LiteralPath $DownloadPath) -eq $false) { New-Item -ItemType Directory -Force -Path $DownloadPath | Out-Null }

        $FileStatus = Save-SessionFile -DownloadUrl $DownloadUrl -OutputFilePath $OutputFilePath -WebSession $WebSess -ReDownloadIfHashIsDifferent $ReDownloadIfHashIsDifferent
        if (Test-CancellationRequested) { return }
        Add-DownloadResult -ConfId $ConfId -SessionTitle $EventTitle -FileName $FileName -Status $FileStatus -Path $OutputFilePath
    }
}

# Run the supplied session targets with current GUI options, progress, cancellation, and final export.
function Start-SessionDownload {
    param([Parameter(Mandatory = $true)][object[]]$Targets)

    if ($Script:IsBusy -or ($Script:Form -and ($Script:Form.IsDisposed -or $Script:Form.ShutdownRequested))) { return }
    if (-not $Script:Credential) {
        Set-Status -Text 'Sign in before downloading conference files.'
        Write-Log -Message 'Download blocked: complete Sign In before starting a download.' -Level 'WARN'
        return
    }

    $DownloadLocation = $Script:TxtLocation.Text.Trim().TrimEnd('\')
    if (Test-IsBlank $DownloadLocation) { $DownloadLocation = $Script:DefaultDownloadLocation }
    $ExcludeDetails = [bool]$Script:ChkExclude.Checked
    $ReDownload = [bool]$Script:ChkRehash.Checked

    $Script:CancelRequested = $false
    $Script:Results.Clear()
    Set-UiBusy -Busy $true

    $Script:PbMain.Minimum = 0
    $Script:PbMain.Maximum = [System.Math]::Max($Targets.Count, 1)
    $Script:PbMain.Value = 0

    Write-Log -Message ('Base download folder is ' + $DownloadLocation)
    Write-Log -Message ('Starting download of ' + $Targets.Count + ' session(s).')

    $DownloadAccess = @{}
    $PreflightStatus = $null
    try {
        foreach ($Target in $Targets) {
            if (Test-CancellationRequested) {
                $PreflightStatus = 'Download cancelled during access verification. No conference files were written.'
                return
            }
            if ($DownloadAccess.ContainsKey($Target.ConfId)) { continue }
            Set-Status -Text ('Verifying download access: ' + $Target.ConfId)
            if (-not (Confirm-ConferenceSession -ConfId $Target.ConfId)) {
                if (Test-CancellationRequested) { $PreflightStatus = 'Download cancelled during access verification. No conference files were written.' }
                else { $PreflightStatus = 'Download blocked for ' + $Target.ConfId + ': authenticated access required. No conference files were written.' }
                return
            }
            $DownloadAccess[$Target.ConfId] = $true
        }

        for ($Index = 0; $Index -lt $Targets.Count; $Index++) {
            if ($Script:CancelRequested) { Write-Log -Message 'Download cancelled by user.' -Level 'WARN'; break }

            $Target = $Targets[$Index]
            Set-Status -Text ('Session {0} of {1}: [{2}] {3}' -f ($Index + 1), $Targets.Count, $Target.ConfId, $Target.Title)

            if (-not $DownloadAccess.ContainsKey($Target.ConfId)) {
                Add-DownloadResult -ConfId $Target.ConfId -SessionTitle $Target.Title -FileName '' -Status 'Failed (login)' -Path ''
                $Script:PbMain.Value = $Index + 1
                continue
            }

            Save-ConferenceSession -ConfId $Target.ConfId -EventItem $Target -DownloadLocation $DownloadLocation -ExcludeSessionDetails $ExcludeDetails -ReDownloadIfHashIsDifferent $ReDownload
            if (Test-CancellationRequested) { break }
            $Script:PbMain.Value = $Index + 1
            Update-UiEvents
        }
    }
    catch {
        $Err = $_
        Write-Log -Message ('Unexpected error: ' + $Err.Exception.Message) -Level 'ERROR'
    }
    finally {
        $Downloaded = @($Script:Results | Where-Object { $_.Status -eq 'Downloaded' }).Count
        $Updated = @($Script:Results | Where-Object { $_.Status -eq 'Updated' }).Count
        $Skipped = @($Script:Results | Where-Object { $_.Status -eq 'Skipped' }).Count
        $Failed = @($Script:Results | Where-Object { $_.Status -like 'Failed*' }).Count
        $NoFiles = @($Script:Results | Where-Object { $_.Status -eq 'No files' }).Count
        $Summary = 'Finished. Downloaded: {0} | Updated: {1} | Skipped: {2} | Failed: {3} | Sessions without files: {4}' -f $Downloaded, $Updated, $Skipped, $Failed, $NoFiles
        if (Test-CancellationRequested) { $Summary = 'Download cancelled. Completed files were kept; partial downloads were removed.' }
        if ($PreflightStatus) { $Summary = $PreflightStatus }
        Write-Log -Message $Summary
        Set-Status -Text $Summary
        try {
            if (-not (Test-CancellationRequested)) { Export-DownloadResults }
        }
        catch {
            Write-Log -Message ('Result export failed after downloads completed: ' + $_.Exception.Message) -Level 'ERROR'
        }
        finally {
            Set-UiBusy -Busy $false
        }
    }
}
# ==================================================
# ^^^^^ Download Functions ^^^^^
# ==================================================

# ==================================================
# MARK: Export Functions
# ==================================================
<#
## Export Functions
### Purpose
Persist the current download job's result records as a timestamped report under OutputRoot.
### Report Schema and Naming
Reports contain Timestamp, Conference, Session, File, Status, and Path. Their filename
starts with `yyyy-MM-dd_HHmm-MMS-Download-Results`, followed by `.xlsx` or `.csv`.
Snapshot the generic result list with ToArray() to avoid PowerShell array-binder errors;
the same snapshot implementation works in Windows PowerShell 5.1 and PowerShell 7.
### Preferred Export and Fallback
When ImportExcel is available, import it and export the Results worksheet with automatic
column sizing, a frozen header row, and the currently configured Light21 table style.
If the module is absent or Excel export throws, automatically attempt UTF-8 CSV export.
CSV/export errors are logged; no report is attempted when the result collection is empty.
### Destination
Reports always use `OutputRoot`, not the GUI's session-content download folder.
#>
# Export a nonempty result collection to Excel when possible, otherwise fall back to CSV.
function Export-DownloadResults {
    if ($Script:Results.Count -eq 0 -or (Test-CancellationRequested)) { return }

    $Stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $BaseName = $Stamp + '-MMS-Download-Results'
    $Data = $Script:Results.ToArray()
    $Exported = $false

    if (Get-Module -ListAvailable -Name ImportExcel) {
        try {
            Import-Module -Name ImportExcel -ErrorAction Stop
            $XlsxPath = Join-Path -Path $Script:OutputRoot -ChildPath ($BaseName + '.xlsx')
            $Data | Export-Excel -Path $XlsxPath -WorksheetName 'Results' -TableStyle Light21 -AutoSize -FreezeTopRow -ErrorAction Stop
            Write-Log -Message ('Results exported to ' + $XlsxPath)
            $Exported = $true
        }
        catch {
            $Err = $_
            Write-Log -Message ('Export-Excel failed (' + $Err.Exception.Message + '). Falling back to CSV.') -Level 'WARN'
        }
    }

    if (-not $Exported) {
        try {
            $CsvPath = Join-Path -Path $Script:OutputRoot -ChildPath ($BaseName + '.csv')
            $Data | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
            Write-Log -Message ('Results exported to ' + $CsvPath)
        }
        catch {
            $Err = $_
            Write-Log -Message ('CSV export failed: ' + $Err.Exception.Message) -Level 'ERROR'
        }
    }
}
# ==================================================
# ^^^^^ Export Functions ^^^^^
# ==================================================

# ==================================================
# MARK: GUI State Functions
# ==================================================
<#
## GUI State Functions
### Purpose
Synchronize checkbox state with underlying objects and rebuild list views without recursive handlers.
### Selection and Filtering
`Sync-ConferenceChecks` and `Sync-SessionChecks` copy displayed checkbox states to each
item's tagged model object before a refresh. `Update-ConferenceListView` keeps catalog
order and applies year/access visibility settings, except that manual entries stay visible.
`Update-SessionListView` applies year visibility, checked-conference scope, and the title/ID
text filter. With no conferences checked, it shows loaded sessions not hidden by year.
Session checkbox state belongs to the model; changing text filters does not clear it.
### Operation Controls
`Set-UiBusy` guards concurrent operations, disables the listed action/credential controls,
and enables Cancel while busy. Download actions additionally require an active verified
credential, including after an operation finishes. List rebuilding temporarily sets SuppressEvents and uses
BeginUpdate/EndUpdate. `Get-CheckedConferences` additionally requires an accessible status
for bulk conference downloads; it is not the target selector used by Query Selected.
#>
# Save conference checkbox state from visible rows to their catalog objects.
function Sync-ConferenceChecks {
    if (-not $Script:LvConf) { return }
    foreach ($Item in $Script:LvConf.Items) { $Item.Tag.Checked = [bool]$Item.Checked }
}

# Save displayed session checkbox state before rebuilding or starting a download.
function Sync-SessionChecks {
    if (-not $Script:LvSess) { return }
    foreach ($Item in $Script:LvSess.Items) { $Item.Tag.Checked = [bool]$Item.Checked }
}

# Rebuild the ordered catalog display using Show past years/access filters, preserving manual entries.
function Update-ConferenceListView {
    if (-not $Script:LvConf -or $Script:LvConf.IsDisposed -or $Script:Form.ShutdownRequested) { return }

    $Script:SuppressEvents = $true
    Sync-ConferenceChecks
    $Script:LvConf.BeginUpdate()
    $Script:LvConf.Items.Clear()
    foreach ($Conf in $Script:ConferenceCatalog) {
        if (Test-IsHiddenConference -Conf $Conf) { continue }
        $IsAccessible = ([string]$Conf.Access -like 'Yes*')
        if (-not $IsAccessible -and -not $Script:ChkShowAll.Checked -and -not $Conf.Manual) { continue }

        $Item = New-Object -TypeName System.Windows.Forms.ListViewItem -ArgumentList $Conf.Name
        [void]$Item.SubItems.Add($Conf.Id)
        [void]$Item.SubItems.Add([string]$Conf.Access)
        [void]$Item.SubItems.Add([string]$Conf.SessionCount)
        if (-not $IsAccessible) { $Item.ForeColor = [System.Drawing.Color]::Gray }
        $Item.Tag = $Conf
        $Item.Checked = [bool]$Conf.Checked
        [void]$Script:LvConf.Items.Add($Item)
    }
    $Script:LvConf.EndUpdate()
    $Script:SuppressEvents = $false
    Update-UiEvents
}

# Show loaded sessions in the checked/year-visible conference scope, then apply the title/ID filter.
function Update-SessionListView {
    if (-not $Script:LvSess -or $Script:LvSess.IsDisposed -or $Script:Form.ShutdownRequested) { return }

    $Script:SuppressEvents = $true
    Sync-SessionChecks

    $CheckedIds = @{}
    $HiddenIds = @{}
    foreach ($Conf in $Script:ConferenceCatalog) {
        if (Test-IsHiddenConference -Conf $Conf) { $HiddenIds[$Conf.Id] = $true; continue }
        if ($Conf.Checked) { $CheckedIds[$Conf.Id] = $true }
    }

    $Filter = $Script:TxtFilter.Text.Trim()
    $ShownCount = 0

    $Script:LvSess.BeginUpdate()
    $Script:LvSess.Items.Clear()
    foreach ($Sess in $Script:SessionList) {
        if ($HiddenIds.ContainsKey($Sess.ConfId)) { continue }
        if ($CheckedIds.Count -gt 0 -and -not $CheckedIds.ContainsKey($Sess.ConfId)) { continue }
        if ($Filter) {
            $InTitle = ($Sess.Title.IndexOf($Filter, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
            $InConf = ($Sess.ConfId.IndexOf($Filter, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
            if (-not $InTitle -and -not $InConf) { continue }
        }
        $Item = New-Object -TypeName System.Windows.Forms.ListViewItem -ArgumentList $Sess.ConfId
        [void]$Item.SubItems.Add($Sess.Title)
        $Item.Tag = $Sess
        $Item.Checked = [bool]$Sess.Checked
        [void]$Script:LvSess.Items.Add($Item)
        $ShownCount++
    }
    $Script:LvSess.EndUpdate()
    $Script:SuppressEvents = $false

    $SelectedCount = @($Script:SessionList | Where-Object { $_.Checked }).Count
    $Script:LblSessCount.Text = 'Showing {0} of {1} session(s) | {2} selected' -f $ShownCount, $Script:SessionList.Count, $SelectedCount
    Update-UiEvents
}

# Guard overlapping operations and switch the listed action/credential controls and Cancel button.
function Set-UiBusy {
    param([bool]$Busy)

    $Script:IsBusy = $Busy
    if ($Script:Form -is [MmsContentForm]) {
        $Script:Form.OperationActive = $Busy
        if (-not $Busy) { $Script:Form.CompleteOperation() }
        if ($Script:Form.IsDisposed -or $Script:Form.ShutdownRequested) { $Script:CancelRequested = $true; return }
    }
    foreach ($Ctrl in @($Script:BtnSignIn, $Script:TxtUser, $Script:TxtPass, $Script:BtnQuery, $Script:BtnDiscover, $Script:BtnAddConf, $Script:BtnDlSelected, $Script:BtnDlAll, $Script:BtnBrowse, $Script:ChkPast)) {
        if ($Ctrl -and -not $Ctrl.IsDisposed) { $Ctrl.Enabled = -not $Busy }
    }
    foreach ($Ctrl in @($Script:BtnDlSelected, $Script:BtnDlAll)) {
        if ($Ctrl -and -not $Ctrl.IsDisposed) { $Ctrl.Enabled = (-not $Busy -and $null -ne $Script:Credential) }
    }
    if ($Script:BtnCancel -and -not $Script:BtnCancel.IsDisposed) { $Script:BtnCancel.Enabled = $Busy }
    Update-UiEvents
}

# Return checked, accessible conferences not hidden by year for the bulk-download action.
function Get-CheckedConferences {
    Sync-ConferenceChecks
    return @($Script:ConferenceCatalog | Where-Object { $_.Checked -and ([string]$_.Access -like 'Yes*') -and -not (Test-IsHiddenConference -Conf $_) })
}
# ==================================================
# ^^^^^ GUI State Functions ^^^^^
# ==================================================

# ==================================================
# MARK: Build GUI
# ==================================================
<#
## Build GUI
### Purpose
Construct the main Windows Forms interface and retain references used by state functions/handlers.
### Layout
The default window is 1200 x 840 with a 1000 x 760 minimum. The top group contains
credentials and conference actions; conference and session lists occupy the middle;
download options, progress, status, and the multiline activity log appear below them.
Anchors resize selected controls with the window; coordinates and column widths are explicit.
### Initial State
Show past years and Show unavailable conferences start checked, exposing all catalog entries.
No sessions are preloaded. The password textbox is masked, the default download destination
is populated, and Cancel starts disabled. Event handlers are registered in the next section;
constructing the controls does not itself sign in, discover IDs, or query schedules.
#>
$Script:Form = New-Object -TypeName MmsContentForm
$Script:Form.Text = 'MMS Session Content Downloader (Sched)'
$Script:Form.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1200, 840
$Script:Form.MinimumSize = New-Object -TypeName System.Drawing.Size -ArgumentList 1000, 760
$Script:Form.StartPosition = 'CenterScreen'
$Script:Form.Font = New-Object -TypeName System.Drawing.Font -ArgumentList 'Segoe UI', 9

# ---------- Credentials / Query group ----------
<#
### Credentials and Conference Actions
Place Sign In immediately after the masked password field. Discover IDs appends candidates
without loading sessions; Query Selected is in the conference panel. The second row
also provides ID-only manual entry, Add and Query, and the Show past years visibility toggle.
#>
$GrpCreds = New-Object -TypeName System.Windows.Forms.GroupBox
$GrpCreds.Text = 'Sched Credentials and Conference Discovery'
$GrpCreds.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 8
$GrpCreds.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1160, 88
$GrpCreds.Anchor = 'Top, Left, Right'
$Script:Form.Controls.Add($GrpCreds)

$LblUser = New-Object -TypeName System.Windows.Forms.Label
$LblUser.Text = 'Username:'
$LblUser.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 12, 28
$LblUser.AutoSize = $true
$GrpCreds.Controls.Add($LblUser)

$Script:TxtUser = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtUser.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 85, 25
$Script:TxtUser.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 230, 24
$GrpCreds.Controls.Add($Script:TxtUser)

$LblPass = New-Object -TypeName System.Windows.Forms.Label
$LblPass.Text = 'Password:'
$LblPass.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 335, 28
$LblPass.AutoSize = $true
$GrpCreds.Controls.Add($LblPass)

$Script:TxtPass = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtPass.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 405, 25
$Script:TxtPass.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 230, 24
$Script:TxtPass.UseSystemPasswordChar = $true
$GrpCreds.Controls.Add($Script:TxtPass)

$Script:BtnSignIn = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnSignIn.Text = 'Sign In'
$Script:BtnSignIn.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 645, 22
$Script:BtnSignIn.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 90, 30
$GrpCreds.Controls.Add($Script:BtnSignIn)

$Script:BtnDiscover = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnDiscover.Text = 'Discover IDs'
$Script:BtnDiscover.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 760, 56
$Script:BtnDiscover.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 110, 30
$GrpCreds.Controls.Add($Script:BtnDiscover)

$LblCustom = New-Object -TypeName System.Windows.Forms.Label
$LblCustom.Text = 'Add ID manually (2027atmoa, midway):'
$LblCustom.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 12, 61
$LblCustom.AutoSize = $true
$GrpCreds.Controls.Add($LblCustom)

$Script:TxtCustom = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtCustom.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 290, 58
$Script:TxtCustom.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 160, 24
$GrpCreds.Controls.Add($Script:TxtCustom)

$Script:BtnAddConf = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnAddConf.Text = 'Add and Query'
$Script:BtnAddConf.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 460, 56
$Script:BtnAddConf.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 110, 27
$GrpCreds.Controls.Add($Script:BtnAddConf)

$Script:ChkPast = New-Object -TypeName System.Windows.Forms.CheckBox
$Script:ChkPast.Text = 'Show past years'
$Script:ChkPast.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 600, 60
$Script:ChkPast.AutoSize = $true
$Script:ChkPast.Checked = $true
$GrpCreds.Controls.Add($Script:ChkPast)

# ---------- Conferences group ----------
<#
### Conference Inventory
Display Conference, ID, Access, and Sessions columns with checkbox selection. Rows keep
catalog insertion order, so new manual/discovered IDs appear at the end. Select All/None
operate on visible rows; Show unavailable conferences includes Not queried and failed entries.
Query Selected sits to the right of that visibility toggle and loads checked conferences.
#>
$GrpConf = New-Object -TypeName System.Windows.Forms.GroupBox
$GrpConf.Text = 'Conferences'
$GrpConf.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 104
$GrpConf.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 560, 330
$GrpConf.Anchor = 'Top, Left'
$Script:Form.Controls.Add($GrpConf)

$Script:LvConf = New-Object -TypeName System.Windows.Forms.ListView
$Script:LvConf.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 22
$Script:LvConf.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 540, 260
$Script:LvConf.View = 'Details'
$Script:LvConf.CheckBoxes = $true
$Script:LvConf.FullRowSelect = $true
$Script:LvConf.GridLines = $true
$Script:LvConf.HideSelection = $false
[void]$Script:LvConf.Columns.Add('Conference', 210)
[void]$Script:LvConf.Columns.Add('ID', 85)
[void]$Script:LvConf.Columns.Add('Access', 165)
[void]$Script:LvConf.Columns.Add('Sessions', 60)
$GrpConf.Controls.Add($Script:LvConf)

$BtnConfAll = New-Object -TypeName System.Windows.Forms.Button
$BtnConfAll.Text = 'Select All'
$BtnConfAll.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 290
$BtnConfAll.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 90, 28
$GrpConf.Controls.Add($BtnConfAll)

$BtnConfNone = New-Object -TypeName System.Windows.Forms.Button
$BtnConfNone.Text = 'Select None'
$BtnConfNone.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 106, 290
$BtnConfNone.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 90, 28
$GrpConf.Controls.Add($BtnConfNone)

$Script:ChkShowAll = New-Object -TypeName System.Windows.Forms.CheckBox
$Script:ChkShowAll.Text = 'Show unavailable conferences'
$Script:ChkShowAll.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 210, 295
$Script:ChkShowAll.AutoSize = $true
$Script:ChkShowAll.Checked = $true
$GrpConf.Controls.Add($Script:ChkShowAll)

$Script:BtnQuery = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnQuery.Text = 'Query Selected'
$Script:BtnQuery.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 420, 290
$Script:BtnQuery.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 130, 28
$GrpConf.Controls.Add($Script:BtnQuery)

# ---------- Sessions group ----------
<#
### Loaded Sessions
Display conference IDs and session titles with download checkboxes. The filter narrows
displayed titles/IDs; Select All Shown checks visible rows, while Clear Selection clears
all loaded session selections. The count label reports displayed, loaded, and checked totals.
#>
$GrpSess = New-Object -TypeName System.Windows.Forms.GroupBox
$GrpSess.Text = 'Sessions (check the ones to download)'
$GrpSess.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 580, 104
$GrpSess.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 590, 330
$GrpSess.Anchor = 'Top, Left, Right'
$Script:Form.Controls.Add($GrpSess)

$LblFilter = New-Object -TypeName System.Windows.Forms.Label
$LblFilter.Text = 'Filter:'
$LblFilter.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 27
$LblFilter.AutoSize = $true
$GrpSess.Controls.Add($LblFilter)

$Script:TxtFilter = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtFilter.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 55, 24
$Script:TxtFilter.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 525, 24
$Script:TxtFilter.Anchor = 'Top, Left, Right'
$GrpSess.Controls.Add($Script:TxtFilter)

$Script:LvSess = New-Object -TypeName System.Windows.Forms.ListView
$Script:LvSess.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 54
$Script:LvSess.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 570, 228
$Script:LvSess.Anchor = 'Top, Left, Right'
$Script:LvSess.View = 'Details'
$Script:LvSess.CheckBoxes = $true
$Script:LvSess.FullRowSelect = $true
$Script:LvSess.GridLines = $true
$Script:LvSess.HideSelection = $false
[void]$Script:LvSess.Columns.Add('Conf', 85)
[void]$Script:LvSess.Columns.Add('Session Title', 460)
$GrpSess.Controls.Add($Script:LvSess)

$BtnSessAll = New-Object -TypeName System.Windows.Forms.Button
$BtnSessAll.Text = 'Select All Shown'
$BtnSessAll.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 290
$BtnSessAll.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 120, 28
$GrpSess.Controls.Add($BtnSessAll)

$BtnSessNone = New-Object -TypeName System.Windows.Forms.Button
$BtnSessNone.Text = 'Clear Selection'
$BtnSessNone.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 136, 290
$BtnSessNone.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 110, 28
$GrpSess.Controls.Add($BtnSessNone)

$Script:LblSessCount = New-Object -TypeName System.Windows.Forms.Label
$Script:LblSessCount.Text = 'No sessions loaded yet'
$Script:LblSessCount.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 255, 296
$Script:LblSessCount.AutoSize = $true
$GrpSess.Controls.Add($Script:LblSessCount)

# ---------- Download options group ----------
<#
### Download Options and Actions
Choose the content destination, browse/open it, omit Session Info.txt, or compare existing
file hashes. Separate actions download checked sessions or every loaded session from checked
accessible conferences. Cancel is enabled by busy-state handling, not by folder selection.
#>
$GrpDl = New-Object -TypeName System.Windows.Forms.GroupBox
$GrpDl.Text = 'Download'
$GrpDl.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 442
$GrpDl.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1160, 110
$GrpDl.Anchor = 'Top, Left, Right'
$Script:Form.Controls.Add($GrpDl)

$LblLocation = New-Object -TypeName System.Windows.Forms.Label
$LblLocation.Text = 'Download folder:'
$LblLocation.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 12, 27
$LblLocation.AutoSize = $true
$GrpDl.Controls.Add($LblLocation)

$Script:TxtLocation = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtLocation.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 110, 24
$Script:TxtLocation.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 560, 24
$Script:TxtLocation.Text = $Script:DefaultDownloadLocation
$GrpDl.Controls.Add($Script:TxtLocation)

$Script:BtnBrowse = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnBrowse.Text = 'Browse...'
$Script:BtnBrowse.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 680, 22
$Script:BtnBrowse.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 80, 27
$GrpDl.Controls.Add($Script:BtnBrowse)

$BtnOpenFolder = New-Object -TypeName System.Windows.Forms.Button
$BtnOpenFolder.Text = 'Open Folder'
$BtnOpenFolder.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 766, 22
$BtnOpenFolder.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 90, 27
$GrpDl.Controls.Add($BtnOpenFolder)

$Script:ChkExclude = New-Object -TypeName System.Windows.Forms.CheckBox
$Script:ChkExclude.Text = 'Exclude "Session Info.txt"'
$Script:ChkExclude.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 880, 25
$Script:ChkExclude.AutoSize = $true
$GrpDl.Controls.Add($Script:ChkExclude)

$Script:ChkRehash = New-Object -TypeName System.Windows.Forms.CheckBox
$Script:ChkRehash.Text = 'Re-download if file hash differs'
$Script:ChkRehash.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 880, 50
$Script:ChkRehash.AutoSize = $true
$GrpDl.Controls.Add($Script:ChkRehash)

$Script:BtnDlSelected = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnDlSelected.Text = 'Download SELECTED Sessions'
$Script:BtnDlSelected.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 12, 62
$Script:BtnDlSelected.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 230, 36
$Script:BtnDlSelected.Enabled = $false
$GrpDl.Controls.Add($Script:BtnDlSelected)

$Script:BtnDlAll = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnDlAll.Text = 'Download ALL Sessions of Checked Conferences'
$Script:BtnDlAll.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 250, 62
$Script:BtnDlAll.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 330, 36
$Script:BtnDlAll.Enabled = $false
$GrpDl.Controls.Add($Script:BtnDlAll)

$Script:BtnCancel = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnCancel.Text = 'Cancel'
$Script:BtnCancel.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 590, 62
$Script:BtnCancel.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 90, 36
$Script:BtnCancel.Enabled = $false
$GrpDl.Controls.Add($Script:BtnCancel)

# ---------- Progress / status / log ----------
<#
### Progress, Current Status, and Activity Log
Provide session/conference work-item progress, a replaceable status label, and a read-only
scrolling log. The log mirrors disk messages and expands vertically with the window.
Progress counts completed work items rather than transferred bytes.
#>
$Script:PbMain = New-Object -TypeName System.Windows.Forms.ProgressBar
$Script:PbMain.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 562
$Script:PbMain.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1160, 20
$Script:PbMain.Anchor = 'Top, Left, Right'
$Script:Form.Controls.Add($Script:PbMain)

$Script:LblStatus = New-Object -TypeName System.Windows.Forms.Label
$Script:LblStatus.Text = 'Starting ...'
$Script:LblStatus.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 588
$Script:LblStatus.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1160, 20
$Script:LblStatus.Anchor = 'Top, Left, Right'
$Script:Form.Controls.Add($Script:LblStatus)

$Script:TxtLog = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtLog.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 612
$Script:TxtLog.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1160, 180
$Script:TxtLog.Anchor = 'Top, Bottom, Left, Right'
$Script:TxtLog.Multiline = $true
$Script:TxtLog.ReadOnly = $true
$Script:TxtLog.ScrollBars = 'Vertical'
$Script:TxtLog.Font = New-Object -TypeName System.Drawing.Font -ArgumentList 'Consolas', 9
$Script:Form.Controls.Add($Script:TxtLog)
# ==================================================
# ^^^^^ Build GUI ^^^^^
# ==================================================

# ==================================================
# MARK: Event Handlers
# ==================================================
<#
## Event Handlers
### Purpose
Connect controls to authentication, queries, visibility updates, downloads, and the window lifecycle.
### Authentication
Sign In verifies against the first checked conference or 2026atmoa. Only Authenticated
mode activates the credential; success clears HTTP/schedule caches but retains loaded
session objects. Editing username/password resets authentication and clears those caches.
Download buttons are disabled until verified Sign In, and disabled again on credential edits.
Every download job authenticates all distinct target conferences before writing session content.
### Query and Selection Workflows
Query Selected uses checked entries; Discover IDs performs discovery only. Manual queries
normalize/validate an ID, append only new entries, preserve other results, and check the entry.
Conference checks filter loaded sessions; session checks independently select download targets.
### Visibility and Lifecycle
Show past years changes visibility without network queries. Hiding non-manual past entries
clears their conference/session selections while retaining loaded objects. Form Shown logs
startup and displays the catalog only. Cancel stops the active operation, including pending
HTTP work. Native window closing records shutdown and defers disposal until work unwinds;
there is no PowerShell FormClosing callback to invoke after its pipeline has stopped.
### Download Target Differences
Selected-session downloads use all checked session models, even outside the current filter.
Bulk conference downloads require checked accessible conferences, ignore the session text
filter, and display a confirmation before starting the shared download-job function.
#>
# ---------- In-window Sign In ----------
# Verify credentials for one conference; public fallback is not considered successful sign-in.
$Script:BtnSignIn.Add_Click({
        if ($Script:IsBusy) { return }
        $NewCredential = Get-CurrentCredential
        if (-not $NewCredential) {
            Set-Status -Text 'Enter a username and password.'
            return
        }
        Sync-ConferenceChecks
        $SignInConf = @($Script:ConferenceCatalog | Where-Object { $_.Checked } | Select-Object -First 1)[0]
        if (-not $SignInConf) { $SignInConf = Get-ConferenceInfo -ConfId '2026atmoa' }
        $Script:CancelRequested = $false
        Set-UiBusy -Busy $true
        try {
            Set-Status -Text ('Signing in to ' + $SignInConf.Name + ' ...')
            $Connect = Connect-SchedConference -ConfId $SignInConf.Id -Type $SignInConf.Type -Credential $NewCredential
            if (Test-CancellationRequested) { Set-Status -Text 'Sign-in cancelled.'; return }
            if ($Connect.Success -and $Connect.Mode -eq 'Authenticated') {
                $Script:Credential = $NewCredential
                $Script:Sessions.Clear()
                $Script:EventCache.Clear()
                $Script:BtnSignIn.Text = 'Signed In'
                Set-Status -Text ('Signed in to ' + $SignInConf.Name + ' as ' + $NewCredential.UserName)
                Write-Log -Message ('Sign-in verified for ' + $SignInConf.Id)
            }
            else {
                $Script:Sessions.Remove($SignInConf.Id)
                Set-Status -Text ('Sign-in not verified for ' + $SignInConf.Name + ': ' + $Connect.Message)
            }
        }
        catch {
            Set-Status -Text ('Sign-in error: ' + $_.Exception.Message)
        }
        finally {
            Set-UiBusy -Busy $false
        }
    })

# ---------- Credential edits invalidate authentication ----------
# Reset the credential/button and request caches without discarding the loaded session inventory.
$CredentialChanged = {
    $Script:Credential = $null
    $Script:Sessions.Clear()
    $Script:EventCache.Clear()
    $Script:BtnSignIn.Text = 'Sign In'
    $Script:BtnDlSelected.Enabled = $false
    $Script:BtnDlAll.Enabled = $false
}
$Script:TxtUser.Add_TextChanged($CredentialChanged)
$Script:TxtPass.Add_TextChanged($CredentialChanged)

# ---------- Query selected conferences / discover IDs ----------
# Keep explicit schedule queries separate from candidate discovery to avoid automatic bulk requests.
$Script:BtnQuery.Add_Click({ Invoke-ConferenceDiscovery -SelectedOnly })
$Script:BtnDiscover.Add_Click({ Invoke-ConferenceDiscovery -DiscoverOnly })

# ---------- Append or reuse a manual ID, select it, and query it ----------
# Refresh only this ID; keep the existing catalog order and other conferences' loaded sessions.
$Script:BtnAddConf.Add_Click({
        if ($Script:IsBusy) { return }
        $NewId = $Script:TxtCustom.Text.Trim().ToLower()
        if ($NewId -notmatch '^[a-z0-9]+$') {
            [void][System.Windows.Forms.MessageBox]::Show('Conference IDs may only contain letters and numbers (for example 2027atmoa).', 'MMS Session Downloader', 'OK', 'Warning')
            return
        }
        $Conf = Get-ConferenceInfo -ConfId $NewId
        if (-not $Conf) {
            $Conf = New-ConferenceEntry -Id $NewId -Name ('MMS ' + $NewId + ' (manual)') -Type 'Discovered'
            # Assign the current year to new yearless IDs; the Manual flag separately overrides visibility.
            if ($Conf.Year -eq 0) { $Conf.Year = (Get-Date).Year }
            $Script:ConferenceCatalog.Add($Conf)
        }
        $Conf.Manual = $true
        Sync-ConferenceChecks
        $Conf.Checked = $true
        foreach ($Item in $Script:LvConf.Items) {
            if ($Item.Tag.Id -eq $NewId) { $Item.Checked = $true }
        }
        $Script:TxtCustom.Text = ''

        Set-UiBusy -Busy $true
        try {
            $Script:CancelRequested = $false
            for ($Index = $Script:SessionList.Count - 1; $Index -ge 0; $Index--) {
                if ($Script:SessionList[$Index].ConfId -eq $NewId) { $Script:SessionList.RemoveAt($Index) }
            }
            Test-ConferenceAccess -Conf $Conf -Credential $Script:Credential
            Update-ConferenceListView
            Update-SessionListView
            Set-Status -Text ($NewId + ': ' + $Conf.Access)
        }
        finally {
            Set-UiBusy -Busy $false
        }
    })

# ---------- Conference checks drive the session list ----------
# Checkbox changes update the loaded-session display; they do not query Sched by themselves.
$Script:LvConf.Add_ItemChecked({
        if ($Script:SuppressEvents) { return }
        Sync-ConferenceChecks
        Update-SessionListView
    })

$BtnConfAll.Add_Click({
        $Script:SuppressEvents = $true
        foreach ($Item in $Script:LvConf.Items) { $Item.Checked = $true }
        $Script:SuppressEvents = $false
        Sync-ConferenceChecks
        Update-SessionListView
    })

$BtnConfNone.Add_Click({
        $Script:SuppressEvents = $true
        foreach ($Item in $Script:LvConf.Items) { $Item.Checked = $false }
        $Script:SuppressEvents = $false
        Sync-ConferenceChecks
        Update-SessionListView
    })

$Script:ChkShowAll.Add_CheckedChanged({ Update-ConferenceListView })

# ---------- Show past years visibility ----------
# Toggle display only; retain past results and preserve manually queried entries and their selections.
$Script:ChkPast.Add_CheckedChanged({
        if ($Script:ChkPast.Checked) {
            # Redisplay older catalog/session objects without querying or refreshing their schedules.
            Update-ConferenceListView
            Update-SessionListView
        }
        else {
            # Clear selections for year-hidden non-manual conferences/sessions, keeping loaded results.
            $Script:SuppressEvents = $true
            foreach ($Item in $Script:LvConf.Items) {
                if (Test-IsHiddenConference -Conf $Item.Tag) { $Item.Checked = $false }
            }
            foreach ($Item in $Script:LvSess.Items) {
                $ItemConf = Get-ConferenceInfo -ConfId $Item.Tag.ConfId
                if ($ItemConf -and (Test-IsHiddenConference -Conf $ItemConf)) { $Item.Checked = $false }
            }
            $Script:SuppressEvents = $false
            Sync-ConferenceChecks
            Sync-SessionChecks
            foreach ($Sess in $Script:SessionList) {
                $SessConf = Get-ConferenceInfo -ConfId $Sess.ConfId
                if ($SessConf -and (Test-IsHiddenConference -Conf $SessConf)) { $Sess.Checked = $false }
            }
            Update-ConferenceListView
            Update-SessionListView
            Set-Status -Text 'Past years hidden. Their results are kept.'
        }
    })

# ---------- Session list helpers ----------
# Text changes only filter display; Select All Shown checks visible rows and Clear Selection clears all.
$Script:TxtFilter.Add_TextChanged({ Update-SessionListView })

$Script:LvSess.Add_ItemChecked({
        if ($Script:SuppressEvents) { return }
        $_.Item.Tag.Checked = [bool]$_.Item.Checked
        $SelectedCount = @($Script:SessionList | Where-Object { $_.Checked }).Count
        $Script:LblSessCount.Text = 'Showing {0} of {1} session(s) | {2} selected' -f $Script:LvSess.Items.Count, $Script:SessionList.Count, $SelectedCount
    })

$BtnSessAll.Add_Click({
        $Script:SuppressEvents = $true
        foreach ($Item in $Script:LvSess.Items) { $Item.Checked = $true }
        $Script:SuppressEvents = $false
        Update-SessionListView
    })

$BtnSessNone.Add_Click({
        foreach ($Sess in $Script:SessionList) { $Sess.Checked = $false }
        $Script:SuppressEvents = $true
        foreach ($Item in $Script:LvSess.Items) { $Item.Checked = $false }
        $Script:SuppressEvents = $false
        Update-SessionListView
    })

# ---------- Browse / Open folder ----------
# Browse selects a base content destination; Open Folder requires an already existing directory.
$Script:BtnBrowse.Add_Click({
        $Dialog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
        $Dialog.Description = 'Select the base download folder'
        if (Test-Path -Path $Script:TxtLocation.Text) { $Dialog.SelectedPath = $Script:TxtLocation.Text }
        if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $Script:TxtLocation.Text = $Dialog.SelectedPath }
    })

$BtnOpenFolder.Add_Click({
        $FolderPath = $Script:TxtLocation.Text.Trim()
        if (Test-Path -Path $FolderPath) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $FolderPath + '"') }
        else { [void][System.Windows.Forms.MessageBox]::Show('The download folder does not exist yet.', 'MMS Session Downloader', 'OK', 'Information') }
    })

# ---------- Download SELECTED sessions ----------
# Collect all checked session models, including selections outside the current display filter.
$Script:BtnDlSelected.Add_Click({
        if ($Script:IsBusy) { return }

        Sync-SessionChecks
        $Targets = @($Script:SessionList | Where-Object { $_.Checked })
        if ($Targets.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('No sessions are checked. Check the sessions you want in the Sessions list.', 'MMS Session Downloader', 'OK', 'Information')
            return
        }
        Start-SessionDownload -Targets $Targets
    })

# ---------- Download ALL sessions of checked conferences ----------
# Collect loaded sessions from eligible checked conferences and require bulk-download confirmation.
$Script:BtnDlAll.Add_Click({
        if ($Script:IsBusy) { return }

        $Confs = @(Get-CheckedConferences)
        if ($Confs.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('Check at least one accessible conference first.', 'MMS Session Downloader', 'OK', 'Information')
            return
        }

        $ConfIds = @{}
        foreach ($Conf in $Confs) { $ConfIds[$Conf.Id] = $true }
        $Targets = @($Script:SessionList | Where-Object { $ConfIds.ContainsKey($_.ConfId) })

        if ($Targets.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('No sessions were found for the checked conferences.', 'MMS Session Downloader', 'OK', 'Information')
            return
        }

        $Answer = [System.Windows.Forms.MessageBox]::Show(('Download ALL files for {0} session(s) across {1} conference(s)?' -f $Targets.Count, $Confs.Count), 'MMS Session Downloader', 'YesNo', 'Question')
        if ($Answer -eq 'Yes') { Start-SessionDownload -Targets $Targets }
    })

# ---------- Cancel ----------
# Stop active work, including the cancellable request worker, while keeping the main window open.
$Script:BtnCancel.Add_Click({
        $Script:CancelRequested = $true
        Write-Log -Message 'Cancel requested. Stopping after the current operation ...' -Level 'WARN'
    })

# ---------- Form events ----------
# Native form closing tracks shutdown without invoking a stopped PowerShell pipeline.

# On initial display, log startup and show the preloaded catalog without authentication or network queries.
$Script:Form.Add_Shown({
        Write-Log -Message ('MMS Session Content Downloader started. Log file: ' + $Script:LogFile)
        Update-ConferenceListView
        $Script:Form.Activate()
        Set-Status -Text 'Public content'
    })
# ==================================================
# ^^^^^ Event Handlers ^^^^^
# ==================================================

# ==================================================
# MARK: Launch GUI
# ==================================================
<#
## Launch GUI
### Purpose
Start the main form's modal message loop after all controls and event handlers are ready.
### Startup Behavior
`ShowDialog()` raises the Shown handler, which displays the catalog and Public content
status. No legacy credential prompt, conference discovery, or schedule query is called here.
Authentication and requests occur only when their corresponding main-window actions run.
### Cleanup
The modal loop is protected by try/finally and treats PipelineStoppedException as shutdown.
Dispose the form even if execution is interrupted. Native closing defers disposal while an
operation is active, allowing request/partial-file cleanup to unwind before the dialog exits.
This section does not start exports; cancelled download jobs skip their result export.
#>
try {
    [void]$Script:Form.ShowDialog()
}
catch [System.Management.Automation.PipelineStoppedException] {
}
finally {
    $Script:CancelRequested = $true
    if ($Script:Form -and -not $Script:Form.IsDisposed) { $Script:Form.Dispose() }
}
# ==================================================
# ^^^^^ Launch GUI ^^^^^
# ==================================================

