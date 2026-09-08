#requires -Version 7.2

<#
.SYNOPSIS
    Repairs the MKV files that Find-PlexMkcleanCandidates.ps1 flagged, by remuxing
    away the mkclean content compression that stops the Plex for Windows app from
    playing them.

.DESCRIPTION
    Version 1.0.5

    Every candidate is taken through five steps, and no step is allowed to destroy
    the evidence of the one before it:

      1. The file is copied to $BackupRootFolder, mirroring its original layout:
         "Q:\Horror\Film\x.mkv" becomes "<BackupRootFolder>\Q\Horror\Film\x.mkv".
         The drive colon cannot appear in a folder name and is dropped; a UNC
         path loses its leading backslashes the same way. Mirroring rather than
         flattening matters because two films can easily carry the same name.

         An existing backup is never overwritten. Where the name is taken,
         " (1)", " (2)" and so on go in before the extension until a free one is
         found, so the pristine original of a file repaired twice survives
         whatever the second run does to it.

      2. The copy is verified. Size and last-write time are compared by default;
         -VerifyHash compares a SHA-256 of both sides instead, which reads both
         files in full and roughly doubles the network traffic.

      3. mkvmerge remuxes the BACKUP file - never the original - to a temporary
         file in the original's own folder:

             mkvmerge --ui-language en --compression -1:none -o <original>.remux.tmp <backup>

         --compression -1:none applies to every track and makes it impossible for
         the output to repeat the exact fault being repaired. Writing to a
         temporary file in the same folder means Plex never sees a half-written
         file, the replacement is a rename on the same volume, and an interrupted
         run leaves the original untouched.

      4. The temporary file is verified before it replaces anything: the HEVC
         track must have lost its content compression, mkvinfo must be able to
         read the HEVC profile again (the "Unknown @L0.0" that Plex chokes on has
         to be gone), and the track count and duration must still match the
         backup. Only then is the original replaced and its timestamps restored.

      5. The outcome is written back to the report file.

    Free space on the source volume
    -------------------------------
    The remux output is written beside the original, so the source volume has to
    hold an extra copy of every repair under way: step 5 deletes each original the
    moment its replacement is ready, which frees the room for the next one. How
    many are under way at once is -ThrottleLimit, or -BatchSize where a batch is
    the smaller of the two. That is what the pre-flight check requires - room for
    that many of the volume's largest candidates, not for the whole source folder
    tree.

    A NAS share with a recycle bin breaks that assumption. The delete in step 4
    does not free anything; the original is moved aside and kept in the recycle bin, so the volume
    ends up having to hold the entire source folder tree a second time, and a long run stops
    for lack of space partway through. The script warns when free space is below
    the source folder tree total, but it does not refuse to run: on an ordinary drive the
    warning is harmless.

    Synology NAS: Control Panel > Shared Folder > Edit > Enable Recycle Bin >
    (Clear Checkbox).

    Where the result goes
    ---------------------
    Five columns carry the outcome: RemuxStatus, RemuxTime, RemuxHevcProfile,
    RemuxSizeBytes and BackupPath. Find-PlexMkcleanCandidates.ps1 1.0.4 and later
    creates them empty, so they already exist in a current report .csv file. 
    In an older one they are appended.

    An .xlsx report is normally a view that Power Query builds over the .csv beside
    it, so where the result has to be written depends on whether that link is still
    in place. The script reads the table's SourceType and decides for itself:

      xlSrcQuery (3)  The workbook is still driven by its query. Writing into the
                      table would be pointless - the next refresh overwrites it -
                      so the columns go into the CSV that the query reads, and the
                      workbook is refreshed afterwards. Power Query passes columns
                      it does not name straight through, so they arrive intact and
                      correctly aligned.

      xlSrcRange (1)  The table has been unlinked (Table Design > Unlink) and is
                      now static. There is no query left to overwrite anything, so
                      the values are written straight into the table, matched on
                      Root + File Path Under Root - never on row position.

    Which CSV feeds a linked workbook is resolved from the query's own M code: a
    literal "....csv" in it wins. If the M code computes the name instead, the
    sibling CSV with the same base name as the workbook is used. When the two
    disagree the script says so, because a workbook saved under a new name keeps
    pointing at the old CSV and loads the wrong data without any error at all.

    Repairing from the same report twice
    ------------------------------------
    A row whose RemuxStatus already reads "Fixed..." is skipped, so a second run
    over the same report does not put an already repaired file through the whole
    cycle a second time for nothing.

    Until 1.0.4 this guard was load-bearing: step 1 overwrote unconditionally, so
    a second pass copied the repaired file straight over its own pristine backup
    and left no untouched original anywhere. Step 1 now numbers a colliding
    backup instead - " (1)", " (2)" - so the original survives either way and the
    guard is back to being what it looks like: a way of not repeating work.

    -Resume is no help here: it reads the .partial file, and a run that got as
    far as writing its results back deletes that file. The report itself is the
    only record left, which is why the guard reads it rather than the partial.

    -IncludeFixed turns the guard off. A report with no RemuxStatus column cannot
    be guarded at all - there is nothing to read - so that case is called out with
    a warning rather than passed over quietly: it means the report being used for repair
    is not the one the earlier run wrote its results to.

    Parallelism
    -----------
    Source and backup usually live on the same NAS, so a single 1 GiB file costs
    roughly 2 GiB of traffic for the copy, another 2 GiB for -VerifyHash, and 2 GiB
    more for the remux - all over one link. Beyond two or three files at a time the
    link is already saturated and further parallelism buys nothing while making
    every individual file repair slower. Hence the deliberately modest default of
    $ThrottleLimit = 2. $BatchSize is a different lever: it is how many files are
    handed to the parallel engine at a time, and it should stay comfortably above
    $ThrottleLimit so a slot that finishes early is immediately given new work.

    Results are appended to a .partial file as each file completes - not once per
    batch - so an interrupted or crashed run loses nothing and can be continued
    with -Resume.

    Process priority
    ----------------
    -Priority sets the Windows priority class of this script, and every mkvmerge
    and mkvinfo it starts inherits it: a child process inherits its parent's class
    only when that class is Idle or BelowNormal. Raising the priority is therefore
    not offered - it would apply to the PowerShell process and never reach the
    tools that do the work, which is a promise the parameter could not keep.

    Note what it does and does not do. Priority is CPU scheduling. This run is
    almost entirely network- and disk-bound, so a lower priority keeps mkvmerge
    and the SHA-256 hashing out of the way of everything else on the machine, but
    it does nothing at all for the load on the NAS. $ThrottleLimit is the only
    lever that reduces that.

.PARAMETER SourceData
    The report to read the candidates from: the .xlsx workbook, or the .csv it is
    built from. Mandatory: the reports are per source folder, so there is no value that
    would be right more often than it would be wrong, and the cost of being
    pointed at the wrong one could be hours of NAS traffic.

    Every Candidate=True row is repaired, whether or not it is also a
    StrictCandidate. A candidate that is not strict is not a weaker hit - it
    usually only means mkvinfo could not list the file's tracks - so there is
    nothing to gain by holding it back.

.PARAMETER BackupRootFolder
    Where the untouched originals are kept. The folder structure below it mirrors
    Root and the file path under it. Nothing here is ever deleted by this script.
    There is no default: the folder has to be named on every run.

.PARAMETER MKVToolNix
    Folder holding mkvmerge.exe and mkvinfo.exe. Defaults to
    'C:\Program Files\MKVToolNix'.

.PARAMETER ThrottleLimit
    Files repaired at the same time. See "Parallelism" above; 1 makes the run
    strictly sequential.

.PARAMETER BatchSize
    Files handed to the parallel engine per batch. Keep it above $ThrottleLimit.

.PARAMETER VerifyHash
    Compare a SHA-256 of the backup and the original instead of size and
    last-write time. Reads both files in full.

.PARAMETER Priority
    Windows priority class for this script and the MKVToolNix processes it starts.
    Idle or BelowNormal to stay out of the way of everything else on the machine;
    Normal is the default. See "Process priority" above for why nothing higher is
    offered.

.PARAMETER Resume
    Continue a previous run: files already recorded in the .partial file are
    skipped.

.PARAMETER IncludeFixed
    Repair rows whose RemuxStatus already reads "Fixed..." instead of skipping
    them. Each one gets a second backup alongside the first - "x (1).mkv" beside
    "x.mkv" - holding the already repaired file. The original backup is left
    alone. See "Repairing the same report twice" above.

.EXAMPLE
    .\Fix-PlexMkcleanCandidates.ps1 -SourceData '.\Prod\Plex-mkclean-candidates - Horror.xlsx' -WhatIf
    Lists exactly which files would be backed up, remuxed and replaced, and
    touches nothing.

.EXAMPLE
    .\Fix-PlexMkcleanCandidates.ps1 -SourceData '.\Prod\Plex-mkclean-candidates - Horror.xlsx'
    Repairs every Candidate=True row in that report that is not already marked
    Fixed.

.EXAMPLE
    .\Fix-PlexMkcleanCandidates.ps1 -SourceData '.\Prod\Plex-mkclean-candidates - Horror.csv' -ThrottleLimit 1 -VerifyHash
    One file at a time, with the backup checked by SHA-256 before the original
    is overwritten.

.EXAMPLE
    .\Fix-PlexMkcleanCandidates.ps1 -SourceData '.\Prod\Plex-mkclean-candidates - Horror.csv' -Priority BelowNormal
    Repairs everything while staying out of the way of whatever else the machine
    is doing.

.EXAMPLE
    .\Fix-PlexMkcleanCandidates.ps1 -SourceData '.\Prod\Plex-mkclean-candidates - Horror.csv' -Resume
    Continues a run that was interrupted, skipping files already recorded in
    the .partial file.

.NOTES
    Version: 1.0.5

    This script overwrites media files. Every original is copied to
    $BackupRootFolder and the copy is verified before the original is touched, and
    the backups are never deleted - delete them yourself once Plex has played the
    repaired files without problems.

    Version history:
      1.0.5  The three places that decided for the user now ask him, the same way
             Find-PlexMkcleanCandidates.ps1 1.0.9 does and through the same
             Read-UserChoice. Partial results from an interrupted run were
             announced and deleted in the same breath, so the warning arrived
             after the work was already gone; resuming, discarding and stopping
             are now offered. A report that could not be written - which in
             practice means Excel has it open - was written beside the intended
             one under a timestamped name, leaving the workbook's query pointed
             at the stale copy; closing Excel and retrying is now the first
             option. A workbook that was open simply ended the run with two
             warnings; it can now be closed and the write retried, without
             repeating an hour of remuxing. All three answer themselves, safely,
             when there is no console to answer them - so an unattended run
             still cannot hang on a question.
             The free-space check on a source volume asked for room for the
             largest candidate, which would only have been right had the repairs
             run one at a time. Up to -ThrottleLimit of them are under way at
             once, each holding its own remux output until the original it
             replaces is deleted, so the check now asks for the sum of that many
             of the volume's largest candidates.
             Get-FreeSpaceBytes and Restore-RunState are gone from here and come
             from PlexMkclean.psm1 1.0.3 instead, where the identical copies in
             the sister scripts join them. Restore-RunState took a new
             Get-RunState with it, because the host state it puts back can no
             longer be read from $script: variables across the module boundary.
             The module import moved up above the first line that changes
             anything about the host, since Get-RunState has to record the state
             before it is disturbed.
             The eight Restore-RunState calls that stood in front of a throw
             are gone. The script-scope trap already restores on any
             terminating error, so they only did the same work a moment
             earlier. The five in front of an early return stay: a return is a
             normal exit the trap never sees, and they all sit above the
             try/finally, so nothing else would put the run state back. The
             comment on the trap states that rule now.
             The file had also picked up LF line endings somewhere along the
             way; it is CRLF again, as .gitattributes and the three sister
             files have it.
             The defaults that pointed into this machine are gone.
             -MKVToolNix is the ordinary 'C:\Program Files\MKVToolNix', and
             -BackupRootFolder no longer names a share of its own: it is empty,
             so the folder has to be given on every run. A check at the top of
             the pre-flight says that outright, since an empty value otherwise
             travelled as far as Test-Path and failed there as a binding error
             naming -Path.
      1.0.4  -WhatIf no longer dies on a backup root that does not exist yet.
             New-Item honours -WhatIf and creates nothing, while the Resolve-Path
             beside it resolves only a path that is already there, so a dry run
             failed on the very folder a real run would have made for itself.
             That call and the one inside Get-FreeSpaceBytes now resolve without
             requiring the path to exist, and the writability probe - which has
             to write a real file - is skipped when the folder is still missing
             under -WhatIf, rather than reported as unwritable.
             An existing backup is no longer overwritten - see step 1 of the
             description. The change itself is in PlexMkclean.psm1 1.0.1, so the
             single-file script gets it too; the guard here that skips rows
             already marked Fixed is now a convenience rather than the only
             thing between a second run and the loss of an original.
             The repair engine reaches the parallel runspaces by importing
             PlexMkclean.psm1 inside the -Parallel block, rather than by
             shipping the text of Repair-MkvCandidate into each runspace with
             ${function:...}.ToString(). A runspace has a session state of its
             own, so the module has to arrive somehow; importing it brings the
             whole module instead of one nameless function body, which means
             the engine may now call private helpers and be refactored like
             ordinary code. The old way forbade that silently - a call to a
             sister function parsed cleanly and then failed per file, out in a
             worker, as a non-terminating error that read like a bad media file
             rather than a bug. The cost is one module load per runspace in the
             pool, not per file, which is nothing beside a remux.
      1.0.3  Repair-MkvCandidate moved verbatim to PlexMkclean.psm1, so the same
             repair engine can serve a one-off file as well as a report. The
             script imports the module and goes on serialising the function into
             its runspaces exactly as before.
             The comment-based help had been dead since it was written: one line
             of an example began with ".partial", which the help parser reads as
             an unknown keyword and rejects the whole block over, so -? showed
             nothing but the generated syntax line. Rewrapped, and a section on
             free space and the recycle bin trap added to .DESCRIPTION - the
             default help view shows DESCRIPTION but not NOTES, which is where
             that warning would otherwise have been invisible.
      1.0.2  The progress line carries its own 'ETA hh:mm:ss' in the status text
             instead of -SecondsRemaining, which PowerShell renders as a bare
             number with nothing on it to say what it means. A source volume
             with less free space than the whole source folder tree is now warned about: the
             check beside it only reserves room for one file at a time, which is
             the wrong assumption on a share whose recycle bin keeps the
             originals rather than freeing their space.
      1.0.1  Candidates the report already marks as Fixed are skipped, so
             repairing a source folder tree in more than one sitting no longer overwrites a
             good backup with the file that was repaired from it; -IncludeFixed
             repairs them anyway. A report with no RemuxStatus column at all is
             called out, since the guard cannot work on one. -SourceData lost
             its default and is mandatory instead: the reports are per source folder tree, so
             there is no report that would be right to fall back on.
      1.0.0  First version. Backup, verified copy, remux from the backup,
             verified result, in-place replacement, report write-back to either
             the CSV behind a linked workbook or an unlinked table, and
             -Priority for keeping the run out of the way of the rest of the
             machine.
#>

# Mandatory is reserved for a parameter that is important and has no sensible
# default: -SourceData here, -Source in Find-PlexMkcleanCandidates.ps1. Both
# are about which source folder(s) a run is about, and there is no value that 
# would be right more often than it would be wrong. 
# Everything else keeps its default, because a mandatory parameter never 
# uses one - PowerShell prompts instead, which would make the value on it 
# a lie.
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    # The report to read: the .xlsx workbook, or the .csv it is built from.
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceData,

    # Where the untouched originals are kept.
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$BackupRootFolder = '',

    # Folder holding mkvmerge.exe and mkvinfo.exe.
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$MKVToolNix = 'C:\Program Files\MKVToolNix',

    # Files repaired at the same time. Source and backup are usually on the same NAS,
    # so a high number only makes each file slower. See .DESCRIPTION.
    [Parameter()]
    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 2,

    # Files handed to the parallel engine per batch. Keep it above $ThrottleLimit
    # so a slot that finishes early gets new work immediately.
    [Parameter()]
    [ValidateRange(1, 1000)]
    [int]$BatchSize = 10,

    # Verify the backup by SHA-256 rather than by size and last-write time.
    [Parameter()]
    [switch]$VerifyHash,

    # Windows priority class for this script and, by inheritance, every mkvmerge
    # and mkvinfo it starts. Only the polite end of the scale is offered - see the
    # note under "Process priority" in .DESCRIPTION.
    [Parameter()]
    [ValidateSet('Idle', 'BelowNormal', 'Normal')]
    [string]$Priority = 'Normal',

    # Continue a previous run: files already present in the .partial file are skipped.
    [Parameter()]
    [switch]$Resume,

    # Repair candidates the report already marks as Fixed. Off by default: it is
    # work already done, and it leaves a numbered second backup behind for every
    # file it touches.
    [Parameter()]
    [switch]$IncludeFixed
)

$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.5'

$MkvMerge = Join-Path $MKVToolNix 'mkvmerge.exe'
$MkvInfo  = Join-Path $MKVToolNix 'mkvinfo.exe'

$RemuxColumns = @('RemuxStatus', 'RemuxTime', 'RemuxHevcProfile', 'RemuxSizeBytes', 'BackupPath')

# The Repair-MkvCandidate function lives in PlexMkclean.psm1 beside this script,
# shared with Find-PlexMkcleanCandidates.ps1 and the single-file tooling, as do
# the Read-UserChoice and Get-FreeSpaceBytes helpers used further down.
# This import serves the main thread only - each parallel runspace imports the
# module for itself.
# The path is resolved here because $PSScriptRoot is empty inside a -Parallel
# block, so the repair loop below could not work it out for itself.
# It happens this early because Get-RunState comes from the module too, and
# nothing may be changed about the host before that has run.
$modulePath = Join-Path $PSScriptRoot 'PlexMkclean.psm1'
Import-Module $modulePath -Force

# Everything this script is about to change about its host, so Restore-RunState
# can put it back at every exit path.
$runState = Get-RunState

# The progress bar is drawn in bold yellow by default. Use light cyan instead;
# Restore-RunState puts the original back when the script is done.
$PSStyle.Progress.Style = $PSStyle.Foreground.BrightCyan

# Any terminating error that reaches script scope leaves the host as the script
# found it before the error is passed on. Without this, an aborted run would
# leave the lowered process priority behind in an interactive session, where it
# would silently apply to everything started afterwards. Every throw below is
# covered by this, so none of them restore for themselves. An early return is
# not - that is a normal exit and no error - so those few places restore before
# they return, and the finally at the end covers the normal path. Ctrl+C runs
# none of the three.
trap { Restore-RunState $runState; break }


Write-Host ''
Write-Host "Fix-PlexMkcleanCandidates $ScriptVersion"

if ($Priority -ne 'Normal') {
    # Child processes inherit Idle and BelowNormal from their parent, so this one
    # line reaches every mkvmerge and mkvinfo the run starts.
    $runState.Process.PriorityClass = [System.Diagnostics.ProcessPriorityClass]$Priority
    Write-Host "Process priority: $Priority (mkvmerge and mkvinfo inherit it)"
}

#region Report helpers

function Get-CanonicalColumnName {
    <#
        The Excel table's headers carry line breaks ("Scan<lf>Status",
        "File Path Under Root") while the CSV's do not, so both are folded to one
        whitespace-free lower-case key before they are matched.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Header)

    ($Header -replace '\s', '').ToLowerInvariant()
}

function ConvertTo-ReportBoolean {
    param($Value)

    if ($Value -is [bool]) { return $Value }
    return ("$Value").Trim() -in @('true', '1', 'yes')
}

function ConvertTo-CandidateSet {
    <#
        Folds either report shape - Excel table rows or Import-Csv rows - into the
        same handful of properties the rest of the script works with.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)

    if ($Rows.Count -eq 0) { return @() }

    $wanted = @{
        'scanstatus'        = 'ScanStatus'
        'candidate'         = 'Candidate'
        'strictcandidate'   = 'StrictCandidate'
        'sizebytes'         = 'SizeBytes'
        'remuxstatus'       = 'RemuxStatus'
        'root'              = 'Root'
        'filename'          = 'FileName'
        'filepathunderroot' = 'FileName'
    }

    $lookup = @{}
    foreach ($property in $Rows[0].PSObject.Properties) {
        $key = Get-CanonicalColumnName $property.Name
        if ($wanted.ContainsKey($key)) { $lookup[$wanted[$key]] = $property.Name }
    }

    foreach ($required in @('Root', 'FileName', 'Candidate')) {
        if (-not $lookup.ContainsKey($required)) {
            throw "The report has no $required column; it does not look like a Find-PlexMkcleanCandidates report."
        }
    }

    $Rows | ForEach-Object {
        [pscustomobject]@{
            Root            = [string]$_.($lookup['Root'])
            FileName        = [string]$_.($lookup['FileName'])
            Candidate       = ConvertTo-ReportBoolean $_.($lookup['Candidate'])
            StrictCandidate = if ($lookup.ContainsKey('StrictCandidate')) { ConvertTo-ReportBoolean $_.($lookup['StrictCandidate']) } else { $false }
            SizeBytes       = if ($lookup.ContainsKey('SizeBytes')) { [int64]0 + ($_.($lookup['SizeBytes']) -as [int64]) } else { [int64]0 }
            ScanStatus      = if ($lookup.ContainsKey('ScanStatus')) { [string]$_.($lookup['ScanStatus']) } else { '' }
            RemuxStatus     = if ($lookup.ContainsKey('RemuxStatus')) { [string]$_.($lookup['RemuxStatus']) } else { '' }
        }
    }
}

function Get-CandidateKey {
    param([Parameter(Mandatory)][psobject]$Row)

    (Join-Path $Row.Root $Row.FileName).ToLowerInvariant()
}

function Test-FileIsLocked {
    <#
        Excel keeps the workbook open for the whole session, so an exclusive open
        is a reliable "somebody has this file open" test. Used before writing,
        never before reading - reading is done read-only and works either way.
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }

    try {
        $stream = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $stream.Dispose()
        return $false
    }
    catch {
        return $true
    }
}

function Test-FolderIsWritable {
    param([Parameter(Mandatory)][string]$Folder)

    $probe = Join-Path $Folder ('write-probe-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
    try {
        # Plain .NET on both halves. Remove-Item would honour -WhatIf and leave
        # the probe file behind on every folder a dry run touched.
        [System.IO.File]::WriteAllText($probe, 'probe')
        return $true
    }
    catch {
        return $false
    }
    finally {
        try { [System.IO.File]::Delete($probe) } catch { }
    }
}

function Read-ExcelReport {
    <#
        Reads the candidate table out of a workbook, read-only so the workbook may
        be open in Excel at the same time, and reports how the table gets its data
        so the caller knows where the results have to be written back.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $excel    = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        # UpdateLinks 0, ReadOnly $true.
        $workbook = $excel.Workbooks.Open($Path, 0, $true)

        $table   = $null
        $headers = @()

        foreach ($sheet in $workbook.Worksheets) {
            foreach ($listObject in $sheet.ListObjects) {
                $columnCount  = $listObject.ListColumns.Count
                $headerValues = $listObject.HeaderRowRange.Value2

                $names = for ($c = 1; $c -le $columnCount; $c++) {
                    Get-CanonicalColumnName "$($headerValues[1, $c])"
                }
                $names = @($names)

                if ($names -contains 'root' -and
                    (($names -contains 'filepathunderroot') -or ($names -contains 'filename'))) {
                    $table   = $listObject
                    $headers = $names
                    break
                }
            }
            if ($null -ne $table) { break }
        }

        if ($null -eq $table) {
            throw "No table with Root and File Path Under Root columns was found in $Path"
        }

        $rows = [System.Collections.Generic.List[psobject]]::new()

        if ($table.ListRows.Count -gt 0) {
            $values = $table.DataBodyRange.Value2
            for ($r = 1; $r -le $table.ListRows.Count; $r++) {
                $row = [ordered]@{}
                for ($c = 1; $c -le $headers.Count; $c++) {
                    $row[$headers[$c - 1]] = $values[$r, $c]
                }
                $rows.Add([pscustomobject]$row)
            }
        }

        # 3 = xlSrcQuery (still fed by Power Query), 1 = xlSrcRange (unlinked).
        $sourceType = [int]$table.SourceType

        $csvPath = $null
        $csvFrom = $null

        if ($sourceType -eq 3) {
            $folder  = Split-Path -Path $Path -Parent
            $sibling = Join-Path $folder ((Split-Path -Path $Path -LeafBase) + '.csv')

            $literal = $null
            try {
                foreach ($query in $workbook.Queries) {
                    # Comments are stripped first and the match is anchored to
                    # File.Contents: a comment mentioning a file name is not the
                    # source, and M code that documents itself will say one.
                    $formula = [regex]::Replace($query.Formula, '/\*.*?\*/', '', 'Singleline')
                    $formula = [regex]::Replace($formula, '//[^\r\n]*', '')

                    $match = [regex]::Match($formula, 'File\.Contents\s*\([^)]*?"(?<name>[^"]+\.csv)"', 'IgnoreCase')
                    if ($match.Success) { $literal = $match.Groups['name'].Value; break }
                }
            }
            catch {
                # Workbook.Queries is not available on every Excel build; the
                # naming convention below covers it.
            }

            if ($literal) {
                $csvPath = if ([System.IO.Path]::IsPathRooted($literal)) { $literal } else { Join-Path $folder $literal }
                $csvFrom = "the workbook's own query"

                if ((Get-CanonicalColumnName $csvPath) -ne (Get-CanonicalColumnName $sibling)) {
                    Write-Warning "The query reads '$([System.IO.Path]::GetFileName($csvPath))' but the workbook is named '$([System.IO.Path]::GetFileNameWithoutExtension($Path))'."
                    Write-Warning 'A workbook saved under a new name keeps pointing at the old CSV and loads the wrong data without any error. See the .DESCRIPTION.'
                }
            }
            else {
                $csvPath = $sibling
                $csvFrom = 'the workbook name'
            }
        }

        return [pscustomobject]@{
            Rows       = $rows.ToArray()
            SourceType = $sourceType
            TableName  = $table.Name
            SheetName  = $table.Parent.Name
            CsvPath    = $csvPath
            CsvFrom    = $csvFrom
        }
    }
    finally {
        if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
        if ($null -ne $excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
    }
}

function Write-ResultsToCsv {
    <#
        Writes the remux columns into the CSV, matched on Root + FileName. Returns
        the path actually written, which may differ from $CsvPath if the file was
        locked.
    #>
    param(
        [Parameter(Mandatory)][string]$CsvPath,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Results
    )

    $rows = @(Import-Csv -LiteralPath $CsvPath)
    if ($rows.Count -eq 0) {
        throw "The CSV behind the report is empty: $CsvPath"
    }

    # Reports written before Find-PlexMkcleanCandidates.ps1 1.0.4 have no remux
    # columns. Append them rather than refuse to write; they land at the end
    # instead of before Root, which only affects where they show up in the table.
    $existing = $rows[0].PSObject.Properties.Name
    $missing  = @($RemuxColumns | Where-Object { $_ -notin $existing })

    if ($missing.Count -gt 0) {
        Write-Warning "The CSV predates the remux columns; adding: $($missing -join ', ')"
        foreach ($row in $rows) {
            foreach ($column in $missing) {
                $row | Add-Member -NotePropertyName $column -NotePropertyValue '' -Force
            }
        }
    }

    $index = @{}
    foreach ($row in $rows) {
        $index[(Get-CandidateKey $row)] = $row
    }

    $applied = 0
    $orphans = 0

    foreach ($result in $Results) {
        $key = Get-CandidateKey $result
        if (-not $index.ContainsKey($key)) { $orphans++; continue }

        $row = $index[$key]
        foreach ($column in $RemuxColumns) {
            $row.$column = [string]$result.$column
        }
        $applied++
    }

    if ($orphans -gt 0) {
        Write-Warning "$orphans repaired file(s) had no matching row in $CsvPath and could not be recorded there."
    }

    $written = $null
    $target  = $CsvPath

    while (-not $written) {
        try {
            $rows | Export-Csv -LiteralPath $target -NoTypeInformation -Encoding utf8BOM
            $written = $target
        }
        catch {
            # Nearly always the report is open in Excel, which holds it
            # exclusively. Nothing is lost whichever way this is answered - the
            # caller keeps the .partial file whenever Path comes back empty -
            # but a second report leaves the workbook's query pointing at the
            # stale one, so closing Excel and retrying is offered first.
            $firstAttempt = $target -eq $CsvPath

            # Writing beside it is offered once only. The timestamped name is one
            # this run invented, so if that cannot be written either then a third
            # name will not help - and an unattended run, which always takes the
            # default, would otherwise circle here for ever.
            $option = if ($firstAttempt) { '&Retry', '&Beside it', '&Cancel' }
                      else               { '&Retry', '&Cancel' }

            $help = if ($firstAttempt) {
                        'Close whatever is holding the file, then try the write again.',
                        'Write the results to a second file, with a timestamp in its name.',
                        'Write no report now. The results stay in the .partial file.'
                    }
                    else {
                        'Close whatever is holding the file, then try the write again.',
                        'Write no report now. The results stay in the .partial file.'
                    }

            # Enter retries, because somebody standing here has just been told
            # the file is locked and the useful thing to do is close Excel. With
            # nobody standing here, retrying would only meet the same lock, so
            # the answer that keeps the results is taken instead.
            $answer = Read-UserChoice -Option $option -Help $help -Default 'Retry' `
                -Unattended $(if ($firstAttempt) { 'Beside it' } else { 'Cancel' }) `
                -Title   'The report could not be written - usually because Excel has it open' `
                -Message "$target could not be written ($($_.Exception.Message)):"

            if ($answer -eq 'Cancel') { break }

            if ($answer -eq 'Beside it') {
                $stamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
                $target = Join-Path (Split-Path -Path $CsvPath -Parent) `
                    ('{0}-{1}{2}' -f [IO.Path]::GetFileNameWithoutExtension($CsvPath), $stamp, [IO.Path]::GetExtension($CsvPath))
            }
        }
    }

    # Path is $null when nothing could be written. The caller reads that as "the
    # results are still only in the .partial file" and keeps it.
    return [pscustomobject]@{ Path = $written; Applied = $applied }
}

function Write-ResultsToWorkbook {
    <#
        For an unlinked (xlSrcRange) table: writes the remux columns straight into
        the table, matched on Root + File Path Under Root. Never by row position -
        a re-scan reorders the report, and a positional write would then put every
        result against the wrong film.
    #>
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Results
    )

    $excel    = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($WorkbookPath)

        $table   = $null
        $headers = @()

        foreach ($sheet in $workbook.Worksheets) {
            foreach ($listObject in $sheet.ListObjects) {
                $headerValues = $listObject.HeaderRowRange.Value2
                $names = @(for ($c = 1; $c -le $listObject.ListColumns.Count; $c++) {
                    Get-CanonicalColumnName "$($headerValues[1, $c])"
                })
                if ($names -contains 'root' -and
                    (($names -contains 'filepathunderroot') -or ($names -contains 'filename'))) {
                    $table   = $listObject
                    $headers = $names
                    break
                }
            }
            if ($null -ne $table) { break }
        }

        if ($null -eq $table) {
            throw "No candidate table was found in $WorkbookPath"
        }

        # Add any remux column the table does not have yet, at the right-hand end.
        foreach ($column in $RemuxColumns) {
            if ((Get-CanonicalColumnName $column) -notin $headers) {
                $added = $table.ListColumns.Add()
                $added.Name = $column
                $headers += (Get-CanonicalColumnName $column)
            }
        }

        $columnIndex = @{}
        for ($c = 1; $c -le $table.ListColumns.Count; $c++) {
            $columnIndex[(Get-CanonicalColumnName $table.ListColumns.Item($c).Name)] = $c
        }

        $rootColumn = $columnIndex['root']
        $fileColumn = if ($columnIndex.ContainsKey('filepathunderroot')) { $columnIndex['filepathunderroot'] } else { $columnIndex['filename'] }

        $rowIndex = @{}
        $values   = $table.DataBodyRange.Value2
        for ($r = 1; $r -le $table.ListRows.Count; $r++) {
            $key = (Join-Path "$($values[$r, $rootColumn])" "$($values[$r, $fileColumn])").ToLowerInvariant()
            $rowIndex[$key] = $r
        }

        $applied = 0
        $orphans = 0

        foreach ($result in $Results) {
            $key = Get-CandidateKey $result
            if (-not $rowIndex.ContainsKey($key)) { $orphans++; continue }

            $r = $rowIndex[$key]
            foreach ($column in $RemuxColumns) {
                $c = $columnIndex[(Get-CanonicalColumnName $column)]
                $table.DataBodyRange.Cells($r, $c).Value2 = [string]$result.$column
            }
            $applied++
        }

        if ($orphans -gt 0) {
            Write-Warning "$orphans repaired file(s) had no matching row in the workbook and could not be recorded there."
        }

        $workbook.Save()
        return [pscustomobject]@{ Path = $WorkbookPath; Applied = $applied }
    }
    finally {
        if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
        if ($null -ne $excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
    }
}

function Update-WorkbookQuery {
    <#
        Refreshes a still-linked workbook so the values just written to the CSV
        become visible, and waits for Power Query to finish before saving.
    #>
    param([Parameter(Mandatory)][string]$WorkbookPath)

    $excel    = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $workbook.RefreshAll()
        $excel.CalculateUntilAsyncQueriesDone()
        $workbook.Save()
        return $true
    }
    catch {
        Write-Warning "Could not refresh the workbook automatically ($($_.Exception.Message)). Open it and use Data > Refresh All."
        return $false
    }
    finally {
        if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
        if ($null -ne $excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
    }
}

#endregion Report helpers

#region Pre-flight

# -BackupRootFolder is checked first and by hand, because it has no default and
# cannot sensibly be given one - it has to be a folder with room for a copy of
# every file that gets repaired, which only the caller knows. Left empty it would
# otherwise travel as far as the Test-Path below and fail there as a binding
# error naming -Path, which says nothing about which parameter was forgotten.
if ([string]::IsNullOrWhiteSpace($BackupRootFolder)) {
    throw 'BackupRootFolder is required. Pass -BackupRootFolder naming the folder that is to hold the untouched originals.'
}

foreach ($tool in @($MkvMerge, $MkvInfo)) {
    if (-not (Test-Path -LiteralPath $tool -PathType Leaf)) {
        throw "MKVToolNix executable was not found: $tool"
    }
}

if (-not (Test-Path -LiteralPath $SourceData -PathType Leaf)) {
    throw "Source report was not found: $SourceData"
}

$sourceDataPath = (Resolve-Path -LiteralPath $SourceData).ProviderPath
$sourceExtension = [System.IO.Path]::GetExtension($sourceDataPath).ToLowerInvariant()

Write-Host "Reading report: $sourceDataPath"

$workbookPath = $null
$csvPath      = $null
$writeTarget  = $null
$reportRows   = @()

switch ($sourceExtension) {
    { $_ -in @('.xlsx', '.xlsm', '.xlsb') } {
        $workbookPath = $sourceDataPath
        $report = Read-ExcelReport -Path $workbookPath
        $reportRows = $report.Rows

        Write-Host "  Table '$($report.TableName)' on sheet '$($report.SheetName)', $($reportRows.Count) row(s)."

        if ($report.SourceType -eq 3) {
            $writeTarget = 'Csv'
            $csvPath = $report.CsvPath
            Write-Host "  The table is still fed by Power Query, so results go to the CSV it reads,"
            Write-Host "  resolved from $($report.CsvFrom): $csvPath"

            if (-not (Test-Path -LiteralPath $csvPath -PathType Leaf)) {
                throw "The CSV behind the workbook was not found: $csvPath"
            }
        }
        else {
            $writeTarget = 'Workbook'
            Write-Host '  The table has been unlinked from its query, so results are written into it directly.'
        }
    }
    '.csv' {
        $csvPath = $sourceDataPath
        $writeTarget = 'Csv'
        $reportRows = @(Import-Csv -LiteralPath $csvPath)
        Write-Host "  $($reportRows.Count) row(s)."
    }
    default {
        throw "Unsupported report type '$sourceExtension'. Pass the .xlsx workbook or the .csv it is built from."
    }
}

$allRows    = @(ConvertTo-CandidateSet -Rows $reportRows)
$candidates = @($allRows | Where-Object { $_.Candidate })

Write-Host "  $($candidates.Count) candidate(s) of $($allRows.Count) row(s)."

# A row the report already marks as Fixed has had its original replaced and its
# backup taken, so repairing it again is work for nothing. Since 1.0.4 it is no
# longer dangerous - step 1 numbers a colliding backup rather than overwriting
# it - but it still costs a full copy and remux and leaves a spare backup behind.
# -Resume cannot cover this, because a run that wrote its results back deletes
# the .partial file it reads.
# Without the column there is nothing to read, and every already repaired file
# would be repaired again in silence. That is not a hypothetical: a source folder's .csv
# and its unlinked .xlsx drift apart the moment the results are written to one of
# them, and the one without the results looks like a perfectly good report.
$hasRemuxStatusColumn = $reportRows.Count -gt 0 -and @(
    $reportRows[0].PSObject.Properties |
        Where-Object { (Get-CanonicalColumnName $_.Name) -eq 'remuxstatus' }
).Count -gt 0

if (-not $hasRemuxStatusColumn) {
    Write-Warning 'The report has no RemuxStatus column, so files repaired by an earlier run cannot be recognised and would be repaired again, each leaving a second numbered backup behind. Use the report that run wrote its results to.'
}

$alreadyFixed = @($candidates | Where-Object { $_.RemuxStatus -like 'Fixed*' })

if ($alreadyFixed.Count -gt 0) {
    if ($IncludeFixed) {
        Write-Warning "$($alreadyFixed.Count) file(s) are already marked Fixed and will be repaired again (-IncludeFixed); each gets a second, numbered backup and the first one is left alone."
    }
    else {
        Write-Host "  $($alreadyFixed.Count) already repaired, skipped. -IncludeFixed repairs them again."
        $candidates = @($candidates | Where-Object { $_.RemuxStatus -notlike 'Fixed*' })
    }
}

if ($candidates.Count -eq 0) {
    Write-Host 'Nothing to do.'
    Write-Host ''
    Restore-RunState $runState
    return
}

#endregion Pre-flight

#region Access and space checks

Write-Host 'Checking access and free space...'

if (-not (Test-Path -LiteralPath $BackupRootFolder -PathType Container)) {
    New-Item -ItemType Directory -Path $BackupRootFolder -Force | Out-Null

    # New-Item honours -WhatIf and creates nothing, so the folder is tested
    # again rather than assumed: announcing a folder that was not made is
    # exactly the kind of thing a dry run must not do.
    if (Test-Path -LiteralPath $BackupRootFolder -PathType Container) {
        Write-Host "  Created backup root: $BackupRootFolder"
    }
}

# Resolved to a full path before it is used: a relative backup root would be
# recorded in the report's BackupPath column as something that only means
# anything from the folder the script happened to be run in.
#
# GetUnresolvedProviderPathFromPSPath rather than Resolve-Path, which resolves
# only a path that already exists. Under -WhatIf the New-Item above created
# nothing, so Resolve-Path threw and the dry run died on the one folder the real
# run would have made for itself.
$BackupRootFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($BackupRootFolder)

if (Test-Path -LiteralPath $BackupRootFolder -PathType Container) {
    if (-not (Test-FolderIsWritable $BackupRootFolder)) {
        throw "The backup root is not writable: $BackupRootFolder"
    }
}
elseif ($WhatIfPreference) {
    # The probe writes a real file, so it needs a real folder. Reporting a
    # missing folder as a problem would make the dry run describe a run that
    # would never have happened.
    Write-Host '  Backup root does not exist yet; the real run would create it.'
}
else {
    throw "The backup root could not be created: $BackupRootFolder"
}

$missingFiles  = [System.Collections.Generic.List[string]]::new()
$readOnlyDirs  = [System.Collections.Generic.List[string]]::new()
$checkedDirs   = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$totalBytes    = [int64]0
$sourceSizes   = @{}
$sourceVolumes = @{}
$sourceTotals  = @{}

foreach ($candidate in $candidates) {
    $path = Join-Path $candidate.Root $candidate.FileName

    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $missingFiles.Add($path)
        continue
    }

    $file = Get-Item -LiteralPath $path
    $totalBytes += $file.Length

    # The remuxed file is written next to the original, so its volume has to
    # hold a second copy of every file being repaired at that moment. How many
    # that is depends on -ThrottleLimit, which is not a per-file question: the
    # sizes are collected here and the peak worked out once they are all in.
    $volume = [System.IO.Path]::GetPathRoot($file.FullName)
    if (-not $sourceSizes.ContainsKey($volume)) {
        $sourceSizes[$volume] = [System.Collections.Generic.List[int64]]::new()
    }
    $sourceSizes[$volume].Add($file.Length)

    # Kept alongside the largest file, for the recycle bin warning below. A
    # missing key reads as $null, which [int64] turns into 0.
    $sourceTotals[$volume] = [int64]$sourceTotals[$volume] + $file.Length

    $folder = Split-Path -Path $path -Parent
    if ($checkedDirs.Add($folder) -and -not (Test-FolderIsWritable $folder)) {
        $readOnlyDirs.Add($folder)
    }
}

foreach ($missing in $missingFiles) {
    Write-Warning "Candidate file no longer exists and is skipped: $missing"
}
foreach ($folder in $readOnlyDirs) {
    Write-Warning "Source folder is not writable, its files are skipped: $folder"
}

if ($readOnlyDirs.Count -gt 0) {
    $blocked = [System.Collections.Generic.HashSet[string]]::new($readOnlyDirs, [StringComparer]::OrdinalIgnoreCase)
    $candidates = @($candidates | Where-Object { -not $blocked.Contains((Split-Path -Path (Join-Path $_.Root $_.FileName) -Parent)) })
}
if ($missingFiles.Count -gt 0) {
    $gone = [System.Collections.Generic.HashSet[string]]::new($missingFiles, [StringComparer]::OrdinalIgnoreCase)
    $candidates = @($candidates | Where-Object { -not $gone.Contains((Join-Path $_.Root $_.FileName)) })
}

# Repairs run concurrently, so more than one remux output can exist at a time.
# The ceiling is -ThrottleLimit, held down by -BatchSize if a batch is the
# smaller of the two, since the parallel engine is handed one batch at a time.
# Each in-flight repair holds its own output until step 5 replaces the original,
# so the peak a volume must carry is the sum of that many of its largest files -
# not the largest one alone. Which files actually run together is unknowable
# here; the largest are the worst case, and a space check is worth nothing if it
# is not the worst case.
$concurrentRepairs = [math]::Max(1, [math]::Min($ThrottleLimit, $BatchSize))

foreach ($volume in $sourceSizes.Keys) {
    $peak = $sourceSizes[$volume] | Sort-Object -Descending | Select-Object -First $concurrentRepairs
    $sourceVolumes[$volume] = [int64](($peak | Measure-Object -Sum).Sum)
}

$backupFree = Get-FreeSpaceBytes $BackupRootFolder
if ($null -ne $backupFree -and $backupFree -lt $totalBytes) {
    throw ('The backup root has {0:N1} GiB free but the candidates need {1:N1} GiB: {2}' -f
        ($backupFree / 1GB), ($totalBytes / 1GB), $BackupRootFolder)
}

foreach ($volume in $sourceVolumes.Keys) {
    $free = Get-FreeSpaceBytes $volume
    if ($null -eq $free) { continue }

    if ($free -lt $sourceVolumes[$volume]) {
        throw ('{0} has {1:N1} GiB free, too little for the {2} remux output(s) that can be under way at once ({3:N1} GiB).' -f
            $volume, ($free / 1GB), $concurrentRepairs, ($sourceVolumes[$volume] / 1GB))
    }

    # Room for the files under way at once is all the check above asks for,
    # because step 5 deletes each original the moment its replacement is ready.
    # A share with a recycle bin does not free that space - it moves the original
    # aside - and then the volume has to hold the whole source folder tree a
    # second time.
    if ($free -lt $sourceTotals[$volume]) {
        Write-Warning ('{0} has {1:N1} GiB free but holds {2:N1} GiB of candidates. Each original frees its own space as it is replaced, so this is normally enough - unless the volume is a NAS share with a recycle bin, which keeps the originals instead of deleting them. Then the run stops for lack of space partway through. Synology NAS: Control Panel > Shared Folder > Edit > Enable Recycle Bin > (Clear Checkbox).' -f
            $volume, ($free / 1GB), ($sourceTotals[$volume] / 1GB))
    }
}

Write-Host ('  {0} file(s), {1:N1} GiB to copy to {2}.' -f $candidates.Count, ($totalBytes / 1GB), $BackupRootFolder)

if ($candidates.Count -eq 0) {
    Write-Host 'Nothing left to do.'
    Write-Host ''
    Restore-RunState $runState
    return
}

#endregion Access and space checks

#region Resume

$reportFolder = Split-Path -Path $sourceDataPath -Parent
$partialFile  = Join-Path $reportFolder ((Split-Path -Path $sourceDataPath -LeafBase) + '-remux.partial.csv')

$previousResults = @()

if (Test-Path -LiteralPath $partialFile -PathType Leaf) {
    # The file holds every repair the interrupted run finished but never wrote
    # back to the report, so discarding it does not undo those repairs - it
    # loses the only record that they happened. Hence the question. It is asked
    # here rather than earlier because a report with nothing left to repair
    # returns above, and that run has no business touching the file at all.
    $answer = if ($Resume) { 'Resume' } else {
        Read-UserChoice -Default 'Cancel' `
            -Title   'An unfinished run left results behind' `
            -Message "$partialFile holds the results of a run that did not finish. What should happen to it?" `
            -Option  '&Resume', '&Discard', '&Cancel' `
            -Help    'Continue that run. Files already recorded in it are not repaired again - the same as -Resume.',
                     'Delete it and start over. Those files are repaired already, but nothing records it, so each one is repaired again and leaves a second, numbered backup.',
                     'Stop now and leave the file alone.'
    }

    if ($answer -eq 'Cancel') {
        Write-Host "Stopped. $partialFile is untouched - run again with -Resume to continue it."
        Write-Host ''
        Restore-RunState $runState
        return
    }

    if ($answer -eq 'Resume') {
        $previousResults = @(Import-Csv -LiteralPath $partialFile)
        Write-Host "Resuming: $($previousResults.Count) file(s) already recorded in $partialFile"
    }
    else {
        Write-Host "Discarding previous partial results: $partialFile"
        Write-Host '  The files it recorded are repaired a second time; each keeps its first backup and gets a numbered one.'
        Remove-Item -LiteralPath $partialFile -Force
    }
}

$alreadyDone = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($previous in $previousResults) {
    [void]$alreadyDone.Add((Get-CandidateKey $previous))
}

$pending = @($candidates | Where-Object { -not $alreadyDone.Contains((Get-CandidateKey $_)) })

# -WhatIf and -Confirm are settled here, in the main thread: $PSCmdlet is not
# available inside a parallel runspace.
$pending = @($pending | Where-Object {
    $PSCmdlet.ShouldProcess((Join-Path $_.Root $_.FileName), 'Back up, remux and replace')
})

$total = $pending.Count

if ($total -eq 0) {
    if ($WhatIfPreference) {
        Write-Host ''
        Write-Host 'Nothing was changed (-WhatIf).'
        Write-Host ''
        Restore-RunState $runState
        return
    }

    if ($previousResults.Count -eq 0) {
        Write-Host 'Nothing left to repair.'
        Write-Host ''
        Restore-RunState $runState
        return
    }

    # Everything was repaired by an earlier run, but its results are still sitting
    # in the partial file - which is exactly what happens when the write-back
    # failed because the workbook was open. Carry on to the write-back rather than
    # leaving the user with a -Resume that reports nothing to do and does nothing.
    Write-Host 'Nothing left to repair; writing the results of the previous run to the report.'
}

#endregion Resume

#region Repair

if ($total -gt 0) {
    Write-Host ''
    Write-Host "Repairing $total file(s), $ThrottleLimit at a time."
    Write-Host ''
}

$verifyHashFlag = [bool]$VerifyHash

$resultQueue = [System.Collections.Concurrent.ConcurrentQueue[psobject]]::new()
$phase       = [System.Collections.Concurrent.ConcurrentDictionary[string, string]]::new()

$completed   = [System.Collections.Generic.List[psobject]]::new()
$fixedCount  = 0
$failedCount = 0
$stopwatch   = [System.Diagnostics.Stopwatch]::StartNew()

# Drains everything the runspaces have finished: writes each result to the partial
# file as it arrives - not once per batch - and prints one line per file just
# above the progress bar.
$drain = {
    $item = $null
    while ($resultQueue.TryDequeue([ref]$item)) {
        $completed.Add($item)
        $item | Export-Csv -LiteralPath $partialFile -NoTypeInformation -Encoding utf8BOM -Append

        $leaf = Split-Path -Path $item.FileName -Leaf
        if ($leaf.Length -gt 70) { $leaf = $leaf.Substring(0, 67) + '...' }

        if ($item.RemuxStatus -like 'Fixed*') {
            $script:fixedCount++
            Write-Host ('  OK      {0}  ({1}s)' -f $leaf, $item.Seconds) -ForegroundColor Green
        }
        else {
            $script:failedCount++
            Write-Host ('  FAILED  {0}' -f $leaf) -ForegroundColor Red
            Write-Host ('          {0}' -f $item.RemuxStatus) -ForegroundColor Red
        }
    }
}

$showProgress = {
    $done   = $completed.Count
    $status = "$done of $total | Fixed: $fixedCount | Failed: $failedCount"

    # The estimate is written into the status text rather than passed as
    # -SecondsRemaining, which PowerShell renders as a bare number at the end of
    # the line with nothing to say what it counts. It is whole-job time, not
    # per-file: an average over the files already finished, so -ThrottleLimit is
    # already accounted for. Hours are taken from TotalHours, since a TimeSpan
    # longer than a day would otherwise report its hours modulo 24.
    if ($done -gt 0) {
        $remaining = [timespan]::FromSeconds(($stopwatch.Elapsed.TotalSeconds / $done) * ($total - $done))
        $status += ' | ETA ' + ('{0:00}:{1:00}:{2:00}' -f
            [math]::Floor($remaining.TotalHours), $remaining.Minutes, $remaining.Seconds)
    }

    # Last, because the Minimal progress view truncates at the console width and
    # these are the longest and least important part of the line.
    $inFlight = @($phase.Values)
    if ($inFlight.Count -gt 0) { $status += ' | ' + ($inFlight -join ', ') }

    Write-Progress `
        -Activity 'Repairing MKV files' `
        -Status $status `
        -PercentComplete (($done / [math]::Max($total, 1)) * 100)
}

try {
    for ($offset = 0; $offset -lt $total; $offset += $BatchSize) {
        $last  = [math]::Min($offset + $BatchSize - 1, $total - 1)
        $batch = $pending[$offset..$last]

        $job = $batch | ForEach-Object -ThrottleLimit $ThrottleLimit -AsJob -Parallel {
            # A parallel runspace starts with a session state of its own, so
            # nothing this script imported is in scope here. The import runs once
            # per runspace in the pool rather than once per file: it is a no-op
            # from the second file that runspace handles onwards. Never -Force,
            # which would reload the module for every file.
            Import-Module $using:modulePath

            Repair-MkvCandidate `
                -Candidate $_ `
                -BackupRootFolder $using:BackupRootFolder `
                -MkvMerge $using:MkvMerge `
                -MkvInfo $using:MkvInfo `
                -VerifyHash $using:verifyHashFlag `
                -ResultQueue $using:resultQueue `
                -Phase $using:phase
        }

        try {
            while ($job.State -in @('NotStarted', 'Running')) {
                & $drain
                & $showProgress
                Start-Sleep -Milliseconds 400
            }
        }
        finally {
            & $drain
            & $showProgress
            # The worker traps everything itself, so this only surfaces a runspace
            # that failed before the worker got a chance to run.
            Receive-Job -Job $job -ErrorAction SilentlyContinue | Out-Null
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
    }
}
finally {
    # Put the console back the way it was, even if the run is interrupted.
    Write-Progress -Activity 'Repairing MKV files' -Completed
    Restore-RunState $runState
}

$stopwatch.Stop()

#endregion Repair

#region Write back

$allResults = @($previousResults) + @($completed)
$writtenTo  = $null
$applied    = 0

if ($allResults.Count -gt 0) {
    Write-Host ''

    if ($writeTarget -eq 'Workbook') {
        # Hours of remuxing can be behind these results, and the only thing in
        # the way is a workbook someone left open. Ending the run over that and
        # making him start again with -Resume is a poor trade for a question he
        # can answer by clicking a window shut.
        while ($true) {
            if (-not (Test-FileIsLocked $workbookPath)) {
                $write = Write-ResultsToWorkbook -WorkbookPath $workbookPath -Results $allResults
                $writtenTo = $write.Path
                $applied   = $write.Applied
                break
            }

            $answer = Read-UserChoice -Default 'Retry' -Unattended 'Cancel' `
                -Title   'The workbook cannot be opened' `
                -Message "The results cannot be written to $workbookPath. It is most likely open in another program - a read-only file or a denied permission looks the same from here." `
                -Option  '&Retry', '&Cancel' `
                -Help    'Close the workbook or solve the access problem, then try the write again.',
                         'Write nothing now. The results are kept and -Resume writes them back later.'

            if ($answer -eq 'Cancel') {
                Write-Warning "The results were not written to the workbook this time. They are safe in $partialFile - once the workbook can be written, run again with -Resume."
                break
            }
        }
    }
    else {
        $write = Write-ResultsToCsv -CsvPath $csvPath -Results $allResults
        $writtenTo = $write.Path
        $applied   = $write.Applied

        if ($workbookPath -and $writtenTo -eq $csvPath) {
            if (Test-FileIsLocked $workbookPath) {
                Write-Host 'The workbook is open; use Data > Refresh All to bring the results in.'
            }
            elseif (Update-WorkbookQuery -WorkbookPath $workbookPath) {
                Write-Host "Refreshed the workbook: $workbookPath"
            }
        }
    }

    # Only once the results are safely in the report is the crash-safety net
    # allowed to go away.
    if ($writtenTo) {
        Remove-Item -LiteralPath $partialFile -Force -ErrorAction SilentlyContinue
    }
}

#endregion Write back

#region Report

$fixed  = @($allResults | Where-Object { $_.RemuxStatus -like 'Fixed*' })
$failed = @($allResults | Where-Object { $_.RemuxStatus -notlike 'Fixed*' })

$bytesBefore = ($fixed | ForEach-Object { [int64]0 + ($_.SizeBeforeBytes -as [int64]) } | Measure-Object -Sum).Sum
$bytesAfter  = ($fixed | ForEach-Object { [int64]0 + ($_.RemuxSizeBytes  -as [int64]) } | Measure-Object -Sum).Sum

Write-Host ''
Write-Host "Candidates in report:      $($candidates.Count)"
Write-Host "Processed this run:        $($completed.Count)"
Write-Host "Repaired:                  $($fixed.Count)"
Write-Host "Failed:                    $($failed.Count)"
Write-Host ('Size before / after:       {0:N2} GiB / {1:N2} GiB' -f ($bytesBefore / 1GB), ($bytesAfter / 1GB))
Write-Host "Elapsed:                   $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
Write-Host "Backups kept in:           $BackupRootFolder"

if ($writtenTo) {
    Write-Host "Results written to:        $writtenTo ($applied row(s) updated)"
}
elseif ($allResults.Count -gt 0) {
    Write-Host "Results written to:        nothing - they are in $partialFile; run again with -Resume"
}

if ($failed.Count -gt 0) {
    Write-Host ''
    Write-Host 'Failures:' -ForegroundColor Red
    foreach ($failure in $failed) {
        Write-Host ("  {0}" -f (Join-Path $failure.Root $failure.FileName)) -ForegroundColor Red
        Write-Host ("    {0}" -f $failure.RemuxStatus) -ForegroundColor Red
    }
}

Write-Host ''
Write-Host 'Every original was copied to the backup root before it was touched, and nothing there'
Write-Host 'is deleted by this script. Play the repaired files in Plex, then delete the backups.'
Write-Host ''

#endregion Report
