#requires -Version 7.2

<#
.SYNOPSIS
    Finds MKV files that Plex may fail to play because mkclean compressed the
    CodecPrivate data of an HEVC video track.

.DESCRIPTION
    Version 1.0.9

    Two checks are made per file, both read-only:

      1. mkvmerge -J (JSON identification) gives the container's writing
         application and, per track, the codec ID and the content encoding
         algorithms. This lets the "is HEVC" and "is compressed" tests be tied
         to the SAME track, which a flat text search cannot do.

      2. Only for files that are mkclean + HEVC, mkvinfo is run to read the
         content encoding scope and the HEVC profile of that same track. A
         profile of "Unknown @L0.0" means mkvinfo could not parse the compressed
         CodecPrivate - the same thing that trips up Plex. mkvmerge decompresses
         it transparently, which is why it cannot report the problem itself.

    Candidate       = mkclean + an HEVC video track that carries content
                      compression. Decided from mkvmerge alone, because mkvinfo
                      cannot list the tracks of every mkclean'd file.
    StrictCandidate = Candidate, corroborated by mkvinfo reading the profile as
                      Unknown. A candidate that is not strict is not a weaker
                      hit; it usually means mkvinfo could not see the track.

    The path is reported in two columns: Root holds the drive and the first
    folder (the genre folder, e.g. "Q:\Horror"), FileName holds everything below
    it. That makes the report easy to filter and group by genre.

    The last five columns - RemuxStatus, RemuxTime, RemuxHevcProfile,
    RemuxSizeBytes and BackupPath - are written empty and stay empty here. They
    are filled in by Fix-PlexMkcleanCandidates.ps1, which repairs the
    candidates. They are created in this script rather than added later because
    the report is normally viewed through an Excel workbook whose Power Query
    builds the table from this CSV: a column that exists from the first scan is
    part of the query's own output and survives every refresh, whereas one added
    to the workbook by hand is positional and drifts out of alignment as soon as
    a re-scan changes the row order.

    Results are appended to a .partial file as the scan runs, so an interrupted
    or crashed scan loses nothing and can be continued with -Resume. A run that
    finds one waiting asks what to do with it instead of deciding by itself.

.EXAMPLE
    .\Find-PlexMkcleanCandidates.ps1 -Source 'Q:\War' `
        -ReportFile '.\Prod\Plex-mkclean-candidates - War.csv'
    Scans one genre folder into its own report. The workbook beside that CSV
    picks it up by name, and Fix-PlexMkcleanCandidates.ps1 writes its results
    back into it.

.EXAMPLE
    .\Find-PlexMkcleanCandidates.ps1 -Source 'Q:\Horror', 'Q:\Animation', 'X:\'
    Scans several roots in one run. A root that is unavailable is reported and
    skipped; files found under more than one root are only scanned once.

.EXAMPLE
    .\Find-PlexMkcleanCandidates.ps1 -Source 'X:\' -Resume
    Continues a scan that was interrupted, skipping files already recorded.

.NOTES
    Version: 1.0.9
    This script never modifies, remuxes, renames or deletes a media file.

    Version history:
      1.0.9  The two places that decided for the user now ask him. A .partial
             file left by an interrupted scan was announced and deleted in the
             same breath, which is no help at all - by the time the warning is
             on screen the work is gone; resuming it, discarding it and stopping
             are now offered, and offered before the enumeration rather than
             after it, so the question does not arrive a quarter of an hour
             late. A report that could not be written - which in practice means
             Excel has it open - was written beside the intended one under a
             timestamped name; closing Excel and retrying is now the first
             option, and the .partial file is kept until the report is actually
             on disk. Both questions answer themselves, safely, when there is no
             console to answer them.
             Restore-RunState is gone from here and comes from PlexMkclean.psm1
             1.0.3 instead, together with the new Get-RunState that captures
             what it puts back. All three scripts had grown a copy of it. The
             module import moved up above the first line that changes anything
             about the host, since Get-RunState has to record the state before
             it is disturbed.
             The comment on the trap now says which exits it covers: every
             throw, but not an early return, which is a normal exit the trap
             never sees. The two calls left here are of that second kind and
             stay where they are.
             The two defaults that pointed into this machine are gone.
             -MKVToolNix is the ordinary 'C:\Program Files\MKVToolNix', and
             -ReportFile is a bare file name, so the report is written in the
             current directory rather than in one particular project folder.
      1.0.8  The $ReportFile default follows the project folder, which was
             renamed to video-plex-mkclean-repair. The old default pointed at a
             path that no longer exists, so a run without -ReportFile would have
             failed on writing rather than on starting.
      1.0.7  The scan engine reaches the parallel runspaces by importing
             PlexMkclean.psm1 inside the -Parallel block, rather than by
             shipping the text of Test-MkvCleanFile into each runspace with
             ${function:...}.ToString(). A runspace has a session state of its
             own, so the module has to arrive somehow; importing it brings the
             whole module instead of one nameless function body, which means
             the engine may now call private helpers and be refactored like
             ordinary code. The old way forbade that silently - a call to a
             sister function parsed cleanly and then failed per file, out in a
             worker, as a non-terminating error that read like a bad media file
             rather than a bug. The cost is one module load per runspace in the
             pool, not per file, which is nothing beside an mkvmerge run.
      1.0.6  Test-MkvCleanFile moved verbatim to PlexMkclean.psm1, so the same
             scan engine can serve a one-off file as well as a whole tree. The
             script imports the module and goes on serialising the function into
             its runspaces exactly as before. $ScriptVersion had been left at
             1.0.4 while the header already said 1.0.5, so the banner reported
             the wrong version; both now read 1.0.6.
      1.0.5  $Source lost its 'Q:\' default and is mandatory instead. A scan is
             now normally scoped to one genre folder, so the root has to be
             stated rather than silently becoming the whole drive. Everything
             else keeps its default, $MKVToolNix and $ReportFile included.
      1.0.4  Five empty remux columns added to the report: RemuxStatus,
             RemuxTime, RemuxHevcProfile, RemuxSizeBytes and BackupPath.
             Fix-PlexMkcleanCandidates.ps1 fills them in. $Priority added, to
             keep a long scan out of the way of the rest of the machine.
      1.0.3  $Source takes a list of roots. Path split into Root + FileName
             columns. Version number added to the header and the banner.
             Progress bar drawn in light cyan instead of yellow.
      1.0.2  Candidacy decided from mkvmerge alone; mkvinfo corroborates it, so
             files whose tracks mkvinfo cannot list are no longer discarded.
      1.0.1  Corrupt files are recognized instead of being reported as clean.
      1.0.0  Rewrite: per-track matching via mkvmerge -J, parallel scanning,
             crash-safe partial report, culture-invariant output.
#>

# Mandatory is reserved for a parameter that is important and has no sensible
# default: -Source here, -SourceData in Fix-PlexMkcleanCandidates.ps1. Both
# are about which source folder(s) a run is about, and there is no value that 
# would be right more often than it would be wrong. 
# Everything else keeps its default, because a mandatory parameter never 
# uses one - PowerShell prompts instead, which would make the value on it 
# a lie.
[CmdletBinding()]
param(
    # One or more root folders to scan. Roots may overlap; each file is scanned once.
    # Mandatory, and deliberately without a default: a forgotten -Source might
    # mean the whole of drive or share, which could be an hour of scanning nobody asked for.
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Source,

    # Folder holding mkvmerge.exe and mkvinfo.exe.
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$MKVToolNix = 'C:\Program Files\MKVToolNix',

    [Parameter()]
    [string]$ReportFile = 'Plex-mkclean-candidates.csv',

    # Regex patterns matched against the full path; matching files are skipped.
    # Pass an empty array to scan everything, including the NAS recycle bin.
    [Parameter()]
    [string[]]$ExcludePattern = @('\\#recycle\\', '\\@eaDir\\', '\\\$RECYCLE\.BIN\\'),

    [Parameter()]
    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 8,

    # Files per parallel batch. Each completed batch is flushed to disk.
    [Parameter()]
    [ValidateRange(1, 1000)]
    [int]$BatchSize = 50,

    # Windows priority class for this script and, by inheritance, every mkvmerge
    # and mkvinfo it starts. Nothing above Normal is offered: a child process
    # inherits its parent's class only when that class is Idle or BelowNormal, so
    # a higher setting would never reach the tools that do the work.
    [Parameter()]
    [ValidateSet('Idle', 'BelowNormal', 'Normal')]
    [string]$Priority = 'Normal',

    # Continue a previous run: files already present in the .partial file are skipped.
    [Parameter()]
    [switch]$Resume
)

$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.9'

$MkvInfo  = Join-Path $MKVToolNix 'mkvinfo.exe'
$MkvMerge = Join-Path $MKVToolNix 'mkvmerge.exe'

# Test-MkvCleanFile lives in PlexMkclean.psm1 beside this script, shared with
# Fix-PlexMkcleanCandidates.ps1 and the single-file tooling. This import serves
# the main thread only - each parallel runspace imports the module for itself.
# The path is resolved here because $PSScriptRoot is empty inside a -Parallel
# block, so the scan below could not work it out for itself.
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
Write-Host "Find-PlexMkcleanCandidates $ScriptVersion"

if ($Priority -ne 'Normal') {
    # Idle and BelowNormal are inherited by child processes, so this one line
    # reaches every mkvmerge and mkvinfo the scan starts.
    $runState.Process.PriorityClass = [System.Diagnostics.ProcessPriorityClass]$Priority
    Write-Host "Process priority: $Priority (mkvmerge and mkvinfo inherit it)"
}

#region Pre-flight

foreach ($tool in @($MkvMerge, $MkvInfo)) {
    if (-not (Test-Path -LiteralPath $tool -PathType Leaf)) {
        throw "MKVToolNix executable was not found: $tool"
    }
}

$sourceRoots  = [System.Collections.Generic.List[string]]::new()
$missingRoots = [System.Collections.Generic.List[string]]::new()

foreach ($root in $Source) {
    if (Test-Path -LiteralPath $root -PathType Container) {
        $sourceRoots.Add($root)
    }
    else {
        $missingRoots.Add($root)
    }
}

if ($sourceRoots.Count -eq 0) {
    throw "None of the source folders were found: $($Source -join ', ')"
}

$reportFolder = Split-Path -Path $ReportFile -Parent
if ($reportFolder -and -not (Test-Path -LiteralPath $reportFolder -PathType Container)) {
    New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
}

$partialFile = "$ReportFile.partial"

# What becomes of a .partial file left behind by an interrupted scan. It holds
# everything that scan got through, so the question is put here - before the
# enumeration, which on a large share is minutes of work - rather than after it.
$alreadyDone = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

if (Test-Path -LiteralPath $partialFile -PathType Leaf) {
    $answer = if ($Resume) { 'Resume' } else {
        Read-UserChoice -Default 'Cancel' `
            -Title   'An unfinished scan is waiting' `
            -Message "$partialFile holds the results of a scan that did not finish. What should happen to it?" `
            -Option  '&Resume', '&Discard', '&Cancel' `
            -Help    'Continue that scan. Files already recorded in it are not scanned again - the same as -Resume.',
                     'Delete it and scan everything from the beginning.',
                     'Stop now and leave the file alone.'
    }

    if ($answer -eq 'Cancel') {
        Write-Host "Stopped. $partialFile is untouched - run again with -Resume to continue it."
        Write-Host ''
        Restore-RunState $runState
        return
    }

    if ($answer -eq 'Resume') {
        Import-Csv -LiteralPath $partialFile | ForEach-Object {
            [void]$alreadyDone.Add((Join-Path $_.Root $_.FileName))
        }
        Write-Host "Resuming: $($alreadyDone.Count) file(s) already recorded in $partialFile"
    }
    else {
        Write-Host "Discarding previous partial report: $partialFile"
        Remove-Item -LiteralPath $partialFile -Force
    }
}

#endregion Pre-flight

#region Enumerate

Write-Host "Enumerating .mkv files under: $($sourceRoots -join ', ')"

foreach ($root in $missingRoots) {
    Write-Warning "Source folder was not found and is skipped: $root"
}

$files      = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
$enumErrors = [System.Collections.Generic.List[object]]::new()

# Roots may overlap (like Q:\ and Q:\Horror), so keep track of what has been added.
$seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

foreach ($root in $sourceRoots) {
    # -ErrorAction SilentlyContinue is essential here: $ErrorActionPreference = 'Stop'
    # would otherwise turn a single unreadable folder into a terminating error and
    # abandon the whole scan before a single file is examined.
    $found = Get-ChildItem -LiteralPath $root -Recurse -File -Force -Filter '*.mkv' `
        -ErrorAction SilentlyContinue -ErrorVariable rootErrors |
        Where-Object { $_.Extension -ieq '.mkv' }   # -Filter can over-match via 8.3 names

    foreach ($file in $found) {
        if ($seen.Add($file.FullName)) { $files.Add($file) }
    }

    foreach ($rootError in $rootErrors) { $enumErrors.Add($rootError) }
}

$foundCount = $files.Count

$kept = if ($ExcludePattern.Count -gt 0) {
    @($files | Where-Object {
        $path = $_.FullName
        -not ($ExcludePattern | Where-Object { $path -match $_ })
    })
}
else {
    @($files)
}

$excludedCount = $foundCount - $kept.Count

if ($enumErrors.Count -gt 0) {
    Write-Warning "$($enumErrors.Count) path(s) could not be enumerated and were skipped:"
    $enumErrors | Select-Object -First 10 | ForEach-Object {
        Write-Warning "  $($_.TargetObject)"
    }
    if ($enumErrors.Count -gt 10) { Write-Warning "  ... and $($enumErrors.Count - 10) more" }
}

$pending = @($kept | Where-Object { -not $alreadyDone.Contains($_.FullName) })

$total = $pending.Count
Write-Host "Found $foundCount file(s); $excludedCount excluded; $total to scan."

if ($total -eq 0 -and $alreadyDone.Count -eq 0) {
    Write-Host 'Nothing to do.'
    Write-Host ''
    Restore-RunState $runState
    return
}

#endregion Enumerate

#region Scan

$broadCandidateCount  = 0
$strictCandidateCount = 0
$problemCount         = 0
$scanned              = 0
$stopwatch            = [System.Diagnostics.Stopwatch]::StartNew()

try {
    for ($offset = 0; $offset -lt $total; $offset += $BatchSize) {
        $last  = [math]::Min($offset + $BatchSize - 1, $total - 1)
        $batch = $pending[$offset..$last]

        $batchResults = $batch | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            # A parallel runspace starts with a session state of its own, so
            # nothing this script imported is in scope here. The import runs once
            # per runspace in the pool rather than once per file: it is a no-op
            # from the second file that runspace handles onwards. Never -Force,
            # which would reload the module for every file.
            Import-Module $using:modulePath

            Test-MkvCleanFile -File $_ -MkvMerge $using:MkvMerge -MkvInfo $using:MkvInfo
        }

        # Flush after every batch: an aborted scan keeps everything done so far.
        $batchResults | Export-Csv -LiteralPath $partialFile -NoTypeInformation -Encoding utf8BOM -Append

        foreach ($result in $batchResults) {
            if ($result.Candidate)             { $broadCandidateCount++ }
            if ($result.StrictCandidate)       { $strictCandidateCount++ }
            if ($result.ScanStatus -ne 'Scanned') { $problemCount++ }
        }

        $scanned += $batch.Count
        $elapsed = $stopwatch.Elapsed.TotalSeconds
        $remaining = if ($scanned -gt 0) { ($elapsed / $scanned) * ($total - $scanned) } else { 0 }

        Write-Progress `
            -Activity 'Scanning MKV files' `
            -Status "${scanned} of ${total} | Broad: ${broadCandidateCount} | Strict: ${strictCandidateCount} | Problems: ${problemCount}" `
            -PercentComplete (($scanned / [math]::Max($total, 1)) * 100) `
            -SecondsRemaining ([int]$remaining)
    }
}
finally {
    # Put the console back the way it was, even if the scan is interrupted.
    Write-Progress -Activity 'Scanning MKV files' -Completed
    Restore-RunState $runState
}

$stopwatch.Stop()

#endregion Scan

#region Report

$all = @(Import-Csv -LiteralPath $partialFile)

$sorted = $all | Sort-Object -Property @(
    @{ Expression = { $_.StrictCandidate -eq 'True' }; Descending = $true },
    @{ Expression = { $_.Candidate -eq 'True' };       Descending = $true },
    @{ Expression = 'Root';                            Descending = $false },
    @{ Expression = 'FileName';                        Descending = $false }
)

$finalPath = $ReportFile
$written   = $false

while (-not $written) {
    try {
        $sorted | Export-Csv -LiteralPath $finalPath -NoTypeInformation -Encoding utf8BOM
        $written = $true
    }
    catch {
        # Nearly always the report is open in Excel, which holds it exclusively.
        # Nothing is lost whichever way this is answered - the .partial file
        # stays on disk until the write succeeds - but a second report leaves
        # the workbook's query pointing at the stale one, so closing Excel and
        # retrying is offered first.
        $firstAttempt = $finalPath -eq $ReportFile

        # Writing beside it is offered once only. The timestamped name is one
        # this run invented, so if that cannot be written either then a third
        # name will not help - and an unattended run, which always takes the
        # default, would otherwise circle here for ever.
        $option = if ($firstAttempt) { '&Retry', '&Beside it', '&Cancel' }
                  else               { '&Retry', '&Cancel' }

        $help = if ($firstAttempt) {
                    'Close whatever is holding the file, then try the write again.',
                    'Write the results to a second file, with a timestamp in its name.',
                    'Write no report. The results stay in the .partial file.'
                }
                else {
                    'Close whatever is holding the file, then try the write again.',
                    'Write no report. The results stay in the .partial file.'
                }

        # Enter retries, because somebody standing here has just been told the
        # file is locked and the useful thing to do is close Excel. With nobody
        # standing here, retrying would only meet the same lock, so the answer
        # that keeps the scan is taken instead.
        $answer = Read-UserChoice -Option $option -Help $help -Default 'Retry' `
            -Unattended $(if ($firstAttempt) { 'Beside it' } else { 'Cancel' }) `
            -Title   'The report could not be written - usually because Excel has it open' `
            -Message "$finalPath could not be written ($($_.Exception.Message)):"

        if ($answer -eq 'Cancel') { break }

        if ($answer -eq 'Beside it') {
            $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
            $finalPath = Join-Path (Split-Path -Path $ReportFile -Parent) `
                ("{0}-{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($ReportFile), $stamp, [IO.Path]::GetExtension($ReportFile))
        }
    }
}

# Only once the report is on disk. Until then the .partial file is the sole copy
# of the scan, and -Resume can pick the run up and try the report again.
if ($written) {
    Remove-Item -LiteralPath $partialFile -Force -ErrorAction SilentlyContinue
}

$all |
    Group-Object -Property ScanStatus, Candidate, StrictCandidate |
    Sort-Object -Property Name |
    Select-Object Name, Count |
    Format-Table -AutoSize

Write-Host ''
Write-Host "Files found:               $foundCount"
Write-Host "Excluded by pattern:       $excludedCount"
Write-Host "Rows in report:            $($all.Count)"
Write-Host "Broad candidates:          $(($all | Where-Object Candidate -eq 'True').Count)"
Write-Host "Strict candidates:         $(($all | Where-Object StrictCandidate -eq 'True').Count)"
Write-Host "Files with problems:       $(($all | Where-Object ScanStatus -ne 'Scanned').Count)"
Write-Host "Elapsed:                   $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
if ($written) { Write-Host "CSV report written to:     $finalPath" }
else           { Write-Host "CSV report:                not written - the results are in $partialFile" }
Write-Host ''
Write-Host 'This script does not modify, remux, rename, or delete any media file.'
Write-Host 'Review Candidate=True rows in the CSV, then repair them with Fix-PlexMkcleanCandidates.ps1.'
Write-Host ''

#endregion Report
