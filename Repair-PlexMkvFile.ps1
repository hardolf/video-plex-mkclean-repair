#requires -Version 7.2

<#
.SYNOPSIS
    Repairs individual or a small set of MKV files that mkclean made unplayable
    in the Plex for Windows app, named directly instead of through a scan report.

.DESCRIPTION
    Version 1.0.3

    Find-PlexMkcleanCandidates.ps1 and Fix-PlexMkcleanCandidates.ps1 work on a
    whole tree through an Excel report. This script is for the single file or
    small set of files that turns up afterwards: it takes paths on the command
    line, scans each file, and repairs only the files that actually need it.

    Both engines come from PlexMkclean.psm1, so a file repaired here goes through
    exactly the same five steps as one repaired from a report - backup, verify,
    remux, verify, swap - and its backup lands in the same place under the backup
    root. Nothing is written to any .csv or .xlsx report.

    Files are handled one at a time. This is a tool for a handful of files that
    someone is watching; use Fix-PlexMkcleanCandidates.ps1 for e.g. a whole genre,
    where parallel remuxing and a crash-safe partial report matter.

    A file the scan does not recognise as a candidate is reported and left alone.
    -Force repairs it anyway.

.PARAMETER Path
    One or more paths, wildcards allowed. A folder is expanded to the .mkv files
    directly inside it, or to the whole tree below it with -Recurse.

    Film and episode names routinely contain square brackets - for example
    "Blackadder (1983) [x265].mkv" - and [ ] are wildcard characters to -Path,
    which reads them as a character class and matches nothing at all. A pattern
    that holds [ or ] but no * or ? is therefore retried as a literal path, with
    a warning saying so. That is safe because * and ? are illegal in a Windows
    file name, so brackets without them cannot be a wildcard anyone meant.

    The fallback covers the common case; -LiteralPath remains the way to say it
    outright, and the only way to be certain when a name holds both brackets and
    a real wildcard character.

.PARAMETER LiteralPath
    Same as -Path, but the value is used exactly as typed, with no wildcard
    expansion. This is the right parameter for names containing [ ] characters.

.PARAMETER Recurse
    Expand a folder to every .mkv below it, not just the ones directly inside.

.PARAMETER BackupRootFolder
    Where the untouched original is copied before anything else happens. The
    original layout is mirrored below it - "Q:\Horror\Film\x.mkv" becomes
    "<BackupRootFolder>\Q\Horror\Film\x.mkv" - which is the same path
    Fix-PlexMkcleanCandidates.ps1 would have used for that file. Nothing here
    is ever deleted by this script.

    An existing backup is never overwritten: " (1)", " (2)" and so on go in
    before the extension until a free name is found. This matters most with
    -Force, which will happily repair an already repaired file - without the
    numbering, that second run would otherwise copy the repaired file over the
    pristine original and leave no untouched copy anywhere.

.PARAMETER Force
    Repair a file even when the scan does not call it a candidate. Intended for a
    file that misbehaves in Plex although the detection says it should not.

.PARAMETER VerifyHash
    Also compare SHA-256 of the backup against the original. Correct but slow: it
    reads every byte of both files.

.PARAMETER Priority
    Windows priority class for this script and, by inheritance, every mkvmerge and
    mkvinfo it starts. Nothing above Normal is offered: a child process inherits
    its parent's class only when that class is Idle or BelowNormal.

.EXAMPLE
    .\Repair-PlexMkvFile.ps1 -LiteralPath 'V:\TV Shows\Blackadder (1983)\S04E01 [x265].mkv'
    The everyday case: one file, named exactly. -LiteralPath because of the
    square brackets.

.EXAMPLE
    .\Repair-PlexMkvFile.ps1 -Path 'Q:\Horror\The Thing (1982)\*.mkv'
    Every MKV in one film folder.

.EXAMPLE
    .\Repair-PlexMkvFile.ps1 -Path 'V:\TV Shows\Slow Horses (2022)' -Recurse -WhatIf
    Scans the whole series and reports which files would be repaired, without
    touching anything.

.NOTES
    Version: 1.0.3

    This script overwrites media files. Every original is copied to
    $BackupRootFolder and the copy is verified before the original is overwritten,
    and nothing under the backup root is ever deleted here.

    The pre-flight is deliberately lighter than the one in
    Fix-PlexMkcleanCandidates.ps1: tools, backup root and free space are checked,
    but the source folder is not tested for write access and the file is not
    tested for locks in advance. Those helpers still live inside that script; a
    later version can share them through the module. A locked or read-only file is
    reported by the engine instead - after the backup has been made, and before
    the original is touched.

    Version history:
      1.0.3  "Size before" in the summary was always 0,00 GiB. The engine result
             carries the original size as SizeBeforeBytes, and this script read it
             as SizeBytes - the name that column has in a scan row, not in a repair
             result. A missing property is $null, which casts to 0 without error,
             so the total stayed at zero however many files were repaired.
             Fix-PlexMkcleanCandidates.ps1 reads the right name and was never
             affected.

      1.0.2  Comment fix: the note on the de-duplicating HashSet still claimed
             that repairing a file twice would overwrite its own backup, which
             PlexMkclean.psm1 1.0.1 stopped being true. The de-duplication earns
             its place either way - it saves a pointless copy and remux, and a
             spare numbered backup - but the stated reason was wrong.
             Get-FreeSpaceBytes is gone from here and comes from
             PlexMkclean.psm1 1.0.3 instead, which is what its own comment had
             said should happen ever since the copy was made. Restore-RunState
             followed it, together with the new Get-RunState that captures what
             it puts back; all three scripts had grown a copy. The module import
             moved up above the first line that changes anything about the host,
             since Get-RunState has to record the state before it is disturbed.
             The four Restore-RunState calls that stood in front of a throw
             are gone. The script-scope trap already restores on any
             terminating error, so they only did the same work a moment
             earlier; the trap and the finally at the end of the repair loop
             are now the only two places that restore, and the comment on the
             trap says so.
             The defaults that pointed into this machine are gone.
             -MKVToolNix is the ordinary 'C:\Program Files\MKVToolNix', and
             -BackupRootFolder is empty, so the folder has to be given on
             every run. A check at the top of the pre-flight says that outright,
             since an empty value otherwise travelled as far as Test-Path and
             failed there as a binding error naming -Path.
      1.0.1  -WhatIf no longer dies on a backup root that does not exist yet.
             New-Item honours -WhatIf and creates nothing, while the Resolve-Path
             beside it resolves only a path that is already there, so a dry run
             failed on the very folder a real run would have made for itself.
             That call and the one inside Get-FreeSpaceBytes now resolve without
             requiring the path to exist. The .EXAMPLE with -Recurse -WhatIf had
             therefore never worked.
             A -Path pattern holding [ or ] but no * or ? is retried as a
             literal path when it matches nothing, so the bracketed names this
             script is usually pointed at stop being reported as missing. The
             warning says when the retry was used, and -LiteralPath still says
             it outright.
             Backups are no longer overwritten either - that change is in
             PlexMkclean.psm1 1.0.1 and closes the hole where -Force on an
             already repaired file destroyed its own pristine backup.
      1.0.0  First version. Both engines come from PlexMkclean.psm1 1.0.0.
#>

# Mandatory is reserved for a parameter that is important and has no sensible
# default. -Path and -LiteralPath qualify: naming the files is the whole point of
# the script. They are mutually exclusive, so they sit in their own parameter
# sets and PowerShell enforces the choice.
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Path')]
param(
    [Parameter(Mandatory, Position = 0, ParameterSetName = 'Path')]
    [ValidateNotNullOrEmpty()]
    [string[]]$Path,

    [Parameter(Mandatory, ParameterSetName = 'LiteralPath')]
    [ValidateNotNullOrEmpty()]
    [string[]]$LiteralPath,

    [Parameter()]
    [switch]$Recurse,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$BackupRootFolder = '',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$MKVToolNix = 'C:\Program Files\MKVToolNix',

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [switch]$VerifyHash,

    [Parameter()]
    [ValidateSet('Idle', 'BelowNormal', 'Normal')]
    [string]$Priority = 'Normal'
)

$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.3'

$MkvMerge = Join-Path $MKVToolNix 'mkvmerge.exe'
$MkvInfo  = Join-Path $MKVToolNix 'mkvinfo.exe'

# Both engines live in PlexMkclean.psm1 beside this script, shared with
# Find-PlexMkcleanCandidates.ps1 and Fix-PlexMkcleanCandidates.ps1. So do
# Get-FreeSpaceBytes, used by the pre-flight check below, and the Get-RunState
# and Restore-RunState pair used right here - which is why the import comes
# before anything is changed about the host.
Import-Module (Join-Path $PSScriptRoot 'PlexMkclean.psm1') -Force

# Everything this script is about to change about its host, so Restore-RunState
# can put it back at every exit path.
$runState = Get-RunState

# Any terminating error that reaches script scope leaves the host as the script
# found it before the error is passed on. Without this, an aborted run would
# leave the lowered process priority behind in an interactive session, where it
# would silently apply to everything started afterwards. Every throw below is
# covered by this, so none of them restore for themselves; the finally at the
# end of the repair loop covers the normal path. Ctrl+C runs neither.
trap { Restore-RunState $runState; break }

Write-Host ''
Write-Host "Repair-PlexMkvFile $ScriptVersion"

if ($Priority -ne 'Normal') {
    $runState.Process.PriorityClass = [System.Diagnostics.ProcessPriorityClass]$Priority
    Write-Host "Process priority: $Priority (mkvmerge and mkvinfo inherit it)"
}

#region Resolve the input

$literal  = $PSCmdlet.ParameterSetName -eq 'LiteralPath'
$patterns = if ($literal) { $LiteralPath } else { $Path }

# A HashSet on the full path, because two patterns can easily name the same file.
# Repairing it twice is no longer dangerous - PlexMkclean.psm1 1.0.1 numbers a
# colliding backup rather than overwriting it - but it is still a pointless copy
# and remux, and it leaves a spare numbered backup behind.
$seen  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$files = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

foreach ($pattern in $patterns) {
    $items = @()
    try {
        $items = if ($literal) {
            @(Get-Item -LiteralPath $pattern -ErrorAction Stop)
        }
        else {
            @(Get-Item -Path $pattern -ErrorAction Stop)
        }
    }
    catch {
        # A pattern naming a missing file throws; one holding a wildcard that
        # matches nothing comes back empty and silent. Both land on the retry
        # below, so neither is reported before it has had its second chance.
        $items = @()
    }

    # -Path reads [ and ] as a character class, so a film called
    # "Alien [1979].mkv" matches nothing at all - the commonest way this script
    # is handed a name it then claims not to find. Brackets with no * or ? beside
    # them cannot be a wildcard the caller meant, because * and ? are illegal in
    # a Windows file name, so the pattern is retried as a literal path. If the
    # brackets really were a character class that matched nothing, the retry
    # finds nothing either and costs one file system call.
    $viaLiteralFallback = $false

    if (-not $literal -and $items.Count -eq 0 -and
        $pattern -match '[\[\]]' -and $pattern -notmatch '[*?]') {
        try {
            $items = @(Get-Item -LiteralPath $pattern -ErrorAction Stop)
            $viaLiteralFallback = $items.Count -gt 0
        }
        catch {
            $items = @()
        }
    }

    if ($items.Count -eq 0) {
        Write-Warning "Nothing matched: $pattern"
        continue
    }

    if ($viaLiteralFallback) {
        Write-Warning "Read as a literal path because of the [ ] in it: $pattern"
    }

    foreach ($item in $items) {
        if ($item.PSIsContainer) {
            $found = @(Get-ChildItem -LiteralPath $item.FullName -Filter '*.mkv' -File -Recurse:$Recurse)
            if ($found.Count -eq 0) { Write-Warning "No .mkv files under: $($item.FullName)" }
            foreach ($f in $found) {
                if ($seen.Add($f.FullName)) { $files.Add($f) }
            }
        }
        elseif ($item.Extension -eq '.mkv') {
            if ($seen.Add($item.FullName)) { $files.Add([System.IO.FileInfo]$item) }
        }
        else {
            Write-Warning "Not an .mkv file, skipped: $($item.FullName)"
        }
    }
}

if ($files.Count -eq 0) {
    throw 'No .mkv files matched. If a name contains [ ] characters, use -LiteralPath instead of -Path.'
}

Write-Host ('Files to examine:          {0}' -f $files.Count)

#endregion Resolve the input

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

if (-not (Test-Path -LiteralPath $BackupRootFolder -PathType Container)) {
    New-Item -ItemType Directory -Path $BackupRootFolder -Force | Out-Null

    # New-Item honours -WhatIf and creates nothing, so the folder is tested
    # again rather than assumed: announcing a folder that was not made is
    # exactly the kind of thing a dry run must not do.
    if (Test-Path -LiteralPath $BackupRootFolder -PathType Container) {
        Write-Host ('Created backup root:       {0}' -f $BackupRootFolder)
    }
}

# GetUnresolvedProviderPathFromPSPath rather than Resolve-Path, which resolves
# only a path that already exists. Under -WhatIf the New-Item above created
# nothing, so Resolve-Path threw and the dry run died on the one folder the real
# run would have made for itself.
$BackupRootFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($BackupRootFolder)

if ($WhatIfPreference -and -not (Test-Path -LiteralPath $BackupRootFolder -PathType Container)) {
    Write-Host ('Backup root:               {0} (would be created)' -f $BackupRootFolder)
}

$totalBytes   = [int64]0
$largestBytes = [int64]0
$sourceTotals = @{}

foreach ($file in $files) {
    $totalBytes += $file.Length
    if ($file.Length -gt $largestBytes) { $largestBytes = $file.Length }

    # A missing key reads as $null, which [int64] turns into 0.
    $volume = [System.IO.Path]::GetPathRoot($file.FullName)
    $sourceTotals[$volume] = [int64]$sourceTotals[$volume] + $file.Length
}

$backupFree = Get-FreeSpaceBytes $BackupRootFolder
if ($null -ne $backupFree -and $backupFree -lt $totalBytes) {
    throw ('The backup root has {0:N1} GiB free but the files need {1:N1} GiB: {2}' -f
        ($backupFree / 1GB), ($totalBytes / 1GB), $BackupRootFolder)
}

foreach ($volume in $sourceTotals.Keys) {
    $free = Get-FreeSpaceBytes $volume
    if ($null -eq $free) { continue }

    # The remux output is written beside the original, so one file at a time has
    # to fit. Room for all of them is only needed when the volume does not free
    # the original's space as it is replaced - e.g. a NAS share with a recycle bin.
    if ($free -lt $largestBytes) {
        throw ('{0} has {1:N1} GiB free, too little for the temporary remux output ({2:N1} GiB).' -f
            $volume, ($free / 1GB), ($largestBytes / 1GB))
    }

    if ($free -lt $sourceTotals[$volume]) {
        Write-Warning ('{0} has {1:N1} GiB free but holds {2:N1} GiB of files. Normally enough, since each original frees its own space as it is replaced - unless the volume is a NAS share with a recycle bin, which keeps the originals instead. Synology NAS: Control Panel > Shared Folder > Edit > Enable Recycle Bin > (Clear Checkbox).' -f
            $volume, ($free / 1GB), ($sourceTotals[$volume] / 1GB))
    }
}

Write-Host ('Backups go to:             {0}' -f $BackupRootFolder)
Write-Host ''

#endregion Pre-flight

#region Repair

# Repair-MkvCandidate reports through a queue rather than a return value, because
# in Fix-PlexMkcleanCandidates.ps1 it runs inside parallel runspaces. Called
# directly, as here, the queue holds exactly one result per call.
$resultQueue = [System.Collections.Concurrent.ConcurrentQueue[psobject]]::new()
$phase       = [System.Collections.Concurrent.ConcurrentDictionary[string, string]]::new()

$repaired   = 0
$failed     = 0
$notNeeded  = 0
$unreadable = 0
$number     = 0

$bytesBefore = [int64]0
$bytesAfter  = [int64]0

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    foreach ($file in $files) {
        $number++
        Write-Host ('[{0}/{1}] {2}' -f $number, $files.Count, $file.Name)

        $row = Test-MkvCleanFile -File $file -MkvMerge $MkvMerge -MkvInfo $MkvInfo

        # 'Scanned' on its own, or 'Scanned; mkvinfo could not list the tracks' -
        # the latter is still a usable verdict, because candidacy is decided from
        # mkvmerge alone. Anything else is the scan reporting why it failed.
        if ($row.ScanStatus -notlike 'Scanned*') {
            $unreadable++
            Write-Host ('  UNREADABLE  {0}' -f $row.ScanStatus) -ForegroundColor Yellow
            continue
        }

        if (-not $row.Candidate -and -not $Force) {
            $notNeeded++
            Write-Host '  NOT NEEDED  no compressed HEVC track; Plex should play it. -Force remuxes it anyway.' -ForegroundColor DarkGray
            continue
        }

        if (-not $row.Candidate) {
            Write-Host '  FORCED      the scan does not call this a candidate.' -ForegroundColor Yellow
        }

        if (-not $PSCmdlet.ShouldProcess($file.FullName, 'Back up and remux')) { continue }

        Repair-MkvCandidate `
            -Candidate $row `
            -BackupRootFolder $BackupRootFolder `
            -MkvMerge $MkvMerge `
            -MkvInfo $MkvInfo `
            -VerifyHash ([bool]$VerifyHash) `
            -ResultQueue $resultQueue `
            -Phase $phase

        $result = $null
        if (-not $resultQueue.TryDequeue([ref]$result)) {
            $failed++
            Write-Host '  FAILED      the engine returned no result.' -ForegroundColor Red
            continue
        }

        if ($result.RemuxStatus -like 'Fixed*') {
            $repaired++
            $bytesBefore += [int64]$result.SizeBeforeBytes
            $bytesAfter  += [int64]$result.RemuxSizeBytes
            Write-Host ('  OK          {0}s, backup: {1}' -f $result.Seconds, $result.BackupPath) -ForegroundColor Green
        }
        else {
            $failed++
            Write-Host ('  FAILED      {0}' -f $result.RemuxStatus) -ForegroundColor Red
        }
    }
}
finally {
    $stopwatch.Stop()
    Restore-RunState $runState
}

#endregion Repair

#region Summary

$elapsed = $stopwatch.Elapsed

Write-Host ''
Write-Host ('Files examined:            {0}' -f $files.Count)
Write-Host ('Repaired:                  {0}' -f $repaired)
Write-Host ('Failed:                    {0}' -f $failed)
Write-Host ('Not candidates:            {0}' -f $notNeeded)
Write-Host ('Unreadable:                {0}' -f $unreadable)
if ($repaired -gt 0) {
    Write-Host ('Size before / after:       {0:N2} GiB / {1:N2} GiB' -f ($bytesBefore / 1GB), ($bytesAfter / 1GB))
}
Write-Host ('Elapsed:                   {0:00}:{1:00}:{2:00}' -f
    [math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds)
Write-Host ('Backups kept in:           {0}' -f $BackupRootFolder)

if ($repaired -gt 0) {
    Write-Host ''
    Write-Host 'Every original was copied to the backup root before it was touched, and nothing there'
    Write-Host 'is deleted by this script. Play the repaired files in Plex, then delete the backups.'
}
Write-Host ''

#endregion Summary
