<#
.SYNOPSIS
    The scan and repair engines behind Find-PlexMkcleanCandidates.ps1 and
    Fix-PlexMkcleanCandidates.ps1.

.DESCRIPTION
    Version 1.0.3

    Two functions, each the whole of what one script does to a single file:

      Test-MkvCleanFile   decides whether an .mkv was written by mkclean and
                          carries a compressed HEVC track, which is what the
                          Plex for Windows app cannot play.

      Repair-MkvCandidate takes such a file through backup, verification, remux
                          and swap, and reports the outcome.

    They live here rather than in the scripts because the same engines have to
    serve a report-driven run over a whole source tree e.g. a whole genre
    and a one-off run over a single file.

    The rest of what is exported is not an engine but is shared for the same
    reason: every one of them was needed by more than one script, and none of
    them could be owned by a single script.

      Read-UserChoice        the console prompt used where the decision is the
                             user's to make. It answers itself from a
                             caller-stated default when no console is there, so
                             an unattended run reaches a decision instead of
                             hanging on a question.

      Get-RunState and       capture and put back what a script changes about
      Restore-RunState       its host - process priority, progress style - so
                             an interrupted run does not leave a lowered
                             priority behind in an interactive session.

      Get-FreeSpaceBytes     the pre-flight space check made before writing a
                             backup and a remux output the size of a film.

    Both run inside parallel runspaces, which is worth knowing before editing
    them. A runspace is one instance of the PowerShell engine with a session
    state of its own - its own function table, variables, preferences and list
    of loaded modules - so a worker inherits nothing from the script that
    started it. The scripts therefore call Import-Module inside their -Parallel
    blocks, which puts this whole module into each runspace and leaves the
    functions free to call each other and any private helper added here.

    Two consequences survive that import and apply to anything written here:

      Preferences do not cross. $ErrorActionPreference and
      $PSNativeCommandUseErrorActionPreference are set inside each function
      rather than relied on from the caller.

      Script scope is per runspace, not per process. A $script: variable in
      this module is not shared state, the way a static field would be in C#;
      every runspace gets its own copy. State that has to be shared is passed
      in as a parameter - the ConcurrentQueue and ConcurrentDictionary that
      Repair-MkvCandidate reports through are exactly that.

.NOTES
    Version: 1.0.3

    Version history:
      1.0.3  Read-UserChoice added, so a script can ask instead of announcing
             what it has already done. Find-PlexMkcleanCandidates.ps1 1.0.9 and
             Fix-PlexMkcleanCandidates.ps1 1.0.5 use it in the five places that
             warned and acted in the same breath - discarding the results of an
             interrupted run, and giving up on a report or workbook Excel had
             open - where the script could see the problem coming and the user
             had a way out of it. It is not an engine, but it belongs beside
             them: it is the one thing both scripts needed and neither could
             own. What Enter picks and what is taken with nobody watching are
             separate parameters, because they are rarely the same answer.
             Get-FreeSpaceBytes moved here for the same reason, out of the two
             identical copies that had grown in Fix-PlexMkcleanCandidates.ps1
             and Repair-PlexMkvFile.ps1. Its comment had promised the move since
             the copy was made; a published repository is a poor place to leave
             that promise unkept.
             Restore-RunState moved here too, out of all three scripts, and
             became a pair: Get-RunState captures the host state and returns it,
             Restore-RunState takes it back as a parameter. It could not be
             moved as it stood, because it read the values from $script:
             variables - which, in a module, are the module's own and not the
             calling script's, so the copy would have compiled and then quietly
             restored nothing.
      1.0.2  BackupPath is recorded after the copy rather than before it. A copy
             that failed deleted the file it had just claimed but left the path
             in the result, so the report named a backup that did not exist - in
             the one column a reader would trust to find the untouched original.
             A failure in step 2 still records it, because there the backup does
             exist and is exactly what needs looking at.
      1.0.1  Repair-MkvCandidate no longer overwrites an existing backup. It
             inserts " (1)", " (2)" and so on before the extension until a free
             name is found, claiming the name with CreateNew so the check and
             the claim cannot come apart under parallelism. Repairing a file
             twice used to replace its pristine original with the already-
             repaired copy - Fix-PlexMkcleanCandidates.ps1 was protected by its
             report, Repair-PlexMkvFile.ps1 -Force by nothing at all.
             The backup path is now derived from the source path in one step
             instead of being reassembled from Candidate.Root and
             Candidate.FileName. The result is identical for drive, root-level
             and UNC paths.
             The functions no longer have to be self-contained. Both scripts
             import this module inside their -Parallel blocks instead of
             serialising a single function into each runspace with
             ${function:...}.ToString(), so private helpers and calls between
             the two engines are now allowed. Nothing inside either function
             was changed - the ban is simply lifted, and the notes above say
             what still holds.
      1.0.0  Test-MkvCleanFile moved here verbatim from
             Find-PlexMkcleanCandidates.ps1 1.0.5, and Repair-MkvCandidate from
             Fix-PlexMkcleanCandidates.ps1 1.0.2. Nothing inside either function
             was changed; both scripts import this module and go on serialising
             the functions into their runspaces exactly as before.
#>

#requires -Version 7.2

#region Scan engine

function Test-MkvCleanFile {
    <#
        Runs inside a parallel runspace, which imports this whole module, so
        calls to other functions here are fine. What does not cross into that
        runspace is the caller's state: preferences are set below rather than
        inherited, and anything shared has to arrive as a parameter.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)][string]$MkvMerge,
        [Parameter(Mandatory)][string]$MkvInfo
    )

    # A non-zero exit code from mkvmerge/mkvinfo must not become a terminating
    # error: we want to record it in the report instead. This preference
    # defaults to $true on some hosts, which would otherwise send every warning
    # exit straight to the catch block below.
    $PSNativeCommandUseErrorActionPreference = $false
    $ErrorActionPreference = 'Continue'

    $inv = [cultureinfo]::InvariantCulture

    # Split "Q:\Horror\Some Film (2023)\film.mkv" into "Q:\Horror" and the rest,
    # so the report can be filtered and grouped by source root folder, e.g. genre folder.
    $pathRoot = [System.IO.Path]::GetPathRoot($File.FullName)   # 'Q:\' or '\\nas\share\'
    $belowRoot = $File.FullName.Substring($pathRoot.Length)
    $separator = $belowRoot.IndexOf([System.IO.Path]::DirectorySeparatorChar)

    if ($separator -ge 0) {
        $root = $pathRoot + $belowRoot.Substring(0, $separator)
        $name = $belowRoot.Substring($separator + 1)
    }
    else {
        # A file sitting directly in the root has no folder to split off.
        $root = $pathRoot
        $name = $belowRoot
    }

    # Numbers and dates are formatted invariantly so the CSV reads the same on a
    # da-DK machine (where 2.529 would otherwise be written "2,529") as anywhere else.
    $row = [ordered]@{
        ScanStatus            = 'Scanned'
        Candidate             = $false
        StrictCandidate       = $false
        IsMkclean             = $false
        HasHevc               = $false
        HevcTrackCompressed   = $false
        CodecPrivateScope2    = $false
        UnknownHevcProfile    = $false
        ContentCompression    = $false
        HevcProfile           = ''
        HevcTrackId           = ''
        CompressionAlgorithms = ''
        WritingApplication    = ''
        MkvMergeExitCode      = ''
        MkvInfoExitCode       = ''
        SizeGiB               = ([math]::Round($File.Length / 1GB, 3)).ToString($inv)
        SizeBytes             = $File.Length
        LastWriteTime         = $File.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss', $inv)
        # Written empty and left empty here. Fix-PlexMkcleanCandidates.ps1 fills
        # these in when it repairs a candidate; see the note in .DESCRIPTION for
        # why they are created by the scan rather than added to the workbook.
        RemuxStatus           = ''
        RemuxTime             = ''
        RemuxHevcProfile      = ''
        RemuxSizeBytes        = ''
        BackupPath            = ''
        Root                  = $root
        FileName              = $name
    }

    try {
        # --- Pass 1: structured identification -------------------------------
        $mergeOut = & $MkvMerge --ui-language en --identification-format json --identify $File.FullName 2>&1
        $row.MkvMergeExitCode = $LASTEXITCODE
        $mergeText = ($mergeOut | ForEach-Object { "$_" }) -join "`n"

        # Slice out the JSON document: a warning printed alongside it would
        # otherwise break ConvertFrom-Json.
        $jsonStart = $mergeText.IndexOf('{')
        $jsonEnd   = $mergeText.LastIndexOf('}')

        if ($jsonStart -lt 0 -or $jsonEnd -le $jsonStart) {
            throw "mkvmerge produced no JSON (exit code $($row.MkvMergeExitCode))"
        }

        $json = $mergeText.Substring($jsonStart, $jsonEnd - $jsonStart + 1) | ConvertFrom-Json

        # mkvmerge exits 0 even for a file it cannot make sense of, reporting
        # "recognized": false instead. Without this check a corrupt file would be
        # filed as a perfectly ordinary non-candidate.
        $toolErrors = @($json.errors | Where-Object { "$_" -ne '' })

        if ($toolErrors.Count -gt 0) {
            throw "mkvmerge: $($toolErrors -join '; ')"
        }
        if (-not $json.container.recognized) {
            throw 'Not a recognized container'
        }
        if (-not $json.container.supported) {
            throw "Unsupported container: $($json.container.type)"
        }
        if ($row.MkvMergeExitCode -ge 2) {
            throw "mkvmerge exit code $($row.MkvMergeExitCode)"
        }

        $writingApp = [string]$json.container.properties.writing_application
        $muxingApp  = [string]$json.container.properties.muxing_application
        $row.WritingApplication = $writingApp
        $row.IsMkclean = ($writingApp -match 'mkclean') -or ($muxingApp -match 'mkclean')

        $tracks = @($json.tracks)

        # File-level compression flag, kept for comparison with the v1 report.
        $anyAlgorithms = @(
            $tracks |
                ForEach-Object { [string]$_.properties.content_encoding_algorithms } |
                Where-Object { $_ -ne '' }
        )
        $row.ContentCompression = $anyAlgorithms.Count -gt 0

        $hevcTracks = @(
            $tracks | Where-Object {
                $_.type -eq 'video' -and $_.properties.codec_id -eq 'V_MPEGH/ISO/HEVC'
            }
        )
        $row.HasHevc = $hevcTracks.Count -gt 0

        if ($row.HasHevc) {
            $hevc = $hevcTracks[0]
            $row.HevcTrackId = [string]$hevc.id

            # Algorithm 0 is zlib, 3 is header removal. mkclean 0.8.7 zlib-compresses
            # the CodecPrivate of the video track, so any value here is significant.
            $algorithms = [string]$hevc.properties.content_encoding_algorithms
            $row.CompressionAlgorithms = $algorithms
            $row.HevcTrackCompressed = $algorithms -ne ''
        }

        # mkvmerge can always read the track headers, so candidacy is decided here.
        # mkvinfo only corroborates it below.
        $row.Candidate = $row.IsMkclean -and $row.HasHevc -and $row.HevcTrackCompressed

        # --- Pass 2: corroboration via the HEVC profile -----------------------
        # Only worth the second file open when the file is actually mkclean'd HEVC.
        if ($row.IsMkclean -and $row.HasHevc) {

            # --ui-language en: mkvinfo's output is translated, and the patterns
            # below only match the English strings.
            $infoOut = & $MkvInfo --ui-language en $File.FullName 2>&1
            $row.MkvInfoExitCode = $LASTEXITCODE

            if ($row.MkvInfoExitCode -ge 2) {
                throw "mkvinfo exit code $($row.MkvInfoExitCode)"
            }

            # Split the element tree into per-track blocks so that scope,
            # compression and profile are read from the HEVC track and not from a
            # subtitle track that happens to be compressed too. The nesting depth
            # of an element is the column its '+' sits in.
            $blocks  = [System.Collections.Generic.List[object]]::new()
            $current = $null

            foreach ($line in $infoOut) {
                $text = "$line"
                $plus = $text.IndexOf('+')
                if ($plus -lt 0) { continue }

                $element = $text.Substring($plus + 1).Trim()

                if ($element -eq 'Track') {
                    if ($null -ne $current) { $blocks.Add($current) }
                    $current = [pscustomobject]@{
                        Depth = $plus
                        Lines = [System.Collections.Generic.List[string]]::new()
                    }
                    continue
                }

                if ($null -ne $current) {
                    if ($plus -le $current.Depth) {
                        # Left the Track element (Clusters, Tags, ...).
                        $blocks.Add($current)
                        $current = $null
                    }
                    else {
                        $current.Lines.Add($element)
                    }
                }
            }
            if ($null -ne $current) { $blocks.Add($current) }

            $hevcBlock = $blocks | Where-Object {
                $_.Lines -match '^Codec ID:\s*V_MPEGH/ISO/HEVC\s*$'
            } | Select-Object -First 1

            if ($null -eq $hevcBlock) {
                # Seen in the wild: mkvinfo never emits a Tracks element for some
                # mkclean'd files, even with --continue over the whole file. The
                # mkvmerge verdict above still stands; it just goes unconfirmed.
                $row.ScanStatus = 'Scanned; mkvinfo could not list the tracks'
                return [pscustomobject]$row
            }

            $row.CodecPrivateScope2 = [bool]($hevcBlock.Lines -match '^Scope:\s*2\b')

            # mkvinfo appends the parsed profile to the private data line:
            #   Codec's private data: size 1234 (HEVC profile: Main@L4.0)
            # "Unknown @L0.0" means it could not read the compressed CodecPrivate,
            # which is exactly what Plex chokes on.
            $profileLine = $hevcBlock.Lines |
                Where-Object { $_ -match 'HEVC profile:' } |
                Select-Object -First 1

            if ($profileLine -match 'HEVC profile:\s*([^)]+)') {
                $row.HevcProfile = $Matches[1].Trim()
                $row.UnknownHevcProfile = $row.HevcProfile -match '^Unknown'
            }

            # An unreadable profile is the direct consequence of the CodecPrivate
            # being compressed, so it confirms what mkvmerge already reported.
            $row.StrictCandidate = $row.Candidate -and $row.UnknownHevcProfile
        }
    }
    catch {
        # Every failure path above throws a plain message; unexpected exceptions
        # land here too, so a problem file is never filed as a clean scan.
        $row.ScanStatus = $_.Exception.Message
    }

    return [pscustomobject]$row
}

#endregion Scan engine

#region Repair engine

function Repair-MkvCandidate {
    <#
        Runs inside a parallel runspace, which imports this whole module, so
        calls to other functions here are fine. What does not cross into that
        runspace is the caller's state: preferences are set below rather than
        inherited, and anything shared has to arrive as a parameter. The result
        is pushed onto $ResultQueue rather than returned, so the main thread can
        write it to disk and report it the moment the file is done instead of at
        the end of a batch.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][psobject]$Candidate,
        [Parameter(Mandatory)][string]$BackupRootFolder,
        [Parameter(Mandatory)][string]$MkvMerge,
        [Parameter(Mandatory)][string]$MkvInfo,
        [Parameter(Mandatory)][bool]$VerifyHash,
        [Parameter(Mandatory)][System.Collections.Concurrent.ConcurrentQueue[psobject]]$ResultQueue,
        [Parameter(Mandatory)][System.Collections.Concurrent.ConcurrentDictionary[string, string]]$Phase
    )

    # A non-zero exit code from mkvmerge/mkvinfo must not become a terminating
    # error: we want to record it in the report instead. This preference defaults
    # to $true on some hosts, which would otherwise send every warning exit
    # straight to the catch block below.
    $PSNativeCommandUseErrorActionPreference = $false
    $ErrorActionPreference = 'Continue'

    $inv = [cultureinfo]::InvariantCulture

    $sourcePath = Join-Path $Candidate.Root $Candidate.FileName
    $tempPath   = "$sourcePath.remux.tmp"
    $phaseKey   = $sourcePath

    # Numbers and dates are formatted invariantly so the CSV reads the same on a
    # da-DK machine as anywhere else.
    $result = [ordered]@{
        Root             = $Candidate.Root
        FileName         = $Candidate.FileName
        RemuxStatus      = ''
        RemuxTime        = ''
        RemuxHevcProfile = ''
        RemuxSizeBytes   = ''
        BackupPath       = ''
        SizeBeforeBytes  = ''
        Seconds          = ''
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # mkvmerge -J, with the JSON sliced out of whatever else the tool printed. A
    # warning alongside the document would otherwise break ConvertFrom-Json.
    $identify = {
        param($Exe, $Path)

        $output   = & $Exe --ui-language en --identification-format json --identify $Path 2>&1
        $exitCode = $LASTEXITCODE
        $text     = ($output | ForEach-Object { "$_" }) -join "`n"

        $jsonStart = $text.IndexOf('{')
        $jsonEnd   = $text.LastIndexOf('}')

        if ($jsonStart -lt 0 -or $jsonEnd -le $jsonStart) {
            throw "mkvmerge produced no JSON for $Path (exit code $exitCode)"
        }

        $text.Substring($jsonStart, $jsonEnd - $jsonStart + 1) | ConvertFrom-Json
    }

    try {
        $Phase[$phaseKey] = 'Copying'

        $sourceFile = Get-Item -LiteralPath $sourcePath -ErrorAction Stop
        $result.SizeBeforeBytes = $sourceFile.Length

        # Restored onto the repaired file, so the Plex library's sort order and
        # "date added" do not jump for a file whose content is unchanged.
        $originalWrite  = $sourceFile.LastWriteTime
        $originalCreate = $sourceFile.CreationTime

        # --- 1. Back up, mirroring the original layout -----------------------
        # The backup path is the whole source path with the drive colon or the
        # UNC prefix taken off, hung under the backup root:
        #
        #   Q:\Horror\Film\x.mkv    ->  <root>\Q\Horror\Film\x.mkv
        #   \\nas\movies\Film\x.mkv ->  <root>\nas\movies\Film\x.mkv
        #
        # A colon cannot appear in a folder name, hence dropping it; mirroring
        # rather than flattening is what stops two films with the same file name
        # from overwriting each other's backup.
        $mirrored     = (($sourcePath -replace '^\\\\', '') -replace ':', '').Trim('\')
        $backupPath   = Join-Path $BackupRootFolder $mirrored
        $backupFolder = Split-Path -Path $backupPath -Parent

        if (-not (Test-Path -LiteralPath $backupFolder -PathType Container)) {
            New-Item -ItemType Directory -Path $backupFolder -Force -ErrorAction Stop | Out-Null
        }

        # An existing backup is never overwritten: " (1)", " (2)" and so on go in
        # before the extension until a free name is found. Repairing the same
        # file a second time would otherwise replace the pristine original with
        # the already-repaired copy, which is the one loss nothing can undo.
        # Fix-PlexMkcleanCandidates.ps1 is guarded against that by its report,
        # but a -Force run of the single-file script has no report to consult.
        #
        # The name is taken with CreateNew rather than tested with Test-Path, so
        # that the check and the claim are one operation: two workers racing for
        # the same name cannot both be told it is free.
        $backupBase      = [System.IO.Path]::GetFileNameWithoutExtension($backupPath)
        $backupExtension = [System.IO.Path]::GetExtension($backupPath)
        $suffix          = 0

        while ($true) {
            $attempt = if ($suffix -eq 0) {
                $backupPath
            }
            else {
                Join-Path $backupFolder ('{0} ({1}){2}' -f $backupBase, $suffix, $backupExtension)
            }

            try {
                [System.IO.File]::Open(
                    $attempt,
                    [System.IO.FileMode]::CreateNew,
                    [System.IO.FileAccess]::Write,
                    [System.IO.FileShare]::None).Dispose()

                $backupPath = $attempt
                break
            }
            catch [System.IO.DirectoryNotFoundException] {
                # An IOException too, but not one another number would fix.
                throw
            }
            catch [System.IO.IOException] {
                # The name is taken. Anything else - no permission, a folder
                # sitting where the file should go - is a real failure and is
                # left to the outer catch.
                $suffix++
                if ($suffix -gt 999) {
                    throw "no free backup name for $backupPath after 999 attempts"
                }
            }
        }

        try {
            Copy-Item -LiteralPath $sourcePath -Destination $backupPath -Force -ErrorAction Stop
        }
        catch {
            # The name was claimed with an empty file. A failed copy must not
            # leave that behind, or the next run would treat it as a backup and
            # start numbering from (1).
            Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
            throw
        }

        # Recorded only once the copy has actually happened. Set any earlier and
        # a failed copy would leave the report pointing at the file just deleted
        # above - a path to nothing, in the one column whose whole purpose is to
        # say where the untouched original went. A failure in step 2 below keeps
        # it, because there the backup does exist and is what has to be examined.
        $result.BackupPath = $backupPath

        # --- 2. Verify the backup --------------------------------------------
        # This is the only safety net that exists once step 5 overwrites the
        # original, so a failure here has to stop the file dead.
        $Phase[$phaseKey] = 'Verifying copy'

        $backupFile = Get-Item -LiteralPath $backupPath -ErrorAction Stop

        if ($backupFile.Length -ne $sourceFile.Length) {
            throw "the backup is $($backupFile.Length) bytes but the source is $($sourceFile.Length)"
        }

        # Copy-Item carries the last-write time across; a couple of seconds of
        # slack covers the coarser timestamp resolution of some file systems.
        $skew = [math]::Abs(($backupFile.LastWriteTimeUtc - $sourceFile.LastWriteTimeUtc).TotalSeconds)
        if ($skew -gt 2) {
            throw 'the backup last-write time does not match the source'
        }

        if ($VerifyHash) {
            $sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
            $backupHash = (Get-FileHash -LiteralPath $backupPath -Algorithm SHA256).Hash
            if ($sourceHash -ne $backupHash) {
                throw 'the SHA-256 of the backup does not match the source'
            }
        }

        # --- 3. Remux the backup into place ----------------------------------
        $Phase[$phaseKey] = 'Remuxing'

        # Read the backup's own track layout, so the result is compared against
        # the exact bytes it was made from.
        $before         = & $identify $MkvMerge $backupPath
        $beforeTracks   = @($before.tracks).Count
        $beforeDuration = [double]$before.container.properties.duration

        if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction Stop
        }

        # --compression -1:none covers every track. mkvmerge does not compress
        # CodecPrivate on its own, but stating it makes it impossible for the
        # output to reproduce the very fault being repaired.
        $mergeOutput = & $MkvMerge --ui-language en --compression -1:none --output $tempPath $backupPath 2>&1
        $mergeCode   = $LASTEXITCODE

        if ($mergeCode -ge 2) {
            $mergeText = (($mergeOutput | ForEach-Object { "$_" }) -join ' ').Trim()
            if ($mergeText.Length -gt 200) { $mergeText = $mergeText.Substring(0, 200) + '...' }
            throw "mkvmerge exit code ${mergeCode}: $mergeText"
        }
        if (-not (Test-Path -LiteralPath $tempPath -PathType Leaf)) {
            throw "mkvmerge exited $mergeCode but wrote no output file"
        }

        # --- 4. Verify the result before it replaces anything ----------------
        $Phase[$phaseKey] = 'Verifying remux'

        $after = & $identify $MkvMerge $tempPath

        $toolErrors = @($after.errors | Where-Object { "$_" -ne '' })
        if ($toolErrors.Count -gt 0) {
            throw "the remuxed file is faulty: $($toolErrors -join '; ')"
        }
        if (-not $after.container.recognized) {
            throw 'the remuxed file is not a recognized container'
        }

        $afterTracks = @($after.tracks).Count
        if ($afterTracks -ne $beforeTracks) {
            throw "the track count changed: $beforeTracks before, $afterTracks after"
        }

        # Duration is reported in nanoseconds. A second of slack absorbs the
        # rounding that a rewritten timestamp scale can introduce.
        $afterDuration = [double]$after.container.properties.duration
        if ($beforeDuration -gt 0 -and [math]::Abs($afterDuration - $beforeDuration) -gt 1e9) {
            $beforeSeconds = [math]::Round($beforeDuration / 1e9, 1)
            $afterSeconds  = [math]::Round($afterDuration / 1e9, 1)
            throw "the duration changed: ${beforeSeconds}s before, ${afterSeconds}s after"
        }

        $hevcTracks = @(
            $after.tracks | Where-Object {
                $_.type -eq 'video' -and $_.properties.codec_id -eq 'V_MPEGH/ISO/HEVC'
            }
        )
        if ($hevcTracks.Count -eq 0) {
            throw 'the remuxed file has no HEVC video track'
        }

        $algorithms = [string]$hevcTracks[0].properties.content_encoding_algorithms
        if ($algorithms -ne '') {
            throw "the HEVC track is still compressed (algorithms: $algorithms)"
        }

        # The direct proof: mkvinfo has to be able to parse the CodecPrivate
        # again. "Unknown @L0.0" is what it reports when it cannot - and that is
        # precisely what Plex trips over.
        $infoOutput = & $MkvInfo --ui-language en $tempPath 2>&1
        $infoCode   = $LASTEXITCODE

        if ($infoCode -ge 2) {
            throw "mkvinfo exit code $infoCode on the remuxed file"
        }

        # Split the element tree into per-track blocks, exactly as
        # Find-PlexMkcleanCandidates.ps1 does, so the profile is read from the
        # HEVC track and not from whichever track prints a profile line first.
        # The nesting depth of an element is the column its '+' sits in.
        $blocks  = [System.Collections.Generic.List[object]]::new()
        $current = $null

        foreach ($line in $infoOutput) {
            $text = "$line"
            $plus = $text.IndexOf('+')
            if ($plus -lt 0) { continue }

            $element = $text.Substring($plus + 1).Trim()

            if ($element -eq 'Track') {
                if ($null -ne $current) { $blocks.Add($current) }
                $current = [pscustomobject]@{
                    Depth = $plus
                    Lines = [System.Collections.Generic.List[string]]::new()
                }
                continue
            }

            if ($null -ne $current) {
                if ($plus -le $current.Depth) {
                    $blocks.Add($current)
                    $current = $null
                }
                else {
                    $current.Lines.Add($element)
                }
            }
        }
        if ($null -ne $current) { $blocks.Add($current) }

        $hevcBlock = $blocks | Where-Object {
            $_.Lines -match '^Codec ID:\s*V_MPEGH/ISO/HEVC\s*$'
        } | Select-Object -First 1

        $profileText = ''
        if ($null -ne $hevcBlock) {
            $profileLine = $hevcBlock.Lines |
                Where-Object { $_ -match 'HEVC profile:' } |
                Select-Object -First 1

            if ($profileLine -match 'HEVC profile:\s*([^)]+)') {
                $profileText = $Matches[1].Trim()
            }
        }
        $result.RemuxHevcProfile = $profileText

        if ($profileText -match '^Unknown') {
            throw "the HEVC profile is still unreadable ($profileText)"
        }

        # --- 5. Replace the original -----------------------------------------
        $Phase[$phaseKey] = 'Replacing'

        Remove-Item -LiteralPath $sourcePath -Force -ErrorAction Stop

        try {
            Move-Item -LiteralPath $tempPath -Destination $sourcePath -ErrorAction Stop
        }
        catch {
            # The only window in which the original is gone. Say exactly where
            # both halves are, so it can be put right by hand.
            throw ("the original was removed but the repaired file could not take its place. " +
                   "The repaired file is '$tempPath' and the untouched original is '$backupPath'. " +
                   "($($_.Exception.Message))")
        }

        $newFile = Get-Item -LiteralPath $sourcePath -ErrorAction Stop
        $newFile.LastWriteTime = $originalWrite
        $newFile.CreationTime  = $originalCreate

        $result.RemuxSizeBytes = $newFile.Length
        $result.RemuxStatus = if ($profileText -eq '') {
            # Same tolerance as the scan: mkvinfo occasionally emits no Tracks
            # element at all. Everything mkvmerge could check did pass.
            'Fixed; mkvinfo could not list the tracks of the result'
        }
        else {
            'Fixed'
        }
    }
    catch {
        $result.RemuxStatus = "Failed: $($_.Exception.Message)"

        # Leave nothing half-finished behind - but only while the original is
        # still there, because otherwise the temporary file is the repaired copy
        # and the message above tells the user to go and get it.
        if ((Test-Path -LiteralPath $tempPath -PathType Leaf) -and
            (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
    finally {
        $stopwatch.Stop()
        $result.Seconds   = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1).ToString($inv)
        $result.RemuxTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $inv)

        $discarded = ''
        [void]$Phase.TryRemove($phaseKey, [ref]$discarded)

        $ResultQueue.Enqueue([pscustomobject]$result)
    }
}

#endregion Repair engine

#region Console

function Test-HostCanPrompt {
    <#
        True when there is somebody at the other end who could answer a question.

        PromptForChoice reads the console, so redirected input - a scheduled
        task, a CI runner, another program driving the shell - means nobody is
        there to type, and the read either fails or hands back end-of-file for as
        long as it is asked. [Environment]::UserInteractive is no use for this;
        it reports True in every one of those cases.

        -NonInteractive is checked as well, because redirection alone misses the
        case that matters most to somebody starting a long run and walking away:
        a script launched from his own terminal with only its output sent to a
        log still has a console attached, so it would sit on a question he cannot
        see. -NonInteractive is the documented way to say nobody is watching, and
        it is honoured here whether or not anything is redirected. PowerShell
        allows the switch to be abbreviated to any unambiguous prefix, of which
        -non is the shortest.

        Private on purpose. Callers want Read-UserChoice, which acts on the
        answer this gives.
    #>
    if (-not $Host.UI) { return $false }

    try {
        if ([System.Console]::IsInputRedirected) { return $false }
    }
    catch { return $false }

    foreach ($argument in [System.Environment]::GetCommandLineArgs()) {
        if ($argument -match '^-{1,2}non') { return $false }
    }

    return $true
}

function Read-UserChoice {
    <#
    .SYNOPSIS
        Asks the user to pick one of several named options, and falls back to a
        stated default when there is nobody to ask.

    .DESCRIPTION
        A wrapper over $Host.UI.PromptForChoice that returns the chosen label as
        text rather than as an index, so the calling code reads as English:

            if ((Read-UserChoice ... -Default 'Cancel') -eq 'Retry') { ... }

        The reason for the wrapper is the unattended case. A prompt that blocks
        for ever is worse than no prompt at all, so when the host cannot ask -
        see Test-HostCanPrompt - the question and the answer taken are both
        written as warnings, and an answer is returned unasked.

        Which answer that is, is deliberately a parameter of its own. -Default is
        what Enter picks, and wants to be the option somebody reached for the
        keyboard in order to get. -Unattended is what is taken with nobody in the
        room, and wants to be the option that destroys nothing. They are often
        not the same: a user who has just closed Excel means Retry, whereas a
        scheduled run must not answer Retry to a file that will still be locked a
        millisecond later. Where they do coincide, -Unattended can be left out
        and -Default serves for both.

    .PARAMETER Message
        The question. Shown above the options, and written as a warning when the
        host cannot prompt.

    .PARAMETER Option
        The options, in the order they are offered, each with & before the letter
        that selects it: '&Retry', '&Discard', '&Cancel'. The label without the &
        is what this function returns.

    .PARAMETER Default
        The label - without its & - that Enter selects. Must be one of -Option.

    .PARAMETER Unattended
        The label returned unasked when the host cannot prompt. Must be one of
        -Option. Defaults to -Default, which is right only where the answer a
        user would want is also the answer that is safe to take without one.

    .PARAMETER Help
        One line per option, shown when the user answers '?'. Optional, and
        matched to -Option by position.

    .PARAMETER Title
        Optional caption above the question.

    .OUTPUTS
        System.String. The chosen label, without its &.

    .EXAMPLE
        Read-UserChoice -Default 'Cancel' -Option '&Retry', '&Cancel' `
            -Message 'The file could not be written. Try again?'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Mandatory)]
        [ValidateCount(2, 9)]
        [string[]]$Option,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Default,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Unattended,

        [Parameter()]
        [string[]]$Help = @(),

        [Parameter()]
        [string]$Title = ''
    )

    $labels = @($Option | ForEach-Object { $_ -replace '&', '' })

    if (-not $PSBoundParameters.ContainsKey('Unattended')) { $Unattended = $Default }

    # Ordinal, so a caller's typo is caught here. Left to PromptForChoice it
    # would surface as an out-of-range index, and on the unattended path it would
    # not surface at all - the wrong answer would simply be taken in silence.
    # Both labels are checked every time, so a mistake in the one that only shows
    # itself on an unattended run is found on the first interactive one.
    $defaultIndex = [Array]::IndexOf($labels, $Default)
    if ($defaultIndex -lt 0) {
        throw "Read-UserChoice: -Default '$Default' is not one of: $($labels -join ', ')"
    }

    if ([Array]::IndexOf($labels, $Unattended) -lt 0) {
        throw "Read-UserChoice: -Unattended '$Unattended' is not one of: $($labels -join ', ')"
    }

    if (-not (Test-HostCanPrompt)) {
        Write-Warning $Message
        Write-Warning "There is no console to answer this, so '$Unattended' is assumed."
        return $Unattended
    }

    $descriptions = @(
        for ($i = 0; $i -lt $Option.Count; $i++) {
            $text = if ($i -lt $Help.Count) { $Help[$i] } else { $labels[$i] }
            [System.Management.Automation.Host.ChoiceDescription]::new($Option[$i], $text)
        }
    )

    $chosen = $Host.UI.PromptForChoice(
        $Title,
        $Message,
        [System.Management.Automation.Host.ChoiceDescription[]]$descriptions,
        $defaultIndex)

    return $labels[$chosen]
}

#endregion Console

#region Run state

function Get-RunState {
    <#
        Captures what a script is about to change about the host it is running
        in, so Restore-RunState can put it back: the console progress style and
        the process priority class.

        The state is handed to the caller instead of being kept here, and that
        is the whole reason this is a pair of functions rather than one. A
        $script: variable in a module belongs to the module, not to the script
        that called it, so a copy kept here would be shared by every script in
        the session and would survive the run that stored it. The caller holds
        the object and hands it back.

        Call it before changing anything - it reads the current values, not the
        defaults.
    #>
    $process = [System.Diagnostics.Process]::GetCurrentProcess()

    [pscustomobject]@{
        Process       = $process
        Priority      = $process.PriorityClass
        ProgressStyle = $PSStyle.Progress.Style
    }
}

function Restore-RunState {
    <#
        Puts back what Get-RunState captured, and is meant to run at every exit
        path - the end of the script, a trap, a finally. An interactive session
        would otherwise keep a lowered process priority long after the run has
        finished, where it would silently apply to everything started
        afterwards, and a progress style the script changed for itself.

        Both are put back unconditionally, so it is safe to call more than once
        and safe in a script that never changed either: writing a value back
        where it already is costs nothing. The priority is guarded because the
        process may refuse to raise it back, which is not worth failing over at
        the end of a run that otherwise went well.
    #>
    param([Parameter(Mandatory)][psobject]$State)

    $PSStyle.Progress.Style = $State.ProgressStyle

    if ($State.Process.PriorityClass -ne $State.Priority) {
        try { $State.Process.PriorityClass = $State.Priority } catch { }
    }
}

#endregion Run state

#region File system

function Get-FreeSpaceBytes {
    <#
        Free bytes on the volume a path lives on, or $null when that cannot be
        worked out - a UNC path, a drive letter that is not a local volume - in
        which case the caller skips its space check rather than refusing to run.

        Not an engine either, and here for the same reason Read-UserChoice is:
        every script that writes a file this large looks at the free space
        first, and none of them could own the helper alone.

        A relative path resolves against the caller's working directory even
        though this runs in the module's session state, because the provider's
        current location belongs to the runspace rather than to a session state.
    #>
    param([Parameter(Mandatory)][string]$Path)

    try {
        # Unresolved, so the volume of a folder that does not exist yet can
        # still be found: under -WhatIf the backup root has not been created,
        # and Resolve-Path would throw and silently cost the dry run its free
        # space check.
        $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $root = [System.IO.Path]::GetPathRoot($full)
        if ($root -match '^[A-Za-z]:\\$') {
            return ([System.IO.DriveInfo]::new($root)).AvailableFreeSpace
        }
    }
    catch {
        # Deliberately quiet: an unknown free space is not a reason to stop.
    }
    return $null
}

#endregion File system

Export-ModuleMember -Function Test-MkvCleanFile, Repair-MkvCandidate,
                              Read-UserChoice, Get-RunState, Restore-RunState,
                              Get-FreeSpaceBytes
