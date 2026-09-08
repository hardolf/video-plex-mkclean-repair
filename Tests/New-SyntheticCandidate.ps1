#requires -Version 7.2

<#
.SYNOPSIS
    Builds a small .mkv that Find-PlexMkcleanCandidates.ps1 reports as a
    candidate, so the tools can be tried on a file nobody minds losing.

.DESCRIPTION
    Version 1.0.0

    These scripts overwrite media files. Before pointing them at a library, it is
    worth watching them work on something disposable: build a file here, scan it,
    repair it, and confirm that the backup is where the tools said it would be.

    Three steps, and each one supplies a different part of what the scan looks
    for:

      1. ffmpeg writes a one-second HEVC test pattern - a real video track, so
         mkvmerge and mkvinfo have something to identify.
      2. mkvmerge remuxes it with zlib compression on the video track, which is
         what makes ContentCompression true.
      3. mkvpropedit stamps the writing application as mkclean, which is what
         makes IsMkclean true.

    -pix_fmt yuv420p in step 1 is not decoration. Without it x265 chooses Main
    4:4:4, mkvinfo cannot name that profile, and the file comes back
    StrictCandidate=True with HevcProfile 'Unknown @L1.0' - strict for entirely
    the wrong reason, which makes the file useless as a fixture because it would
    pass a test that ought to fail. With it the same file reads 'Main @L1.0' and
    StrictCandidate=False, which is the honest answer.

    WHAT THIS FILE IS NOT
    ---------------------
    It is Candidate=True, and it is never a genuine StrictCandidate. The
    difference is where the compression sits. mkvmerge --compression compresses
    the track's frames; real mkclean compresses the CodecPrivate block, which is
    what makes mkvinfo report 'Unknown @L0.0' and what actually breaks playback
    in the Plex app. Running mkclean itself over this file does not reproduce it
    either.

    So the strict-detection path, and the "the profile can be read again" check in
    step 4 of the repair, are NOT exercised by anything built here. Only a file
    that a real mkclean has been through exercises those. See Tests\README.md.

.PARAMETER OutputFolder
    Where the file is written. Created if it does not exist. Defaults to a
    plexmkclean-sample folder under the temporary directory, so nothing lands in
    the repository by accident.

.PARAMETER MKVToolNix
    Folder holding mkvmerge.exe and mkvpropedit.exe.

.PARAMETER FFmpeg
    ffmpeg.exe, or just 'ffmpeg' when it is on the PATH.

.PARAMETER DurationSeconds
    Length of the test pattern. One second is enough to be identified and keeps
    the file at a few kilobytes; a longer one is only useful for watching the
    progress bar move.

.PARAMETER Name
    Base name of the file, without the extension. Square brackets are worth
    trying here - "Some Film (1982) [x265]" - because that is the case -Path has
    to fall back to a literal path for.

.EXAMPLE
    .\New-SyntheticCandidate.ps1

    Writes one candidate to the temporary folder and prints its path.

.EXAMPLE
    $sample = .\New-SyntheticCandidate.ps1 -OutputFolder 'D:\Scratch\mkv'
    ..\Find-PlexMkcleanCandidates.ps1 -Source 'D:\Scratch\mkv' -ReportFile 'D:\Scratch\scan.csv'
    ..\Repair-PlexMkvFile.ps1 -LiteralPath $sample -BackupRootFolder 'D:\Scratch\Backup' -WhatIf

    The whole walkthrough: build a candidate, scan the folder it is in, then see
    what a repair would do to it without doing it.

.NOTES
    Version: 1.0.0

    Version history:
      1.0.0  First version. The recipe itself is older than the script and was
             worked out against Find-PlexMkcleanCandidates.ps1 1.0.9.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder = (Join-Path ([System.IO.Path]::GetTempPath()) 'plexmkclean-sample'),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$MKVToolNix = 'C:\Program Files\MKVToolNix',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$FFmpeg = 'ffmpeg',

    [Parameter()]
    [ValidateRange(1, 60)]
    [int]$DurationSeconds = 1,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Name = 'synthetic-candidate'
)

$ScriptVersion = '1.0.0'

$ErrorActionPreference = 'Stop'

#region Pre-flight

$mkvMerge    = Join-Path $MKVToolNix 'mkvmerge.exe'
$mkvPropEdit = Join-Path $MKVToolNix 'mkvpropedit.exe'

foreach ($tool in @($mkvMerge, $mkvPropEdit)) {
    if (-not (Test-Path -LiteralPath $tool -PathType Leaf)) {
        throw "MKVToolNix executable was not found: $tool"
    }
}

if (-not (Get-Command $FFmpeg -ErrorAction SilentlyContinue)) {
    throw "ffmpeg was not found: $FFmpeg. Pass -FFmpeg with the full path to ffmpeg.exe, or put it on the PATH."
}

if (-not (Test-Path -LiteralPath $OutputFolder -PathType Container)) {
    [void](New-Item -ItemType Directory -Path $OutputFolder -Force)
}

#endregion Pre-flight

#region Build

Write-Host ''
Write-Host "New-SyntheticCandidate $ScriptVersion"

# 128x96 at 5 frames a second: large enough for x265 to encode and for mkvinfo to
# identify, small enough that the finished file is a few kilobytes.
$raw   = Join-Path $OutputFolder ('{0}.raw.mkv' -f $Name)
$final = Join-Path $OutputFolder ('{0}.mkv' -f $Name)

Write-Host '  1/3  ffmpeg: HEVC test pattern'
& $FFmpeg -hide_banner -loglevel error -y `
    -f lavfi -i "testsrc=size=128x96:rate=5:duration=$DurationSeconds" `
    -c:v libx265 -pix_fmt yuv420p `
    $raw

if ($LASTEXITCODE -ne 0) { throw "ffmpeg failed with exit code $LASTEXITCODE." }

Write-Host '  2/3  mkvmerge: zlib compression on the video track'
& $mkvMerge --compression 0:zlib -o $final $raw | Out-Null

# mkvmerge returns 1 for warnings, which are routine on a test pattern and mean
# nothing here; 2 and above are real failures.
if ($LASTEXITCODE -gt 1) { throw "mkvmerge failed with exit code $LASTEXITCODE." }

Write-Host '  3/3  mkvpropedit: writing application set to mkclean'
& $mkvPropEdit $final --edit info --set 'writing-application=mkclean 0.8.7' | Out-Null

if ($LASTEXITCODE -gt 1) { throw "mkvpropedit failed with exit code $LASTEXITCODE." }

Remove-Item -LiteralPath $raw -Force

#endregion Build

#region Result

$file = Get-Item -LiteralPath $final

Write-Host ''
Write-Host ('Built {0} ({1:N0} bytes)' -f $file.FullName, $file.Length) -ForegroundColor Green
Write-Host 'Candidate=True when scanned. Never a genuine StrictCandidate - see the .DESCRIPTION.'
Write-Host ''

# The path, so the file can be piped or captured:
#   $sample = .\New-SyntheticCandidate.ps1
$file.FullName

#endregion Result
