#requires -Version 7.2

<#
.SYNOPSIS
    Checks the invariants that must hold for the scripts in this repository,
    without media files, without MKVToolNix and without any module that does not
    ship with PowerShell.

.DESCRIPTION
    Version 1.0.1

    Six checks, none of which needs anything installed:

      1. Every .ps1 and .psm1 parses.

      2. The comment-based help of the three scripts is intact. This is the check
         that earns its place on its own: a single stray keyword inside an
         .EXAMPLE kills the whole help block silently, and Get-Help then answers
         with the file name and nothing else. It has happened here before, in
         Fix-PlexMkcleanCandidates.ps1.

      3. PlexMkclean.psm1 exports exactly the functions it is meant to export -
         no more, so a helper cannot leak into the caller's session by accident,
         and no fewer, so a script cannot lose the function it calls.

      4. No parameter default points at one particular machine. A default may
         name C:\Program Files, because that is where Windows software lives
         everywhere; anything else absolute is somebody's own disk and has no
         business in a published script.

      5. Every file is UTF-8 with a BOM and CRLF throughout, as .gitattributes
         requires. Fix-PlexMkcleanCandidates.ps1 had quietly become LF-only.

      6. The run state - process priority and the console progress style - is put
         back on every exit path. Both kinds are exercised, because they are not
         restored by the same mechanism: a throw, which the script-scope trap
         catches, and an early return, which it does not and which therefore has
         to restore for itself.

    Check 6 runs the three scripts for real, but only as far as their earliest
    exits. It stands in for MKVToolNix with two empty files named mkvmerge.exe
    and mkvinfo.exe - the pre-flight only asks whether they exist - and points
    the scripts at an empty folder and a report with a header and no rows, so
    every run stops before either executable would be started. Nothing is read
    from or written to the project folder, apart from the one header line the
    report fixture is built from; every fixture lives in a temporary directory
    that is removed afterwards.

    What these checks cannot cover is set out in Tests\README.md, along with what
    a fuller test would need and why it is not here.

.PARAMETER ProjectRoot
    The repository root. Defaults to the folder above this one, which is where it
    is when the Tests folder has not been moved.

.EXAMPLE
    .\Invoke-StaticChecks.ps1

    Runs every check and prints one line per invariant, with the detail of any
    failure underneath it. Exits 0 when all pass and 1 when any fails, so it can
    be driven from a build step or a git hook.

.EXAMPLE
    .\Invoke-StaticChecks.ps1 -ProjectRoot 'C:\Code\video-plex-mkclean-repair'

    Checks a copy of the repository somewhere else.

.NOTES
    Version: 1.0.0

    Deliberately plain PowerShell rather than Pester. Six checks do not repay a
    module every reader would have to install first, and the Pester that comes
    with Windows is 3.4.0, whose syntax modern Pester no longer accepts - so
    "just run the tests" would mean an install before it meant anything else.
    If this file ever grows past a couple of dozen assertions, that trade turns
    around and it should become a Pester suite.

    Version history:
      1.0.1  The caller's $PSDefaultParameterValues is cleared for the run. A
             profile on this machine supplied -BackupRootFolder to
             Fix-PlexMkcleanCandidates.ps1, so the case that checks it gives
             up without one saw it given one, and reported a clean return.
      1.0.0  First version.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot)
)

$ScriptVersion = '1.0.1'

$ErrorActionPreference = 'Stop'

# The scripts are run below the way a stranger would run them, so nothing the
# calling session has arranged for itself may reach them. $PSDefaultParameterValues
# is the one that matters: PowerShell hands a value from it to a script as though
# it had been typed on the command line, which is indistinguishable from the
# caller having passed it. A profile doing exactly that for -BackupRootFolder is
# what made the run-state check report a pass that was not one. Assigning here
# shadows the global for this script and everything it calls, and leaves the
# caller's own copy alone.
$PSDefaultParameterValues = @{}

#region Harness

$results = [System.Collections.Generic.List[object]]::new()

function Test-Invariant {
    <#
        Runs one check and records the outcome. The body returns a string per
        failure and nothing at all when the check passes, so a check that finds
        several problems reports all of them rather than only the first.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Body
    )

    try {
        $failures = @(& $Body)
    }
    catch {
        # A check that cannot run is a failure of the check, and saying so is
        # more use than letting the exception end the whole file.
        $failures = @("the check itself failed: $($_.Exception.Message)")
    }

    $results.Add([pscustomobject]@{ Name = $Name; Failures = $failures })

    if ($failures.Count -eq 0) {
        Write-Host '  PASS  ' -ForegroundColor Green -NoNewline
        Write-Host $Name
    }
    else {
        Write-Host '  FAIL  ' -ForegroundColor Red -NoNewline
        Write-Host $Name
        foreach ($failure in $failures) {
            Write-Host ('          {0}' -f $failure) -ForegroundColor Red
        }
    }
}

#endregion Harness

#region Fixtures

function New-TestFixture {
    <#
        A temporary directory holding everything the run-state check needs, and
        nothing that has to be true of the machine it runs on.

        MKVToolNix is stood in for by two empty files with the right names. That
        works because the pre-flight only asks whether they exist, and every case
        in the check stops before either would be started - so no real MKVToolNix
        is needed to test what happens when a script gives up.

        The report fixture takes its header from the real report in the project
        folder rather than repeating the column names here, so it cannot drift
        away from the schema the scripts read. Only the first line is used; what
        the file holds below it does not matter.
    #>
    param([Parameter(Mandatory)][string]$ProjectRoot)

    $root = Join-Path ([System.IO.Path]::GetTempPath()) (
        'plexmkclean-tests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))

    $tools  = Join-Path $root 'MKVToolNix'
    $source = Join-Path $root 'Source'
    $backup = Join-Path $root 'Backup'

    foreach ($folder in @($root, $tools, $source, $backup)) {
        [void](New-Item -ItemType Directory -Path $folder -Force)
    }

    foreach ($executable in @('mkvmerge.exe', 'mkvinfo.exe')) {
        [void](New-Item -ItemType File -Path (Join-Path $tools $executable) -Force)
    }

    $header = Get-Content -LiteralPath (Join-Path $ProjectRoot 'Plex-mkclean-candidates.csv') -TotalCount 1
    $report = Join-Path $root 'no-candidates.csv'
    Set-Content -LiteralPath $report -Value $header -Encoding utf8BOM

    [pscustomobject]@{
        Root         = $root
        ToolsFolder  = $tools
        SourceFolder = $source
        BackupFolder = $backup
        ReportFile   = $report
    }
}

#endregion Fixtures

#region What is checked

# The three scripts, by name, because checks 2 and 3 are about these files in
# particular rather than about whatever happens to be lying in the folder.
$scriptNames = @(
    'Find-PlexMkcleanCandidates.ps1'
    'Fix-PlexMkcleanCandidates.ps1'
    'Repair-PlexMkvFile.ps1'
)

# What PlexMkclean.psm1 is meant to expose. This list is the contract: changing
# it is a decision, not a side effect, which is exactly why it is written out.
$expectedExports = @(
    'Get-FreeSpaceBytes'
    'Get-RunState'
    'Read-UserChoice'
    'Repair-MkvCandidate'
    'Restore-RunState'
    'Test-MkvCleanFile'
) | Sort-Object

# Checks 1, 4 and 5 apply to every file of code in the repository, this one
# included. Work\ and Prod\ are scratch and log folders whose contents are not
# version controlled, and Old\ is an archive, so none of them is scanned.
$codeFiles = @(
    Get-ChildItem -LiteralPath $ProjectRoot -File | Where-Object { $_.Extension -in '.ps1', '.psm1' }
    Get-ChildItem -LiteralPath $PSScriptRoot -File | Where-Object { $_.Extension -in '.ps1', '.psm1' }
) | Sort-Object FullName

Write-Host ''
Write-Host "Invoke-StaticChecks $ScriptVersion"
Write-Host ('Project root:  {0}' -f $ProjectRoot)
Write-Host ('Files of code: {0}' -f $codeFiles.Count)
Write-Host ''

#endregion What is checked

#region Checks

Test-Invariant 'Every .ps1 and .psm1 parses' {
    $failures = @()

    foreach ($file in $codeFiles) {
        $parseErrors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile(
            $file.FullName, [ref]$null, [ref]$parseErrors)

        foreach ($parseError in $parseErrors) {
            $failures += '{0}:{1} {2}' -f
                $file.Name, $parseError.Extent.StartLineNumber, $parseError.Message
        }
    }

    $failures
}

Test-Invariant 'Comment-based help is intact' {
    $failures = @()

    foreach ($name in $scriptNames) {
        # Get-Help on a script needs a path it recognises as one, which a bare
        # file name is not.
        $help = Get-Help (Join-Path $ProjectRoot $name) -ErrorAction SilentlyContinue

        if (-not $help) {
            $failures += "$name : Get-Help returned nothing at all"
            continue
        }

        # A broken help block does not produce an error. Get-Help falls back to
        # the file name as the synopsis and leaves the rest empty, so it is the
        # description and the examples that give the fault away.
        if (-not $help.Description)                { $failures += "$name : no description - the help block is broken" }
        if (@($help.Examples.Example).Count -lt 1) { $failures += "$name : no examples - the help block is broken" }
    }

    $failures
}

Test-Invariant 'PlexMkclean.psm1 exports exactly the intended functions' {
    $module = Import-Module (Join-Path $ProjectRoot 'PlexMkclean.psm1') -Force -PassThru

    try {
        $exported = @($module.ExportedFunctions.Keys | Sort-Object)

        @(
            @($expectedExports | Where-Object { $_ -notin $exported }) |
                ForEach-Object { "missing from the exports: $_" }
            @($exported | Where-Object { $_ -notin $expectedExports }) |
                ForEach-Object { "exported but not on the list: $_" }
        )
    }
    finally {
        Remove-Module -ModuleInfo $module -Force -ErrorAction SilentlyContinue
    }
}

Test-Invariant 'No parameter default points at one particular machine' {
    $failures = @()

    foreach ($file in $codeFiles) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            $file.FullName, [ref]$null, [ref]$null)

        $parameters = $ast.FindAll(
            { $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)

        foreach ($parameter in $parameters) {
            if ($null -eq $parameter.DefaultValue) { continue }

            # Every string literal inside the default, so an array default such
            # as -ExcludePattern is covered element by element.
            $literals = $parameter.DefaultValue.FindAll(
                { $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

            foreach ($literal in $literals) {
                # A drive letter, or a UNC name that starts with a letter or a
                # digit. The second half of that matters: -ExcludePattern holds
                # regexes beginning with two backslashes, and they are patterns,
                # not paths.
                $looksLikeAPath = $literal.Value -match '^[A-Za-z]:[\\/]' -or
                                  $literal.Value -match '^\\\\[A-Za-z0-9]'

                $isStandardInstallLocation = $literal.Value -match '^[Cc]:\\Program Files'

                if ($looksLikeAPath -and -not $isStandardInstallLocation) {
                    $failures += '{0}:{1} -{2} defaults to {3}' -f
                        $file.Name,
                        $parameter.Extent.StartLineNumber,
                        $parameter.Name.VariablePath.UserPath,
                        $literal.Value
                }
            }
        }
    }

    $failures
}

Test-Invariant 'Every file is UTF-8 with a BOM and CRLF throughout' {
    $failures = @()

    foreach ($file in $codeFiles) {
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

        if ($bytes.Length -lt 3 -or
            $bytes[0] -ne 0xEF -or $bytes[1] -ne 0xBB -or $bytes[2] -ne 0xBF) {
            $failures += "$($file.Name) : no UTF-8 BOM"
        }

        $loneLineFeeds = 0
        for ($i = 0; $i -lt $bytes.Length; $i++) {
            if ($bytes[$i] -eq 0x0A -and ($i -eq 0 -or $bytes[$i - 1] -ne 0x0D)) {
                $loneLineFeeds++
            }
        }

        if ($loneLineFeeds -gt 0) {
            $failures += "$($file.Name) : $loneLineFeeds line ending(s) are LF, not CRLF"
        }
    }

    $failures
}

Test-Invariant 'The run state is restored on every exit path' {
    $fixture = New-TestFixture -ProjectRoot $ProjectRoot
    $failures = @()

    $process = [System.Diagnostics.Process]::GetCurrentProcess()

    # A colour nothing else uses, so a script that fails to put the style back
    # cannot pass by accidentally leaving the right value behind.
    $probeStyle = $PSStyle.Foreground.BrightMagenta

    # Every case asks for BelowNormal, so each run really does change the two
    # things the check is about before it exits.
    $cases = @(
        @{
            Name   = 'Find- on a missing MKVToolNix (throw, caught by the trap)'
            Script = 'Find-PlexMkcleanCandidates.ps1'
            Throws = $true
            Args   = @{
                Source     = $fixture.SourceFolder
                MKVToolNix = Join-Path $fixture.Root 'no-such-folder'
                ReportFile = Join-Path $fixture.Root 'find-report.csv'
                Priority   = 'BelowNormal'
            }
        }
        @{
            Name   = 'Find- on an empty source folder (return, which the trap never sees)'
            Script = 'Find-PlexMkcleanCandidates.ps1'
            Throws = $false
            Args   = @{
                Source     = $fixture.SourceFolder
                MKVToolNix = $fixture.ToolsFolder
                ReportFile = Join-Path $fixture.Root 'find-report.csv'
                Priority   = 'BelowNormal'
            }
        }
        @{
            Name   = 'Fix- without -BackupRootFolder (throw)'
            Script = 'Fix-PlexMkcleanCandidates.ps1'
            Throws = $true
            Args   = @{
                SourceData = $fixture.ReportFile
                MKVToolNix = $fixture.ToolsFolder
                Priority   = 'BelowNormal'
            }
        }
        @{
            Name   = 'Fix- on a report holding no candidates (return)'
            Script = 'Fix-PlexMkcleanCandidates.ps1'
            Throws = $false
            Args   = @{
                SourceData       = $fixture.ReportFile
                BackupRootFolder = $fixture.BackupFolder
                MKVToolNix       = $fixture.ToolsFolder
                Priority         = 'BelowNormal'
            }
        }
        @{
            Name   = 'Repair- when nothing matches -Path (throw)'
            Script = 'Repair-PlexMkvFile.ps1'
            Throws = $true
            Args   = @{
                Path             = Join-Path $fixture.SourceFolder 'no-such-file.mkv'
                BackupRootFolder = $fixture.BackupFolder
                MKVToolNix       = $fixture.ToolsFolder
                Priority         = 'BelowNormal'
            }
        }
    )

    $styleBefore = $PSStyle.Progress.Style

    # The scripts write their reports where they are told to, but running them
    # from the fixture keeps anything they might resolve relatively out of the
    # project folder as well.
    Push-Location -LiteralPath $fixture.Root

    try {
        foreach ($case in $cases) {
            $scriptPath     = Join-Path $ProjectRoot $case.Script
            $priorityBefore = $process.PriorityClass
            $splat          = $case.Args
            $threw          = $false

            $PSStyle.Progress.Style = $probeStyle

            try   { & $scriptPath @splat *> $null }
            catch { $threw = $true }

            if ($threw -ne $case.Throws) {
                $expected = if ($case.Throws) { 'a terminating error' } else { 'a clean return' }
                $actual   = if ($threw)       { 'a terminating error' } else { 'a clean return' }
                $failures += '{0}: expected {1}, got {2}' -f $case.Name, $expected, $actual
            }

            if ($process.PriorityClass -ne $priorityBefore) {
                $failures += '{0}: priority left at {1}, was {2}' -f
                    $case.Name, $process.PriorityClass, $priorityBefore

                # Put it back, or every later case reports the same fault again.
                $process.PriorityClass = $priorityBefore
            }

            if ($PSStyle.Progress.Style -ne $probeStyle) {
                $failures += '{0}: the progress style was not put back' -f $case.Name
            }
        }
    }
    finally {
        Pop-Location
        $PSStyle.Progress.Style = $styleBefore
        Remove-Item -LiteralPath $fixture.Root -Recurse -Force -ErrorAction SilentlyContinue
    }

    $failures
}

#endregion Checks

#region Summary

$failed = @($results | Where-Object { $_.Failures.Count -gt 0 })

Write-Host ''

if ($failed.Count -eq 0) {
    Write-Host ('All {0} checks passed.' -f $results.Count) -ForegroundColor Green
    Write-Host ''
    exit 0
}

Write-Host ('{0} of {1} checks failed.' -f $failed.Count, $results.Count) -ForegroundColor Red
Write-Host ''
exit 1

#endregion Summary
