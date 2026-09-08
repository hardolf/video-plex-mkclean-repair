# Tests

Two scripts with quite different jobs, and one file they produced.

## `Invoke-StaticChecks.ps1`

```powershell
.\Invoke-StaticChecks.ps1
```

Six invariants about the scripts themselves. No MKVToolNix, no ffmpeg, no media
files, no module to install first - it runs on a bare PowerShell 7.2 in about two
seconds, and exits 0 or 1 so a build step or a git hook can use it.

| # | Check | Why it is here |
|---|-------|----------------|
| 1 | Everything parses | The floor. |
| 2 | Comment-based help is intact | One stray keyword inside an `.EXAMPLE` kills the whole help block **silently** - `Get-Help` then answers with the file name and nothing else. It has happened here. |
| 3 | The module exports exactly its intended functions | No helper leaks into the caller's session; no script loses the function it calls. |
| 4 | No parameter default points at one particular machine | `C:\Program Files` is allowed, because that is where Windows software lives everywhere. Everything else absolute is somebody's own disk. |
| 5 | UTF-8 with BOM, CRLF throughout | What `.gitattributes` asks for. `Fix-PlexMkcleanCandidates.ps1` had quietly become LF-only. |
| 6 | The run state is restored on every exit path | Process priority and the console progress style. Both kinds of exit are exercised, because they are not restored by the same mechanism. |

Check 6 is the only one that runs anything. It drives the three scripts as far as
their earliest exits, standing in for MKVToolNix with two empty files named
`mkvmerge.exe` and `mkvinfo.exe` - the pre-flight only asks whether they exist,
and every case stops before either would be started. Fixtures live in a temporary
directory that is removed afterwards; the project folder is only ever read from,
and only for the one header line the report fixture is built from.

Each of the six was verified by breaking the thing it watches and confirming the
check goes red: an unclosed brace, a bogus `.FOO` keyword inside an `.EXAMPLE`, a
renamed export, a `Y:\` default, a file converted to LF, and a deleted
`Restore-RunState` in front of an early `return`. Worth knowing about check 2: the
**description** is what gives a broken help block away. A broken block still
reports one example - the whole thing collapses into a single blob - so an
example count alone would not have caught it.

Both kinds of exit matter because they are handled differently. A `throw` is
caught by the script-scope `trap`, which restores and stops. An early `return` is
a normal exit that the trap never sees, and those few places have to restore for
themselves - so a change that removes one of *those* calls would break the
scripts silently, while removing one before a `throw` changes nothing.

## `New-SyntheticCandidate.ps1`

```powershell
$sample = .\New-SyntheticCandidate.ps1
```

Builds a small `.mkv` that the scan reports as a candidate. Needs **ffmpeg** and
**MKVToolNix**.

This one is less a test than a way to earn trust in a tool that overwrites media
files. Build a candidate, scan the folder, run a repair with `-WhatIf`, then run
it for real and confirm the backup is where the tools said it would be - all on a
file nobody minds losing, before pointing anything at a library.

## `synthetic-candidate.mkv`

One output of the script above, committed rather than generated, so the tools can be
tried straight out of a clone - no ffmpeg, no MKVToolNix, and nothing pointed at your
own library:

```powershell
.\Find-PlexMkcleanCandidates.ps1 -Source .\Tests
```

reports it as `Candidate=True`. Three things follow from it living in the repository:

- **Copy it before repairing it.** `Fix` and `Repair` overwrite in place, so repairing
  the fixture turns it into a modified tracked file.
  `git checkout -- Tests/synthetic-candidate.mkv` puts it back, but it is easier to
  repair a copy.
- **Do not rebuild it out of habit.** Every rebuild is another blob in the history for
  good. At 11 KB nothing is at stake, but there is no reason to add one unless the
  fixture has to change.
- **It is not the real fault.** Everything under *What none of this covers* applies to
  this file as much as to a freshly built one. A clean run against it says the scan and
  the repair work end to end; it does not say the strict-detection path does.

## What none of this covers

**The real fault.** Files built here are `Candidate=True` and are *never* a
genuine `StrictCandidate`. `mkvmerge --compression` compresses the track's
frames; real mkclean compresses the **CodecPrivate** block, and that is what
makes `mkvinfo` report `Unknown @L0.0` and what actually breaks playback in the
Plex app. Running mkclean itself over a synthetic file does not reproduce it.

So two things stay untested by anything in this folder:

- the strict-detection path in the scan, and
- the "the profile can be read again" check in step 4 of the repair.

Only a file a real mkclean has been through exercises those, and building one
means finding software that produces the fault rather than simulating it.

**The parallel engine**, the Excel and Power Query paths, backup verification,
and the repair itself are not covered either. They need real media files and real
runs, and the honest place for that is a run against files you have a copy of.

**Interruption.** `Ctrl+C` runs neither the `trap` nor, reliably, a `finally`, so
a run stopped that way can leave the process priority lowered. That is a
PowerShell limitation rather than a bug here, and there is nothing to assert
about it.

## Why not Pester

Six checks do not repay a module every reader has to install first. The Pester
that ships with Windows is 3.4.0, whose syntax modern Pester no longer accepts,
so "just run the tests" would mean `Install-Module Pester -Force
-SkipPublisherCheck` before it meant anything else.

If `Invoke-StaticChecks.ps1` ever grows past a couple of dozen assertions, that
trade turns around and it should become a Pester suite.
