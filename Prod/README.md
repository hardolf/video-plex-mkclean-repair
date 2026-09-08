# Prod

Reports from production runs against real media - a log of what has actually been
run, and what it found and repaired.

Contents are deliberately not in version control: the reports name real files on
real drives, and they are a record of one person's library rather than anything
another user of these scripts would want.

Point `-ReportFile` here for a run that matters:

```powershell
.\Find-PlexMkcleanCandidates.ps1 -Source 'Q:\Horror' `
    -ReportFile '.\Prod\Plex-mkclean-candidates - Horror.csv'
```
