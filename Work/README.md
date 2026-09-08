# Work

Scratch space. Trial runs, throwaway reports, experiments while changing the
scripts. Nothing here is precious - overwriting or deleting any of it should
never matter.

Contents are not in version control. Anything worth keeping belongs in `Prod`
(a real run) or in the repository itself (a change to the tools).

Two scripts here carry the values that are only true on this machine, so the
published scripts one folder up can keep the defaults a stranger should get:

- `Run-Parent.ps1` runs one of them with those values passed as arguments. This
  is the normal way, because an argument is visible at the call site and cannot
  reach anything else.
- `Set-Defaults.ps1` writes them into `$PSDefaultParameterValues` instead, which
  is convenient and invisible: from then on every command in that session may be
  given one without saying so, the test harness included. Opt in per session,
  and `-Clear` puts the session back.
