# Upstream

| | |
|---|---|
| Upstream repo | https://github.com/UnknowCao/dsh-outlook |
| Base commit | `5a2ec7d6bbdcaed48a8b1dc04d18dc0bb0c64659` (v2.11.0) |
| Remote name | `upstream` |
| Licence | MIT - see LICENSE and NOTICE |

## Status: the fork was a starting point, not a relationship

This project is no longer tracking upstream. It began from that commit, kept the
COM layer it could not have re-derived (the Exchange quirks documented in
`lib/olk.ps1`), and has since grown its own panel, composer, draft store, card
and Sidebar tab. The upstream version line was retired at 1.0.0.

The remote and the pinned commit stay here for two reasons: the MIT notice
requires the attribution to be traceable, and a diff against `5a2ec7d` is still
the fastest way to see what the COM layer looked like before this project
started changing it. Merging upstream is not planned; if a fix there ever looks
worth taking, take it by hand and note it in `CHANGELOG.md`.

## Why a pinned commit, not a moving `main`

The upstream COM layer encodes hard-won Exchange quirks that are not written
down anywhere else (Jet vs DASL filtering, `IncludeRecurrences` counts, PS 5.1
stdin decoding). Tracking `main` would pull in rewrites of those paths faster
than they can be validated against a real mailbox. Pinning the commit makes
every "did upstream fix this?" question answerable by a diff.

## Keeping the merge cheap

- `lib/olk.ps1`: changes are additive - one helper plus new `switch` cases.
  Do not reformat or refactor existing cases; that is where upstream fixes
  will land.
- `index.js`: changes are additive - new `ctx.tools.register({...})` blocks
  inserted before the closing log line.
- `lib/watch.ps1` is currently untouched so upstream watcher fixes apply
  cleanly. It is scheduled for replacement by the resident bridge.

## Transport quirk on this workstation

`git fetch upstream` fails with
`schannel: AcquireCredentialsHandle failed: SEC_E_NO_CREDENTIALS (0x8009030e)`.
Use the OpenSSL backend for every network operation:

    git -c http.sslBackend=openssl fetch upstream
    git -c http.sslBackend=openssl log --oneline upstream/main

## Reviewing an upstream change

    git -c http.sslBackend=openssl fetch upstream
    git log --oneline 5a2ec7d..upstream/main
    git diff 5a2ec7d upstream/main -- lib/olk.ps1

Apply anything that touches a quirk in the vendored file by hand, then move
the pin in this file and note it in CHANGELOG.md.