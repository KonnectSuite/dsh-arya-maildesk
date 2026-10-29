# Upstream

| | |
|---|---|
| Upstream repo | https://github.com/UnknowCao/dsh-outlook |
| Base commit | `5a2ec7d6bbdcaed48a8b1dc04d18dc0bb0c64659` (v2.11.0) |
| Remote name | `upstream` |
| Licence | MIT - see LICENSE and NOTICE |

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