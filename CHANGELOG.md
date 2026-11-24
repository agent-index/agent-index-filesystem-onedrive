# Agent-Index Filesystem — OneDrive / SharePoint adapter — Changelog

Earlier versions are described in `adapter.json` `description`.

## [2.7.3] — 2026-11-24 — Credential-store lock around token refresh (bug c1a2)

**PATCH — maintenance release from `release/2.7` (cut at v2.7.2).** Bug `20261118-68dff8bf-143550-c1a2`; the same commits as #7, which stays open for the 4.x line. Pairs with **agent-index-core 3.31.1**.

### Fixed
- **Parallel refreshes no longer overwrite each other's token.** A refresh of the saved sign-in takes a lock on the credential store first. A process that finds another one already saved a different token uses it instead of refreshing again. A stale lock is taken over atomically.
- **A busy lock is not "expired".** It is reported as `authenticated: false`, reason `credential_lock_busy`, `needs_auth: false`, so core 3.31.1's session-start doesn't start a sign-in.
- **A rejected sign-in (`invalid_grant`)** is reported as `NOT_AUTHENTICATED`, reason `expired`, with a plain "sign in again on this computer" message.

### Changed
- Notes the rule that a member's saved sign-in never leaves their device.

### Build
- Bundle built on framework `agent-index-filesystem` `fe0efab`; identical to the bundle tested on the OneDrive 3.x test org (`sha256:eeea9a3f…4adf`).
