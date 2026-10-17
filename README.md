# agent-index-filesystem-onedrive

Microsoft OneDrive and SharePoint adapter for the agent-index remote filesystem. Connects the `aifs_*` tool interface to OneDrive and SharePoint document libraries via the Microsoft Graph API.

## Overview

This adapter implements the `BackendAdapter` interface from `@agent-index/filesystem` against the Microsoft Graph API. It supports personal OneDrive, organizational OneDrive for Business, and SharePoint document libraries. Path-based access is native to the Graph API, so no ID resolution is needed (unlike the Google Drive adapter).

Members never interact with this package directly. The pre-built bundle is included in the bootstrap zip during org setup and runs as an on-demand exec process invoked by the shell wrapper.

## Features

- Personal OneDrive and OneDrive for Business support
- SharePoint document library support via site ID
- OAuth2 per-member authentication via Azure AD with automatic token refresh
- Path-based access via Microsoft Graph API
- All 9 `aifs_*` tools supported

## Connection Config

Set by the org admin during `create-org`:

| Field | Required | Description |
|---|---|---|
| `drive_id` | No | OneDrive or SharePoint document library drive ID. Omit for user's default drive. |
| `site_id` | No | SharePoint site ID. Required for SharePoint, omit for personal OneDrive. |
| `tenant_id` | Yes | Azure AD tenant ID for the org's Microsoft 365 tenant. |
| `client_id` | Yes | OAuth 2.0 application (client) ID from Azure AD app registration. |

## Addressing and sharing notes

- **Absolute paths** (`/shared/...`) address the configured drive (`site_id`/`drive_id`), or `/me/drive` when none is configured. They never fall back to another drive.
- **Bare id anchors** (`id:{itemId}[/rel]`) address the caller's own OneDrive (`/me/drive`). If the item is not there (404) and the connection has a configured `drive_id`, the adapter retries once on the configured drive and uses it for the rest of the call, so org items addressed by bare id (for example `id:{org_config_id}`) still resolve. `aifs_stat` `drive_id` reports the drive where the item was found.
- **Drive-qualified anchors** (`id:{driveId}:{itemId}[/rel]`) always address that drive, with no fallback.
- **`aifs_share` with an email outside the tenant** (domain not one of the tenant's verified domains) sends Graph's sharing invitation (`sendInvitation: true`) so the guest can redeem it. Tenant users get a silent grant (`sendInvitation: false`). The result includes `external` and `invitation_sent`.
- **Share errors:** `EXTERNAL_SHARE_BLOCKED` means the tenant or site sharing policy doesn't allow sharing outside the organization. Graph's own text is in `reason`, plus `graph_status` and `graph_code`. When Graph gives only a generic refusal (`sharingFailed` / "There was a problem sharing…") with no policy wording, the error also carries `policy_inferred: true` and its message says the external-sharing setting is the usual cause rather than a certain one (`policy_inferred` is `false` when Graph names the policy). `INVALID_SUBJECT` is used only for a malformed address or a tenant recipient that doesn't resolve. Any other failure is `BACKEND_ERROR` or `ACCESS_DENIED`, with Graph's status, error code, and message included.
- **`aifs_get_permissions` `inherited_from`** is filled from Graph's `inheritedFrom` (a known path, a logical path, or `id:{driveId}:{itemId}`) when Graph reports it. OneDrive and SharePoint often leave it out, for example for access through site membership. In that case it is `null`, which does not prove the grant is direct.

## Development

```bash
npm install              # Install dependencies
npm run build            # Bundle, checksum, and stamp adapter.json
npm run build:bundle     # esbuild only (no metadata stamp)
npm test                 # Run tests
```

The `npm run build` command produces `dist/aifs-exec.bundle.js` (a self-contained single-file exec bundle) and updates `adapter.json` with the build timestamp and checksum. Commit both files together.

## Repository Structure

```
├── adapter.json            # Adapter metadata, connection schema, build info
├── package.json            # Source dependencies and build scripts
├── scripts/
│   └── build.js            # Build pipeline (bundle + checksum + stamp)
├── src/
│   ├── index.js            # Entry point
│   └── adapters/
│       └── onedrive.js     # BackendAdapter implementation
└── dist/
    ├── aifs-exec.bundle.js # Pre-built exec bundle (committed to repo)
    └── aifs-exec.sh        # Shell wrapper
```

## License

Proprietary — Copyright (c) 2026 Agent Index Inc. All rights reserved. See [LICENSE](LICENSE) for details.
