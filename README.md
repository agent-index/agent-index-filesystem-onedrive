# agent-index-filesystem-onedrive

Microsoft OneDrive and SharePoint adapter for the agent-index remote filesystem. Connects the `aifs_*` MCP tool interface to OneDrive and SharePoint document libraries via the Microsoft Graph API.

## Overview

This adapter implements the `BackendAdapter` interface from `@agent-index/filesystem` against the Microsoft Graph API. It supports personal OneDrive, organizational OneDrive for Business, and SharePoint document libraries. Path-based access is native to the Graph API, so no ID resolution is needed (unlike the Google Drive adapter).

Members never interact with this package directly. The pre-built bundle is included in the bootstrap zip during org setup and runs as a background MCP server process inside Cowork.

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

## Development

```bash
npm install              # Install dependencies
npm run build            # Bundle, checksum, and stamp adapter.json
npm run build:bundle     # esbuild only (no metadata stamp)
npm test                 # Run tests
```

The `npm run build` command produces `dist/server.bundle.js` (a self-contained single-file MCP server) and updates `adapter.json` with the build timestamp and checksum. Commit both files together.

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
    └── server.bundle.js    # Pre-built bundle (committed to repo)
```

## License

Proprietary — Copyright (c) 2026 Agent Index Inc. All rights reserved. See [LICENSE](LICENSE) for details.
