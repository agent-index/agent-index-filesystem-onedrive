#!/usr/bin/env node
// discover-sharepoint.mjs — resolve a SharePoint site_id + default document-
// library drive_id, using the SAME credential resolution the adapter uses
// (framework loadConfig + adapter.initialize), so there is no hand-built
// credential path to get wrong on Windows/Git Bash.
//
// Usage (AIFS_CONFIG_PATH must point at an authenticated onedrive config):
//   AIFS_CONFIG_PATH=.../agent-index.json node scripts/discover-sharepoint.mjs <site-url>
//   e.g. ... discover-sharepoint.mjs https://agentindex.sharepoint.com/sites/AgentIndexDev
// A full URL is taken on purpose: Git Bash / MSYS rewrites bare POSIX-looking
// args (e.g. "/sites/...") into Windows paths, but leaves "https://..." alone.
// Prints: {"site_id":"...","drive_id":"..."}  on success (exit 0), else an error (exit 1).

import { initEnvironment, loadConfig } from '@agent-index/filesystem';
import { OneDriveAdapter } from '../src/adapters/onedrive.js';

const siteArg = process.argv[2];
if (!siteArg) {
  console.error('usage: discover-sharepoint.mjs <site-url>  (e.g. https://contoso.sharepoint.com/sites/Team)');
  process.exit(1);
}
let host, rel;
try {
  const u = new URL(siteArg);
  host = u.hostname;
  rel = u.pathname.replace(/\/+$/, '');
} catch {
  console.error(`invalid site URL: "${siteArg}"`);
  process.exit(1);
}
if (!rel || rel === '/') {
  console.error(`site URL must include a path, e.g. https://${host}/sites/<name>`);
  process.exit(1);
}

initEnvironment();

try {
  const cfg = await loadConfig();
  const adapter = new OneDriveAdapter();
  await adapter.initialize(cfg.connection, cfg.auth.credentialStore);
  // getAuthStatus loads/refreshes and populates adapter.tokens.
  const status = await adapter.getAuthStatus();
  const token = adapter.tokens?.access_token;
  if (!status?.authenticated || !token) {
    console.error('not authenticated — run probe-onedrive.sh to sign in first');
    process.exit(1);
  }

  const h = { Authorization: `Bearer ${token}` };

  // Site-by-path uses the colon form `/sites/{host}:/{server-relative-path}`.
  // No $select on the colon-addressed path (Graph returns 400 if present).
  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${host}:${rel}`;
  let siteId = null;
  let r = await fetch(siteUrl, { headers: h });
  if (r.ok) {
    siteId = (await r.json()).id;
  } else {
    const body = (await r.text()).slice(0, 300);
    // Fallback: search sites by the final path segment and match on webUrl.
    const name = rel.split('/').filter(Boolean).pop() || '';
    const sr = await fetch(`https://graph.microsoft.com/v1.0/sites?search=${encodeURIComponent(name)}`, { headers: h });
    if (sr.ok) {
      const hits = (await sr.json()).value || [];
      const want = `${host}${rel}`.toLowerCase();
      const hit = hits.find(s => (s.webUrl || '').toLowerCase().endsWith(want)) || hits.find(s => (s.webUrl || '').toLowerCase().includes(name.toLowerCase()));
      if (hit) siteId = hit.id;
    }
    if (!siteId) {
      console.error(`site lookup failed: ${r.status} ${body}\n  tried: ${siteUrl}\n  (search fallback for "${name}" found no match)`);
      process.exit(1);
    }
  }

  r = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive?$select=id,name`, { headers: h });
  if (!r.ok) {
    console.error(`drive lookup failed: ${r.status} ${(await r.text()).slice(0, 300)}`);
    process.exit(1);
  }
  const driveId = (await r.json()).id;

  console.log(JSON.stringify({ site_id: siteId, drive_id: driveId }));
} catch (err) {
  console.error(err?.message || String(err));
  process.exit(1);
}
