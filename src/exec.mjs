#!/usr/bin/env node

/**
 * aifs-exec — On-demand single-invocation executor for the AIFS filesystem
 * (Microsoft OneDrive / SharePoint adapter).
 *
 * Executes exactly one tool call per invocation and exits. Called via bash:
 *   node aifs-exec.bundle.js aifs_read '{"path":"/projects/foo/project.md"}'
 *
 * Mirrors the gdrive executor: default-quiet output, byte-exact read stdout,
 * persisted path cache, content/content_file/content_stdin write payloads,
 * and if_revision pass-through for revision-aware writes.
 */

import {
  initEnvironment,
  loadConfig,
  // CI-025 (shared library 2.3.0): org path resolution and session state.
  resolveOrgMap,
  canonicalizeArgs,
  touchesMappedPrefix,
  isProtectedOrgPath,
  missingSegments,
  needsOrgMap,
  IDEMPOTENT_TOOLS,
  readPendingAuth,
  writePendingAuth,
  clearPendingAuth,
  reusedAuthResponse,
  writeIdentity,
  readIdentity,
  bootstrapState,
} from '@agent-index/filesystem';
import { AifsError } from '@agent-index/filesystem/errors';
import { OneDriveAdapter } from './adapters/onedrive.js';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { join, dirname } from 'node:path';

const VERBOSE = process.env.AIFS_VERBOSE === '1' || process.argv.includes('--verbose');
const DEBUG_FIELDS = new Set(['debug', 'raw_response', '_trace', '_timing']);

function stripDebugFields(value) {
  if (VERBOSE || value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) return value.map(stripDebugFields);
  const out = {};
  for (const [k, v] of Object.entries(value)) {
    if (DEBUG_FIELDS.has(k)) continue;
    out[k] = stripDebugFields(v);
  }
  return out;
}

function normalizePathArgs(args) {
  if (!args || typeof args !== 'object') return args;
  const out = { ...args };
  // NOTE: do not touch id:{itemId} anchors (no backslashes) — only logical paths.
  const fix = (p) => (typeof p === 'string' && !p.startsWith('id:') ? p.replace(/\\/g, '/') : p);
  for (const key of ['path', 'source', 'destination', 'scope']) {
    if (typeof out[key] === 'string') out[key] = fix(out[key]);
  }
  if (Array.isArray(out.entries)) out.entries = out.entries.map((e) => (e && typeof e.path === 'string' ? { ...e, path: fix(e.path) } : e));
  if (Array.isArray(out.paths)) out.paths = out.paths.map(fix);
  return out;
}

// ─── Path cache persistence ───────────────────────────────────────────

const PATH_CACHE_FILENAME = 'path-cache.json';
const PATH_CACHE_MAX_AGE_MS = 30 * 60 * 1000;
const PATH_CACHE_MAX_ENTRIES = 2000;

async function loadPathCache(credentialStore) {
  const cachePath = join(credentialStore, PATH_CACHE_FILENAME);
  try {
    const data = JSON.parse(await readFile(cachePath, 'utf-8'));
    if (data.timestamp && (Date.now() - data.timestamp) > PATH_CACHE_MAX_AGE_MS) return new Map();
    const map = new Map(Object.entries(data.entries || {}));
    // CI-025: keep the first-written time so re-saving doesn't extend the expiry.
    map._createdAt = typeof data.timestamp === 'number' ? data.timestamp : Date.now();
    return map;
  } catch {
    return new Map();
  }
}

async function savePathCache(credentialStore, pathCache, createdAt = Date.now()) {
  const cachePath = join(credentialStore, PATH_CACHE_FILENAME);
  const entries = {};
  let count = 0;
  for (const [key, value] of pathCache) {
    if (count >= PATH_CACHE_MAX_ENTRIES) break;
    entries[key] = value;
    count++;
  }
  try {
    await mkdir(dirname(cachePath), { recursive: true });
    await writeFile(cachePath, JSON.stringify({ timestamp: createdAt, entries }), 'utf-8');
  } catch { /* non-fatal */ }
}

// ─── Tool routing ─────────────────────────────────────────────────────

function requireArgs(toolName, args, required) {
  const missing = [];
  for (const entry of required) {
    const [key, kind] = Array.isArray(entry) ? entry : [entry, null];
    const v = args[key];
    if (v === undefined || v === null) { missing.push(key); continue; }
    if (kind === 'path' && (typeof v !== 'string' || v === '')) missing.push(key);
  }
  if (missing.length > 0) {
    throw new AifsError('INVALID_ARGS',
      `${toolName}: missing or empty required argument(s): ${missing.join(', ')}`,
      { tool: toolName, missing });
  }
}

async function routeToolCall(adapter, toolName, args, ctx = {}) {
  switch (toolName) {
    case 'aifs_read':
      requireArgs(toolName, args, [['path', 'path']]);
      return adapter.read(args.path);

    case 'aifs_write': {
      requireArgs(toolName, args, [['path', 'path']]);
      let content = args.content;
      if (content === undefined || content === null) {
        if (typeof args.content_file === 'string' && args.content_file.length > 0) {
          const payload = await readFile(args.content_file);
          content = args.encoding === 'base64' ? payload.toString('base64') : payload.toString('utf-8');
        } else if (args.content_stdin === true) {
          const chunks = [];
          for await (const chunk of process.stdin) chunks.push(chunk);
          const payload = Buffer.concat(chunks);
          content = args.encoding === 'base64' ? payload.toString('base64') : payload.toString('utf-8');
        } else {
          requireArgs(toolName, args, ['content']);
        }
      }
      if (args.encoding === 'base64' && !content.startsWith('base64:')) {
        content = 'base64:' + content;
      }
      const writeOptions = {};
      if (typeof args.if_revision === 'string' && args.if_revision.length > 0) {
        writeOptions.ifRevision = args.if_revision;
      }
      const res = await adapter.write(args.path, content, writeOptions);
      return { success: true, path: args.path, revision: res?.revision ?? null };
    }

    case 'aifs_list': {
      requireArgs(toolName, args, [['path', 'path']]);
      const entries = await adapter.list(args.path, args.recursive ?? false);
      return { entries };
    }

    case 'aifs_exists':
      requireArgs(toolName, args, [['path', 'path']]);
      return adapter.exists(args.path);

    case 'aifs_stat':
      requireArgs(toolName, args, [['path', 'path']]);
      return adapter.stat(args.path);

    case 'aifs_delete': {
      requireArgs(toolName, args, [['path', 'path']]);
      await adapter.delete(args.path);
      return { success: true };
    }

    case 'aifs_copy': {
      requireArgs(toolName, args, [['source', 'path'], ['destination', 'path']]);
      await adapter.copy(args.source, args.destination);
      return { success: true };
    }

    case 'aifs_auth_status': {
      // CI-025 (§5): also report the bootstrap state, from local files only.
      const status = await adapter.getAuthStatus();
      let email = status?.user_identity && status.user_identity !== 'unknown' ? status.user_identity : null;
      if (email && ctx.credentialStore) await writeIdentity(ctx.credentialStore, email);
      if (!email && ctx.credentialStore) email = await readIdentity(ctx.credentialStore);
      const b = bootstrapState({ configPath: process.env.AIFS_CONFIG_PATH, authenticated: !!status?.authenticated, email });
      return {
        ...status,
        bootstrap: b.bootstrap,
        ...(b.member_hash ? { member_hash: b.member_hash } : {}),
        ...(b.reason ? { bootstrap_reason: b.reason } : {}),
      };
    }

    case 'aifs_authenticate': {
      // pkcerestart (bug 20260615-8d20ea22-pkcerestart): if the caller passes an
      // auth_code but forgets action:"complete", infer complete — defaulting to
      // "start" silently restarted the flow and (pre-fix) rotated the verifier,
      // invalidating the code. With the adapter's verifier-reuse this is already
      // harmless, but inferring intent removes the footgun entirely.
      const action = args.action || (args.auth_code ? 'complete' : 'start');
      if (action === 'start') {
        // CI-025 (§4): a second start returns the sign-in already in progress
        // (under 10 minutes old; the adapter also reuses its PKCE verifier for
        // that window), unless force:true.
        if (!args.force && ctx.credentialStore) {
          const pending = await readPendingAuth(ctx.credentialStore);
          // Reuse only while the PKCE verifier that link was made with still exists;
          // otherwise complete would fail ("Missing PKCE verifier").
          const verifier = pending && typeof adapter._readVerifier === 'function' ? await adapter._readVerifier() : null;
          if (pending && verifier?.code_verifier) return reusedAuthResponse(pending);
        }
        if (args.force && typeof adapter._clearVerifier === 'function') await adapter._clearVerifier();
        const started = await adapter.startAuth();
        if (ctx.credentialStore && started?.auth_url) await writePendingAuth(ctx.credentialStore, started);
        return started;
      }
      if (action === 'complete') {
        const done = await adapter.completeAuth(args.auth_code);
        if (ctx.credentialStore) {
          await clearPendingAuth(ctx.credentialStore);
          if (done?.user_identity) await writeIdentity(ctx.credentialStore, done.user_identity);
        }
        return done;
      }
      throw new AifsError('INVALID_ARGS', `Unknown auth action: ${action}`,
        { tool: toolName, valid_actions: ['start', 'complete'] });
    }

    // ─── adapter-internal helper (onedrive only) — used by create-org ──
    case 'aifs_resolve_site': {
      requireArgs(toolName, args, ['site_url']);
      return adapter.resolveSite(args.site_url);
    }

    // Resolve a member reference (email/UPN/objectId) to the tenant's grantable
    // identity {id, upn, mail}. Used by invite-member to persist sharing_identity
    // (identitymap). On gdrive the analog is a passthrough.
    case 'aifs_resolve_identity': {
      // Accept the common aliases callers reach for (email/subject/identity/upn/…)
      // rather than requiring exactly `ref` — requiring the one exact key cost a
      // wasted round-trip in a real invite (resolve-ref-arg hiccup). Coalesce, then
      // still requireArgs so a truly empty call gets the clear missing-arg error.
      const ref = args.ref ?? args.email ?? args.subject ?? args.identity
        ?? args.upn ?? args.user ?? args.member ?? args.recipient;
      requireArgs(toolName, { ref }, ['ref']);
      return adapter.resolveIdentity(ref);
    }

    // ─── v2.0 access-control ops — ACL fast-follow (currently NOT_IMPLEMENTED) ──
    case 'aifs_share': {
      requireArgs(toolName, args, [['path', 'path'], 'subject', 'role']);
      return adapter.share(args.path, args.subject, args.role, {});
    }
    case 'aifs_unshare': {
      requireArgs(toolName, args, [['path', 'path'], 'subject']);
      return adapter.unshare(args.path, args.subject);
    }
    case 'aifs_get_permissions': {
      requireArgs(toolName, args, [['path', 'path']]);
      return adapter.getPermissions(args.path, {});
    }
    case 'aifs_search': {
      requireArgs(toolName, args, ['scope']);
      return adapter.search({ scope: args.scope, type: args.type, nameContains: args.name_contains, maxResults: args.max_results });
    }
    case 'aifs_transfer_ownership': {
      requireArgs(toolName, args, [['path', 'path'], 'new_owner']);
      return adapter.transferOwnership(args.path, args.new_owner);
    }

    // ─── Batch ops (bulkuploadserial) — many files in ONE process ──────
    case 'aifs_write_batch': {
      requireArgs(toolName, args, ['entries']);
      if (!Array.isArray(args.entries) || args.entries.length === 0) {
        throw new AifsError('INVALID_ARGS', `${toolName}: 'entries' must be a non-empty array of {path, content|content_file}`, { tool: toolName });
      }
      const resolved = [];
      for (let i = 0; i < args.entries.length; i++) {
        const e = args.entries[i];
        if (!e || typeof e.path !== 'string' || e.path === '') {
          throw new AifsError('INVALID_ARGS', `${toolName}: entries[${i}] is missing a 'path'`, { tool: toolName, index: i });
        }
        let content = e.content;
        if (content === undefined || content === null) {
          if (typeof e.content_file === 'string' && e.content_file.length > 0) {
            const payload = await readFile(e.content_file);
            content = e.encoding === 'base64' ? payload.toString('base64') : payload.toString('utf-8');
          } else {
            throw new AifsError('INVALID_ARGS', `${toolName}: entries[${i}] ('${e.path}') has neither 'content' nor 'content_file'`, { tool: toolName, index: i });
          }
        }
        if (e.encoding === 'base64' && !content.startsWith('base64:')) content = 'base64:' + content;
        resolved.push({ path: e.path.replace(/\\/g, '/'), content });
      }
      return adapter.writeBatch(resolved);
    }

    case 'aifs_stat_batch': {
      requireArgs(toolName, args, ['paths']);
      if (!Array.isArray(args.paths) || args.paths.length === 0) {
        throw new AifsError('INVALID_ARGS', `${toolName}: 'paths' must be a non-empty array`, { tool: toolName });
      }
      const paths = args.paths.map(p => (typeof p === 'string' ? p.replace(/\\/g, '/') : p));
      return adapter.statBatch(paths);
    }

    default:
      throw new AifsError('UNKNOWN_TOOL', `Unknown tool: ${toolName}`, { tool: toolName });
  }
}

// ─── Main ─────────────────────────────────────────────────────────────

async function main() {
  const args = process.argv.slice(2);
  if (args.length === 0 || args[0] === '--help' || args[0] === '-h') {
    console.log(JSON.stringify({
      usage: 'aifs-exec <tool_name> [json_args]',
      tools: ['aifs_read', 'aifs_write', 'aifs_list', 'aifs_exists', 'aifs_stat',
        'aifs_delete', 'aifs_copy', 'aifs_auth_status', 'aifs_authenticate', 'aifs_resolve_site',
        'aifs_resolve_identity', 'aifs_share', 'aifs_unshare', 'aifs_get_permissions',
        'aifs_search', 'aifs_transfer_ownership',
        'aifs_write_batch', 'aifs_stat_batch'],
      notes: [
        'Org locations are addressed by plain path: /org-config.json, /members-registry.json, /shared/..., /{collection}/... The executor finds them from org-config on the org site drive (2.8.0). Drive-qualified id:{drive_id}:{item_id} addresses still work and pass through unchanged.',
        'aifs_authenticate {action?: "start"|"complete", auth_code?, force?}: start returns the sign-in already in progress (under 10 minutes old) instead of a new link; pass force:true to start over.',
        'aifs_auth_status also returns bootstrap: ready | needs_auth | needs_setup, from local files only.',
      ],
    }, null, 2));
    process.exit(0);
  }

  const toolName = args[0];
  let toolArgs = {};
  if (args[1] && !args[1].startsWith('--')) {
    try { toolArgs = JSON.parse(args[1]); }
    catch (err) {
      console.log(JSON.stringify({ error: 'INVALID_ARGS', message: `Failed to parse JSON arguments: ${err.message}`, input_preview: args[1].slice(0, 120) }));
      process.exit(1);
    }
  }
  toolArgs = normalizePathArgs(toolArgs);

  initEnvironment();

  let config;
  try { config = await loadConfig(); }
  catch (err) { console.log(JSON.stringify({ error: 'CONFIG_ERROR', message: err.message })); process.exit(1); }

  if (config.backend !== 'onedrive') {
    console.log(JSON.stringify({ error: 'CONFIG_ERROR', message: `This package only supports "onedrive" backend. Config specifies "${config.backend}".` }));
    process.exit(1);
  }

  const adapter = new OneDriveAdapter();
  try { await adapter.initialize(config.connection, config.auth.credentialStore); }
  catch (err) { console.log(JSON.stringify({ error: 'INIT_ERROR', message: `Adapter initialization failed: ${err.message}` })); process.exit(1); }

  const credentialStore = config.auth.credentialStore;
  const cachedPaths = await loadPathCache(credentialStore);
  const cacheCreatedAt = cachedPaths._createdAt || Date.now();
  for (const [path, entry] of cachedPaths) {
    if (!adapter.pathCache.has(path)) adapter.pathCache.set(path, entry);
  }

  // CI-025 (technical design §2-§3): org map. On OneDrive plain paths already
  // work; seeding registers the org's own item ids on the site drive so a bare
  // id: for an org item skips the own-drive probe. resolveOrgMap never throws.
  const warn = (m) => { try { process.stderr.write(m + '\n'); } catch { /* ignore */ } };
  let orgMap = null;
  let orgMapReason = null;
  if (needsOrgMap(toolName, toolArgs)) {
    try {
      const ctx = { config, adapter, credentialStore, warn };
      let r = await resolveOrgMap(ctx);
      orgMap = r.map;
      if (!orgMap) {
        orgMapReason = r.reason;
        warn(`aifs: org map unavailable (${r.reason}); paths resolved without it`);
      } else {
        for (const seg of missingSegments(orgMap, toolName, toolArgs)) {
          r = await resolveOrgMap({ ...ctx, mode: 'missing-prefix', segment: seg });
          if (r.map) orgMap = r.map;
        }
        toolArgs = canonicalizeArgs(orgMap, toolName, toolArgs);
        adapter.seedOrgPrefixes(orgMap.prefixes);
      }
    } catch (err) {
      // Never let org path resolution fail a call (technical design §2.5).
      orgMap = null;
      orgMapReason = `org_map_error:${err?.message || 'error'}`;
      warn(`aifs: org map unavailable (${orgMapReason}); paths resolved without it`);
    }
  }

  let exitCode = 0;
  try {
    if (toolName === 'aifs_delete' && orgMap && isProtectedOrgPath(orgMap, toolArgs?.path)) {
      throw new AifsError('PROTECTED_ORG_PATH',
        `aifs_delete refused: "${toolArgs.path}" is an org root (the shared root, a collection folder, org-config or the members registry).`,
        { path: toolArgs.path });
    }
    let result;
    try {
      result = await routeToolCall(adapter, toolName, toolArgs, { credentialStore });
    } catch (err) {
      const notFound = err?.code === 'FILE_NOT_FOUND' || err?.code === 'PATH_NOT_FOUND';
      if (!(orgMap && notFound && IDEMPOTENT_TOOLS.has(toolName) && touchesMappedPrefix(orgMap, toolName, toolArgs))) throw err;
      const r = await resolveOrgMap({ config, adapter, credentialStore, mode: 'not-found', warn });
      if (!r.rebuilt || !r.map) throw err;
      orgMap = r.map;
      adapter.seedOrgPrefixes(orgMap.prefixes);
      result = await routeToolCall(adapter, toolName, toolArgs, { credentialStore });
    }
    const stripped = typeof result === 'string' ? result : stripDebugFields(result);
    if (typeof stripped === 'string') {
      // Byte-exact: file content is emitted without an appended newline
      // (parity with gdrive 2.6.0 F4 — console.log would poison hashes/diffs).
      process.stdout.write(stripped);
    } else {
      console.log(JSON.stringify(stripped, null, 2));
    }
  } catch (err) {
    exitCode = 1;
    if (err instanceof AifsError) {
      const resp = stripDebugFields(err.toResponse());
      // §2.5: a call that ran without the org map and still failed says why the map was missing.
      if (orgMapReason && resp && typeof resp === 'object') resp.org_map = orgMapReason;
      console.log(JSON.stringify(resp, null, 2));
    } else {
      console.log(JSON.stringify({ error: 'BACKEND_ERROR', message: err.message }));
    }
  } finally {
    // CI-025: save the path cache after errors too (it was skipped before).
    await savePathCache(credentialStore, adapter.pathCache, cacheCreatedAt);
  }
  if (exitCode) process.exit(exitCode);
}

main();
