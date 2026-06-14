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

import { initEnvironment, loadConfig } from '@agent-index/filesystem';
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
  for (const key of ['path', 'source', 'destination']) {
    if (typeof out[key] === 'string' && !out[key].startsWith('id:')) {
      out[key] = out[key].replace(/\\/g, '/');
    }
  }
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
    return new Map(Object.entries(data.entries || {}));
  } catch {
    return new Map();
  }
}

async function savePathCache(credentialStore, pathCache) {
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
    await writeFile(cachePath, JSON.stringify({ timestamp: Date.now(), entries }), 'utf-8');
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

async function routeToolCall(adapter, toolName, args) {
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

    case 'aifs_auth_status':
      return adapter.getAuthStatus();

    case 'aifs_authenticate': {
      const action = args.action || 'start';
      if (action === 'start') return adapter.startAuth();
      if (action === 'complete') return adapter.completeAuth(args.auth_code);
      throw new AifsError('INVALID_ARGS', `Unknown auth action: ${action}`,
        { tool: toolName, valid_actions: ['start', 'complete'] });
    }

    // ─── adapter-internal helper (onedrive only) — used by create-org ──
    case 'aifs_resolve_site': {
      requireArgs(toolName, args, ['site_url']);
      return adapter.resolveSite(args.site_url);
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
        'aifs_share', 'aifs_unshare', 'aifs_get_permissions', 'aifs_search', 'aifs_transfer_ownership'],
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

  const cachedPaths = await loadPathCache(config.auth.credentialStore);
  for (const [path, entry] of cachedPaths) {
    if (!adapter.pathCache.has(path)) adapter.pathCache.set(path, entry);
  }

  try {
    const result = await routeToolCall(adapter, toolName, toolArgs);
    const stripped = typeof result === 'string' ? result : stripDebugFields(result);
    if (typeof stripped === 'string') {
      // Byte-exact: file content is emitted without an appended newline
      // (parity with gdrive 2.6.0 F4 — console.log would poison hashes/diffs).
      process.stdout.write(stripped);
    } else {
      console.log(JSON.stringify(stripped, null, 2));
    }
    await savePathCache(config.auth.credentialStore, adapter.pathCache);
  } catch (err) {
    if (err instanceof AifsError) {
      console.log(JSON.stringify(stripDebugFields(err.toResponse()), null, 2));
      process.exit(1);
    }
    console.log(JSON.stringify({ error: 'BACKEND_ERROR', message: err.message }));
    process.exit(1);
  }
}

main();
