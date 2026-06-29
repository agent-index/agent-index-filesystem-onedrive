import { readFile, writeFile, mkdir, rm } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { randomBytes, createHash } from 'node:crypto';
import { tmpdir } from 'node:os';
import {
  AifsError,
  FileNotFoundError,
  PathNotFoundError,
  AccessDeniedError,
  NotAuthenticatedError,
  NotEmptyError,
  AuthFailedError,
  BackendError,
  RevisionConflictError,
  NotImplementedError,
  NotProvisionedError,
  InvalidRoleError,
  InvalidSubjectError,
  InvalidScopeError,
} from '@agent-index/filesystem/errors';

// ─── Platform-reliability helpers (parity with gdrive 2.6.0) ──────────

export const AIFS_SENTINEL = 'AIFS:FILE-END';

/**
 * Returns the sentinel encoding kind ('md' | 'hash' | 'slash' | 'json') if
 * the content's last non-whitespace text is a recognized AIFS:FILE-END
 * marker, else null. Binary ("base64:") content is never sentinel-checked.
 */
export function detectSentinel(content) {
  if (typeof content !== 'string' || content.length === 0) return null;
  if (content.startsWith('base64:')) return null;
  const tail = content.slice(-400).replace(/\s+$/, '');
  if (tail.endsWith(`<!-- ${AIFS_SENTINEL} -->`)) return 'md';
  if (tail.endsWith(`// ${AIFS_SENTINEL}`)) return 'slash';
  if (tail.endsWith(`# ${AIFS_SENTINEL}`)) return 'hash';
  if (/"_file_end"\s*:\s*"AIFS:FILE-END"\s*[}\]\s]*$/.test(tail)) return 'json';
  return null;
}

export const aifsSleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
export const READ_RETRY_BACKOFF_MS = [500, 1000, 2000];

/** Simple-upload ceiling. Graph requires an upload session above ~4 MB. */
const SIMPLE_UPLOAD_MAX_BYTES = 4 * 1024 * 1024;
/** Upload-session chunk size — must be a multiple of 320 KiB per Graph. */
const UPLOAD_CHUNK_BYTES = 5 * 320 * 1024; // 1600 KiB

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';
const REDIRECT_URI = 'http://localhost:3939/'; // matches the registered app (single-tenant public client)
// User.Read.All (delegated, admin-consent-required) is needed by resolveIdentity:
// GET /users/{id} and the $filter on proxyAddresses both require directory read,
// which plain User.Read ("read your OWN profile") does NOT grant. Without it the
// identitymap resolver 403s on every lookup (bug 20260620-8d20ea22-identityperm).
const SCOPES = 'User.Read User.Read.All Files.ReadWrite.All Sites.ReadWrite.All offline_access';

/** Accept either a raw OAuth code or a pasted callback URL; return the code. */
function _extractAuthCode(input) {
  if (input == null) return undefined;
  if (typeof input !== 'string') return input;
  const trimmed = input.trim();
  if (!trimmed) return undefined;
  if (/^https?:\/\//i.test(trimmed) || trimmed.startsWith('/')) {
    try {
      const u = new URL(trimmed, 'http://localhost');
      const code = u.searchParams.get('code');
      if (code) return code;
    } catch { /* fall through */ }
  }
  if (trimmed.includes('code=') && !trimmed.includes(' ')) {
    const match = trimmed.match(/[?&]?code=([^&\s]+)/);
    if (match) return decodeURIComponent(match[1]);
  }
  return trimmed;
}

/**
 * Microsoft OneDrive / SharePoint backend adapter for the AIFS MCP server.
 *
 * Uses the Microsoft Graph driveItem API. Unlike Google Drive (ID-only), Graph
 * supports path-addressing natively — `/drives/{id}/root:/path:` and
 * `/drives/{id}/items/{id}:/rel:` — so this adapter resolves logical AIFS paths
 * directly to Graph addresses without a path->ID walk.
 *
 * Addressing convention (core-ops phase):
 *   - Absolute paths (`/shared/...`)  -> the org-remote SharePoint document
 *     library (the configured site_id/drive_id), or the user's default drive.
 *   - `id:{itemId}/rel` ID anchors    -> the member's OWN OneDrive (`/me/drive`),
 *     for member-space addressing where the caller is granted but cannot
 *     enumerate from a root (standards section "Addressing"). Cross-drive
 *     shared-with-me anchors are an ACL fast-follow concern.
 *
 * Connection config (agent-index.json -> remote_filesystem.connection):
 *   { "tenant_id", "client_id", "site_id"?, "drive_id"? }   — NO client_secret.
 *
 * Auth: single-tenant public client, OAuth auth-code + PKCE, loopback
 * redirect, no secret. Requires "Allow public client flows = Yes" on the app.
 */
export class OneDriveAdapter {
  constructor() {
    this.connection = null;
    this.credentialPath = null;
    this.pkcePath = null;
    this.pkceTmpPath = null;
    this.tokens = null;
    this._codeVerifier = null;
    this._ownDriveOk = false; // cached: member's own OneDrive confirmed provisioned
    // path cache: normalized logical path -> { id, type, etag, ctag }
    this.pathCache = new Map();
  }

  async initialize(connection, credentialStore) {
    this.connection = connection;

    if (!connection.client_id) {
      throw new BackendError('OneDrive connection config missing "client_id"');
    }
    if (!connection.tenant_id) {
      throw new BackendError('OneDrive connection config missing "tenant_id"');
    }
    // Public client by design — a secret means the app was mis-registered as a
    // confidential client. Fail loud (tech-design D1 / dev-environment finding).
    if (connection.client_secret) {
      throw new BackendError(
        'OneDrive connection carries a "client_secret", but this adapter is a public client ' +
        '(PKCE, no secret). Remove client_secret and register the app with "Allow public client flows = Yes".'
      );
    }

    this.credentialPath = join(credentialStore, 'onedrive.json');
    // PKCE verifier persistence between the (separate-process) start/complete
    // auth invocations. PRIMARY is a sandbox-local tmp path (shared across exec
    // processes in the same sandbox, NOT subject to the workspace-mount
    // write-then-immediate-read race that lost the verifier in ms-install-4 —
    // bug 20260615-8d20ea22-pkcerestart). The workspace path is kept as a
    // fallback for environments where tmp isn't shared. Keyed by tenant+client
    // so concurrent installs in one sandbox don't collide. See startAuth/completeAuth.
    this.pkcePath = join(credentialStore, 'onedrive-pkce.json');
    const pkceKey = createHash('sha256').update(`${connection.tenant_id}:${connection.client_id}`).digest('hex').slice(0, 16);
    this.pkceTmpPath = join(tmpdir(), `aifs-onedrive-pkce-${pkceKey}.json`);
    try {
      this.tokens = JSON.parse(await readFile(this.credentialPath, 'utf-8'));
    } catch {
      this.tokens = null;
    }
  }

  // ─── Auth ──────────────────────────────────────────────────────────

  async getAuthStatus() {
    const base = { backend: 'onedrive' };
    if (!this.tokens || !this.tokens.access_token) {
      return { authenticated: false, ...base, reason: 'no_credential' };
    }
    if (this.tokens.expires_at && this.tokens.expires_at < Date.now()) {
      if (this.tokens.refresh_token) {
        try {
          await this._refreshToken();
          return {
            authenticated: true, ...base,
            user_identity: await this._getUserEmail(),
            expires_at: new Date(this.tokens.expires_at).toISOString(),
          };
        } catch {
          return { authenticated: false, ...base, reason: 'expired' };
        }
      }
      return { authenticated: false, ...base, reason: 'expired' };
    }
    return {
      authenticated: true, ...base,
      user_identity: await this._getUserEmail(),
      expires_at: this.tokens.expires_at
        ? new Date(this.tokens.expires_at).toISOString() : undefined,
    };
  }

  _authUrl(verifier) {
    this._codeVerifier = verifier || randomBytes(32).toString('base64url');
    const codeChallenge = createHash('sha256').update(this._codeVerifier).digest('base64url');
    const params = new URLSearchParams({
      client_id: this.connection.client_id,
      response_type: 'code',
      redirect_uri: REDIRECT_URI,
      scope: SCOPES,
      code_challenge: codeChallenge,
      code_challenge_method: 'S256',
      prompt: 'select_account',
    });
    const tenant = this.connection.tenant_id; // single-tenant
    return `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize?${params.toString()}`;
  }

  /** PKCE verifier persistence paths: sandbox-local tmp first (reliable across
   *  exec processes), workspace second (fallback). */
  _pkcePaths() {
    return [this.pkceTmpPath, this.pkcePath].filter(Boolean);
  }

  /** Persist the verifier to every path (best-effort). Returns true if any write landed. */
  async _persistVerifier(verifier) {
    const payload = JSON.stringify({ code_verifier: verifier, created: Date.now() });
    let wrote = false;
    for (const p of this._pkcePaths()) {
      try {
        await mkdir(dirname(p), { recursive: true });
        await writeFile(p, payload, 'utf-8');
        wrote = true;
      } catch (err) {
        console.error(`[aifs] Warning: could not persist PKCE verifier to ${p}: ${err.message}`);
      }
    }
    return wrote;
  }

  /** Read the persisted verifier from the first path that has a valid one. */
  async _readVerifier() {
    for (const p of this._pkcePaths()) {
      try {
        const v = JSON.parse(await readFile(p, 'utf-8'));
        if (v?.code_verifier) return v;
      } catch { /* try next path */ }
    }
    return null;
  }

  /** Remove the persisted verifier from all paths (post-completion cleanup). */
  async _clearVerifier() {
    for (const p of this._pkcePaths()) {
      try { await rm(p, { force: true }); } catch { /* */ }
    }
  }

  async startAuth() {
    // pkcerestart (bug 20260615-8d20ea22-pkcerestart): TWO defenses.
    // (1) Persistence: write the verifier to a sandbox-local tmp path (primary)
    //     so it reliably survives between the separate start/complete exec
    //     processes — the workspace-mount write was lost to a write-then-read
    //     race in ms-install-4 ("verifier didn't carry over"), which is the real
    //     root cause. (2) Reuse: if a still-fresh verifier is already persisted,
    //     REUSE it (same challenge) so an accidentally re-issued `start` doesn't
    //     rotate it and invalidate an already-obtained code. Codes are ~10 min.
    let verifier = null;
    let reused = false;
    const existing = await this._readVerifier();
    const FRESH_MS = 10 * 60 * 1000;
    if (existing?.code_verifier && existing.created && (Date.now() - existing.created) < FRESH_MS) {
      verifier = existing.code_verifier;
      reused = true;
    }

    const authUrl = this._authUrl(verifier); // verifier=null → fresh

    if (!reused) {
      await this._persistVerifier(this._codeVerifier);
    }
    return {
      status: 'awaiting_code',
      auth_url: authUrl,
      message:
        (reused
          ? 'A sign-in is already in progress — resuming it (your earlier code, if any, is still valid). '
          : '') +
        'Open this URL in your browser and sign in with your Microsoft 365 account. After granting ' +
        'access, the page will try to redirect to http://localhost:3939/ and fail to load — that is ' +
        'EXPECTED. Copy the FULL URL from your browser address bar (it contains "?code=...") and pass it ' +
        'back as auth_code to the "complete" step. If sign-in fails with a public-client error, enable ' +
        '"Allow public client flows = Yes" (Entra -> App registrations -> your app -> Authentication -> Advanced settings).',
    };
  }

  async completeAuth(authCode) {
    authCode = _extractAuthCode(authCode);
    if (!authCode) throw new AuthFailedError('No authorization code provided');

    // Recover the PKCE verifier persisted by startAuth (separate process) —
    // tmp-first, then workspace fallback (see _readVerifier / pkcerestart).
    let verifier = this._codeVerifier;
    if (!verifier) {
      const v = await this._readVerifier();
      verifier = v?.code_verifier || null;
    }
    if (!verifier) {
      throw new AuthFailedError(
        'Missing PKCE verifier. Run the authenticate "start" step again, then "complete" with the ' +
        'same credential store (do not delete .agent-index/credentials between the two steps).'
      );
    }

    const tenant = this.connection.tenant_id;
    const tokenUrl = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: this.connection.client_id,
      grant_type: 'authorization_code',
      code: authCode,
      redirect_uri: REDIRECT_URI,
      code_verifier: verifier,
      scope: SCOPES,
    });

    const res = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      const desc = err.error_description || err.error || res.statusText;
      if (/AADSTS7000218|public client|client_assertion|client_secret/i.test(desc)) {
        throw new AuthFailedError(
          'Token exchange failed: the app rejected the public-client (no-secret) flow. ' +
          'Enable "Allow public client flows = Yes" on the app registration (Entra -> App registrations -> ' +
          'your app -> Authentication -> Advanced settings).',
          { retryable: false }
        );
      }
      if (/expired|invalid_grant/i.test(desc)) {
        throw new AuthFailedError(
          'The authorization code expired or was already used (codes are single-use). Run authentication again.',
          { retryable: true }
        );
      }
      throw new AuthFailedError(`Token exchange failed: ${desc}`);
    }

    const data = await res.json();
    this.tokens = {
      access_token: data.access_token,
      refresh_token: data.refresh_token,
      expires_at: Date.now() + data.expires_in * 1000,
    };
    await this._writeCredential(this.tokens);
    await this._clearVerifier(); // remove verifier from tmp + workspace
    const email = await this._getUserEmail();
    return {
      status: 'authenticated',
      user_identity: email,
      message: `Successfully authenticated to Microsoft 365 as ${email}.`,
    };
  }

  // ─── File operations ─────────────────────────────────────────────────

  async read(path) {
    await this._ensureAuth();
    await this._guardOwnDrive(path);
    const addr = this._addr(path);

    const fetchBuffer = async () => {
      const res = await this._graph(addr.content, { rawResponse: true });
      return Buffer.from(await res.arrayBuffer());
    };

    try {
      let buffer = await fetchBuffer();

      // flakyread parity: never return empty for a file metadata says is non-empty.
      if (buffer.length === 0) {
        let declaredSize = 0;
        try {
          const metaRes = await this._graph(`${addr.meta}?$select=size`);
          declaredSize = Number((await metaRes.json())?.size ?? 0);
        } catch { /* treat as genuinely empty */ }
        if (declaredSize > 0) {
          for (const delay of READ_RETRY_BACKOFF_MS) {
            await aifsSleep(delay);
            buffer = await fetchBuffer();
            if (buffer.length > 0) break;
          }
          if (buffer.length === 0) {
            throw new AifsError(
              'AIFS_READ_UNRELIABLE',
              `read: backend returned empty content for "${path}" but metadata reports ${declaredSize} bytes ` +
              `(retried ${READ_RETRY_BACKOFF_MS.length}x). Transient backend failure — retry; do NOT treat as empty.`,
              { path, declared_size: declaredSize, retries: READ_RETRY_BACKOFF_MS.length }
            );
          }
        }
      }

      const text = buffer.toString('utf-8');
      if (text.includes('\0')) return 'base64:' + buffer.toString('base64');
      return text;
    } catch (err) {
      if (err instanceof AifsError) throw err;
      this._handleGraphError(err, path);
    }
  }

  async write(path, content, options = {}) {
    await this._ensureAuth();
    await this._guardOwnDrive(path);
    const addr = this._addr(path);

    // Ensure parent directories exist — Graph does NOT auto-create them on a
    // path-addressed PUT (it 404s). This is the single most common silent
    // failure if omitted.
    await this._ensureDir(this._parentPath(path));

    // Revision-aware write (O1 = cTag): compare current cTag to ifRevision,
    // then PUT with If-Match: eTag as a backend backstop.
    let etagForIfMatch = null;
    if (options.ifRevision) {
      try {
        const metaRes = await this._graph(`${addr.meta}?$select=cTag,eTag`, { allowNotFound: true });
        if (metaRes.status !== 404) {
          const meta = await metaRes.json();
          const currentRevision = meta.cTag || null;
          if (currentRevision !== options.ifRevision) {
            throw new RevisionConflictError(path, options.ifRevision, currentRevision);
          }
          etagForIfMatch = meta.eTag || null;
        }
      } catch (err) {
        if (err instanceof RevisionConflictError) throw err;
        if (err.status !== 404) this._handleGraphError(err, path);
      }
    }

    const isBinary = content.startsWith('base64:');
    const payload = isBinary ? Buffer.from(content.slice(7), 'base64') : Buffer.from(content, 'utf-8');
    const contentType = isBinary ? 'application/octet-stream' : 'text/plain';
    const sentinelKind = detectSentinel(content);

    const doWrite = async () => {
      if (payload.length > SIMPLE_UPLOAD_MAX_BYTES) {
        return this._uploadLarge(addr, payload, etagForIfMatch);
      }
      const headers = { 'Content-Type': contentType };
      if (etagForIfMatch) headers['if-match'] = etagForIfMatch;
      const res = await this._graph(addr.content, { method: 'PUT', headers, body: payload });
      return res.json();
    };

    try {
      let item = await doWrite();

      // Universal integrity check (bug 20260614-8d20ea22-writeverify): every
      // write is verified by size — not just sentinel-bearing files. The PUT /
      // upload-session response carries the stored byte count, so this is free
      // and catches truncated / partial uploads for ALL content (config, JSON,
      // binaries). The sentinel re-read below is the opt-in stronger tail check.
      if (item && typeof item.size === 'number' && item.size !== payload.length) {
        throw new AifsError(
          'AIFS_WRITE_VERIFY_FAILED',
          `write: sent ${payload.length} bytes to "${path}" but the backend stored ${item.size} — ` +
          `the upload was truncated or partial; do not trust the remote copy, re-write from source.`,
          { path, expected_bytes: payload.length, actual_bytes: item.size }
        );
      }

      if (sentinelKind) {
        const verifyOnce = async () => {
          const vRes = await this._graph(addr.content, { rawResponse: true });
          const back = Buffer.from(await vRes.arrayBuffer()).toString('utf-8');
          return detectSentinel(back) === sentinelKind;
        };
        if (!(await verifyOnce())) {
          item = await doWrite();
          if (!(await verifyOnce())) {
            throw new AifsError(
              'AIFS_WRITE_VERIFY_FAILED',
              `write: AIFS:FILE-END sentinel did not survive the upload of "${path}" (retried once). ` +
              `The remote copy is likely tail-truncated — do not trust it; re-write from the canonical source.`,
              { path, sentinel_kind: sentinelKind }
            );
          }
        }
      }

      if (item?.id) {
        this.pathCache.set(this._normalizePath(path), {
          id: item.id, type: 'file', etag: item.eTag, ctag: item.cTag,
        });
      }
      return { revision: item?.cTag || null };
    } catch (err) {
      if (err instanceof AifsError) throw err;
      this._handleGraphError(err, path);
    }
  }

  async list(path, recursive = false) {
    await this._ensureAuth();
    await this._guardOwnDrive(path);
    const addr = this._addr(path);
    try {
      const entries = [];
      let url = addr.child;
      do {
        const res = await this._graph(url);
        const data = await res.json();
        for (const item of data.value || []) {
          const isDir = !!item.folder;
          const entry = { name: item.name, type: isDir ? 'directory' : 'file' };
          if (!isDir) { entry.size = item.size || 0; entry.modified = item.lastModifiedDateTime; }
          const np = this._normalizePath(path);
          const entryPath = np === '/' ? `/${item.name}` : `${np}/${item.name}`;
          this.pathCache.set(entryPath, { id: item.id, type: entry.type, etag: item.eTag, ctag: item.cTag });
          entries.push(entry);
          if (recursive && isDir) {
            const sub = await this.list(entryPath, true);
            for (const s of sub) entries.push({ ...s, name: `${item.name}/${s.name}` });
          }
        }
        url = data['@odata.nextLink'] || null;
      } while (url);
      return entries;
    } catch (err) {
      if (err.status === 404) throw new PathNotFoundError(path);
      this._handleGraphError(err, path);
    }
  }

  async exists(path) {
    if (typeof path !== 'string' || !path) {
      throw new AifsError('INVALID_ARGS', 'exists: "path" must be a non-empty string', { path });
    }
    await this._ensureAuth();
    await this._guardOwnDrive(path);
    const addr = this._addr(path);
    try {
      const res = await this._graph(`${addr.meta}?$select=id,folder,eTag,cTag`, { allowNotFound: true });
      if (res.status === 404) {
        this.pathCache.delete(this._normalizePath(path));
        return { exists: false };
      }
      const data = await res.json();
      const isDir = !!data.folder;
      this.pathCache.set(this._normalizePath(path), {
        id: data.id, type: isDir ? 'directory' : 'file', etag: data.eTag, ctag: data.cTag,
      });
      return { exists: true, type: isDir ? 'directory' : 'file' };
    } catch (err) {
      if (err.status === 404) return { exists: false };
      this._handleGraphError(err, path);
    }
  }

  async stat(path) {
    await this._ensureAuth();
    await this._guardOwnDrive(path);
    const addr = this._addr(path);
    try {
      const res = await this._graph(
        `${addr.meta}?$select=id,size,lastModifiedDateTime,createdDateTime,cTag,eTag,folder,file,parentReference,webUrl`
      );
      const data = await res.json();
      return {
        id: data.id,
        // The item's home drive ID. Pointer-writers capture this so the discovery
        // pointer can carry a fully-qualified cross-drive reference (id:{drive_id}:{id}),
        // letting another member open content shared from this drive (C.1.3 crossdriveread).
        drive_id: data.parentReference?.driveId || null,
        // Browser-openable URL for the item (Graph webUrl). Lets invite-member put a real
        // clickable link to the bootstrap zip in the welcome email instead of a bare aifs
        // path (C.1.3.3 bootstraplinkunavailable). May be absent for some item types.
        web_url: data.webUrl || null,
        size: data.size || 0,
        modified: data.lastModifiedDateTime,
        created: data.createdDateTime,
        revision: data.cTag || null,
        is_dir: !!data.folder,
      };
    } catch (err) {
      if (err.status === 404) throw new FileNotFoundError(path);
      this._handleGraphError(err, path);
    }
  }

  async delete(path) {
    if (typeof path !== 'string' || !path) {
      throw new AifsError('INVALID_ARGS', 'delete: "path" must be a non-empty string', { path });
    }
    await this._ensureAuth();
    await this._guardOwnDrive(path);
    const itemId = await this._resolveItemId(path);
    if (!itemId) throw new FileNotFoundError(path);

    // Non-recursive contract (O2): refuse a non-empty directory even though
    // Graph DELETE is server-side recursive — keeps behavior identical across
    // backends. Hard-delete workflows must remove contents first.
    const cached = this.pathCache.get(this._normalizePath(path));
    if (cached && cached.type === 'directory') {
      const children = await this.list(path, false);
      if (children.length > 0) throw new NotEmptyError(path);
    }

    try {
      await this._graph(`${this._driveBaseFor(path)}/items/${itemId}`, { method: 'DELETE' });
      this.pathCache.delete(this._normalizePath(path));
    } catch (err) {
      if (err.status === 404) throw new FileNotFoundError(path);
      this._handleGraphError(err, path);
    }
  }

  async copy(source, destination) {
    if (typeof source !== 'string' || !source) {
      throw new AifsError('INVALID_ARGS', 'copy: "source" must be a non-empty string', { source });
    }
    if (typeof destination !== 'string' || !destination) {
      throw new AifsError('INVALID_ARGS', 'copy: "destination" must be a non-empty string', { destination });
    }
    await this._ensureAuth();
    await this._guardOwnDrive(source);

    const sourceId = await this._resolveItemId(source);
    if (!sourceId) throw new FileNotFoundError(source);

    const destParent = this._parentPath(destination);
    const destName = this._fileName(destination);
    await this._ensureDir(destParent);
    const parentMeta = await this._statRaw(destParent);
    if (!parentMeta) throw new PathNotFoundError(destParent);

    try {
      // Graph copy is async: 202 + Location monitor URL. Poll to completion so
      // the synchronous AIFS contract holds.
      const res = await this._graph(`${this._driveBaseFor(source)}/items/${sourceId}/copy`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parentReference: { driveId: parentMeta.parentReference?.driveId, id: parentMeta.id },
          name: destName,
        }),
        rawResponse: true,
      });
      const monitor = res.headers.get('location');
      if (res.status === 202 && monitor) {
        await this._pollCopy(monitor, destination);
      }
    } catch (err) {
      this._handleGraphError(err, source);
    }
  }

  async _pollCopy(monitorUrl, destination) {
    const deadline = Date.now() + 60_000;
    let delay = 500;
    for (;;) {
      const res = await fetch(monitorUrl); // monitor URL is pre-authenticated
      const data = await res.json().catch(() => ({}));
      const status = data.status;
      if (status === 'completed' || res.status === 200 || res.status === 303) {
        this.pathCache.delete(this._normalizePath(destination));
        return;
      }
      if (status === 'failed') {
        throw new BackendError(`copy failed: ${data.error?.message || 'backend reported failure'}`);
      }
      if (Date.now() > deadline) {
        throw new BackendError(`copy did not complete within 60s (monitor: ${monitorUrl})`);
      }
      await aifsSleep(delay);
      delay = Math.min(delay * 1.5, 4000);
    }
  }

  // ─── ACL ops (Release B, contract 2.0) ───────────────────────────────
  //
  // Sharing is ADDITIVE ONLY. OneDrive/SharePoint inheritance is never
  // broken: `inherit:false` is deprecated (decision 2026-06-15-deprecate-
  // inherit-false) — if a caller passes it, we ignore it (grant additively)
  // and emit a one-line deprecation notice to stderr. The real limited-
  // visibility pattern is structural inheritance + owned content in the
  // member's own space, which works identically on gdrive and onedrive.
  //
  // Privileged ops (share/unshare/transferOwnership) are only ever invoked
  // through the permission-change-helper (user-clicked Accept, member token);
  // the adapter just implements the Graph call. getPermissions is read-only
  // and directly agent-callable (the verified-outcome gate).

  /** AIFS role → Graph permission roles (additive grant). commenter→read (OneDrive has no commenter). */
  _aifsRoleToGraphRoles(aifsRole) {
    const map = { reader: ['read'], commenter: ['read'], writer: ['write'] };
    if (!Object.prototype.hasOwnProperty.call(map, aifsRole)) {
      throw new InvalidRoleError(aifsRole);
    }
    return map[aifsRole];
  }

  /** Graph permission roles[] → AIFS role. write/owner/full-control → writer; else reader. */
  _graphRolesToAifs(roles) {
    const r = (roles || []).map((x) => String(x).toLowerCase());
    if (r.some((x) => x === 'write' || x === 'owner' || x.includes('full control') || x.startsWith('sp.full'))) {
      return 'writer';
    }
    return 'reader';
  }

  /** Best-effort subject string for a Graph permission object. */
  _permSubject(p) {
    const g = p.grantedToV2 || {};
    const u = g.user || g.siteUser;
    if (u) return u.email || u.loginName || u.displayName || u.id || 'unknown';
    if (g.group) return g.group.email || g.group.displayName || g.group.id || 'group';
    if (g.siteGroup) return g.siteGroup.loginName || g.siteGroup.displayName || g.siteGroup.id || 'group';
    if (Array.isArray(p.grantedToIdentitiesV2) && p.grantedToIdentitiesV2.length) {
      const ids = p.grantedToIdentitiesV2
        .map((i) => i?.user?.email || i?.user?.loginName || i?.user?.displayName)
        .filter(Boolean);
      if (ids.length) return ids.join(', ');
    }
    if (p.grantedTo?.user) return p.grantedTo.user.email || p.grantedTo.user.displayName || 'unknown';
    if (p.link) return `link:${p.link.scope || 'unknown'}`;
    return 'unknown';
  }

  /** Does a Graph permission object grant access to the given subject (email/loginName/objectId)? */
  _permMatchesSubject(p, subjLower) {
    const cands = [];
    const g = p.grantedToV2 || {};
    for (const k of ['user', 'siteUser', 'group', 'siteGroup']) {
      const o = g[k];
      if (o) cands.push(o.email, o.loginName, o.displayName, o.id);
    }
    if (Array.isArray(p.grantedToIdentitiesV2)) {
      for (const i of p.grantedToIdentitiesV2) {
        const o = i?.user || {};
        cands.push(o.email, o.loginName, o.displayName, o.id);
      }
    }
    if (p.grantedTo?.user) cands.push(p.grantedTo.user.email, p.grantedTo.user.displayName, p.grantedTo.user.id);
    return cands.filter(Boolean).some((c) => String(c).toLowerCase() === subjLower);
  }

  /** Reverse-lookup a Graph item id to a known AIFS path via the path cache; null if unknown. */
  _idToPath(itemId) {
    if (!itemId) return null;
    for (const [path, entry] of this.pathCache) {
      if (entry?.id === itemId) return path;
    }
    return null;
  }

  /** Page through a Graph collection, returning all `value` entries. */
  async _graphCollect(firstUrl) {
    const out = [];
    let url = firstUrl;
    while (url) {
      const res = await this._graph(url);
      const data = await res.json();
      out.push(...(data.value || []));
      url = data['@odata.nextLink'] || null;
    }
    return out;
  }

  /**
   * Grant `subject` the `role` at `path` (additive). subject = a member email
   * or an M365 group (email or objectId). Returns { shared, permission_id, path }.
   */
  async share(path, subject, role, options = {}) {
    await this._ensureAuth();
    if (!subject || typeof subject !== 'string') {
      throw new InvalidSubjectError(subject, 'must be an email or group address/objectId');
    }
    const graphRoles = this._aifsRoleToGraphRoles(role);
    if (options.inherit === false) {
      // O5 / decision 2026-06-15-deprecate-inherit-false: never break inheritance.
      process.stderr.write(
        '[aifs] note: inherit:false is deprecated and ignored on onedrive — grant applied additively (parent inheritance unchanged).\n'
      );
    }
    const itemId = await this._resolveItemId(path);
    if (!itemId) throw new PathNotFoundError(path);
    const base = this._driveBaseFor(path);
    const recipient = subject.includes('@') ? { email: subject } : { objectId: subject };
    let res;
    try {
      res = await this._graph(`${base}/items/${itemId}/invite`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          recipients: [recipient],
          roles: graphRoles,
          requireSignIn: true,
          sendInvitation: false, // agent-index sends its own onboarding mail
        }),
      });
    } catch (err) {
      this._handlePermissionError(err, path, subject, role);
    }
    const data = await res.json().catch(() => ({}));
    const permission_id = data?.value?.[0]?.id ?? null;
    return { shared: true, permission_id, path, inherit_disabled: false };
  }

  /**
   * Revoke `subject`'s explicit grant at `path`. Mirrors gdrive: list → match →
   * DELETE by permission id. Returns { unshared: true } if a grant was removed,
   * { unshared: false } if the subject had no explicit permission here (soft).
   */
  async unshare(path, subject) {
    await this._ensureAuth();
    if (!subject || typeof subject !== 'string') {
      throw new InvalidSubjectError(subject, 'must be an email or group address/objectId');
    }
    const itemId = await this._resolveItemId(path);
    if (!itemId) throw new PathNotFoundError(path);
    const base = this._driveBaseFor(path);
    let perms;
    try {
      perms = await this._graphCollect(`${base}/items/${itemId}/permissions`);
    } catch (err) {
      this._handlePermissionError(err, path, subject, null);
    }
    const subjLower = subject.toLowerCase();
    const match = perms.find((p) => this._permMatchesSubject(p, subjLower));
    if (!match) return { unshared: false, path };
    try {
      await this._graph(`${base}/items/${itemId}/permissions/${match.id}`, { method: 'DELETE' });
    } catch (err) {
      this._handlePermissionError(err, path, subject, null);
    }
    return { unshared: true, path };
  }

  /**
   * List permissions at `path`. options.includeInherited (default true) controls
   * whether inherited grants are returned. Read-only; agent-callable.
   */
  async getPermissions(path, options = {}) {
    await this._ensureAuth();
    const includeInherited = options.includeInherited !== false;
    const itemId = await this._resolveItemId(path);
    if (!itemId) throw new PathNotFoundError(path);
    const base = this._driveBaseFor(path);
    let perms;
    try {
      perms = await this._graphCollect(`${base}/items/${itemId}/permissions`);
    } catch (err) {
      this._handlePermissionError(err, path, null, null);
    }
    const result = [];
    for (const p of perms) {
      const inherited = !!p.inheritedFrom;
      if (!includeInherited && inherited) continue;
      let inherited_from = null;
      if (inherited) {
        const src = p.inheritedFrom?.id;
        inherited_from = this._idToPath(src) || (src ? `onedrive-id:${src}` : (p.inheritedFrom?.path || null));
      }
      result.push({
        subject: this._permSubject(p),
        role: this._graphRolesToAifs(p.roles),
        permission_id: p.id || null,
        inherited_from,
        granted_date: null, // Graph permission resource has no creation time
      });
    }
    return { permissions: result };
  }

  /**
   * Permission-aware enumeration under a scope. Portable subset: scope (absolute
   * path), type (folder|file|any), nameContains, maxResults. Mirrors gdrive's
   * contract (returns { results, truncated }). Graph returns only items the
   * caller can access, so permission-awareness is automatic.
   *
   * With nameContains: Graph drive `search(q=...)` (recursive under scope).
   * Without: a children listing of the scope (one level) — sufficient for the
   * path→id and policy-probe usages and permission-aware. Not byte-for-byte
   * identical to gdrive's recursive type-only search; documented as such.
   */
  async search(query) {
    await this._ensureAuth();
    const scope = query?.scope;
    if (!scope || typeof scope !== 'string' || !(scope.startsWith('/') || this._isAnchor(scope))) {
      throw new InvalidScopeError(scope, 'must be an absolute path or id: anchor');
    }
    const type = query.type || 'any';
    if (!['folder', 'file', 'any'].includes(type)) {
      throw new InvalidScopeError(scope, `invalid type "${type}"`);
    }
    const nameContains = query.nameContains || query.name_contains || null;
    const maxResults = Math.min(query.maxResults || query.max_results || 100, 1000);

    const base = this._driveBaseFor(scope);
    const norm = this._normalizePath(scope);
    const isRoot = !this._isAnchor(scope) && norm === '/';
    let scopeId = null;
    if (!isRoot) {
      scopeId = await this._resolveItemId(scope);
      if (!scopeId) throw new InvalidScopeError(scope, 'scope path does not exist or is not visible');
    }

    const select = '$select=id,name,folder,file,parentReference,lastModifiedDateTime,createdBy';
    let firstUrl;
    if (nameContains) {
      const q = encodeURIComponent(String(nameContains));
      const stem = isRoot ? `${base}/root` : `${base}/items/${scopeId}`;
      firstUrl = `${stem}/search(q='${q}')?${select}&$top=${maxResults}`;
    } else {
      const stem = isRoot ? `${base}/root` : `${base}/items/${scopeId}`;
      firstUrl = `${stem}/children?${select}&$top=${maxResults}`;
    }

    let items, truncated = false;
    try {
      let url = firstUrl;
      items = [];
      const res = await this._graph(url);
      const data = await res.json();
      items.push(...(data.value || []));
      truncated = !!data['@odata.nextLink'] && items.length >= maxResults;
    } catch (err) {
      this._handleGraphError(err, scope);
    }

    const ncLower = nameContains ? String(nameContains).toLowerCase() : null;
    const results = [];
    for (const f of items) {
      const isFolder = !!f.folder;
      if (type === 'folder' && !isFolder) continue;
      if (type === 'file' && isFolder) continue;
      if (ncLower && !(f.name || '').toLowerCase().includes(ncLower)) continue;
      let p = this._idToPath(f.id);
      if (!p) {
        const baseScope = isRoot ? '' : norm.replace(/\/$/, '');
        p = this._isAnchor(scope) ? `${norm}/${f.name}` : `${baseScope}/${f.name}`;
      }
      results.push({
        path: p,
        type: isFolder ? 'folder' : 'file',
        name: f.name,
        owner: f.createdBy?.user?.email || f.createdBy?.user?.displayName || null,
        modified: f.lastModifiedDateTime || null,
      });
      if (results.length >= maxResults) break;
    }
    return { results, truncated };
  }

  /**
   * Not supported on OneDrive/SharePoint: items are owned by the user or the
   * site, and Microsoft Graph has no per-item ownership-transfer analog to
   * Drive's. Member departure is handled via the standards' `owner_departed`
   * pointer annotation + M365 admin retention/site action — not a live transfer.
   */
  async transferOwnership() {
    throw new NotImplementedError(
      'transferOwnership',
      'OneDrive/SharePoint (items are owned by the user or site; no per-item transfer in Graph — handle member departure via owner_departed + M365 admin retention)'
    );
  }

  /** Map Graph permissions-API errors to typed AIFS errors (richer than file-op mapping). */
  _handlePermissionError(err, path, subject, role) {
    const status = err.status || err.response?.status;
    const message = err?.body?.error?.message || err?.message || '';
    // identitymap diagnosability: Graph returns a generic "sharingFailed / please
    // try again later" when the recipient doesn't resolve in the tenant — it
    // LOOKS transient but isn't. Classify it as an unresolvable subject so the
    // caller stops retrying and fixes the identity (bug 20260617-8d20ea22-identitymap).
    if (/does not exist|invalid.*recipient|invalid.*principal|could not be found|unknown.*user|sharingFailed|problem sharing/i.test(message)) {
      throw new InvalidSubjectError(
        subject,
        'could not be resolved to a grantable identity in this tenant — resolve the recipient (aifs_resolve_identity) and grant the resolved UPN/objectId, not the roster email'
      );
    }
    switch (status) {
      case 401: throw new NotAuthenticatedError('expired');
      case 403: throw new AccessDeniedError(path);
      case 404: throw new PathNotFoundError(path);
      default:
        throw new BackendError(`Microsoft Graph permissions error (${status ?? 'unknown'}): ${message}`, err);
    }
  }

  /**
   * Resolve a member reference (email / UPN / objectId) to the tenant's
   * grantable identity. Returns { id, upn, mail } — `id` (objectId) is the most
   * robust recipient for a Graph invite. Throws InvalidSubjectError if no user
   * matches. This is the onedrive answer to identitymap: the roster email is
   * often NOT the grantable identity, and the resolution is per-user (UPN for
   * one member, a proxy/vanity for another), so we look it up rather than guess.
   * gdrive's analog is a no-op passthrough (the email IS the grantable identity).
   * Callers (invite-member) resolve ONCE at invite and persist the result as the
   * member's `sharing_identity`; share-spec composition then uses that.
   */
  async resolveIdentity(ref) {
    await this._ensureAuth();
    if (!ref || typeof ref !== 'string') throw new InvalidSubjectError(ref, 'empty reference');
    const r = String(ref).replace(/^mailto:/i, '').trim();
    const pick = (u) => ({ id: u.id, upn: u.userPrincipalName || null, mail: u.mail || null });
    // A 403 / Authorization_RequestDenied here does NOT mean "no such user" — it
    // means the app registration can't READ THE DIRECTORY (missing User.Read.All).
    // Swallowing it as a fall-through and then throwing INVALID_SUBJECT is the
    // errormask half of bug 20260620-8d20ea22-identityperm: it sent a real install
    // hunting for an account that demonstrably existed. Detect it and surface a
    // distinct, actionable ACCESS_DENIED instead; only a genuine 404/empty result
    // is an unresolvable subject.
    const isPermissionDenied = (e) =>
      e?.status === 403 ||
      /Authorization_RequestDenied|Insufficient privileges|insufficient.*scope/i.test(e?.body?.error?.message || e?.message || '');
    let denied = false;
    // 1) Direct GET — ref is already a UPN or objectId.
    try {
      const res = await this._graph(`/users/${encodeURIComponent(r)}?$select=id,userPrincipalName,mail`, { allowNotFound: true });
      if (res.status !== 404) { const u = await res.json(); if (u?.id) return pick(u); }
    } catch (e) { if (isPermissionDenied(e)) denied = true; /* else fall through to filter */ }
    // 2) Filter on mail / UPN / proxyAddresses (proxy covers vanity addresses
    //    like bill@agent-index.ai -> BillSalak@...onmicrosoft.com). proxyAddresses
    //    filtering needs the advanced-query header.
    const esc = r.replace(/'/g, "''");
    const filt = `mail eq '${esc}' or userPrincipalName eq '${esc}' or proxyAddresses/any(p:p eq 'SMTP:${esc}') or proxyAddresses/any(p:p eq 'smtp:${esc}')`;
    try {
      const res = await this._graph(
        `/users?$select=id,userPrincipalName,mail&$count=true&$filter=${encodeURIComponent(filt)}`,
        { headers: { ConsistencyLevel: 'eventual' }, allowNotFound: true }
      );
      if (res.status !== 404) { const data = await res.json(); const u = (data.value || [])[0]; if (u?.id) return pick(u); }
    } catch (e) { if (isPermissionDenied(e)) denied = true; /* else fall through to throw */ }
    if (denied) {
      throw new AccessDeniedError(
        r,
        'resolve the member identity — the Entra app registration is missing the delegated Microsoft Graph permission "User.Read.All". Add it to the app, grant admin consent, then re-authenticate (aifs_authenticate). The account most likely exists; this is a consent gap, not a missing user. Operation',
      );
    }
    throw new InvalidSubjectError(ref, 'no matching user found in the tenant (checked UPN, mail, and proxy addresses)');
  }

  // ─── Large upload (session) ──────────────────────────────────────────

  async _uploadLarge(addr, payload, etagForIfMatch) {
    const headers = { 'Content-Type': 'application/json' };
    if (etagForIfMatch) headers['if-match'] = etagForIfMatch;
    const sessionRes = await this._graph(`${addr.content.replace(/\/content$/, '')}/createUploadSession`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
    });
    const { uploadUrl } = await sessionRes.json();
    const total = payload.length;
    let start = 0;
    let lastItem = null;
    while (start < total) {
      const end = Math.min(start + UPLOAD_CHUNK_BYTES, total);
      const chunk = payload.subarray(start, end);
      const res = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': String(chunk.length),
          'Content-Range': `bytes ${start}-${end - 1}/${total}`,
        },
        body: chunk,
      });
      if (!res.ok && res.status !== 202) {
        const e = new Error(`upload session chunk failed: ${res.status} ${res.statusText}`);
        e.status = res.status;
        throw e;
      }
      if (res.status === 200 || res.status === 201) lastItem = await res.json().catch(() => null);
      start = end;
    }
    return lastItem || {};
  }

  // ─── Provisioning guard ──────────────────────────────────────────────

  /** True when the path targets the member's OWN OneDrive (a bare id-anchor, or a
   *  config with no site/drive so the default drive is /me/drive). A cross-drive
   *  anchor (id:{driveId}:{itemId}) targets ANOTHER member's drive, so it is not
   *  own-drive and must not be gated on the caller's own-drive provisioning. */
  _isOwnDrive(path) {
    if (this._isAnchor(path)) return !this._parseAnchor(path).driveId;
    return (!this.connection.site_id && !this.connection.drive_id);
  }

  /** Guard own-drive access: a member's OneDrive provisions lazily (first
   *  office.com sign-in). Surface a clear NOT_PROVISIONED instead of a
   *  misleading FILE_NOT_FOUND. Cached per process; configured SharePoint
   *  org-root paths skip it entirely. */
  async _guardOwnDrive(path) {
    if (!this._isOwnDrive(path) || this._ownDriveOk) return;
    const r = await this._graph('/me/drive?$select=id', { allowNotFound: true });
    if (r.status === 404) {
      // memberlicense (bug 20260615-8d20ea22-memberlicense): a 404 here has two
      // distinct causes with different fixes — (a) no OneDrive license (admin
      // must assign one) vs (b) licensed but never signed in (member self-fixes
      // at office.com). Best-effort: inspect the error body to point at the
      // right remedy instead of always saying "sign in".
      let detail = '';
      try { const b = await r.json(); detail = `${b?.error?.code || ''} ${b?.error?.message || ''}`; } catch { /* */ }
      if (/licen[sc]|not provisioned for|no.*service plan|sharepoint.*licen|tenant.*sharepoint/i.test(detail)) {
        throw new NotProvisionedError(
          "Your account doesn't have a OneDrive license, so a personal space can't be created. Ask your Microsoft 365 admin to assign a license that includes OneDrive/SharePoint (it's included in Business Standard/Premium and E3). You can still use org shared collections via site membership."
        );
      }
      throw new NotProvisionedError(
        "Your OneDrive isn't set up yet. Sign in once at https://office.com (open OneDrive), then re-run setup. (If your admin says you have no OneDrive license, that's the cause instead — ask them to assign one.)"
      );
    }
    this._ownDriveOk = true;
  }

  /** Resolve a SharePoint site URL to { site_id, drive_id } (adapter-internal
   *  helper used by create-org; not a contract op). */
  async resolveSite(siteUrl) {
    await this._ensureAuth();
    let u;
    try { u = new URL(siteUrl); } catch { throw new AifsError('INVALID_ARGS', `resolveSite: invalid site URL "${siteUrl}"`, { siteUrl }); }
    const rel = u.pathname.replace(/\/+$/, '');
    if (!rel || rel === '/') {
      throw new AifsError('INVALID_ARGS', 'resolveSite: site URL must include a /sites/<name> path', { siteUrl });
    }
    try {
      // Colon form, NO $select (Graph 400s with it on the colon-addressed path).
      const site = await (await this._graph(`/sites/${u.hostname}:${rel}`)).json();
      const drive = await (await this._graph(`/sites/${site.id}/drive?$select=id,name`)).json();
      return { site_id: site.id, drive_id: drive.id, site_web_url: site.webUrl, drive_name: drive.name };
    } catch (err) {
      if (err instanceof AifsError) throw err;
      this._handleGraphError(err, siteUrl);
    }
  }

  // ─── Graph plumbing ──────────────────────────────────────────────────

  async _graph(urlOrPath, options = {}) {
    const { method = 'GET', headers = {}, body, allowNotFound = false } = options;
    const url = urlOrPath.startsWith('https://') ? urlOrPath : `${GRAPH_ROOT}${urlOrPath}`;
    const doFetch = () => fetch(url, {
      method,
      headers: { Authorization: `Bearer ${this.tokens.access_token}`, ...headers },
      body: body !== undefined ? body : undefined,
    });
    let res = await doFetch();
    if (allowNotFound && res.status === 404) return res;
    if (!res.ok) {
      // Honor Retry-After on throttling with one bounded retry.
      if (res.status === 429 || res.status === 503) {
        const wait = Number(res.headers.get('retry-after') || 2) * 1000;
        await aifsSleep(Math.min(wait, 10_000));
        res = await doFetch();
        if (allowNotFound && res.status === 404) return res;
        if (res.ok) return res;
      }
      const err = new Error(`Graph API error: ${res.status} ${res.statusText}`);
      err.status = res.status;
      try { err.body = await res.json(); err.message = err.body?.error?.message || err.message; } catch { /* */ }
      throw err;
    }
    return res;
  }

  /** Drive base for absolute paths: SharePoint library or user default drive. */
  _driveBase() {
    if (this.connection.site_id && this.connection.drive_id) {
      return `/sites/${this.connection.site_id}/drives/${this.connection.drive_id}`;
    }
    if (this.connection.drive_id) return `/drives/${this.connection.drive_id}`;
    return '/me/drive';
  }

  /** Drive base for a given logical path: id-anchors live in the member's own OneDrive
   *  (or, for a cross-drive anchor `id:{driveId}:{itemId}`, in the owner's drive). */
  _driveBaseFor(path) {
    if (!this._isAnchor(path)) return this._driveBase();
    const { driveId } = this._parseAnchor(path);
    return driveId ? `/drives/${driveId}` : '/me/drive';
  }

  _isAnchor(path) {
    return typeof path === 'string' && path.startsWith('id:');
  }

  /**
   * Parse an id-anchor into { driveId, id, rel }.
   *   id:{itemId}[/rel]            → { driveId: null, id, rel }  (member's own /me/drive)
   *   id:{driveId}:{itemId}[/rel]  → { driveId, id, rel }        (cross-drive — an item
   *                                   shared from another member's OneDrive; C.1.3 crossdriveread)
   * Drive IDs and item IDs contain no ':' or '/', so the first ':' after `id:`
   * disambiguates the qualified (cross-drive) form from the bare (own-drive) form.
   */
  _parseAnchor(path) {
    const body = path.slice(3); // strip leading 'id:'
    const slash = body.indexOf('/');
    const head = slash === -1 ? body : body.slice(0, slash);
    const relRaw = slash === -1 ? '' : body.slice(slash + 1);
    const rel = relRaw ? this._normalizePath('/' + relRaw).slice(1) : '';
    const colon = head.indexOf(':');
    if (colon !== -1) {
      return { driveId: head.slice(0, colon), id: head.slice(colon + 1), rel };
    }
    return { driveId: null, id: head, rel };
  }

  /**
   * Build the Graph addresses for a logical path:
   *   { meta, child, content }  — metadata GET, /children, /content endpoints.
   */
  _addr(path) {
    if (this._isAnchor(path)) {
      const { driveId, id, rel } = this._parseAnchor(path);
      // Cross-drive anchor routes to /drives/{driveId}; bare anchor to the member's own drive.
      const base = driveId ? `/drives/${driveId}` : '/me/drive';
      if (!rel) {
        return { meta: `${base}/items/${id}`, child: `${base}/items/${id}/children`, content: `${base}/items/${id}/content` };
      }
      const anchored = `${base}/items/${id}:/${rel}`;
      return { meta: anchored, child: `${anchored}:/children`, content: `${anchored}:/content` };
    }
    const db = this._driveBase();
    const norm = this._normalizePath(path);
    if (norm === '/') {
      return { meta: `${db}/root`, child: `${db}/root/children`, content: `${db}/root/content` };
    }
    const anchored = `${db}/root:${norm}`;
    return { meta: anchored, child: `${anchored}:/children`, content: `${anchored}:/content` };
  }

  async _statRaw(path) {
    try {
      const res = await this._graph(`${this._addr(path).meta}?$select=id,parentReference,folder`, { allowNotFound: true });
      if (res.status === 404) return null;
      return res.json();
    } catch { return null; }
  }

  async _resolveItemId(path) {
    const cached = this.pathCache.get(this._normalizePath(path));
    if (cached?.id) return cached.id;
    const meta = await this._statRaw(path);
    if (!meta) return null;
    this.pathCache.set(this._normalizePath(path), {
      id: meta.id, type: meta.folder ? 'directory' : 'file',
    });
    return meta.id;
  }

  /**
   * Ensure a directory (and its ancestors) exists. Graph does not create
   * parents on a path-addressed PUT, so writes must call this on the parent.
   * Idempotent and 409-tolerant for concurrent creators.
   */
  async _ensureDir(dirPath) {
    // Root / anchor-root always exist (the anchor item is the member space root).
    if (this._isAnchor(dirPath)) {
      const m = /^id:([^/]+)(?:\/(.*))?$/.exec(dirPath);
      if (!m[2]) return; // bare id: anchor — assumed to exist
    } else if (this._normalizePath(dirPath) === '/') {
      return;
    }

    if (await this._statRaw(dirPath)) return;

    const parent = this._parentPath(dirPath);
    await this._ensureDir(parent);

    const name = this._fileName(dirPath);
    try {
      await this._graph(`${this._addr(parent).child}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
      });
    } catch (err) {
      if (err.status !== 409) throw err; // 409 = already created by a concurrent writer
    }
  }

  // ─── Token management ────────────────────────────────────────────────

  async _ensureAuth() {
    if (!this.tokens || !this.tokens.access_token) {
      throw new NotAuthenticatedError('no_credential');
    }
    if (this.tokens.expires_at && (this.tokens.expires_at - 300_000) < Date.now()) {
      if (this.tokens.refresh_token) await this._refreshToken();
      else throw new NotAuthenticatedError('expired');
    }
  }

  async _refreshToken() {
    const tenant = this.connection.tenant_id;
    const tokenUrl = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: this.connection.client_id,
      grant_type: 'refresh_token',
      refresh_token: this.tokens.refresh_token,
      scope: SCOPES,
    });
    const res = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new NotAuthenticatedError(
        `Token refresh failed: ${err.error_description || err.error || 'unknown error'}`
      );
    }
    const data = await res.json();
    this.tokens = {
      access_token: data.access_token,
      refresh_token: data.refresh_token || this.tokens.refresh_token,
      expires_at: Date.now() + data.expires_in * 1000,
    };
    await this._writeCredential(this.tokens);
  }

  async _getUserEmail() {
    try {
      const res = await this._graph('/me?$select=mail,userPrincipalName');
      const data = await res.json();
      return data.mail || data.userPrincipalName || 'unknown';
    } catch { return 'unknown'; }
  }

  async _writeCredential(tokens) {
    await mkdir(dirname(this.credentialPath), { recursive: true });
    await writeFile(this.credentialPath, JSON.stringify(tokens, null, 2), 'utf-8');
  }

  // ─── Path helpers (id-anchor aware) ──────────────────────────────────

  _normalizePath(path) {
    if (this._isAnchor(path)) {
      const m = /^id:([^/]+)(?:\/(.*))?$/.exec(path);
      const rel = m[2] ? m[2].replace(/^\/+/, '').replace(/\/+$/, '').replace(/\/+/g, '/') : '';
      return rel ? `id:${m[1]}/${rel}` : `id:${m[1]}`;
    }
    let p = '/' + path.replace(/^\/+/, '').replace(/\/+$/, '');
    p = p.replace(/\/+/g, '/');
    if (p === '') p = '/';
    return p;
  }

  _parentPath(path) {
    if (this._isAnchor(path)) {
      const norm = this._normalizePath(path);
      const m = /^id:([^/]+)(?:\/(.*))?$/.exec(norm);
      if (!m[2]) return norm; // bare anchor is its own parent (root)
      const segs = m[2].split('/');
      segs.pop();
      return segs.length ? `id:${m[1]}/${segs.join('/')}` : `id:${m[1]}`;
    }
    const norm = this._normalizePath(path);
    const i = norm.lastIndexOf('/');
    return i <= 0 ? '/' : norm.slice(0, i);
  }

  _fileName(path) {
    const norm = this._normalizePath(path);
    return norm.slice(norm.lastIndexOf('/') + 1);
  }

  _handleGraphError(err, path) {
    const status = err.status || err.response?.status;
    switch (status) {
      case 401: throw new NotAuthenticatedError('expired');
      case 403: throw new AccessDeniedError(path);
      case 404: throw new FileNotFoundError(path);
      case 412: throw new RevisionConflictError(path, undefined, undefined);
      default:
        throw new BackendError(`Microsoft Graph API error (${status ?? 'unknown'}): ${err.message}`, err);
    }
  }
}
