import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { randomBytes, createHash } from 'node:crypto';
import {
  FileNotFoundError,
  PathNotFoundError,
  AccessDeniedError,
  NotAuthenticatedError,
  NotEmptyError,
  AuthFailedError,
  BackendError,
} from '@agent-index/filesystem/errors';

/**
 * Microsoft OneDrive/SharePoint backend adapter for the AIFS MCP server.
 *
 * Uses the Microsoft Graph API to access OneDrive and SharePoint document libraries.
 * Unlike Google Drive, the Graph API supports path-based access natively via
 * /drive/root:/path/to/file, which simplifies the adapter considerably.
 *
 * Connection config expected in agent-index.json:
 * {
 *   "tenant_id": "...",             // Azure AD tenant ID (or "common" for multi-tenant)
 *   "client_id": "...",             // Azure AD app registration client ID
 *   "drive_id": "...",              // OneDrive/SharePoint drive ID (optional — defaults to user's drive)
 *   "site_id": "...",               // SharePoint site ID (optional — for SharePoint document libraries)
 *   "root_path": "/"               // Root path within the drive (optional — defaults to root)
 * }
 */
export class OneDriveAdapter {
  constructor() {
    this.connection = null;
    this.credentialPath = null;
    this.tokens = null;

    // PKCE state for auth flow
    this._codeVerifier = null;

    // Path cache: maps logical path -> { id, type, etag }
    // Used for operations that need item IDs (copy, delete)
    this.pathCache = new Map();
  }

  /**
   * Initialize the adapter with connection config and credential store path.
   */
  async initialize(connection, credentialStore) {
    this.connection = connection;

    if (!connection.client_id) {
      throw new BackendError('OneDrive connection config missing "client_id"');
    }
    if (!connection.tenant_id) {
      throw new BackendError('OneDrive connection config missing "tenant_id"');
    }

    this.credentialPath = join(credentialStore, 'onedrive.json');

    // Try to load stored credentials
    try {
      this.tokens = JSON.parse(await readFile(this.credentialPath, 'utf-8'));
    } catch {
      // No stored credentials — that's fine, member will authenticate
      this.tokens = null;
    }
  }

  // ─── Auth ────────────────────────────────────────────────────────────

  async getAuthStatus() {
    const base = { backend: 'onedrive' };

    if (!this.tokens || !this.tokens.access_token) {
      return { authenticated: false, ...base, reason: 'no_credential' };
    }

    // Check if token is expired
    if (this.tokens.expires_at && this.tokens.expires_at < Date.now()) {
      if (this.tokens.refresh_token) {
        // Try refreshing
        try {
          await this._refreshToken();
          return {
            authenticated: true,
            ...base,
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
      authenticated: true,
      ...base,
      user_identity: await this._getUserEmail(),
      expires_at: this.tokens.expires_at
        ? new Date(this.tokens.expires_at).toISOString()
        : undefined,
    };
  }

  async startAuth() {
    // Generate PKCE code verifier and challenge
    this._codeVerifier = randomBytes(32).toString('base64url');
    const codeChallenge = createHash('sha256')
      .update(this._codeVerifier)
      .digest('base64url');

    const params = new URLSearchParams({
      client_id: this.connection.client_id,
      response_type: 'code',
      redirect_uri: 'http://localhost:3939/callback',
      scope: 'Files.ReadWrite.All offline_access User.Read',
      code_challenge: codeChallenge,
      code_challenge_method: 'S256',
      prompt: 'consent',
    });

    const tenantId = this.connection.tenant_id || 'common';
    const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${params.toString()}`;

    return {
      status: 'awaiting_code',
      auth_url: authUrl,
      message:
        'Open this URL in your browser, sign in with your Microsoft account, ' +
        'grant access to OneDrive/SharePoint, and paste the authorization code here. ' +
        'The code appears in the URL bar after redirect (the "code" parameter).',
    };
  }

  async completeAuth(authCode) {
    if (!authCode) {
      throw new AuthFailedError('No authorization code provided');
    }

    const tenantId = this.connection.tenant_id || 'common';
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const body = new URLSearchParams({
      client_id: this.connection.client_id,
      grant_type: 'authorization_code',
      code: authCode,
      redirect_uri: 'http://localhost:3939/callback',
      code_verifier: this._codeVerifier || '',
      scope: 'Files.ReadWrite.All offline_access User.Read',
    });

    try {
      const res = await fetch(tokenUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: body.toString(),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new AuthFailedError(
          `Token exchange failed: ${err.error_description || err.error || res.statusText}`
        );
      }

      const data = await res.json();
      this.tokens = {
        access_token: data.access_token,
        refresh_token: data.refresh_token,
        expires_at: Date.now() + (data.expires_in * 1000),
      };

      await this._writeCredential(this.tokens);

      const email = await this._getUserEmail();
      return {
        status: 'authenticated',
        user_identity: email,
        message: `Successfully authenticated to OneDrive as ${email}.`,
      };
    } catch (err) {
      if (err instanceof AuthFailedError) throw err;
      throw new AuthFailedError(`OAuth token exchange failed: ${err.message}`);
    }
  }

  // ─── File Operations ─────────────────────────────────────────────────

  async read(path) {
    await this._ensureAuth();
    const graphPath = this._toGraphPath(path);

    try {
      // Get file content
      const res = await this._graphRequest(`${graphPath}:/content`, {
        rawResponse: true,
      });

      const buffer = Buffer.from(await res.arrayBuffer());

      // Try UTF-8; fall back to base64 for binary
      const text = buffer.toString('utf-8');
      if (text.includes('\0')) {
        return 'base64:' + buffer.toString('base64');
      }
      return text;
    } catch (err) {
      this._handleGraphError(err, path);
    }
  }

  async write(path, content) {
    await this._ensureAuth();
    const graphPath = this._toGraphPath(path);

    // Determine body
    let body;
    let contentType = 'text/plain';
    if (content.startsWith('base64:')) {
      body = Buffer.from(content.slice(7), 'base64');
      contentType = 'application/octet-stream';
    } else {
      body = content;
    }

    try {
      // PUT to :/content creates or overwrites the file
      // Graph API auto-creates parent folders for PUT on path-based endpoints
      const res = await this._graphRequest(`${graphPath}:/content`, {
        method: 'PUT',
        headers: { 'Content-Type': contentType },
        body,
      });

      const data = await res.json();

      // Cache the item
      this.pathCache.set(this._normalizePath(path), {
        id: data.id,
        type: 'file',
        etag: data.eTag,
      });
    } catch (err) {
      this._handleGraphError(err, path);
    }
  }

  async list(path, recursive = false) {
    await this._ensureAuth();
    const graphPath = this._toGraphPath(path);

    try {
      // List children
      const endpoint = path === '/' || path === ''
        ? `${this._driveBase()}/root/children`
        : `${graphPath}:/children`;

      const entries = [];
      let url = endpoint;

      do {
        const res = await this._graphRequest(url);
        const data = await res.json();

        for (const item of data.value || []) {
          const isDir = !!item.folder;
          const entry = {
            name: item.name,
            type: isDir ? 'directory' : 'file',
          };

          if (!isDir) {
            entry.size = item.size || 0;
            entry.modified = item.lastModifiedDateTime;
          }

          // Cache while listing
          const entryPath = path === '/' ? `/${item.name}` : `${this._normalizePath(path)}/${item.name}`;
          this.pathCache.set(entryPath, {
            id: item.id,
            type: isDir ? 'directory' : 'file',
            etag: item.eTag,
          });

          entries.push(entry);

          // Recurse into subdirectories if requested
          if (recursive && isDir) {
            const subEntries = await this.list(entryPath, true);
            for (const sub of subEntries) {
              entries.push({
                ...sub,
                name: `${item.name}/${sub.name}`,
              });
            }
          }
        }

        // Handle pagination
        url = data['@odata.nextLink'] || null;
      } while (url);

      return entries;
    } catch (err) {
      this._handleGraphError(err, path);
    }
  }

  async exists(path) {
    await this._ensureAuth();
    const graphPath = this._toGraphPath(path);

    try {
      const res = await this._graphRequest(graphPath, { allowNotFound: true });

      if (res.status === 404) {
        return { exists: false };
      }

      const data = await res.json();
      const isDir = !!data.folder;

      // Cache
      this.pathCache.set(this._normalizePath(path), {
        id: data.id,
        type: isDir ? 'directory' : 'file',
        etag: data.eTag,
      });

      return { exists: true, type: isDir ? 'directory' : 'file' };
    } catch (err) {
      // 404 is expected for non-existent paths
      if (err.status === 404) {
        return { exists: false };
      }
      this._handleGraphError(err, path);
    }
  }

  async stat(path) {
    await this._ensureAuth();
    const graphPath = this._toGraphPath(path);

    try {
      const res = await this._graphRequest(graphPath);
      const data = await res.json();

      return {
        size: data.size || 0,
        modified: data.lastModifiedDateTime,
        created: data.createdDateTime,
        etag: data.eTag,
      };
    } catch (err) {
      this._handleGraphError(err, path);
    }
  }

  async delete(path) {
    await this._ensureAuth();

    // We need the item ID for delete
    const itemId = await this._resolveItemId(path);
    if (!itemId) {
      throw new FileNotFoundError(path);
    }

    // Check if it's a non-empty directory
    const cached = this.pathCache.get(this._normalizePath(path));
    if (cached && cached.type === 'directory') {
      const children = await this.list(path, false);
      if (children.length > 0) {
        throw new NotEmptyError(path);
      }
    }

    try {
      await this._graphRequest(`${this._driveBase()}/items/${itemId}`, {
        method: 'DELETE',
      });
      this.pathCache.delete(this._normalizePath(path));
    } catch (err) {
      this._handleGraphError(err, path);
    }
  }

  async copy(source, destination) {
    await this._ensureAuth();

    // Need source item ID
    const sourceId = await this._resolveItemId(source);
    if (!sourceId) {
      throw new FileNotFoundError(source);
    }

    // Ensure destination parent exists and get its ID
    const destParentPath = this._parentPath(destination);
    const destFileName = this._fileName(destination);

    // Resolve parent — create it if needed by writing a temp file then deleting
    let parentId;
    const parentCached = this.pathCache.get(this._normalizePath(destParentPath));
    if (parentCached) {
      parentId = parentCached.id;
    } else {
      // Resolve by querying the parent path
      const parentGraphPath = this._toGraphPath(destParentPath);
      try {
        const res = await this._graphRequest(parentGraphPath);
        const data = await res.json();
        parentId = data.id;
        this.pathCache.set(this._normalizePath(destParentPath), {
          id: data.id,
          type: 'directory',
          etag: data.eTag,
        });
      } catch (err) {
        // Parent doesn't exist — create it by writing and reading
        throw new PathNotFoundError(destParentPath);
      }
    }

    try {
      // Graph API copy is async — it returns a monitor URL
      const res = await this._graphRequest(`${this._driveBase()}/items/${sourceId}/copy`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parentReference: { driveId: this._getDriveId(), id: parentId },
          name: destFileName,
        }),
      });

      // Copy is accepted (202) — we don't wait for completion since it's typically fast
      // for small files. For large files, the monitor URL could be polled.
    } catch (err) {
      this._handleGraphError(err, source);
    }
  }

  // ─── Graph API Helpers ──────────────────────────────────────────────

  /**
   * Make an authenticated request to the Microsoft Graph API.
   */
  async _graphRequest(urlOrPath, options = {}) {
    await this._ensureAuth();

    const { method = 'GET', headers = {}, body, rawResponse = false, allowNotFound = false } = options;

    // Build full URL
    let url;
    if (urlOrPath.startsWith('https://')) {
      url = urlOrPath; // Already a full URL (e.g., pagination nextLink)
    } else {
      url = `https://graph.microsoft.com/v1.0${urlOrPath}`;
    }

    const fetchHeaders = {
      Authorization: `Bearer ${this.tokens.access_token}`,
      ...headers,
    };

    const res = await fetch(url, {
      method,
      headers: fetchHeaders,
      body: body !== undefined ? body : undefined,
    });

    if (allowNotFound && res.status === 404) {
      return res;
    }

    if (!res.ok) {
      const err = new Error(`Graph API error: ${res.status} ${res.statusText}`);
      err.status = res.status;
      try {
        err.body = await res.json();
        err.message = err.body?.error?.message || err.message;
      } catch {
        // Ignore JSON parse failure on error response
      }
      throw err;
    }

    if (rawResponse) {
      return res;
    }

    return res;
  }

  /**
   * Build the Graph API drive base path.
   */
  _driveBase() {
    if (this.connection.site_id && this.connection.drive_id) {
      return `/sites/${this.connection.site_id}/drives/${this.connection.drive_id}`;
    }
    if (this.connection.drive_id) {
      return `/drives/${this.connection.drive_id}`;
    }
    // Default to the authenticated user's OneDrive
    return '/me/drive';
  }

  /**
   * Get the drive ID for copy operations.
   */
  _getDriveId() {
    return this.connection.drive_id || null;
  }

  /**
   * Convert a logical AIFS path to a Graph API path.
   * Graph API uses /drive/root:/path/to/file for path-based access.
   */
  _toGraphPath(path) {
    const normalized = this._normalizePath(path);
    if (normalized === '/') {
      return `${this._driveBase()}/root`;
    }
    // Graph API path-based: /drive/root:/path/to/item
    return `${this._driveBase()}/root:${normalized}`;
  }

  /**
   * Resolve a logical path to a OneDrive item ID.
   * Checks cache first, then queries Graph API.
   */
  async _resolveItemId(path) {
    const normalized = this._normalizePath(path);
    const cached = this.pathCache.get(normalized);
    if (cached) {
      return cached.id;
    }

    // Query Graph API for the item
    const graphPath = this._toGraphPath(path);
    try {
      const res = await this._graphRequest(graphPath, { allowNotFound: true });
      if (res.status === 404) {
        return null;
      }
      const data = await res.json();
      const isDir = !!data.folder;
      this.pathCache.set(normalized, {
        id: data.id,
        type: isDir ? 'directory' : 'file',
        etag: data.eTag,
      });
      return data.id;
    } catch (err) {
      if (err.status === 404) return null;
      this._handleGraphError(err, path);
    }
  }

  // ─── Token Management ───────────────────────────────────────────────

  /**
   * Ensure we have a valid access token, refreshing if needed.
   */
  async _ensureAuth() {
    if (!this.tokens || !this.tokens.access_token) {
      throw new NotAuthenticatedError('no_credential');
    }

    // Refresh if expired or about to expire (5 minute buffer)
    if (this.tokens.expires_at && (this.tokens.expires_at - 300000) < Date.now()) {
      if (this.tokens.refresh_token) {
        await this._refreshToken();
      } else {
        throw new NotAuthenticatedError('expired');
      }
    }
  }

  /**
   * Refresh the access token using the stored refresh token.
   */
  async _refreshToken() {
    const tenantId = this.connection.tenant_id || 'common';
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const body = new URLSearchParams({
      client_id: this.connection.client_id,
      grant_type: 'refresh_token',
      refresh_token: this.tokens.refresh_token,
      scope: 'Files.ReadWrite.All offline_access User.Read',
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
      expires_at: Date.now() + (data.expires_in * 1000),
    };

    await this._writeCredential(this.tokens);
  }

  // ─── Helpers ──────────────────────────────────────────────────────────

  _normalizePath(path) {
    let p = '/' + path.replace(/^\/+/, '').replace(/\/+$/, '');
    p = p.replace(/\/+/g, '/');
    if (p === '') p = '/';
    return p;
  }

  _parentPath(path) {
    const normalized = this._normalizePath(path);
    const lastSlash = normalized.lastIndexOf('/');
    if (lastSlash <= 0) return '/';
    return normalized.slice(0, lastSlash);
  }

  _fileName(path) {
    const normalized = this._normalizePath(path);
    const lastSlash = normalized.lastIndexOf('/');
    return normalized.slice(lastSlash + 1);
  }

  async _getUserEmail() {
    try {
      const res = await this._graphRequest('/me', {
        headers: { Accept: 'application/json' },
      });
      const data = await res.json();
      return data.mail || data.userPrincipalName || 'unknown';
    } catch {
      return 'unknown';
    }
  }

  async _writeCredential(tokens) {
    const dir = dirname(this.credentialPath);
    await mkdir(dir, { recursive: true });
    await writeFile(this.credentialPath, JSON.stringify(tokens, null, 2), 'utf-8');
  }

  /**
   * Translate Microsoft Graph API errors to AIFS errors.
   */
  _handleGraphError(err, path) {
    const status = err.status || err.response?.status;

    switch (status) {
      case 401:
        throw new NotAuthenticatedError('expired');
      case 403:
        throw new AccessDeniedError(path);
      case 404:
        throw new FileNotFoundError(path);
      case 409:
        // Conflict — could be write conflict or name collision
        throw new BackendError(`Conflict at path: ${path}. ${err.message}`, err);
      case 412:
        // Precondition failed — ETag mismatch
        throw new BackendError(`Write conflict at path: ${path}. Retry with fresh read.`, err);
      default:
        throw new BackendError(
          `Microsoft Graph API error (${status}): ${err.message}`,
          err
        );
    }
  }
}
