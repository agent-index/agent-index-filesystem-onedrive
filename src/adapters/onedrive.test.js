// Unit tests for OneDriveAdapter pure logic — no network.
// Covers: Graph address construction (absolute + id-anchor), id-anchor-aware
// path helpers, sentinel detection, and the no-secret guard. Live Graph
// behaviour (auth, ops, probes P1–P4) is covered by the stage-4 test plan
// against the real dev tenant.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { NotProvisionedError } from '@agent-index/filesystem/errors';
import { OneDriveAdapter, detectSentinel } from './onedrive.js';

function makeAdapter(connection = { tenant_id: 't', client_id: 'c', site_id: 'S', drive_id: 'D' }) {
  const a = new OneDriveAdapter();
  a.connection = connection;
  return a;
}

test('detectSentinel: recognizes each encoding, ignores binary', () => {
  assert.equal(detectSentinel('body\n<!-- AIFS:FILE-END -->'), 'md');
  assert.equal(detectSentinel('code;\n// AIFS:FILE-END'), 'slash');
  assert.equal(detectSentinel('text\n# AIFS:FILE-END'), 'hash');
  assert.equal(detectSentinel('{"a":1,"_file_end":"AIFS:FILE-END"}'), 'json');
  assert.equal(detectSentinel('no marker here'), null);
  assert.equal(detectSentinel('base64:AAAA'), null);
});

test('_addr: absolute root and nested paths use SharePoint library drive', () => {
  const a = makeAdapter();
  const root = a._addr('/');
  assert.equal(root.meta, '/sites/S/drives/D/root');
  assert.equal(root.child, '/sites/S/drives/D/root/children');

  const nested = a._addr('/shared/projects/x.md');
  assert.equal(nested.meta, '/sites/S/drives/D/root:/shared/projects/x.md');
  assert.equal(nested.content, '/sites/S/drives/D/root:/shared/projects/x.md:/content');
  assert.equal(nested.child, '/sites/S/drives/D/root:/shared/projects/x.md:/children');
});

test('_addr: id-anchors resolve against the member\'s own OneDrive (/me/drive)', () => {
  const a = makeAdapter();
  const bare = a._addr('id:01ABCID');
  assert.equal(bare.meta, '/me/drive/items/01ABCID');
  assert.equal(bare.content, '/me/drive/items/01ABCID/content');

  const rel = a._addr('id:01ABCID/Agent-Index-Private/note.md');
  assert.equal(rel.meta, '/me/drive/items/01ABCID:/Agent-Index-Private/note.md');
  assert.equal(rel.content, '/me/drive/items/01ABCID:/Agent-Index-Private/note.md:/content');
  assert.equal(rel.child, '/me/drive/items/01ABCID:/Agent-Index-Private/note.md:/children');
});

test('_addr: cross-drive anchors route to /drives/{driveId} (crossdriveread, C.1.3)', () => {
  const a = makeAdapter();
  // id:{driveId}:{itemId} — owner's drive, bare item
  const bare = a._addr('id:bDRIVE!ZZ:01ITEMID');
  assert.equal(bare.meta, '/drives/bDRIVE!ZZ/items/01ITEMID');
  assert.equal(bare.content, '/drives/bDRIVE!ZZ/items/01ITEMID/content');
  assert.equal(bare.child, '/drives/bDRIVE!ZZ/items/01ITEMID/children');
  // id:{driveId}:{itemId}/rel — a doc inside a shared folder on the owner's drive
  const rel = a._addr('id:bDRIVE!ZZ:01ITEMID/doc.md');
  assert.equal(rel.meta, '/drives/bDRIVE!ZZ/items/01ITEMID:/doc.md');
  assert.equal(rel.content, '/drives/bDRIVE!ZZ/items/01ITEMID:/doc.md:/content');
  // a bare (own-drive) anchor is unchanged — backward compatible
  assert.equal(a._addr('id:01ITEMID').meta, '/me/drive/items/01ITEMID');
});

test('_parseAnchor + _driveBaseFor: qualified vs bare; cross-drive is not own-drive', () => {
  const a = makeAdapter();
  assert.deepEqual(a._parseAnchor('id:01ITEMID/x/y.md'), { driveId: null, id: '01ITEMID', rel: 'x/y.md' });
  assert.deepEqual(a._parseAnchor('id:bDRV:01ITEMID/x/y.md'), { driveId: 'bDRV', id: '01ITEMID', rel: 'x/y.md' });
  assert.equal(a._driveBaseFor('id:01ITEMID'), '/me/drive');
  assert.equal(a._driveBaseFor('id:bDRV:01ITEMID'), '/drives/bDRV');
  assert.equal(a._isOwnDrive('id:01ITEMID'), true);          // bare = own drive
  assert.equal(a._isOwnDrive('id:bDRV:01ITEMID'), false);    // cross-drive = not own drive
});

test('_driveBase: falls back from site+drive to drive to /me/drive', () => {
  assert.equal(makeAdapter({ site_id: 'S', drive_id: 'D' })._driveBase(), '/sites/S/drives/D');
  assert.equal(makeAdapter({ drive_id: 'D' })._driveBase(), '/drives/D');
  assert.equal(makeAdapter({})._driveBase(), '/me/drive');
});

test('path helpers: absolute', () => {
  const a = makeAdapter();
  assert.equal(a._normalizePath('/a//b/'), '/a/b');
  assert.equal(a._parentPath('/a/b/c.md'), '/a/b');
  assert.equal(a._parentPath('/a'), '/');
  assert.equal(a._fileName('/a/b/c.md'), 'c.md');
});

test('path helpers: id-anchor aware', () => {
  const a = makeAdapter();
  assert.equal(a._normalizePath('id:XID/a//b/'), 'id:XID/a/b');
  assert.equal(a._parentPath('id:XID/a/b/c.md'), 'id:XID/a/b');
  assert.equal(a._parentPath('id:XID/a'), 'id:XID');
  assert.equal(a._parentPath('id:XID'), 'id:XID'); // bare anchor is its own root
  assert.equal(a._fileName('id:XID/a/b/c.md'), 'c.md');
});

test('initialize: rejects a client_secret (public client only)', async () => {
  const a = new OneDriveAdapter();
  await assert.rejects(
    () => a.initialize({ tenant_id: 't', client_id: 'c', client_secret: 'oops' }, '/tmp/creds'),
    /public client/i
  );
});

test('ops require auth (no tokens -> NOT_AUTHENTICATED)', async () => {
  const a = makeAdapter();
  a.tokens = null;
  await assert.rejects(() => a.read('/x'), /Not authenticated/i);
});

test('ACL ops are implemented (Release B) — they require auth, not NOT_IMPLEMENTED', async () => {
  const a = makeAdapter();
  a.tokens = null;
  // share/unshare/getPermissions/search now go through _ensureAuth first, so
  // with no tokens they fail NOT_AUTHENTICATED — proving they're implemented,
  // not stubbed to "not implemented".
  await assert.rejects(() => a.share('/x', 'u@e.com', 'reader'), /Not authenticated/i);
  await assert.rejects(() => a.unshare('/x', 'u@e.com'), /Not authenticated/i);
  await assert.rejects(() => a.getPermissions('/x'), /Not authenticated/i);
  await assert.rejects(() => a.search({ scope: '/' }), /Not authenticated/i);
});

test('transferOwnership is unsupported on onedrive (NOT_IMPLEMENTED)', async () => {
  const a = makeAdapter();
  await assert.rejects(() => a.transferOwnership('/x', 'u@e.com'), /not implemented/i);
});

test('role mapping: AIFS <-> Graph (additive; commenter->read; writer/owner/full-control->writer)', () => {
  const a = makeAdapter();
  assert.deepEqual(a._aifsRoleToGraphRoles('reader'), ['read']);
  assert.deepEqual(a._aifsRoleToGraphRoles('commenter'), ['read']);
  assert.deepEqual(a._aifsRoleToGraphRoles('writer'), ['write']);
  assert.throws(() => a._aifsRoleToGraphRoles('owner'), /not accepted/i);
  assert.equal(a._graphRolesToAifs(['write']), 'writer');
  assert.equal(a._graphRolesToAifs(['owner']), 'writer');
  assert.equal(a._graphRolesToAifs(['sp.full control']), 'writer');
  assert.equal(a._graphRolesToAifs(['read']), 'reader');
});

test('permission subject extraction + case-insensitive match across Graph identity shapes', () => {
  const a = makeAdapter();
  const pUser = { roles: ['read'], grantedToV2: { user: { email: 'A@e.com', id: 'uid1' } } };
  assert.equal(a._permSubject(pUser), 'A@e.com');
  assert.ok(a._permMatchesSubject(pUser, 'a@e.com'));

  const pGroup = { roles: ['write'], grantedToV2: { siteGroup: { loginName: 'Members', id: 'g1' } } };
  assert.equal(a._permSubject(pGroup), 'Members');
  assert.ok(a._permMatchesSubject(pGroup, 'g1'));

  const pMulti = { roles: ['read'], grantedToIdentitiesV2: [{ user: { email: 'X@e.com' } }, { user: { email: 'Y@e.com' } }] };
  assert.match(a._permSubject(pMulti), /X@e\.com/);
  assert.ok(a._permMatchesSubject(pMulti, 'y@e.com'));

  assert.equal(a._permSubject({ roles: ['read'], link: { scope: 'anonymous' } }), 'link:anonymous');
});

test('startAuth persists a PKCE verifier (tmp + workspace) and returns the paste-URL flow', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'od-pkce-'));
  const a = new OneDriveAdapter();
  try {
    await a.initialize({ tenant_id: 'T-START', client_id: 'C-START' }, dir);
    const r = await a.startAuth();
    assert.equal(r.status, 'awaiting_code');
    assert.match(r.auth_url, /login\.microsoftonline\.com\/T-START\/oauth2\/v2\.0\/authorize/);
    assert.match(r.auth_url, /code_challenge_method=S256/);
    // verifier written to BOTH the sandbox-local tmp (primary) and workspace (fallback)
    assert.equal(a._pkcePaths().length, 2);
    const ws = JSON.parse(await readFile(path.join(dir, 'onedrive-pkce.json'), 'utf8'));
    assert.ok(ws.code_verifier && ws.code_verifier.length > 20, 'workspace verifier persisted');
    const viaHelper = await a._readVerifier();
    assert.equal(viaHelper.code_verifier, ws.code_verifier, '_readVerifier round-trips');
  } finally {
    await a._clearVerifier(); await rm(dir, { recursive: true, force: true });
  }
});

test('startAuth REUSES a still-fresh persisted verifier (pkcerestart — re-issued start is harmless)', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'od-pkce-reuse-'));
  const a = new OneDriveAdapter();
  try {
    await a.initialize({ tenant_id: 'T-REUSE', client_id: 'C-REUSE' }, dir);
    const r1 = await a.startAuth();
    const ch1 = new URL(r1.auth_url).searchParams.get('code_challenge');
    const r2 = await a.startAuth(); // accidental second start
    const ch2 = new URL(r2.auth_url).searchParams.get('code_challenge');
    assert.equal(ch1, ch2, 'same challenge -> the first code is still valid');
    assert.match(r2.message, /already in progress|resuming/i);
  } finally {
    await a._clearVerifier(); await rm(dir, { recursive: true, force: true });
  }
});

test('resolveIdentity: requires auth, rejects empty ref', async () => {
  const a = makeAdapter();
  a.tokens = null;
  await assert.rejects(() => a.resolveIdentity('x@e.com'), /Not authenticated/i);
  a.tokens = { access_token: 't', expires_at: Date.now() + 3600_000 };
  await assert.rejects(() => a.resolveIdentity(''), /INVALID_SUBJECT|not a valid identity/i);
});

test('resolveIdentity: a 403 (no directory-read permission) surfaces ACCESS_DENIED, NOT INVALID_SUBJECT (identityperm/errormask)', async () => {
  const a = makeAdapter();
  a.tokens = { access_token: 't', expires_at: Date.now() + 3600_000 };
  // Both the direct GET and the $filter come back 403 Authorization_RequestDenied
  // — the app lacks User.Read.All. The user may well exist; this must NOT be
  // reported as "no matching user".
  a._graph = async () => {
    const e = new Error('Insufficient privileges to complete the operation.');
    e.status = 403;
    e.body = { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges to complete the operation.' } };
    throw e;
  };
  await assert.rejects(
    () => a.resolveIdentity('testproduction@AgentIndex.onmicrosoft.com'),
    (err) => {
      assert.equal(err.code, 'ACCESS_DENIED', 'must be ACCESS_DENIED, not INVALID_SUBJECT');
      assert.match(err.message, /User\.Read\.All/);
      assert.match(err.message, /admin consent|re-authenticate/i);
      assert.doesNotMatch(err.message, /no matching user/i);
      return true;
    }
  );
});

test('resolveIdentity: a genuine empty result (200, no match) is INVALID_SUBJECT', async () => {
  const a = makeAdapter();
  a.tokens = { access_token: 't', expires_at: Date.now() + 3600_000 };
  // Direct GET returns a body with no id; $filter returns an empty value[] — a
  // real "this user does not exist" outcome, distinct from the 403 above.
  a._graph = async (url) => ({
    status: 200,
    json: async () => (/\$filter=/.test(url) ? { value: [] } : {}),
  });
  await assert.rejects(
    () => a.resolveIdentity('ghost@AgentIndex.onmicrosoft.com'),
    (err) => { assert.equal(err.code, 'INVALID_SUBJECT'); return true; }
  );
});

test('startAuth requests the User.Read.All directory-read scope (identityperm)', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'od-scope-'));
  const a = new OneDriveAdapter();
  try {
    await a.initialize({ tenant_id: 'T-SCOPE', client_id: 'C-SCOPE' }, dir);
    const r = await a.startAuth();
    const scope = new URL(r.auth_url).searchParams.get('scope') || '';
    assert.match(scope, /\bUser\.Read\.All\b/, 'resolveIdentity needs directory read');
    assert.match(scope, /\bFiles\.ReadWrite\.All\b/);
    assert.match(scope, /\bSites\.ReadWrite\.All\b/);
    assert.match(scope, /\boffline_access\b/);
  } finally {
    await a._clearVerifier?.();
    await rm(dir, { recursive: true, force: true });
  }
});

test('_handlePermissionError: generic sharingFailed is classified as unresolvable subject (not transient)', () => {
  const a = makeAdapter();
  const err = { status: 400, body: { error: { message: 'sharingFailed: There was a problem sharing, please try again later.' } } };
  assert.throws(() => a._handlePermissionError(err, '/x', 'who@e.com', 'reader'), /INVALID_SUBJECT|not a valid identity/i);
});

test('NotProvisionedError carries NOT_PROVISIONED + needs_provision flag', () => {
  const e = new NotProvisionedError('OneDrive not set up');
  assert.equal(e.code, 'NOT_PROVISIONED');
  assert.equal(e.details.needs_provision, true);
  const r = e.toResponse();
  assert.equal(r.error, 'NOT_PROVISIONED');
  assert.equal(r.needs_provision, true);
});

test('_isOwnDrive: id-anchors and no-site/drive configs are own-drive; SharePoint paths are not', () => {
  const sp = makeAdapter({ site_id: 'S', drive_id: 'D' });
  assert.equal(sp._isOwnDrive('id:XID/foo'), true);     // member space, always /me/drive
  assert.equal(sp._isOwnDrive('/shared/x.md'), false);  // configured SharePoint org root
  const od = makeAdapter({});                            // no site/drive -> default /me/drive
  assert.equal(od._isOwnDrive('/anything'), true);
});

test('resolveSite rejects a non-URL / pathless URL before any network call', async () => {
  const a = makeAdapter();
  a.tokens = { access_token: 't', expires_at: Date.now() + 3600_000 }; // pass _ensureAuth
  await assert.rejects(() => a.resolveSite('not-a-url'), /invalid site URL/i);
  await assert.rejects(() => a.resolveSite('https://contoso.sharepoint.com/'), /\/sites\/<name>/);
});

test('completeAuth without a persisted verifier fails clearly', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'od-noverifier-'));
  try {
    const a = new OneDriveAdapter();
    await a.initialize({ tenant_id: 'TENANT', client_id: 'CID' }, dir);
    await assert.rejects(() => a.completeAuth('some-code'), /PKCE verifier/i);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

// ─── 4.0.1 fix pass: findings 72 (external share), 73 (inherited_from), 63 (bare-id fallback) ──

const AUTHED = { access_token: 't', expires_at: Date.now() + 3600_000 };

/** Mock _graph: routes(url, opts) returns { status, body } or throws an err-shaped object. Records calls. */
function mockGraph(a, routes) {
  const calls = [];
  a._graph = async (url, opts = {}) => {
    calls.push({ url, opts });
    const r = await routes(url, opts);
    if (r?.throw) {
      const e = new Error(r.throw.message || 'Graph API error');
      e.status = r.throw.status;
      e.body = { error: { code: r.throw.code, message: r.throw.message } };
      throw e;
    }
    if (r.status === 404 && !opts.allowNotFound) {
      const e = new Error('Not found'); e.status = 404;
      e.body = { error: { code: 'itemNotFound', message: 'The resource could not be found.' } };
      throw e;
    }
    return { status: r.status ?? 200, json: async () => r.body ?? {} };
  };
  return calls;
}

const ORG = { value: [{ verifiedDomains: [{ name: 'AgentIndex.onmicrosoft.com' }, { name: 'agent-index.ai' }] }] };

function shareRoutes(inviteResult) {
  return (url) => {
    if (url.startsWith('/organization')) return { body: ORG };
    if (url.includes('/invite')) return inviteResult;
    if (url.includes('root:/shared/doc.md')) return { body: { id: 'ITEM1' } };
    return { status: 404 };
  };
}

test('share (finding 72): tenant user -> invite with sendInvitation:false', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  const calls = mockGraph(a, shareRoutes({ body: { value: [{ id: 'P1' }] } }));
  const r = await a.share('/shared/doc.md', 'three@agent-index.ai', 'reader');
  const invite = calls.find((c) => c.url.includes('/invite'));
  const body = JSON.parse(invite.opts.body);
  assert.equal(invite.url, '/sites/S/drives/D/items/ITEM1/invite');
  assert.equal(body.sendInvitation, false);
  assert.equal(body.requireSignIn, true);
  assert.deepEqual(body.recipients, [{ email: 'three@agent-index.ai' }]);
  assert.equal(r.external, false);
  assert.equal(r.permission_id, 'P1');
});

test('share (finding 72): external email -> invite with sendInvitation:true', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  const calls = mockGraph(a, shareRoutes({ body: { value: [{ id: 'P2' }] } }));
  const r = await a.share('/shared/doc.md', 'bilsalak@gmail.com', 'reader');
  const body = JSON.parse(calls.find((c) => c.url.includes('/invite')).opts.body);
  assert.equal(body.sendInvitation, true);
  assert.equal(body.requireSignIn, true);
  assert.equal(r.external, true);
  assert.equal(r.invitation_sent, true);
});

test('share (finding 72): external email, domains unreadable, resolver finds no user -> external', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  const calls = mockGraph(a, (url) => {
    if (url.startsWith('/organization')) return { throw: { status: 403, code: 'Authorization_RequestDenied', message: 'denied' } };
    if (url.startsWith('/users/')) return { status: 404 };
    if (url.startsWith('/users?')) return { body: { value: [] } };
    return shareRoutes({ body: { value: [{ id: 'P3' }] } })(url);
  });
  await a.share('/shared/doc.md', 'someone@outside.example', 'reader');
  const body = JSON.parse(calls.find((c) => c.url.includes('/invite')).opts.body);
  assert.equal(body.sendInvitation, true);
});

test('share (finding 72): external blocked by policy -> EXTERNAL_SHARE_BLOCKED with Graph reason', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  const graphMsg = 'Sharing with external users is not allowed by your organization.';
  mockGraph(a, shareRoutes({ throw: { status: 403, code: 'accessDenied', message: graphMsg } }));
  await assert.rejects(
    () => a.share('/shared/doc.md', 'bilsalak@gmail.com', 'reader'),
    (err) => {
      assert.equal(err.code, 'EXTERNAL_SHARE_BLOCKED');
      assert.match(err.message, /sharing settings don't allow sharing with people outside the organization/);
      assert.doesNotMatch(err.message, /not a valid identity/i);
      const resp = err.toResponse();
      assert.equal(resp.error, 'EXTERNAL_SHARE_BLOCKED');
      assert.equal(resp.reason, graphMsg);
      assert.equal(resp.graph_status, 403);
      assert.equal(resp.graph_code, 'accessDenied');
      return true;
    }
  );
});

test('share (finding 72): external + generic sharingFailed is NOT reported as an invalid identity', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  mockGraph(a, shareRoutes({ throw: { status: 400, code: 'invalidRequest', message: 'sharingFailed: There was a problem sharing.' } }));
  await assert.rejects(() => a.share('/shared/doc.md', 'bilsalak@gmail.com', 'reader'),
    (err) => { assert.equal(err.code, 'EXTERNAL_SHARE_BLOCKED'); return true; });
});

test('share (finding 72): tenant user that does not resolve -> INVALID_SUBJECT with Graph detail', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  mockGraph(a, shareRoutes({ throw: { status: 400, code: 'invalidRequest', message: 'One or more users could not be found.' } }));
  await assert.rejects(() => a.share('/shared/doc.md', 'ghost@agent-index.ai', 'reader'),
    (err) => {
      assert.equal(err.code, 'INVALID_SUBJECT');
      assert.match(err.details.reason, /Graph 400 invalidRequest: One or more users could not be found/);
      return true;
    });
});

test('share (finding 72): malformed email -> INVALID_SUBJECT before any Graph call', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  const calls = mockGraph(a, () => ({ status: 404 }));
  await assert.rejects(() => a.share('/shared/doc.md', 'not an@email', 'reader'),
    (err) => { assert.equal(err.code, 'INVALID_SUBJECT'); assert.match(err.message, /malformed/); return true; });
  assert.equal(calls.length, 0);
});

test('_handlePermissionError (finding 72): other errors keep BACKEND_ERROR but carry Graph status + code + message', () => {
  const a = makeAdapter();
  const err = { status: 500, body: { error: { code: 'generalException', message: 'Boom.' } } };
  assert.throws(() => a._handlePermissionError(err, '/x', 'u@agent-index.ai', 'reader'),
    (e) => { assert.equal(e.code, 'BACKEND_ERROR'); assert.match(e.message, /500 generalException: Boom\./); return true; });
});

test('getPermissions (finding 73): inherited_from populated from inheritedFrom; null when Graph omits it', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  a.pathCache.set('/shared', { id: 'PARENT', type: 'directory' });
  mockGraph(a, (url) => {
    if (url.includes('root:/shared/new')) return { body: { id: 'NEW', folder: {} } };
    if (url.endsWith('/items/NEW/permissions')) return { body: { value: [
      { id: 'p1', roles: ['write'], grantedToV2: { siteGroup: { loginName: 'Members' } }, inheritedFrom: { id: 'PARENT', driveId: 'D' } },
      { id: 'p2', roles: ['read'], grantedToV2: { user: { email: 'x@e.com' } }, inheritedFrom: { driveId: 'D', path: '/drives/D/root:/shared/sub' } },
      { id: 'p3', roles: ['read'], grantedToV2: { user: { email: 'y@e.com' } }, inheritedFrom: { driveId: 'OTHER', id: 'ZZ' } },
      { id: 'p4', roles: ['read'], grantedToV2: { user: { email: 'z@e.com' } } },
    ] } };
    return { status: 404 };
  });
  const { permissions } = await a.getPermissions('/shared/new');
  const by = Object.fromEntries(permissions.map((p) => [p.permission_id, p]));
  assert.equal(by.p1.inherited_from, '/shared');
  assert.equal(by.p2.inherited_from, '/shared/sub');
  assert.equal(by.p3.inherited_from, 'id:OTHER:ZZ');
  assert.equal(by.p4.inherited_from, null);
  assert.equal(by.p1.role, 'writer');
  assert.equal(by.p1.subject, 'Members');
  const direct = await a.getPermissions('/shared/new', { includeInherited: false });
  assert.deepEqual(direct.permissions.map((p) => p.permission_id), ['p4']);
});

test('bare id anchor (finding 63): found on /me/drive -> no retry, stays on /me/drive', async () => {
  const a = makeAdapter(); a.tokens = AUTHED; a._ownDriveOk = true;
  const calls = mockGraph(a, (url) => {
    if (url === '/me/drive/items/OWN?$select=id') return { body: { id: 'OWN' } };
    if (url.startsWith('/me/drive/items/OWN')) return { body: { id: 'OWN', parentReference: { driveId: 'ME' }, size: 3 } };
    return { status: 404 };
  });
  const s = await a.stat('id:OWN');
  assert.equal(s.drive_id, 'ME');
  assert.ok(calls.every((c) => c.url.startsWith('/me/drive')), 'never touched the site drive');
  assert.equal(a._driveBaseFor('id:OWN'), '/me/drive');
});

test('bare id anchor (finding 63): 404 on /me/drive -> retried once on configured site drive', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  const calls = mockGraph(a, (url) => {
    if (url.startsWith('/me/drive')) return { status: 404 };
    if (url === '/sites/S/drives/D/items/ORGCFG?$select=id') return { body: { id: 'ORGCFG' } };
    if (url.startsWith('/sites/S/drives/D/items/ORGCFG?')) return { body: { id: 'ORGCFG', parentReference: { driveId: 'D' }, size: 10 } };
    if (url === '/sites/S/drives/D/items/ORGCFG/content') return { body: {} };
    return { status: 404 };
  });
  const s = await a.stat('id:ORGCFG');
  assert.equal(s.id, 'ORGCFG');
  assert.equal(s.drive_id, 'D', 'drive_id is the drive it was actually found on');
  assert.equal(a._driveBaseFor('id:ORGCFG/sub/x.md'), '/sites/S/drives/D');
  assert.equal(a._isOwnDrive('id:ORGCFG'), false);
  assert.equal(calls.filter((c) => c.url.startsWith('/me/drive/items/ORGCFG')).length, 1, 'one own-drive probe');
  // later ops reuse the pin (no further probes) and address the site drive
  const perms = mockGraph(a, (url) => (url === '/sites/S/drives/D/items/ORGCFG/permissions' ? { body: { value: [] } } : { status: 404 }));
  a.pathCache.set('id:ORGCFG', { id: 'ORGCFG', type: 'file' });
  await a.getPermissions('id:ORGCFG');
  assert.equal(perms.length, 1);
});

test('bare id anchor (finding 63): no configured drive -> no fallback probe', async () => {
  const a = makeAdapter({ tenant_id: 't', client_id: 'c' }); a.tokens = AUTHED;
  const calls = mockGraph(a, () => ({ status: 404 }));
  await a._resolveAnchorDrive('id:X');
  assert.equal(calls.length, 0);
  assert.equal(a._driveBaseFor('id:X'), '/me/drive');
});

test('drive-qualified anchor (finding 63): unchanged, no probes', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  const calls = mockGraph(a, () => ({ status: 404 }));
  await a._resolveAnchorDrive('id:bDRV:01ITEM/doc.md');
  assert.equal(calls.length, 0);
  assert.equal(a._addr('id:bDRV:01ITEM/doc.md').meta, '/drives/bDRV/items/01ITEM:/doc.md');
  // absolute paths never fall back to /me/drive
  assert.equal(a._driveBaseFor('/shared/x.md'), '/sites/S/drives/D');
});

test('bare id anchor (finding 63, review): id:root never falls back to the site drive (unprovisioned own drive)', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  const calls = mockGraph(a, () => ({ status: 404 }));
  await a._resolveAnchorDrive('id:root/Agent-Index-Private-x');
  assert.equal(calls.length, 0, 'no probes for drive-relative aliases');
  assert.equal(a._driveBaseFor('id:root/Agent-Index-Private-x'), '/me/drive');
  assert.equal(a._isOwnDrive('id:root/Agent-Index-Private-x'), true, 'own-drive guard still applies');
  await assert.rejects(() => a.exists('id:root/Agent-Index-Private-x'),
    (err) => { assert.equal(err.code, 'NOT_PROVISIONED'); return true; });
});

test('share (finding 72, review): external + plain 403 accessDenied without policy wording -> ACCESS_DENIED, not EXTERNAL_SHARE_BLOCKED', async () => {
  const a = makeAdapter(); a.tokens = AUTHED;
  mockGraph(a, shareRoutes({ throw: { status: 403, code: 'accessDenied', message: 'Access denied.' } }));
  await assert.rejects(() => a.share('/shared/doc.md', 'bilsalak@gmail.com', 'reader'),
    (err) => { assert.equal(err.code, 'ACCESS_DENIED'); return true; });
});
