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
