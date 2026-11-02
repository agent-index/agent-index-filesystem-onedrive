// CI-025 (2.8.0): org seeding, read size check, nested search paths — unit tests.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OneDriveAdapter } from './onedrive.js';

const CONN = { tenant_id: 't', client_id: 'c', site_id: 'S', drive_id: 'D' };
function makeAdapter() {
  const a = new OneDriveAdapter();
  a.connection = CONN;
  a.tokens = { access_token: 't', expires_at: Date.now() + 3600_000 };
  return a;
}
function mockGraph(a, routes) {
  const calls = [];
  a._graph = async (url, opts = {}) => {
    calls.push(url);
    const r = await routes(url, opts);
    if (r.status === 404 && !opts.allowNotFound) {
      const e = new Error('Not found'); e.status = 404;
      e.body = { error: { code: 'itemNotFound' } }; throw e;
    }
    return {
      status: r.status ?? 200,
      json: async () => r.body ?? {},
      arrayBuffer: async () => new TextEncoder().encode(r.text ?? '').buffer,
    };
  };
  return calls;
}

test('seedOrgPrefixes: an org id resolves on the site drive with no own-drive probe', async () => {
  const a = makeAdapter();
  a.seedOrgPrefixes({ '/org-config.json': { id: 'OCID', kind: 'file' }, '/shared': { id: 'SRID', kind: 'folder' } });
  const calls = mockGraph(a, () => ({ status: 200 }));
  await a._resolveAnchorDrive('id:OCID');
  assert.equal(calls.length, 0, 'no probe for a seeded org id');
  assert.ok(a._addr('id:OCID').meta.startsWith('/sites/S/drives/D/items/OCID'));
});

test('seedOrgPrefixes: an unknown bare id still gets the finding-63 probe', async () => {
  const a = makeAdapter();
  a.seedOrgPrefixes({ '/shared': { id: 'SRID', kind: 'folder' } });
  const calls = mockGraph(a, (url) => (url.startsWith('/me/drive/items/OTHER') ? { status: 404 } : { status: 200 }));
  await a._resolveAnchorDrive('id:OTHER');
  assert.ok(calls.some((u) => u.startsWith('/me/drive/items/OTHER')));
});

test('read: a short body is re-read, then succeeds when sizes match', async () => {
  const a = makeAdapter();
  let n = 0;
  mockGraph(a, (url) => {
    if (url.endsWith(':/content')) { n++; return { text: n === 1 ? 'hello wor' : 'hello world' }; }
    if (url.includes('$select=size')) return { body: { size: 11 } };
    return { status: 200, body: {} };
  });
  assert.equal(await a.read('/shared/x.md'), 'hello world');
  assert.equal(n, 2);
});

test('read: a persistent size mismatch fails AIFS_READ_UNRELIABLE (size_mismatch)', async () => {
  const a = makeAdapter();
  mockGraph(a, (url) => {
    if (url.endsWith(':/content')) return { text: 'short' };
    if (url.includes('$select=size')) return { body: { size: 99 } };
    return { status: 200, body: {} };
  });
  await assert.rejects(() => a.read('/shared/x.md'), (e) => e.code === 'AIFS_READ_UNRELIABLE' && e.details?.reason === 'size_mismatch');
});

test('read: a failed metadata call never fails a read that succeeded', async () => {
  const a = makeAdapter();
  mockGraph(a, (url) => {
    if (url.endsWith(':/content')) return { text: 'body' };
    if (url.includes('$select=size')) throw Object.assign(new Error('boom'), { status: 500 });
    return { status: 200, body: {} };
  });
  assert.equal(await a.read('/shared/x.md'), 'body');
});

test('search: a nested hit from a recursive name search carries its full path', async () => {
  const a = makeAdapter();
  a._resolveItemId = async () => 'SCOPE';
  mockGraph(a, (url) => {
    if (url.includes('/search(')) {
      return { body: { value: [
        { id: 'X', name: 'deep.md', file: {}, parentReference: { path: '/drives/D/root:/shared/projects/a/b' } },
        { id: 'Y', name: 'top.md', file: {}, parentReference: { path: '/drives/D/root:/shared/projects' } },
      ] } };
    }
    return { status: 200, body: {} };
  });
  const out = await a.search({ scope: '/shared/projects', name_contains: 'md' });
  const paths = out.results.map((r) => r.path).sort();
  assert.deepEqual(paths, ['/shared/projects/a/b/deep.md', '/shared/projects/top.md']);
});
