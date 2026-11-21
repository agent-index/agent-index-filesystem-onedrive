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

test('search: nested hits without parentReference.path get their full path by walking parent ids (B15)', async () => {
  const a = makeAdapter();
  a._resolveItemId = async () => 'SCOPE';
  const calls = mockGraph(a, (url) => {
    if (url.includes('/search(')) {
      return { body: { value: [
        { id: 'F1', name: 'update-log.json', file: {}, parentReference: { id: 'UPD', driveId: 'D' } },
        { id: 'F2', name: 'manifest.json', file: {}, parentReference: { id: 'DIST', driveId: 'D' } },
        { id: 'F3', name: 'README.md', file: {}, parentReference: { id: 'A' } },
        { id: 'F4', name: 'README.md', file: {}, parentReference: { id: 'B' } },
        { id: 'F5', name: 'top.json', file: {}, parentReference: { id: 'SCOPE' } },
        { id: 'F6', name: 'stray.json', file: {}, parentReference: { id: 'ELSEWHERE' } },
      ] } };
    }
    const item = (id, name, parent) => ({ body: { id, name, parentReference: { id: parent } } });
    if (url.includes('/items/UPD?')) return item('UPD', 'updates', 'SCOPE');
    if (url.includes('/items/DIST?')) return item('DIST', 'dist', 'SCOPE');
    if (url.includes('/items/A?')) return item('A', 'a', 'P');
    if (url.includes('/items/B?')) return item('B', 'b', 'P');
    if (url.includes('/items/P?')) return item('P', 'projects', 'SCOPE');
    if (url.includes('/items/ELSEWHERE?')) return { body: { id: 'ELSEWHERE', name: 'x', root: {} } };
    return { status: 200, body: {} };
  });
  const out = await a.search({ scope: '/shared', name_contains: 'json' + '' });
  const byId = Object.fromEntries(out.results.map((r) => [r.id, r.path]));
  assert.equal(byId.F1, '/shared/updates/update-log.json');
  assert.equal(byId.F2, '/shared/dist/manifest.json');
  assert.equal(byId.F5, '/shared/top.json');
  assert.equal(byId.F6, undefined, 'a hit that cannot be placed under the scope is left out');
  const md = await a.search({ scope: '/shared', name_contains: 'README' });
  assert.deepEqual(md.results.map((r) => r.path).sort(), ['/shared/projects/a/README.md', '/shared/projects/b/README.md']);
  // each folder looked up at most once per search call
  const pLookups = calls.filter((u) => u.includes('/items/P?')).length;
  assert.ok(pLookups <= 2, `P looked up ${pLookups} times across two searches`);
});
