import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { request } from 'node:http';
import { serveDist } from '../harness/page-probe.mjs';
import { loopbackUrl, localHttpUrl, localRoute } from '../harness/loopback-url.mjs';

test('discovery cannot select a remote host or inject URL credentials', () => {
  assert.equal(loopbackUrl(1234, '/json'), 'http://127.0.0.1:1234/json');
  assert.equal(localRoute('http://127.0.0.1:1234/token/', 'state?x=1'), 'http://127.0.0.1:1234/token/state?x=1');
  for (const port of ['80@evil.example', -1, 0, 65536, 1.5, NaN]) assert.throws(() => loopbackUrl(port));
  for (const url of ['https://127.0.0.1/', 'http://evil.example/', 'http://user@127.0.0.1/']) assert.throws(() => localHttpUrl(url));
  assert.throws(() => localRoute('http://127.0.0.1:1234/', '//evil.example/'));
});

test('probe server serves assets but refuses traversal, including encoded separators', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'xlide-probe-security-'));
  const root = join(scratch, 'public');
  await mkdir(root);
  await writeFile(join(root, 'index.html'), 'public');
  await writeFile(join(scratch, 'secret.txt'), 'private');
  const { server, port } = await serveDist(undefined, root);
  const get = path => new Promise((resolve, reject) => {
    const req = request({ hostname: '127.0.0.1', port, path }, response => {
      let body = '';
      response.on('data', chunk => { body += chunk; });
      response.on('end', () => resolve({ status: response.statusCode, body }));
    });
    req.on('error', reject);
    req.end();
  });
  try {
    assert.deepEqual(await get('/'), { status: 200, body: 'public' });
    for (const path of ['/../secret.txt', '/%2e%2e/secret.txt', '/..%5csecret.txt', '/%zz']) {
      const reply = await get(path);
      assert.notEqual(reply.status, 200, path);
      assert.notEqual(reply.body, 'private', path);
    }
  } finally {
    await new Promise(resolve => server.close(resolve));
    await rm(scratch, { recursive: true, force: true });
  }
});
