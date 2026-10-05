import { afterEach, describe, expect, test } from 'bun:test';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { FileTokenStore } from '../../src/oauth/token-store-file';
import { cleanup, memoryLogger } from '../helpers';
import { codeFor, loopback, proxyFixture, redeem, registerNative, type TokenPair } from './fixture';

/**
 * The proxy's options for servers whose deployed behavior differs from the defaults. Each
 * default keeps the behavior the module had before the option existed.
 */
const directories: string[] = [];
afterEach(async () => {
  await cleanup();
  for (const directory of directories.splice(0)) {
    rmSync(directory, { recursive: true, force: true });
  }
});

async function signIn(f: ReturnType<typeof proxyFixture>) {
  const uri = loopback();
  const client = await registerNative(f, uri);
  const { code } = await codeFor(f, client.client_id, uri);
  const pair = (await (await redeem(f, client.client_id, uri, code)).json()) as TokenPair;
  return { clientId: client.client_id, pair };
}

const refresh = (f: ReturnType<typeof proxyFixture>, token: string, clientId?: string) =>
  f.request(
    '/token',
    new URLSearchParams({
      grant_type: 'refresh_token',
      refresh_token: token,
      ...(clientId && { client_id: clientId }),
    }),
  );

describe('bindGrants', () => {
  test('by default, tokens are bound to their client, and refresh needs that client', async () => {
    const f = proxyFixture();
    const { clientId, pair } = await signIn(f);
    expect((await f.tokens.getByRsAccess(pair.access_token))?.oauth).toEqual({
      clientId,
      resource: 'http://127.0.0.1:3000/mcp',
    });
    expect((await refresh(f, pair.refresh_token)).status).toBe(400);
    expect((await refresh(f, pair.refresh_token, 'someone-else')).status).toBe(400);
    expect((await refresh(f, pair.refresh_token, clientId)).status).toBe(200);
  });

  test('when off, tokens carry no binding, and refresh needs no client_id', async () => {
    const f = proxyFixture({ bindGrants: false });
    const { pair } = await signIn(f);
    const record = await f.tokens.getByRsAccess(pair.access_token);
    expect(record).not.toHaveProperty('oauth');
    expect((await refresh(f, pair.refresh_token)).status).toBe(200);
    // The caller is identified by a hash of the token instead.
    expect((await f.proxy.verifier.verifyAccessToken(pair.access_token)).clientId).toMatch(
      /^rs:[0-9a-f]{64}$/,
    );
  });
});

describe('discoveryAliases', () => {
  const ALIASES = [
    '/.well-known/oauth-protected-resource',
    '/mcp/.well-known/oauth-protected-resource',
    '/mcp/.well-known/oauth-authorization-server',
  ];

  test('by default, the older discovery locations are served', async () => {
    const f = proxyFixture();
    for (const path of ALIASES) expect((await f.request(path)).status).toBe(200);
  });

  test('when off, only the standard locations answer', async () => {
    const f = proxyFixture({ discoveryAliases: false });
    for (const path of ALIASES) expect((await f.request(path)).status).toBe(404);
    expect((await f.request('/.well-known/oauth-protected-resource/mcp')).status).toBe(200);
    expect((await f.request('/.well-known/oauth-authorization-server')).status).toBe(200);
    // The endpoints themselves are still mounted.
    expect((await f.request('/register', { redirect_uris: [] })).status).toBe(400);
  });
});

describe('FileTokenStore after a restart', () => {
  test('a record whose Google token expired is kept, and the verifier refreshes it', async () => {
    const directory = mkdtempSync(join(tmpdir(), 'calendar-options-'));
    directories.push(directory);
    const path = join(directory, 'tokens.json');
    const writer = new FileTokenStore(path, undefined, memoryLogger());
    await writer.storeRsMapping(
      'access',
      { access_token: 'old', refresh_token: 'google-refresh', expires_at: Date.now() - 1000 },
      'refresh',
    );
    writer.flush();

    const store = new FileTokenStore(path, undefined, memoryLogger());
    const f = proxyFixture({ tokens: store });
    expect((await f.proxy.verifier.verifyAccessToken('access')).extra).toEqual({
      providerAccessToken: 'google-fresh',
    });
  });
});
