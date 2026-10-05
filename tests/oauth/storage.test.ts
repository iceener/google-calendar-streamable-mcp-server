import { afterEach, describe, expect, test } from 'bun:test';
import { createCipheriv, createDecipheriv, randomBytes } from 'node:crypto';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { FileOAuthAuthority } from '../../src/oauth/authority-file';
import { createEncryptor, pkceChallenge } from '../../src/oauth/crypto';
import type { RsRecord } from '../../src/oauth/token-store';
import { FileTokenStore } from '../../src/oauth/token-store-file';
import { type KvLike, KvTokenStore } from '../../src/oauth/token-store-kv';
import before from '../fixtures/storage-before.json';
import { cleanup, memoryLogger } from '../helpers';
import { googleTokens, proxyFixture, redeem, VERIFIER } from './fixture';

/**
 * Stored tokens must survive the migration: records the checkpoint code (180f48b) wrote, in
 * `fixtures/storage-before.json`, are read and refreshed here, and new writes are checked
 * against the old layouts with an independent implementation of each codec. The proxy runs
 * as `src/server.ts` configures it: `bindGrants: false`, so records keep the checkpoint's
 * shape, without a client binding.
 */
const KEY = before.key;
const PRODUCTION = 'https://google-calendar.example.workers.dev';
const RESOURCE = `${PRODUCTION}/mcp`;
/** The fields of a record, in the order the checkpoint wrote them. */
const RECORD_FIELDS = ['rs_access_token', 'rs_refresh_token', 'provider', 'created_at'];
const directories: string[] = [];

afterEach(async () => {
  await cleanup();
  for (const directory of directories.splice(0))
    rmSync(directory, { recursive: true, force: true });
});

function temporaryDirectory(): string {
  const directory = mkdtempSync(join(tmpdir(), 'calendar-storage-'));
  directories.push(directory);
  return directory;
}

class FakeKv implements KvLike {
  readonly values = new Map<string, string>();
  constructor(initial: Record<string, string> = {}) {
    for (const [key, value] of Object.entries(initial)) this.values.set(key, value);
  }
  async get(key: string) {
    return this.values.get(key) ?? null;
  }
  async put(key: string, value: string) {
    this.values.set(key, value);
  }
  async delete(key: string) {
    this.values.delete(key);
  }
}

/** The KV and grant layout, independently: base64url(iv[12] || ciphertext || tag[16]). */
function openKvValue(value: string, key = KEY): Record<string, unknown> {
  const bytes = Buffer.from(value, 'base64url');
  const decipher = createDecipheriv(
    'aes-256-gcm',
    Buffer.from(key, 'base64url'),
    bytes.subarray(0, 12),
  );
  decipher.setAuthTag(bytes.subarray(bytes.length - 16));
  return JSON.parse(
    Buffer.concat([
      decipher.update(bytes.subarray(12, bytes.length - 16)),
      decipher.final(),
    ]).toString(),
  );
}

/** The token file layout, independently: base64url(iv[12] || tag[16] || ciphertext). */
function openFile(value: string, key = KEY): unknown {
  const bytes = Buffer.from(value, 'base64url');
  const decipher = createDecipheriv(
    'aes-256-gcm',
    Buffer.from(key, 'base64url'),
    bytes.subarray(0, 12),
  );
  decipher.setAuthTag(bytes.subarray(12, 28));
  return JSON.parse(
    Buffer.concat([decipher.update(bytes.subarray(28)), decipher.final()]).toString(),
  );
}

function kvStore(kv: KvLike, key: string | null = KEY): KvTokenStore {
  return new KvTokenStore(kv, {
    encryptor: key ? createEncryptor(key) : undefined,
    logger: memoryLogger(),
  });
}

/** The proxy as production configures it, on the production origin. */
function productionProxy(options: Parameters<typeof proxyFixture>[0] = {}) {
  return proxyFixture({
    bindGrants: false,
    discoveryAliases: false,
    encryptionKey: KEY,
    ...options,
    env: {
      MCP_PUBLIC_URL: RESOURCE,
      OAUTH_ISSUER_URL: PRODUCTION,
      OAUTH_AUTHORIZATION_URL: `${PRODUCTION}/authorize`,
      OAUTH_TOKEN_URL: `${PRODUCTION}/token`,
      OAUTH_REGISTRATION_URL: `${PRODUCTION}/register`,
      ...options.env,
    },
  });
}

describe('records written before the migration', () => {
  test('KV records are read as they were written, without a client binding', async () => {
    const store = kvStore(new FakeKv(before.kv));
    for (const name of ['current', 'expired', 'updated']) {
      const byAccess = await store.getByRsAccess(`${name}-rs-access`);
      expect(byAccess).toMatchObject({
        rs_access_token: `${name}-rs-access`,
        rs_refresh_token: `${name}-rs-refresh`,
        provider: {
          access_token: `${name}-calendar-access`,
          refresh_token: `${name}-calendar-refresh`,
        },
      });
      expect(byAccess?.oauth).toBeUndefined();
      expect(await store.getByRsRefresh(`${name}-rs-refresh`)).toEqual(byAccess);
    }
  });

  test('a record verifies for this server, and tools get its Google token', async () => {
    const f = productionProxy({ tokens: kvStore(new FakeKv(before.kv)) });
    const caller = await f.proxy.verifier.verifyAccessToken('current-rs-access');
    expect(caller).toMatchObject({
      resource: new URL(RESOURCE),
      scopes: [
        'https://www.googleapis.com/auth/calendar.events',
        'https://www.googleapis.com/auth/calendar.readonly',
      ],
      extra: { providerAccessToken: 'current-calendar-access' },
    });
    // No provenance: the caller is a hash of the token.
    expect(caller.clientId).toMatch(/^rs:[0-9a-f]{64}$/);
    expect(f.google.forms).toHaveLength(0);
  });

  test('a record with an expired Google token is refreshed on use, in the same format', async () => {
    const kv = new FakeKv(before.kv);
    const google = googleTokens();
    const f = productionProxy({ tokens: kvStore(kv), google });

    const caller = await f.proxy.verifier.verifyAccessToken('expired-rs-access');
    expect(caller.extra).toEqual({ providerAccessToken: 'google-fresh' });
    expect(Object.fromEntries(google.forms[0] ?? [])).toEqual({
      grant_type: 'refresh_token',
      refresh_token: 'expired-calendar-refresh',
    });

    // Both keys were rewritten in the checkpoint's layout, readable by the old codec.
    for (const key of ['rs:access:expired-rs-access', 'rs:refresh:expired-rs-refresh']) {
      const record = openKvValue(kv.values.get(key) as string);
      expect(Object.keys(record)).toEqual(RECORD_FIELDS);
      expect(record).toMatchObject({
        rs_access_token: 'expired-rs-access',
        rs_refresh_token: 'expired-rs-refresh',
        provider: { access_token: 'google-fresh', refresh_token: 'google-refresh' },
      });
    }
  });

  test('a refresh at /token needs no client_id, as before', async () => {
    const kv = new FakeKv(before.kv);
    // Google answers a refresh without a new refresh token, so the access token stays.
    const google = googleTokens(() =>
      Response.json({ access_token: 'google-fresh', expires_in: 3600 }),
    );
    const f = productionProxy({ tokens: kvStore(kv), google });
    const refresh = (fields: Record<string, string> = {}) =>
      f.request(
        '/token',
        new URLSearchParams({
          grant_type: 'refresh_token',
          refresh_token: 'expired-rs-refresh',
          ...fields,
        }),
      );

    const first = await refresh();
    expect(first.status).toBe(200);
    expect(await first.json()).toMatchObject({
      access_token: 'expired-rs-access',
      refresh_token: 'expired-rs-refresh',
      token_type: 'bearer',
    });
    expect(f.google.forms).toHaveLength(1);
    // Any client_id is accepted for a record without a binding; another resource is not.
    expect((await refresh({ client_id: 'any-client' })).status).toBe(200);
    expect((await refresh({ resource: 'https://other.example/mcp' })).status).toBe(400);
  });

  test('the encrypted token file opens, and keeps a record whose Google token expired', async () => {
    const directory = temporaryDirectory();
    const path = join(directory, 'rs_tokens.json');
    writeFileSync(path, before.file);

    const store = new FileTokenStore(path, KEY, memoryLogger());
    expect(await store.getByRsAccess('file-rs-access')).toMatchObject({
      rs_refresh_token: 'file-rs-refresh',
      provider: { access_token: 'file-calendar-access', refresh_token: 'file-calendar-refresh' },
    });

    // And the verifier refreshes it, as the checkpoint's Bun server did.
    const google = googleTokens();
    const f = productionProxy({ tokens: store, google });
    expect((await f.proxy.verifier.verifyAccessToken('file-rs-access')).extra).toEqual({
      providerAccessToken: 'google-fresh',
    });
    expect(google.forms[0]?.get('refresh_token')).toBe('file-calendar-refresh');
  });

  test('a code issued before the migration, still within its two minutes, is redeemed', async () => {
    const directory = temporaryDirectory();
    const document = JSON.parse(before.authorityDocument) as {
      client: { client_id: string; redirect_uris: string[] };
      codes: Record<string, { expiresAt: number; codeChallenge: string }>;
    };
    const clientId = document.client.client_id;
    const redirectUri = document.client.redirect_uris[0] as string;
    const grant = document.codes[before.code];
    if (!grant) throw new Error('expected the recorded code');
    expect(before.codeVerifier).toBe(VERIFIER);
    expect(grant.codeChallenge).toBe(await pkceChallenge(VERIFIER));
    // Only the clock moved: the recorded code expired two minutes after it was written.
    grant.expiresAt = Date.now() + 60_000;
    writeFileSync(join(directory, `${clientId}.json`), JSON.stringify(document));

    const f = productionProxy({ authority: new FileOAuthAuthority(directory) });
    const response = await redeem(f, clientId, redirectUri, before.code);
    expect(response.status).toBe(200);
    const pair = (await response.json()) as { access_token: string; scope: string };
    expect(pair.scope).toBe(
      'https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.readonly',
    );
    const record = await f.tokens.getByRsAccess(pair.access_token);
    expect(record?.provider.access_token).toBe('grant-calendar-access');
    expect(record?.oauth).toBeUndefined();
    // Single use, as before.
    expect((await redeem(f, clientId, redirectUri, before.code)).status).toBe(400);
  });

  test('a client registered before the migration can still sign in', async () => {
    const directory = temporaryDirectory();
    const document = JSON.parse(before.authorityDocument) as {
      client: { client_id: string };
    };
    writeFileSync(join(directory, `${document.client.client_id}.json`), before.authorityDocument);
    const f = productionProxy({ authority: new FileOAuthAuthority(directory) });
    const response = await f.request(
      `/authorize?${new URLSearchParams({
        client_id: document.client.client_id,
        response_type: 'code',
        // Native clients may pick another port at sign-in.
        redirect_uri: 'http://127.0.0.1:50000/oauth/callback',
        code_challenge: await pkceChallenge(VERIFIER),
        code_challenge_method: 'S256',
      })}`,
    );
    expect(response.status).toBe(302);
    expect(
      new URL(response.headers.get('location') as string).searchParams.get('redirect_uri'),
    ).toBe(`${PRODUCTION}/oauth/callback`);
  });
});

describe('new writes keep the old layouts', () => {
  test('KV: both keys hold the same record, encrypted as before, without a binding', async () => {
    const kv = new FakeKv();
    await kvStore(kv).storeRsMapping(
      'new-rs-access',
      { access_token: 'new-calendar-access', refresh_token: 'new-calendar-refresh', expires_at: 1 },
      'new-rs-refresh',
    );
    const byAccess = openKvValue(kv.values.get('rs:access:new-rs-access') as string);
    expect(openKvValue(kv.values.get('rs:refresh:new-rs-refresh') as string)).toEqual(byAccess);
    expect(Object.keys(byAccess)).toEqual(RECORD_FIELDS);
  });

  test('KV: a checkpoint record is identical in shape to a new one', async () => {
    const old = openKvValue(before.kv['rs:access:current-rs-access']);
    expect(Object.keys(old)).toEqual(RECORD_FIELDS);
    expect(Object.keys(old.provider as object)).toEqual([
      'access_token',
      'refresh_token',
      'expires_at',
      'scopes',
    ]);
  });

  test('a sign-in through the proxy stores a record without a binding', async () => {
    const kv = new FakeKv();
    const f = proxyFixture({
      bindGrants: false,
      tokens: kvStore(kv),
      encryptionKey: KEY,
      redirectAllowlist: ['alice://oauth/callback'],
    });
    const registered = await f.request('/register', { redirect_uris: ['alice://oauth/callback'] });
    const { client_id } = (await registered.json()) as { client_id: string };
    const authorized = await f.request(
      `/authorize?${new URLSearchParams({
        client_id,
        response_type: 'code',
        redirect_uri: 'alice://oauth/callback',
        code_challenge: await pkceChallenge(VERIFIER),
        code_challenge_method: 'S256',
      })}`,
    );
    const state = new URL(authorized.headers.get('location') as string).searchParams.get('state');
    const back = await f.request(`/oauth/callback?code=google-code&state=${state}`);
    const code = new URL(back.headers.get('location') as string).searchParams.get('code') as string;
    const pair = (await (await redeem(f, client_id, 'alice://oauth/callback', code)).json()) as {
      access_token: string;
      refresh_token: string;
    };
    const stored = openKvValue(kv.values.get(`rs:access:${pair.access_token}`) as string);
    expect(Object.keys(stored)).toEqual(RECORD_FIELDS);
    expect(stored).toMatchObject({ provider: { access_token: 'google-access' } });
    expect(
      (
        await f.request(
          '/token',
          new URLSearchParams({ grant_type: 'refresh_token', refresh_token: pair.refresh_token }),
        )
      ).status,
    ).toBe(200);
  });

  test('KV without a key stores plain JSON, and reads it back', async () => {
    const kv = new FakeKv();
    const store = kvStore(kv, null);
    await store.storeRsMapping('plain-access', { access_token: 'g' }, 'plain-refresh');
    expect(JSON.parse(kv.values.get('rs:access:plain-access') as string)).toMatchObject({
      provider: { access_token: 'g' },
    });
    expect(await store.getByRsRefresh('plain-refresh')).toMatchObject({
      rs_access_token: 'plain-access',
    });
  });

  test('KV: rotating the access token moves the record', async () => {
    const kv = new FakeKv();
    await kvStore(kv).storeRsMapping('old-access', { access_token: 'g' }, 'the-refresh');
    await kvStore(kv).updateByRsRefresh('the-refresh', { access_token: 'g2' }, 'rotated-access');

    const restarted = kvStore(kv);
    expect(await restarted.getByRsAccess('old-access')).toBeNull();
    expect((await restarted.getByRsAccess('rotated-access'))?.provider.access_token).toBe('g2');
    expect((await restarted.getByRsRefresh('the-refresh'))?.rs_access_token).toBe('rotated-access');
  });

  test('KV: a failed write still works in this isolate', async () => {
    const kv: KvLike = {
      get: async () => null,
      put: async () => {
        throw new Error('KV quota exceeded');
      },
      delete: async () => {},
    };
    const logs: Parameters<typeof memoryLogger>[0] = [];
    const store = new KvTokenStore(kv, { encryptor: undefined, logger: memoryLogger(logs) });
    await store.storeRsMapping('a', { access_token: 'g' }, 'r');
    expect((await store.getByRsAccess('a'))?.provider.access_token).toBe('g');
    expect(logs.map((entry) => entry.level)).toEqual(['warning']);
  });

  test('file: version 1, encrypted whole, readable with the old codec and after a restart', async () => {
    const directory = temporaryDirectory();
    const path = join(directory, 'rs_tokens.json');
    const writer = new FileTokenStore(path, KEY, memoryLogger());
    await writer.storeRsMapping(
      'file-access',
      { access_token: 'g', refresh_token: 'r', expires_at: Date.now() - 1000 },
      'file-refresh',
    );
    writer.flush();

    const saved = openFile(readFileSync(path, 'utf8')) as { version: number; records: RsRecord[] };
    expect(saved.version).toBe(1);
    expect(saved.records[0]).toMatchObject({ rs_access_token: 'file-access' });
    const restarted = new FileTokenStore(path, KEY, memoryLogger());
    expect((await restarted.getByRsRefresh('file-refresh'))?.provider.refresh_token).toBe('r');
  });

  test('file: a plaintext version 1 file is read unchanged', async () => {
    const directory = temporaryDirectory();
    const path = join(directory, 'rs_tokens.json');
    const record: RsRecord = {
      rs_access_token: 'existing-access',
      rs_refresh_token: 'existing-refresh',
      provider: {
        access_token: 'g',
        refresh_token: 'r',
        expires_at: Date.now() + 3_600_000,
        scopes: ['s'],
      },
      created_at: Date.now() - 10_000,
    };
    writeFileSync(path, JSON.stringify({ version: 1, encrypted: false, records: [record] }));
    const store = new FileTokenStore(path, undefined, memoryLogger());
    expect(await store.getByRsAccess('existing-access')).toMatchObject(record);
    store.flush();
    expect(JSON.parse(readFileSync(path, 'utf8')).records[0]).toMatchObject(record);
  });

  test('the authority grant layout is the KV layout, prefixed with enc:', async () => {
    const material = before.grantProviderMaterial;
    expect(material.startsWith('enc:')).toBe(true);
    expect(openKvValue(material.slice(4))).toMatchObject({ access_token: 'grant-calendar-access' });
    // The checkpoint codec's output opens with the new code, and the other way round.
    expect(JSON.parse(await createEncryptor(KEY).decrypt(before.sealed))).toEqual({ a: 1 });
    expect(openKvValue(await createEncryptor(KEY).encrypt('{"a":1}'))).toEqual({ a: 1 });
    const iv = randomBytes(12);
    const cipher = createCipheriv('aes-256-gcm', Buffer.from(KEY, 'base64url'), iv);
    const body = Buffer.concat([cipher.update('{"b":2}'), cipher.final()]);
    const old = Buffer.concat([iv, body, cipher.getAuthTag()]).toString('base64url');
    expect(JSON.parse(await createEncryptor(KEY).decrypt(old))).toEqual({ b: 2 });
  });
});
