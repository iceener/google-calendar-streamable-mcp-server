import { afterEach, describe, expect, test } from 'bun:test';
import {
  Client,
  type FetchLike,
  StreamableHTTPClientTransport,
} from '@modelcontextprotocol/client';
import { type AppConfig, parseConfig } from '../src/config/env.js';
import { buildHttpApp, type HttpRuntime } from '../src/http/app.js';
import { createProviderTokenVerifier } from '../src/shared/auth/provider-token-verifier.js';
import { MemoryTokenStore } from '../src/shared/storage/memory.js';

const clients = new Set<Client>();
const runtimes = new Set<HttpRuntime>();
const stores = new Set<MemoryTokenStore>();
const originalFetch = globalThis.fetch;

afterEach(async () => {
  globalThis.fetch = originalFetch;
  await Promise.all([...clients].map((client) => client.close()));
  await Promise.all([...runtimes].map((runtime) => runtime.close()));
  for (const store of stores) store.stopCleanup();
  clients.clear();
  runtimes.clear();
  stores.clear();
});

function config(overrides: Record<string, unknown> = {}): AppConfig {
  return parseConfig({
    NODE_ENV: 'test',
    MCP_PUBLIC_URL: 'http://localhost:3000/mcp',
    MCP_ALLOWED_HOSTS: 'localhost',
    MCP_ALLOWED_ORIGIN_HOSTNAMES: 'localhost',
    AUTH_ENABLED: 'true',
    OAUTH_ISSUER_URL: 'http://localhost:3001',
    OAUTH_REDIRECT_URI: 'http://localhost:3001/oauth/callback',
    PROVIDER_CLIENT_ID: 'google-client',
    PROVIDER_CLIENT_SECRET: 'google-secret',
    ...overrides,
  });
}

function store(): MemoryTokenStore {
  const value = new MemoryTokenStore();
  stores.add(value);
  return value;
}

async function authorize(
  tokenStore: MemoryTokenStore,
  mcpToken: string,
  providerToken: string,
  scopes = config().OAUTH_REQUIRED_SCOPES,
): Promise<void> {
  await tokenStore.storeRsMapping(
    mcpToken,
    {
      access_token: providerToken,
      refresh_token: `${providerToken}-refresh`,
      expires_at: Date.now() + 300_000,
      scopes,
    },
    `${mcpToken}-refresh`,
  );
}

function runtime(tokenStore: MemoryTokenStore, appConfig = config()): HttpRuntime {
  const value = buildHttpApp(appConfig, {
    runtimeName: 'test',
    tokenStore,
    mountOAuthProxy: true,
  });
  runtimes.add(value);
  return value;
}

function runtimeFetch(app: HttpRuntime, token?: string): FetchLike {
  return async (url, init) => {
    const headers = new Headers(init?.headers);
    headers.set('Host', 'localhost:3000');
    if (token) headers.set('Authorization', `Bearer ${token}`);
    return app.fetch(new Request(url, { ...init, headers }));
  };
}

async function connect(
  app: HttpRuntime,
  era: 'modern' | 'legacy',
  token: string,
): Promise<Client> {
  const client = new Client(
    { name: `calendar-${era}-test`, version: '1.0.0' },
    era === 'modern'
      ? { versionNegotiation: { mode: { pin: '2026-07-28' } } }
      : undefined,
  );
  await client.connect(
    new StreamableHTTPClientTransport(new URL('http://localhost:3000/mcp'), {
      fetch: runtimeFetch(app, token),
      authProvider: { token: async () => token },
    }),
  );
  clients.add(client);
  return client;
}

describe('Google Calendar MCP protocol and credential boundary', () => {
  test('serves the exact modern tool contract and provider mock', async () => {
    const tokenStore = store();
    await authorize(tokenStore, 'mcp-caller-token', 'google-provider-token');
    const upstreamTokens: string[] = [];
    globalThis.fetch = (async (input, init) => {
      expect(String(input)).toBe(
        'https://www.googleapis.com/calendar/v3/users/me/calendarList',
      );
      upstreamTokens.push(new Headers(init?.headers).get('Authorization') ?? '');
      return Response.json({
        items: [
          {
            id: 'primary',
            summary: 'Primary',
            primary: true,
            accessRole: 'owner',
            timeZone: 'Europe/Warsaw',
          },
        ],
      });
    }) as typeof globalThis.fetch;

    const client = await connect(runtime(tokenStore), 'modern', 'mcp-caller-token');
    expect(client.getNegotiatedProtocolVersion()).toBe('2026-07-28');
    expect((await client.listTools()).tools.map((tool) => tool.name)).toEqual([
      'list_calendars',
      'search_events',
      'check_availability',
      'create_event',
      'update_event',
      'delete_event',
      'respond_to_event',
    ]);
    const result = await client.callTool({
      name: 'list_calendars',
      arguments: {},
    });
    expect(result.structuredContent).toMatchObject({
      items: [{ id: 'primary', summary: 'Primary', accessRole: 'owner' }],
    });
    expect(upstreamTokens).toEqual(['Bearer google-provider-token']);
    expect(upstreamTokens[0]).not.toContain('mcp-caller-token');
  });

  test('puts only the validated provider access token in AuthInfo.extra', async () => {
    const tokenStore = store();
    await authorize(tokenStore, 'resource-token', 'provider-access-token');
    const authInfo = await createProviderTokenVerifier(
      config(),
      tokenStore,
    ).verifyAccessToken('resource-token');
    expect(authInfo.token).toBe('resource-token');
    expect(authInfo.extra).toEqual({
      providerAccessToken: 'provider-access-token',
    });
    expect(JSON.stringify(authInfo.extra)).not.toContain('refresh');
  });

  test('isolates concurrent MCP principals and never forwards either MCP token', async () => {
    const tokenStore = store();
    await authorize(tokenStore, 'alice-mcp', 'alice-google');
    await authorize(tokenStore, 'bob-mcp', 'bob-google');
    const seen: string[] = [];
    globalThis.fetch = (async (_input, init) => {
      const authorization = new Headers(init?.headers).get('Authorization') ?? '';
      seen.push(authorization);
      const who = authorization.includes('alice') ? 'Alice' : 'Bob';
      return Response.json({
        items: [{ id: who.toLowerCase(), summary: who, accessRole: 'owner' }],
      });
    }) as typeof globalThis.fetch;
    const app = runtime(tokenStore);
    const [alice, bob] = await Promise.all([
      connect(app, 'modern', 'alice-mcp'),
      connect(app, 'modern', 'bob-mcp'),
    ]);
    const [aliceResult, bobResult] = await Promise.all([
      alice.callTool({ name: 'list_calendars', arguments: {} }),
      bob.callTool({ name: 'list_calendars', arguments: {} }),
    ]);
    expect(aliceResult.structuredContent).toMatchObject({
      items: [{ summary: 'Alice' }],
    });
    expect(bobResult.structuredContent).toMatchObject({
      items: [{ summary: 'Bob' }],
    });
    expect(seen.sort()).toEqual(['Bearer alice-google', 'Bearer bob-google']);
    expect(seen.join(' ')).not.toContain('alice-mcp');
    expect(seen.join(' ')).not.toContain('bob-mcp');
  });

  test('preserves event creation and availability behavior', async () => {
    const tokenStore = store();
    await authorize(tokenStore, 'mcp-token', 'provider-token');
    const requests: Array<{ url: string; method: string; body?: string }> = [];
    globalThis.fetch = (async (input, init) => {
      const url = String(input);
      const method = init?.method ?? 'GET';
      requests.push({
        url,
        method,
        ...(typeof init?.body === 'string' ? { body: init.body } : {}),
      });
      if (url.endsWith('/freeBusy')) {
        return Response.json({
          timeMin: '2026-07-28T09:00:00Z',
          timeMax: '2026-07-28T10:00:00Z',
          calendars: { primary: { busy: [] } },
        });
      }
      if (method === 'POST' && url.includes('/calendars/primary/events')) {
        return Response.json({
          id: 'event-1',
          summary: 'Planning',
          start: { dateTime: '2026-07-28T09:00:00Z' },
          end: { dateTime: '2026-07-28T10:00:00Z' },
          htmlLink: 'https://calendar.google.com/event-1',
        });
      }
      throw new Error(`Unexpected Google URL: ${url}`);
    }) as typeof globalThis.fetch;
    const client = await connect(runtime(tokenStore), 'modern', 'mcp-token');
    const availability = await client.callTool({
      name: 'check_availability',
      arguments: {
        timeMin: '2026-07-28T09:00:00Z',
        timeMax: '2026-07-28T10:00:00Z',
      },
    });
    expect(availability.structuredContent).toMatchObject({
      calendars: { primary: { busy: [] } },
    });
    const created = await client.callTool({
      name: 'create_event',
      arguments: {
        summary: 'Planning',
        start: '2026-07-28T09:00:00Z',
        end: '2026-07-28T10:00:00Z',
      },
    });
    expect(created.structuredContent).toMatchObject({
      id: 'event-1',
      summary: 'Planning',
    });
    expect(requests.map((request) => request.method)).toEqual(['POST', 'POST']);
  });

  test('propagates official-client cancellation to Google fetch', async () => {
    const tokenStore = store();
    await authorize(tokenStore, 'cancel-mcp', 'cancel-google');
    let providerAborted = false;
    globalThis.fetch = (async (_input, init) =>
      new Promise<Response>((_resolve, reject) => {
        const signal = init?.signal;
        const onAbort = () => {
          providerAborted = true;
          reject(new DOMException('Aborted', 'AbortError'));
        };
        if (signal?.aborted) onAbort();
        else signal?.addEventListener('abort', onAbort, { once: true });
      })) as typeof globalThis.fetch;
    const client = await connect(runtime(tokenStore), 'modern', 'cancel-mcp');
    const controller = new AbortController();
    const pending = client.callTool(
      { name: 'list_calendars', arguments: {} },
      { signal: controller.signal },
    );
    setTimeout(() => controller.abort(), 20);
    await expect(pending).rejects.toThrow();
    for (let attempt = 0; attempt < 20 && !providerAborted; attempt += 1) {
      await Bun.sleep(5);
    }
    expect(providerAborted).toBe(true);
  });

  test('supports the official legacy client through stateless fallback', async () => {
    const tokenStore = store();
    await authorize(tokenStore, 'legacy-mcp', 'legacy-google');
    const client = await connect(runtime(tokenStore), 'legacy', 'legacy-mcp');
    expect(client.getProtocolEra()).toBe('legacy');
    expect((await client.listTools()).tools).toHaveLength(7);
  });

  test('serves OAuth metadata and rejects missing, invalid, and under-scoped tokens', async () => {
    const tokenStore = store();
    const app = runtime(tokenStore);
    const metadata = await app.fetch(
      new Request('http://localhost:3000/.well-known/oauth-protected-resource/mcp', {
        headers: { Host: 'localhost:3000' },
      }),
    );
    expect(metadata.status).toBe(200);
    expect(await metadata.json()).toMatchObject({
      resource: 'http://localhost:3000/mcp',
      authorization_servers: ['http://localhost:3001'],
    });

    const missing = await app.fetch(
      new Request('http://localhost:3000/mcp', {
        method: 'GET',
        headers: { Host: 'localhost:3000' },
      }),
    );
    expect(missing.status).toBe(401);

    const invalid = await app.fetch(
      new Request('http://localhost:3000/mcp', {
        method: 'GET',
        headers: {
          Host: 'localhost:3000',
          Authorization: 'Bearer invalid-token',
        },
      }),
    );
    expect(invalid.status).toBe(401);

    await authorize(tokenStore, 'under-scoped', 'google-token', []);
    const insufficient = await app.fetch(
      new Request('http://localhost:3000/mcp', {
        method: 'GET',
        headers: {
          Host: 'localhost:3000',
          Authorization: 'Bearer under-scoped',
        },
      }),
    );
    expect(insufficient.status).toBe(403);
    expect(insufficient.headers.get('WWW-Authenticate')).toContain(
      'insufficient_scope',
    );
  });

  test('enforces Host, Origin, method, CORS, and request body limits', async () => {
    const tokenStore = store();
    const app = runtime(
      tokenStore,
      config({ AUTH_ENABLED: 'false', MCP_MAX_REQUEST_BYTES: '1024' }),
    );
    expect(
      (
        await app.fetch(
          new Request('http://localhost:3000/mcp', {
            headers: { Host: 'localhost:3000' },
          }),
        )
      ).status,
    ).toBe(405);
    expect(
      (
        await app.fetch(
          new Request('http://localhost:3000/health', {
            headers: { Host: 'evil.example' },
          }),
        )
      ).status,
    ).toBe(403);
    const badOrigin = await app.fetch(
      new Request('http://localhost:3000/mcp', {
        method: 'POST',
        headers: {
          Host: 'localhost:3000',
          Origin: 'https://evil.example',
          'Content-Type': 'application/json',
        },
        body: '{}',
      }),
    );
    expect(badOrigin.status).toBe(403);
    expect(badOrigin.headers.has('Access-Control-Allow-Origin')).toBe(false);
    const preflight = await app.fetch(
      new Request('http://localhost:3000/mcp', {
        method: 'OPTIONS',
        headers: {
          Host: 'localhost:3000',
          Origin: 'http://localhost:8080',
          'Access-Control-Request-Method': 'POST',
          'Access-Control-Request-Headers': 'authorization, content-type',
        },
      }),
    );
    expect(preflight.status).toBe(204);
    const oversized = await app.fetch(
      new Request('http://localhost:3000/mcp', {
        method: 'POST',
        headers: {
          Host: 'localhost:3000',
          'Content-Type': 'application/json',
        },
        body: 'x'.repeat(1_025),
      }),
    );
    expect(oversized.status).toBe(413);
  });

  test('returns SDK-owned modern header and media-type errors', async () => {
    const app = runtime(store(), config({ AUTH_ENABLED: 'false' }));
    const discover = JSON.stringify({
      jsonrpc: '2.0',
      id: 1,
      method: 'server/discover',
      params: {
        _meta: {
          'io.modelcontextprotocol/protocolVersion': '2026-07-28',
          'io.modelcontextprotocol/clientCapabilities': {},
          'io.modelcontextprotocol/clientInfo': {
            name: 'raw-test',
            version: '1.0.0',
          },
        },
      },
    });
    const mismatch = await app.fetch(
      new Request('http://localhost:3000/mcp', {
        method: 'POST',
        headers: {
          Host: 'localhost:3000',
          Accept: 'application/json, text/event-stream',
          'Content-Type': 'application/json',
          'MCP-Protocol-Version': '2026-07-28',
          'Mcp-Method': 'tools/list',
        },
        body: discover,
      }),
    );
    expect(mismatch.status).toBe(400);
    expect(await mismatch.json()).toMatchObject({ error: { code: -32020 } });
    const mediaType = await app.fetch(
      new Request('http://localhost:3000/mcp', {
        method: 'POST',
        headers: { Host: 'localhost:3000', 'Content-Type': 'text/plain' },
        body: discover,
      }),
    );
    expect(mediaType.status).toBe(415);
  });
});
