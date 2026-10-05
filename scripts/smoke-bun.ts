import assert from 'node:assert/strict';
import { Client, StreamableHTTPClientTransport } from '@modelcontextprotocol/client';
import { serve } from '../src/bun';
import { MemoryOAuthAuthority } from '../src/oauth/authority';
import { MemoryTokenStore } from '../src/oauth/token-store';
import { createApp } from '../src/platform/app';
import { parseConfig } from '../src/platform/config';
import { createLogger } from '../src/platform/logger';
import { createDeps } from '../src/server';
import { TEST_SETTINGS } from '../tests/settings';
import { GOOGLE_HOSTS, MOCK_CALENDAR, startMockGoogle, toMock } from './mock-google';
import { signIn, smoke } from './smoke-client';

/**
 * A tool call that waits on Google sends nothing until Google answers, and Bun's default idle
 * timeout (10 s, enforced within a few seconds) drops a connection that stays quiet longer.
 * 15 s is past that window, so this check fails reliably without the fix in src/bun.ts.
 */
const QUIET_MS = 15_000;

const google = await startMockGoogle('smoke-client', 'smoke-secret');
// The server calls Google with the runtime's fetch: send those calls to the loopback mock.
const realFetch = globalThis.fetch;
globalThis.fetch = Object.assign(
  async (input: string | URL | Request, init?: RequestInit) => {
    const request =
      input instanceof Request ? new Request(input, init) : new Request(String(input), init);
    return GOOGLE_HOSTS.includes(new URL(request.url).hostname)
      ? realFetch(await toMock(request, google))
      : realFetch(request);
  },
  { preconnect: realFetch.preconnect },
);

/** A port the OS just assigned: the proxy needs its own origin before it starts. */
function freePort(): number {
  const probe = Bun.serve({ hostname: '127.0.0.1', port: 0, fetch: () => new Response() });
  const port = probe.port as number;
  probe.stop(true);
  return port;
}

try {
  const port = freePort();
  const origin = `http://127.0.0.1:${port}`;
  const config = parseConfig({
    NODE_ENV: 'test',
    HOST: '127.0.0.1',
    PORT: String(port),
    MCP_PUBLIC_URL: `${origin}/mcp`,
    MCP_MAX_REQUEST_BYTES: '1024',
    AUTH_MODE: 'oauth',
    OAUTH_ISSUER_URL: origin,
    OAUTH_AUTHORIZATION_URL: `${origin}/authorize`,
    OAUTH_TOKEN_URL: `${origin}/token`,
    OAUTH_REGISTRATION_URL: `${origin}/register`,
    OAUTH_SCOPES: 'https://www.googleapis.com/auth/calendar.readonly',
    PROVIDER_CLIENT_ID: 'smoke-client',
    PROVIDER_CLIENT_SECRET: 'smoke-secret',
    TOKENS_ENC_KEY: Buffer.alloc(32, 3).toString('base64url'),
    ...TEST_SETTINGS,
  });
  const deps = createDeps(config, createLogger('warning'), {
    tokens: new MemoryTokenStore(),
    authority: new MemoryOAuthAuthority(),
  });
  const app = createApp(config, { deps });
  const server = serve(config, app);
  try {
    const endpoint = new URL('/mcp', server.url);
    const token = await signIn(new URL(origin), 'bun');
    await smoke(endpoint, 'bun (oauth)', token);
    assert.deepEqual(google.grants, ['authorization_code']);

    // Google Calendar takes longer than Bun's default idle timeout to answer.
    google.delayMs = QUIET_MS;
    const client = new Client(
      { name: 'smoke', version: '1.0.0' },
      { versionNegotiation: { mode: 'auto' } },
    );
    await client.connect(
      new StreamableHTTPClientTransport(endpoint, {
        fetch: (url, init) => {
          const headers = new Headers(init?.headers);
          headers.set('Authorization', `Bearer ${token}`);
          return fetch(url, { ...init, headers });
        },
      }),
    );
    const result = await client.callTool(
      { name: 'list_calendars', arguments: {} },
      { timeout: QUIET_MS + 5_000 },
    );
    await client.close();
    assert.deepEqual(result.structuredContent, { items: [MOCK_CALENDAR] });
    console.info(`bun: a connection quiet for ${QUIET_MS / 1000} s survived the idle timeout`);
  } finally {
    await app.close();
    await server.stop(true);
  }
} finally {
  globalThis.fetch = realFetch;
  await google.close();
}
