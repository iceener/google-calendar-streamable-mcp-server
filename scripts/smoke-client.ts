import assert from 'node:assert/strict';
import { Client, StreamableHTTPClientTransport } from '@modelcontextprotocol/client';
import { MOCK_CALENDAR } from './mock-google';

/**
 * Sign in through the OAuth proxy over real sockets, as a native client does: register,
 * authorize, come back from Google (mocked on loopback), and redeem the code with PKCE.
 * Returns the access token the proxy issued.
 */
export async function signIn(origin: URL, label: string): Promise<string> {
  assert.equal(origin.hostname, '127.0.0.1', 'smoke traffic stays on loopback');
  const at = (path: string) => new URL(path, origin);
  const redirectUri = 'http://127.0.0.1:43210/oauth/callback';

  const metadata = (await (await fetch(at('/.well-known/oauth-authorization-server'))).json()) as {
    issuer: string;
    registration_endpoint: string;
  };
  assert.equal(metadata.issuer, origin.origin);

  const registered = await fetch(metadata.registration_endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ application_type: 'native', redirect_uris: [redirectUri] }),
  });
  assert.equal(registered.status, 201);
  const { client_id: clientId } = (await registered.json()) as { client_id: string };

  const verifier = 'smoke-verifier-'.padEnd(64, 'x');
  const challenge = Buffer.from(
    await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)),
  ).toString('base64url');
  const authorized = await fetch(
    at(
      `/authorize?${new URLSearchParams({
        client_id: clientId,
        response_type: 'code',
        redirect_uri: redirectUri,
        code_challenge: challenge,
        code_challenge_method: 'S256',
        state: 'smoke-state',
      })}`,
    ),
    { redirect: 'manual' },
  );
  assert.equal(authorized.status, 302);
  const google = new URL(authorized.headers.get('location') ?? '');
  assert.equal(
    `${google.origin}${google.pathname}`,
    'https://accounts.google.com/o/oauth2/v2/auth',
  );
  assert.equal(google.searchParams.get('redirect_uri'), at('/oauth/callback').href);
  assert.equal(google.searchParams.get('access_type'), 'offline');

  // The user consents; Google sends the browser back to the callback with a code.
  const callback = await fetch(
    at(`/oauth/callback?code=google-code&state=${google.searchParams.get('state')}`),
    { redirect: 'manual' },
  );
  assert.equal(callback.status, 302);
  const back = new URL(callback.headers.get('location') ?? '');
  assert.equal(`${back.origin}${back.pathname}`, redirectUri);
  assert.equal(back.searchParams.get('state'), 'smoke-state');

  const issued = await fetch(at('/token'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'authorization_code',
      code: back.searchParams.get('code') ?? '',
      client_id: clientId,
      redirect_uri: redirectUri,
      code_verifier: verifier,
    }),
  });
  assert.equal(issued.status, 200);
  const { access_token: token, refresh_token: refreshToken } = (await issued.json()) as {
    access_token: string;
    refresh_token: string;
  };
  assert.notEqual(token, 'google-access', 'the client gets the proxy’s token, not Google’s');

  // Clients refresh without client_id: the records carry no client binding (docs/oauth.md).
  const refreshed = await fetch(at('/token'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ grant_type: 'refresh_token', refresh_token: refreshToken }),
  });
  assert.equal(refreshed.status, 200);
  assert.equal(((await refreshed.json()) as { access_token: string }).access_token, token);
  console.info(`${label}: OAuth round trip through the mocked Google passed`);
  return token;
}

/**
 * Drive a running server over real sockets with the official client, in both protocol
 * eras. Shared by the Bun and Workers smoke tests; it needs no network beyond loopback.
 * With a token, the server must be in OAuth mode: requests without it are refused, requests
 * with it get through, and Google Calendar (mocked) sees the user's Google token.
 */
export async function smoke(endpoint: URL, label: string, token?: string): Promise<void> {
  assert.equal(endpoint.hostname, '127.0.0.1', 'smoke traffic stays on loopback');
  const authorization = token ? { Authorization: `Bearer ${token}` } : undefined;

  const health = await fetch(new URL('/health', endpoint));
  assert.equal(health.status, 200);

  if (token) {
    const anonymous = await fetch(endpoint, { method: 'POST' });
    assert.equal(anonymous.status, 401);
    assert.match(anonymous.headers.get('WWW-Authenticate') ?? '', /resource_metadata=/);
  }

  for (const era of ['modern', 'legacy'] as const) {
    const client = new Client(
      { name: 'smoke', version: '1.0.0' },
      { versionNegotiation: { mode: era === 'modern' ? 'auto' : 'legacy' } },
    );
    await client.connect(
      new StreamableHTTPClientTransport(endpoint, {
        fetch: async (url, init) => {
          const headers = new Headers(init?.headers);
          if (token) headers.set('Authorization', `Bearer ${token}`);
          const response = await fetch(url, { ...init, headers, redirect: 'error' });
          assert.equal(response.headers.has('Mcp-Session-Id'), false, 'the server is stateless');
          return response;
        },
      }),
    );
    try {
      assert.equal(client.getProtocolEra(), era);
      const { tools } = await client.listTools();
      assert.equal(tools[0]?.name, 'list_calendars');
      assert.equal(tools.length, 7);

      const calendars = await client.callTool({ name: 'list_calendars', arguments: {} });
      if (token) {
        assert.deepEqual(calendars.structuredContent, { items: [MOCK_CALENDAR] });
      } else {
        assert.equal(calendars.isError, true);
      }
    } finally {
      await client.close();
    }
    console.info(`${label}: ${era} client passed`);
  }

  const foreignOrigin = await fetch(endpoint, {
    method: 'POST',
    headers: { Origin: 'https://evil.example' },
  });
  assert.equal(foreignOrigin.status, 403);

  const oversized = await fetch(endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...authorization },
    body: new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array(700));
        controller.enqueue(new Uint8Array(700));
        controller.close();
      },
    }),
  });
  assert.equal(oversized.status, 413);
  console.info(`${label}: Origin guard and streamed body limit passed`);
}

if (import.meta.main) {
  const [endpoint, label = 'server', token] = process.argv.slice(2);
  assert.ok(
    endpoint,
    'usage: bun scripts/smoke-client.ts <loopback MCP URL> [label] [token | --sign-in]',
  );
  const url = new URL(endpoint);
  await smoke(url, label, token === '--sign-in' ? await signIn(new URL(url.origin), label) : token);
}
