import { afterEach, expect, test } from 'bun:test';
import { MemoryOAuthAuthority } from '../src/oauth/authority';
import { pkceChallenge } from '../src/oauth/crypto';
import { MemoryTokenStore } from '../src/oauth/token-store';
import { type App, createApp } from '../src/platform/app';
import { parseConfig } from '../src/platform/config';
import { createDeps } from '../src/server';
import before from './fixtures/routes-before.json';
import { cleanup, fakeFetch, memoryLogger, track } from './helpers';
import { PRODUCTION_VARS } from './production-config';

const realFetch = globalThis.fetch;
afterEach(async () => {
  globalThis.fetch = realFetch;
  await cleanup();
});

/**
 * The public routes, compared with what the deployed code answered before the migration
 * (`fixtures/routes-before.json`, recorded in workerd). The server runs here in process,
 * built by `createDeps` exactly as in production, with the production vars from
 * `wrangler.production.example.jsonc` and the five deployed secrets. Google is a fake `fetch`.
 */
const ORIGIN = 'https://google-calendar.example.workers.dev';
const HOST = 'google-calendar.example.workers.dev';

const NATIVE = 'http://127.0.0.1:43210/oauth/callback';
const VERIFIER = 'v'.repeat(43);
/** Fake values for every deployed secret, including the two nothing reads. */
const SECRETS = {
  PROVIDER_CLIENT_ID: 'mock-google-client',
  PROVIDER_CLIENT_SECRET: 'mock-google-secret',
  TOKENS_ENC_KEY: Buffer.alloc(32, 1).toString('base64url'),
  BEARER_TOKEN: 'unused-bearer-secret',
  RESEND_API_KEY: 'unused-resend-secret',
};

/** Unknown paths answer JSON, as the template does everywhere, instead of plain text. */
const JSON_NOT_FOUND = { contentType: 'application/json', error: 'not_found' };
/**
 * The OAuth endpoints answer allowed browser origins: the preflight (which `cors()` answers
 * before the `no-store` header is set; it carries no data), then the response itself.
 */
const OAUTH_PREFLIGHT = {
  status: 204,
  contentType: null,
  allowOrigin: 'https://claude.ai',
  cacheControl: null,
};

/** Differences from the recorded matrix, each deliberate. Keyed by "<label> <METHOD> <path>". */
const CHANGED: Record<string, Partial<Probe>> = {
  'discovery GET /.well-known/oauth-protected-resource': JSON_NOT_FOUND,
  'discovery GET /mcp/.well-known/oauth-protected-resource': JSON_NOT_FOUND,
  'discovery GET /mcp/.well-known/oauth-authorization-server': JSON_NOT_FOUND,
  'not served GET /.well-known/openid-configuration': JSON_NOT_FOUND,
  'not served GET /.well-known/oauth-protected-resource/other': JSON_NOT_FOUND,
  'unknown GET /nope': JSON_NOT_FOUND,
  'unknown POST /nope': JSON_NOT_FOUND,
  'register GET /register': JSON_NOT_FOUND,
  'token GET /token': JSON_NOT_FOUND,
  'revoke GET /revoke': JSON_NOT_FOUND,
  'register preflight, claude.ai OPTIONS /register': OAUTH_PREFLIGHT,
  'token preflight, claude.ai OPTIONS /token': OAUTH_PREFLIGHT,
  'register, web, claude.ai origin POST /register': { allowOrigin: 'https://claude.ai' },
  'token, refresh, claude.ai origin POST /token': { allowOrigin: 'https://claude.ai' },
};

interface Probe {
  label: string;
  method: string;
  path: string;
  status: number;
  contentType: string | null;
  allowOrigin: string | null;
  cacheControl: string | null;
  location?: string;
  error?: unknown;
}

/** The stand-in production vars, from wrangler.production.example.jsonc. */
function productionVars(): Record<string, string> {
  return { ...PRODUCTION_VARS };
}

/** The production Worker's app, with KV and the Durable Object in memory. */
function productionApp(): App {
  const vars = productionVars();
  // Google, as the recording's harness answered it. Nothing else may be called.
  globalThis.fetch = fakeFetch(async (request) => {
    if (request.url === 'https://oauth2.googleapis.com/token') {
      const form = new URLSearchParams(await request.text());
      return Response.json({
        access_token: `google-access-${form.get('grant_type')}`,
        refresh_token: 'google-refresh',
        expires_in: 3600,
        scope: vars.OAUTH_SCOPES,
        token_type: 'Bearer',
      });
    }
    return undefined;
  });
  const config = parseConfig({ ...vars, ...SECRETS });
  const runtime = { tokens: new MemoryTokenStore(), authority: new MemoryOAuthAuthority() };
  return track(createApp(config, { deps: createDeps(config, memoryLogger(), runtime) }));
}

/** The probe sequence the recorded matrix came from (`matrix.mjs`), in the same order. */
async function probeAll(app: App): Promise<{
  matrix: Probe[];
  documents: Record<string, unknown>;
  google: URL;
  clientState: string | null;
  token: Record<string, unknown>;
  refreshed: Record<string, unknown>;
}> {
  const matrix: Probe[] = [];
  const documents: Record<string, unknown> = {};
  const probe = async (
    label: string,
    method: string,
    path: string,
    options: { headers?: Record<string, string>; body?: string } = {},
  ) => {
    const response = await app.fetch(
      new Request(new URL(path, ORIGIN).href, {
        method,
        headers: { Host: HOST, ...options.headers },
        ...(options.body !== undefined && { body: options.body }),
        redirect: 'manual',
      }),
    );
    const location = response.headers.get('location');
    const entry: Probe = {
      label,
      method,
      path,
      status: response.status,
      contentType: response.headers.get('content-type')?.split(';')[0] ?? null,
      allowOrigin: response.headers.get('access-control-allow-origin'),
      cacheControl: response.headers.get('cache-control'),
      ...(location && { location: `${new URL(location).origin}${new URL(location).pathname}` }),
    };
    const text = await response.text();
    if (entry.contentType === 'application/json' && text) {
      const json = JSON.parse(text) as Record<string, unknown>;
      if (json && typeof json === 'object' && 'error' in json) entry.error = json.error;
    }
    matrix.push(entry);
    return { response, text };
  };
  const json = (value: unknown, headers: Record<string, string> = {}) => ({
    headers: { 'Content-Type': 'application/json', ...headers },
    body: JSON.stringify(value),
  });
  const form = (value: Record<string, string>, headers: Record<string, string> = {}) => ({
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', ...headers },
    body: new URLSearchParams(value).toString(),
  });
  const EVIL = { Origin: 'https://evil.example' };
  const CLAUDE = { Origin: 'https://claude.ai' };

  for (const path of [
    '/.well-known/oauth-protected-resource/mcp',
    '/.well-known/oauth-authorization-server',
    '/.well-known/oauth-protected-resource',
    '/mcp/.well-known/oauth-protected-resource',
    '/mcp/.well-known/oauth-authorization-server',
  ]) {
    const { response, text } = await probe('discovery', 'GET', path);
    documents[path] = response.status === 200 ? JSON.parse(text) : null;
    await probe('discovery, foreign origin', 'GET', path, { headers: EVIL });
  }
  for (const path of [
    '/.well-known/oauth-protected-resource/mcp',
    '/.well-known/oauth-authorization-server',
  ]) {
    await probe('discovery', 'HEAD', path);
    await probe('discovery preflight', 'OPTIONS', path, {
      headers: { ...EVIL, 'Access-Control-Request-Method': 'GET' },
    });
    await probe('discovery', 'POST', path);
  }
  await probe('not served', 'GET', '/.well-known/openid-configuration');
  await probe('not served', 'GET', '/.well-known/oauth-protected-resource/other');
  await probe('health', 'GET', '/health');
  await probe('icon', 'GET', '/icon.svg');
  await probe('unknown', 'GET', '/nope');
  await probe('unknown', 'POST', '/nope');

  const list = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/list', params: {} });
  const mcp = { 'Content-Type': 'application/json', Accept: 'application/json, text/event-stream' };
  await probe('mcp, anonymous', 'POST', '/mcp', { headers: mcp, body: list });
  await probe('mcp, anonymous', 'GET', '/mcp');
  await probe('mcp, unknown token', 'POST', '/mcp', {
    headers: { ...mcp, Authorization: 'Bearer not-a-token' },
    body: list,
  });
  await probe('mcp, deployed BEARER_TOKEN secret', 'POST', '/mcp', {
    headers: { ...mcp, Authorization: `Bearer ${SECRETS.BEARER_TOKEN}` },
    body: list,
  });
  await probe('mcp, foreign origin', 'POST', '/mcp', { headers: { ...mcp, ...EVIL }, body: list });
  await probe('mcp preflight, claude.ai', 'OPTIONS', '/mcp', {
    headers: {
      ...CLAUDE,
      'Access-Control-Request-Method': 'POST',
      'Access-Control-Request-Headers': 'authorization, content-type, mcp-protocol-version',
    },
  });

  await probe('register, invalid', 'POST', '/register', json({ redirect_uris: [] }));
  await probe('register', 'GET', '/register');
  await probe(
    'register, foreign origin',
    'POST',
    '/register',
    json({ application_type: 'native', redirect_uris: [NATIVE] }, EVIL),
  );
  await probe(
    'register, web, claude.ai origin',
    'POST',
    '/register',
    json({ redirect_uris: ['https://claude.ai/api/mcp/auth_callback'] }, CLAUDE),
  );
  await probe(
    'register, web, OAUTH_REDIRECT_URI value',
    'POST',
    '/register',
    json({ redirect_uris: ['https://agi.example.net/mcp/oauth/calendar/callback'] }),
  );
  await probe(
    'register, web, unlisted redirect',
    'POST',
    '/register',
    json({ redirect_uris: ['https://evil.example/callback'] }),
  );
  await probe(
    'register, native localhost',
    'POST',
    '/register',
    json({ application_type: 'native', redirect_uris: ['http://localhost:43210/oauth/callback'] }),
  );
  await probe('register preflight, claude.ai', 'OPTIONS', '/register', {
    headers: {
      ...CLAUDE,
      'Access-Control-Request-Method': 'POST',
      'Access-Control-Request-Headers': 'content-type',
    },
  });
  const { text: registered } = await probe(
    'register, native',
    'POST',
    '/register',
    json({
      application_type: 'native',
      token_endpoint_auth_method: 'none',
      redirect_uris: [NATIVE],
      grant_types: ['authorization_code', 'refresh_token'],
      response_types: ['code'],
    }),
  );
  const client = (JSON.parse(registered) as { client_id: string }).client_id;
  // The recording's paths carry its own client ID; compare them without the query.
  const query = {
    client_id: client,
    response_type: 'code',
    redirect_uri: NATIVE,
    code_challenge: await pkceChallenge(VERIFIER),
    code_challenge_method: 'S256',
    state: 'client-state',
    resource: `${ORIGIN}/mcp`,
  };
  await probe('authorize, no parameters', 'GET', '/authorize');
  await probe(
    'authorize, plain PKCE',
    'GET',
    `/authorize?${new URLSearchParams({ ...query, code_challenge_method: 'plain' })}`,
  );
  await probe(
    'authorize, other resource',
    'GET',
    `/authorize?${new URLSearchParams({ ...query, resource: 'https://other.example/mcp' })}`,
  );
  await probe('authorize, foreign origin', 'GET', `/authorize?${new URLSearchParams(query)}`, {
    headers: EVIL,
  });
  const { response: authorized } = await probe(
    'authorize',
    'GET',
    `/authorize?${new URLSearchParams(query)}`,
  );
  const google = new URL(authorized.headers.get('location') as string);
  const state = google.searchParams.get('state');
  await probe('provider callback, no parameters', 'GET', '/oauth/callback');
  await probe(
    'provider callback, provider error',
    'GET',
    '/oauth/callback?error=access_denied&state=x',
  );
  const { response: called } = await probe(
    'provider callback',
    'GET',
    `/oauth/callback?code=google-code&state=${state}`,
  );
  const back = new URL(called.headers.get('location') as string);
  const code = back.searchParams.get('code') as string;
  await probe(
    'provider callback, replay',
    'GET',
    `/oauth/callback?code=google-code&state=${state}`,
  );

  await probe('token, empty', 'POST', '/token', form({}));
  await probe('token', 'GET', '/token');
  await probe('token preflight, claude.ai', 'OPTIONS', '/token', {
    headers: {
      ...CLAUDE,
      'Access-Control-Request-Method': 'POST',
      'Access-Control-Request-Headers': 'content-type',
    },
  });
  await probe(
    'token, foreign origin',
    'POST',
    '/token',
    form({ grant_type: 'authorization_code' }, EVIL),
  );
  await probe('token, unsupported grant', 'POST', '/token', form({ grant_type: 'password' }));
  const grant = { grant_type: 'authorization_code', code, client_id: client, redirect_uri: NATIVE };
  await probe(
    'token, wrong verifier',
    'POST',
    '/token',
    form({ ...grant, code_verifier: 'w'.repeat(43) }),
  );
  const { text: issued } = await probe(
    'token, authorization code',
    'POST',
    '/token',
    form({ ...grant, code_verifier: VERIFIER, resource: `${ORIGIN}/mcp` }),
  );
  const tokens = JSON.parse(issued) as { access_token: string; refresh_token: string };
  await probe('token, code replay', 'POST', '/token', form({ ...grant, code_verifier: VERIFIER }));
  const refresh = { grant_type: 'refresh_token', refresh_token: tokens.refresh_token };
  await probe(
    'token, refresh, unknown token',
    'POST',
    '/token',
    form({ ...refresh, refresh_token: 'not-a-token', client_id: client }),
  );
  await probe(
    'token, refresh, other resource',
    'POST',
    '/token',
    form({ ...refresh, resource: 'https://other.example/mcp' }),
  );
  await probe(
    'token, refresh, other client_id',
    'POST',
    '/token',
    form({ ...refresh, client_id: 'someone-else' }),
  );
  await probe('token, refresh, no client_id', 'POST', '/token', form(refresh));
  const { text: refreshed } = await probe(
    'token, refresh',
    'POST',
    '/token',
    form({ ...refresh, client_id: client }),
  );
  await probe(
    'token, refresh, claude.ai origin',
    'POST',
    '/token',
    form({ ...refresh, client_id: client }, CLAUDE),
  );

  await probe('revoke', 'POST', '/revoke', form({ token: tokens.access_token }));
  await probe('revoke', 'GET', '/revoke');

  const bearer = { Authorization: `Bearer ${tokens.access_token}` };
  await probe('mcp, issued token, 2025 initialize', 'POST', '/mcp', {
    headers: { ...mcp, ...bearer },
    body: JSON.stringify({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2025-11-25',
        capabilities: {},
        clientInfo: { name: 'matrix', version: '1' },
      },
    }),
  });
  await probe('mcp, issued token', 'GET', '/mcp', { headers: bearer });
  await probe('mcp, issued token', 'DELETE', '/mcp', { headers: bearer });
  return {
    matrix,
    documents,
    google,
    clientState: back.searchParams.get('state'),
    token: JSON.parse(issued) as Record<string, unknown>,
    refreshed: JSON.parse(refreshed) as Record<string, unknown>,
  };
}

const withoutQuery = (path: string) => path.split('?')[0];

test('every public route answers as before, except the listed changes', async () => {
  const { matrix } = await probeAll(productionApp());
  const recorded = before.matrix as Probe[];
  const keyOf = (entry: Probe) => `${entry.label} ${entry.method} ${withoutQuery(entry.path)}`;
  expect(matrix.map(keyOf)).toEqual(recorded.map(keyOf));

  const differences = matrix.flatMap((entry, index) => {
    const old = recorded[index] as Probe;
    const key = keyOf(old);
    const expected: Partial<Probe> = { ...old, ...CHANGED[key], path: entry.path };
    if (expected.error === undefined) delete expected.error;
    return JSON.stringify(entry) === JSON.stringify(expected) ? [] : [{ key, entry, expected }];
  });
  expect(differences).toEqual([]);
});

test('every listed change is still a change', () => {
  const recorded = before.matrix as Probe[];
  for (const [key, change] of Object.entries(CHANGED)) {
    const old = recorded.find(
      (entry) => `${entry.label} ${entry.method} ${withoutQuery(entry.path)}` === key,
    );
    expect(old).toBeDefined();
    expect({ ...old, ...change }).not.toEqual(old);
  }
});

test('the discovery documents are field for field what production published', async () => {
  const { documents } = await probeAll(productionApp());
  const { client_id_metadata_document_supported, ...published } = documents[
    '/.well-known/oauth-authorization-server'
  ] as Record<string, unknown>;
  // Added: the one field the template's proxy publishes that the checkpoint left out. Its
  // value is the default, so clients behave the same: they register through /register.
  expect(client_id_metadata_document_supported).toBe(false);
  expect({ ...documents, '/.well-known/oauth-authorization-server': published }).toEqual(
    before.documents,
  );
});

test('Google is asked for the same sign-in as before, with the same callback', async () => {
  const { google, clientState } = await probeAll(productionApp());
  expect(`${google.origin}${google.pathname}`).toBe(before.googleAuthorize.url);
  expect({ ...Object.fromEntries(google.searchParams), state: '<handle>' }).toEqual(
    before.googleAuthorize.query,
  );
  expect(google.searchParams.get('redirect_uri')).toBe(
    'https://google-calendar.example.workers.dev/oauth/callback',
  );
  expect(clientState).toBe(before.clientRedirect.state);
});

test('the token and refresh answers have the same shape and values', async () => {
  const { token, refreshed } = await probeAll(productionApp());
  const opaque = (answer: Record<string, unknown>) => ({
    ...answer,
    access_token: '<opaque>',
    refresh_token: '<opaque>',
  });
  expect(opaque(token)).toEqual(before.tokenResponse);
  expect(opaque(refreshed) as unknown).toEqual({
    ...before.refreshResponse,
    // Seconds until Google's token expires: 3599 or 3600, depending on the clock.
    expires_in: refreshed.expires_in,
  });
  expect([3599, 3600]).toContain(refreshed.expires_in as number);
});
