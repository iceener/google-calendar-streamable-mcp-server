import { expect } from 'bun:test';
import { MemoryOAuthAuthority, type OAuthAuthority } from '../../src/oauth/authority';
import { opaqueToken, pkceChallenge } from '../../src/oauth/crypto';
import type { OAuthProxy, OAuthProxyOptions } from '../../src/oauth/proxy';
import { MemoryTokenStore, type TokenStore } from '../../src/oauth/token-store';
import type { App } from '../../src/platform/app';
import { createServer } from '../../src/server';
import { createCalendarClientFactory } from '../../src/services/google-calendar';
import { CALENDAR_SCOPES } from '../../src/services/google-oauth';
import {
  type FetchRoute,
  fakeFetch,
  type LogEntry,
  memoryLogger,
  testApp,
  testConfig,
  testDeps,
  testProxy,
} from '../helpers';

/** The test server's origin. It is its own authorization server. */
export const ORIGIN = 'http://127.0.0.1:3000';
export const HOST = '127.0.0.1:3000';
export const GOOGLE_TOKEN_URL = 'https://oauth2.googleapis.com/token';
export const CALLBACK = `${ORIGIN}/oauth/callback`;
export const VERIFIER = 'v'.repeat(43);
export const SCOPE = CALENDAR_SCOPES.join(' ');

export const PROXY_ENV: Record<string, string> = {
  AUTH_MODE: 'oauth',
  OAUTH_ISSUER_URL: ORIGIN,
  OAUTH_AUTHORIZATION_URL: `${ORIGIN}/authorize`,
  OAUTH_TOKEN_URL: `${ORIGIN}/token`,
  OAUTH_REGISTRATION_URL: `${ORIGIN}/register`,
  OAUTH_SCOPES: SCOPE,
  MCP_ALLOWED_ORIGIN_HOSTNAMES: '127.0.0.1,client.example',
};

/** Google's token endpoint, answering every code and refresh token with fresh tokens. */
export function googleTokens(
  answer: (form: URLSearchParams) => Response | Promise<Response> = (form) =>
    Response.json({
      access_token: form.get('grant_type') === 'refresh_token' ? 'google-fresh' : 'google-access',
      refresh_token: 'google-refresh',
      expires_in: 3600,
      scope: SCOPE,
    }),
) {
  const forms: URLSearchParams[] = [];
  const route: FetchRoute = async (request) => {
    if (request.url !== GOOGLE_TOKEN_URL) return undefined;
    expect(request.method).toBe('POST');
    expect(request.headers.get('authorization')).toBe(
      `Basic ${btoa('google-client:google-secret')}`,
    );
    expect(request.headers.get('content-type')).toBe('application/x-www-form-urlencoded');
    const form = new URLSearchParams(await request.text());
    forms.push(form);
    return answer(form);
  };
  return Object.assign(fakeFetch(route), { forms });
}

export interface ProxyFixture {
  app: App;
  proxy: OAuthProxy;
  tokens: TokenStore;
  authority: OAuthAuthority;
  google: ReturnType<typeof googleTokens>;
  /** Every request the Calendar tools made. */
  calendar: ReturnType<typeof fakeFetch>;
  logs: LogEntry[];
  /** GET without a body, POST with one: JSON, a form, or a raw string. */
  request(path: string, body?: unknown, headers?: Record<string, string>): Promise<Response>;
}

/** The whole server in OAuth mode, its state in memory, Google answered by `google`. */
export function proxyFixture(
  options: Partial<
    Pick<
      OAuthProxyOptions,
      | 'tokens'
      | 'authority'
      | 'redirectAllowlist'
      | 'encryptionKey'
      | 'credentials'
      | 'provider'
      | 'bindGrants'
      | 'discoveryAliases'
    >
  > & {
    google?: ReturnType<typeof googleTokens>;
    /** Answers the Google Calendar API. By default nothing does. */
    calendar?: FetchRoute;
    env?: Record<string, string>;
  } = {},
): ProxyFixture {
  const config = testConfig({ ...PROXY_ENV, ...options.env });
  const logs: LogEntry[] = [];
  const logger = memoryLogger(logs);
  const google = options.google ?? googleTokens();
  const tokens = options.tokens ?? new MemoryTokenStore();
  const authority = options.authority ?? new MemoryOAuthAuthority();
  const proxy = testProxy({
    config,
    logger,
    tokens,
    authority,
    fetch: google,
    ...(options.redirectAllowlist && { redirectAllowlist: options.redirectAllowlist }),
    ...(options.encryptionKey && { encryptionKey: options.encryptionKey }),
    ...(options.credentials && { credentials: options.credentials }),
    ...(options.provider && { provider: options.provider }),
    ...(options.bindGrants !== undefined && { bindGrants: options.bindGrants }),
    ...(options.discoveryAliases !== undefined && { discoveryAliases: options.discoveryAliases }),
  });
  const calendar = fakeFetch(options.calendar);
  const app = testApp(config, {
    deps: testDeps({
      config,
      logger,
      oauth: proxy,
      calendar: createCalendarClientFactory({ logger, fetch: calendar }),
    }),
    server: createServer,
  });

  const request = (path: string, body?: unknown, headers: Record<string, string> = {}) => {
    const form = body instanceof URLSearchParams;
    return app.fetch(
      new Request(new URL(path, ORIGIN).href, {
        method: body === undefined ? 'GET' : 'POST',
        headers: {
          Host: HOST,
          ...(body !== undefined && {
            'Content-Type': form ? 'application/x-www-form-urlencoded' : 'application/json',
          }),
          ...headers,
        },
        ...(body !== undefined && {
          body: form || typeof body === 'string' ? String(body) : JSON.stringify(body),
        }),
      }),
    );
  };
  return { app, proxy, tokens, authority, google, calendar, logs, request };
}

/** A loopback redirect URI on a port the OS really assigned, as native clients use. */
export function loopback(path = '/oauth/callback'): string {
  const server = Bun.serve({ hostname: '127.0.0.1', port: 0, fetch: () => new Response('ok') });
  const uri = `http://127.0.0.1:${server.port}${path}`;
  server.stop(true);
  return uri;
}

export interface Registered {
  client_id: string;
  redirect_uris: string[];
}

export async function registerNative(f: ProxyFixture, redirectUri: string): Promise<Registered> {
  const response = await f.request('/register', {
    application_type: 'native',
    token_endpoint_auth_method: 'none',
    redirect_uris: [redirectUri],
    grant_types: ['authorization_code', 'refresh_token'],
    response_types: ['code'],
  });
  expect(response.status).toBe(201);
  expect(response.headers.get('cache-control')).toBe('no-store');
  const client = (await response.json()) as Registered;
  expect(client).not.toHaveProperty('registration_access_token');
  expect(client).not.toHaveProperty('registration_client_uri');
  return client;
}

export async function authorize(
  f: ProxyFixture,
  clientId: string,
  redirectUri: string,
  fields: Record<string, string> = {},
): Promise<Response> {
  return f.request(
    `/authorize?${new URLSearchParams({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: redirectUri,
      code_challenge: await pkceChallenge(VERIFIER),
      code_challenge_method: 'S256',
      state: 'original opaque state & / +',
      ...fields,
    })}`,
  );
}

/** Authorize, then come back from Google: the client's code and the transaction handle. */
export async function codeFor(
  f: ProxyFixture,
  clientId: string,
  redirectUri: string,
): Promise<{ code: string; state: string }> {
  const response = await authorize(f, clientId, redirectUri);
  expect(response.status).toBe(302);
  const google = new URL(response.headers.get('location') as string);
  expect(google.searchParams.get('client_id')).toBe('google-client');
  const state = google.searchParams.get('state') as string;
  expect(state).toMatch(/^[A-Za-z0-9_-]{32}\.[A-Za-z0-9_-]{32}$/);
  // Google never sees the client's redirect URI or state.
  expect(google.href).not.toContain(encodeURIComponent(redirectUri));
  expect(google.href).not.toContain('original');
  const callback = await f.request(
    `/oauth/callback?${new URLSearchParams({ code: 'google-code', state })}`,
  );
  expect(callback.status).toBe(302);
  const target = new URL(callback.headers.get('location') as string);
  expect(target.origin + target.pathname).toBe(redirectUri);
  expect(target.searchParams.get('state')).toBe('original opaque state & / +');
  return { code: target.searchParams.get('code') as string, state };
}

export function redeem(
  f: ProxyFixture,
  clientId: string,
  redirectUri: string,
  code: string,
  overrides: Record<string, string> = {},
): Promise<Response> {
  return f.request(
    '/token',
    new URLSearchParams({
      grant_type: 'authorization_code',
      code,
      client_id: clientId,
      redirect_uri: redirectUri,
      code_verifier: VERIFIER,
      ...overrides,
    }),
  );
}

export interface TokenPair {
  access_token: string;
  refresh_token: string;
  scope: string;
}

/** A unique handle in the shape the authority routes on. */
export const handleFor = (clientId: string) => `${clientId}.${opaqueToken(24)}`;
