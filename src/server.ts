/**
 * The project's half of the template. `platform/` relies on exactly these exports:
 * `serverInfo`, `SERVER_ICON_PATH`, `SERVER_ICON_SVG`, `Runtime`, `Deps`, `createDeps`,
 * `createServer`, `createVerifier`, `oauthMetadata` and `routes`. Change what they contain,
 * but keep their names and shapes.
 */
import {
  type CacheHint,
  McpServer,
  type McpServerFactory,
  type OAuthMetadata,
  type OAuthTokenVerifier,
  type ServerCapabilities,
} from '@modelcontextprotocol/server';
import type { Hono } from 'hono';
import type { OAuthAuthority } from './oauth/authority';
import { createOAuthProxy, type OAuthProxy } from './oauth/proxy';
import type { TokenStore } from './oauth/token-store';
import { type Config, ConfigError, type OAuthConfig } from './platform/config';
import type { Logger } from './platform/logger';
import { prompts } from './prompts';
import { resources } from './resources';
import {
  type CalendarClientFactory,
  createCalendarClientFactory,
} from './services/google-calendar';
import { googleOAuth } from './services/google-oauth';
import { tools } from './tools';

/** Who this server is. */
export const serverInfo = {
  name: 'google-calendar-mcp',
  title: 'Google Calendar',
  version: '1.1.0',
  description: 'Manage Google Calendar events, calendars, invitations, and availability.',
};

/** Sent to clients on connect. Many hosts add it to the model's system prompt. */
const instructions =
  'Manage Google Calendar events. Search before mutating and verify changes afterward.';

export const SERVER_ICON_PATH = '/icon.svg';
export const SERVER_ICON_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" role="img" aria-label="Google Calendar"><rect width="64" height="64" rx="12" fill="#fff"/><path d="M14 20h36v32H14z" fill="#4285f4"/><path d="M14 20h36v10H14z" fill="#ea4335"/><path d="M23 12v14m18-14v14" stroke="#34a853" stroke-width="5"/><path d="M24 36h16v10H24z" fill="#fbbc05"/></svg>`;

/**
 * Where the OAuth proxy keeps its state, handed in by the entry points: Workers KV and the
 * `NativeOAuthAuthority` Durable Object from `src/worker.ts`, files from `src/bun.ts`.
 */
export interface Runtime {
  tokens: TokenStore;
  authority: OAuthAuthority;
}

/** Everything tools, resources and prompts may use. Built once, shared by every request. */
export interface Deps {
  config: Config;
  logger: Logger;
  /** A Calendar client acting with one user's Google access token. */
  calendar: CalendarClientFactory;
  /** This server as the OAuth authorization server, in front of Google (docs/oauth.md). */
  oauth: OAuthProxy;
}

export function createDeps(config: Config, logger: Logger, runtime: Runtime): Deps {
  const { settings } = config;
  if (config.environment === 'production' && config.auth.mode === 'oauth') {
    // Production never runs the proxy without its Google client, or with tokens stored in clear.
    const missing = ['PROVIDER_CLIENT_ID', 'PROVIDER_CLIENT_SECRET', 'TOKENS_ENC_KEY'] as const;
    const problems = missing
      .filter((name) => !settings[name])
      .map((name) => `${name} is required in production`);
    if (problems.length > 0) throw new ConfigError(problems);
  }
  return {
    config,
    logger,
    calendar: createCalendarClientFactory({ logger }),
    oauth: createOAuthProxy({
      config,
      logger,
      provider: googleOAuth,
      credentials: {
        clientId: settings.PROVIDER_CLIENT_ID,
        clientSecret: settings.PROVIDER_CLIENT_SECRET,
      },
      redirectAllowlist: settings.PROXY_REDIRECT_ALLOWLIST,
      encryptionKey: settings.TOKENS_ENC_KEY,
      tokens: runtime.tokens,
      authority: runtime.authority,
      // Records have never carried a client binding here, and clients refresh without
      // `client_id`. The old discovery locations were never served. docs/oauth.md.
      bindGrants: false,
      discoveryAliases: false,
    }),
  };
}

/**
 * How bearer tokens are checked when `AUTH_MODE=oauth`: they are the opaque tokens the proxy
 * issued. The verifier looks them up, refreshes the Google token when it is about to expire,
 * and puts it in `authInfo.extra` for the tools.
 */
export function createVerifier(_oauth: OAuthConfig, deps: Deps): OAuthTokenVerifier {
  return deps.oauth.verifier;
}

/**
 * The authorization server metadata (RFC 8414) published at
 * /.well-known/oauth-authorization-server: this server's own endpoints.
 */
export function oauthMetadata(oauth: OAuthConfig, deps: Deps): OAuthMetadata {
  return deps.oauth.metadata(oauth);
}

/** The proxy's endpoints and Google's callback, behind the platform's Host and Origin checks. */
export function routes(app: Hono, deps: Deps): void {
  if (deps.config.auth.mode === 'oauth') deps.oauth.mount(app, deps.config.auth, serverInfo.title);
}

/** Lists only change on deploy, and are the same for every caller. */
const LIST_CACHE: CacheHint = { ttlMs: 60_000, cacheScope: 'public' };

/**
 * The SDK calls this factory once per HTTP request and serves that request with the fresh
 * `McpServer` it returns. Keep it cheap and free of I/O: shared clients live in `deps`.
 */
export function createServer(deps: Deps): McpServerFactory {
  const icons = [
    {
      src: new URL(SERVER_ICON_PATH, deps.config.publicUrl).href,
      mimeType: 'image/svg+xml',
      sizes: ['any'],
      theme: 'light' as const,
    },
  ];
  // Lists change only on deploy, and each request has its own server instance, so there is
  // nothing to notify about. Only what the server has is advertised.
  const capabilities: ServerCapabilities = {
    tools: { listChanged: false },
    ...(prompts.length > 0 && { prompts: { listChanged: false } }),
    ...(resources.length > 0 && { resources: { listChanged: false, subscribe: false } }),
  };

  return () => {
    const server = new McpServer(
      { ...serverInfo, icons },
      {
        instructions,
        capabilities,
        cacheHints: {
          'server/discover': LIST_CACHE,
          'tools/list': LIST_CACHE,
          'prompts/list': LIST_CACHE,
          'resources/list': LIST_CACHE,
          'resources/templates/list': LIST_CACHE,
        },
        // search_events reads at most 250 events; no tool takes more elements than this.
        maxToolInputElements: 1_000,
      },
    );

    for (const primitive of [...tools, ...resources, ...prompts]) {
      primitive.register(server, deps);
    }
    return server;
  };
}
