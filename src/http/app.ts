import { type AuthInfo, oauthMetadataResponse } from '@modelcontextprotocol/server';
import { Hono } from 'hono';
import { buildOAuthRoutes } from '../adapters/http-hono/routes.oauth.js';
import type { AppConfig } from '../config/env.js';
import { SERVER_ICON_SVG } from '../config/metadata.js';
import { createMcpRuntime } from '../core/runtime.js';
import type { TokenStore } from '../shared/storage/interface.js';
import { sharedLogger as logger } from '../shared/utils/logger.js';
import { createAuthServices } from './auth.js';
import { boundedMcpRequest } from './body.js';
import {
  corsPreflightResponse,
  requestSecurityResponse,
  withCors,
} from './security.js';

export interface HttpRuntimeOptions {
  runtimeName: string;
  tokenStore: TokenStore;
  mountOAuthProxy?: boolean;
}

export interface HttpRuntime {
  fetch(request: Request): Promise<Response>;
  close(): Promise<void>;
}

export function buildHttpApp(
  config: AppConfig,
  options: HttpRuntimeOptions,
): HttpRuntime {
  logger.setLevel(config.LOG_LEVEL);
  const mcp = createMcpRuntime(config);
  const auth = createAuthServices(config, options.tokenStore);
  const mcpPath = config.MCP_PUBLIC_URL.pathname;
  const app = new Hono();

  app.use('*', async (context, next) => {
    const request = context.req.raw;
    const rejected = requestSecurityResponse(request, config);
    if (rejected) return rejected;
    if (auth) {
      const metadata = oauthMetadataResponse(request, auth.metadata);
      if (metadata) return metadata;
    }
    await next();
  });

  app.get('/health', (context) =>
    context.json({
      status: 'ok',
      runtime: options.runtimeName,
      protocol: '2026-07-28',
      legacyMode: config.MCP_LEGACY_MODE,
      authEnabled: config.AUTH_ENABLED,
      timestamp: new Date().toISOString(),
    }),
  );

  app.get(
    '/icon.svg',
    () =>
      new Response(SERVER_ICON_SVG, {
        headers: {
          'Content-Type': 'image/svg+xml; charset=utf-8',
          'Cache-Control': 'public, max-age=86400',
          'Content-Security-Policy': "default-src 'none'; style-src 'none'; sandbox",
        },
      }),
  );

  if (options.mountOAuthProxy) {
    app.route('/', buildOAuthRoutes(options.tokenStore, config));
  }

  app.options(mcpPath, (context) => corsPreflightResponse(context.req.raw));
  app.all(mcpPath, async (context) => {
    const request = context.req.raw;
    let authInfo: AuthInfo | undefined;
    if (auth) {
      const result = await auth.gate(request);
      if (result instanceof Response) return withCors(request, result);
      authInfo = result;
    }
    const bounded = await boundedMcpRequest(request, config.MCP_MAX_REQUEST_BYTES);
    if (bounded.rejection) return withCors(request, bounded.rejection);
    const response = await mcp.fetch(
      bounded.request,
      authInfo ? { authInfo } : undefined,
    );
    return withCors(request, response);
  });

  app.notFound((context) => context.text('Not Found', 404));
  return { fetch: async (request) => app.fetch(request), close: mcp.close };
}
