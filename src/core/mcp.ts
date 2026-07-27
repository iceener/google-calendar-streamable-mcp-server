import {
  type McpRequestContext,
  McpServer,
  type ServerCapabilities,
} from '@modelcontextprotocol/server';
import type { AppConfig } from '../config/env.js';
import { serverImplementation } from '../config/metadata.js';
import { providerAccessTokenFromExtra } from '../shared/auth/provider-token-verifier.js';
import { registerTools } from '../shared/tools/registry.js';

function capabilitiesFor(context: McpRequestContext): ServerCapabilities {
  return { tools: { listChanged: context.era === 'modern' } };
}

export function createMcpServer(
  config: AppConfig,
  context: McpRequestContext,
): McpServer {
  const server = new McpServer(serverImplementation(config), {
    instructions: config.MCP_INSTRUCTIONS,
    capabilities: capabilitiesFor(context),
    cacheHints: {
      'server/discover': { ttlMs: 60_000, cacheScope: 'private' },
      'tools/list': { ttlMs: 60_000, cacheScope: 'private' },
    },
  });
  registerTools(server, {
    providerAccessToken: providerAccessTokenFromExtra(context.authInfo?.extra),
  });
  return server;
}
