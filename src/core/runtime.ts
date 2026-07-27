import { createMcpHandler, type McpHttpHandler } from '@modelcontextprotocol/server';
import type { AppConfig } from '../config/env.js';
import { sharedLogger as logger } from '../shared/utils/logger.js';
import { createMcpServer } from './mcp.js';

export type McpRuntime = McpHttpHandler;

export function createMcpRuntime(config: AppConfig): McpRuntime {
  return createMcpHandler((context) => createMcpServer(config, context), {
    legacy: config.MCP_LEGACY_MODE,
    responseMode: 'auto',
    onerror(error) {
      logger.error('mcp', {
        message: 'MCP request failed',
        error: error.message,
      });
    },
  });
}
