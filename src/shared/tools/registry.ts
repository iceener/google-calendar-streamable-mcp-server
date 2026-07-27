import type {
  McpServer,
  ServerContext,
  StandardSchemaWithJSON,
} from '@modelcontextprotocol/server';
import type { ZodObject, ZodType } from 'zod/v4';
import { checkAvailabilityTool } from './check-availability.js';
import { createEventTool } from './create-event.js';
import { deleteEventTool } from './delete-event.js';
import { listCalendarsTool } from './list-calendars.js';
import { respondToEventTool } from './respond-to-event.js';
import { searchEventsTool } from './search-events.js';
import type { SharedToolDefinition } from './types.js';
import { updateEventTool } from './update-event.js';

export interface GoogleCalendarToolDependencies {
  providerAccessToken?: string;
}

export const toolNames = [
  'list_calendars',
  'search_events',
  'check_availability',
  'create_event',
  'update_event',
  'delete_event',
  'respond_to_event',
] as const;

function registerTool<TInput extends ZodObject, TOutput extends ZodType | undefined>(
  server: McpServer,
  tool: SharedToolDefinition<TInput, TOutput>,
  dependencies: GoogleCalendarToolDependencies,
): void {
  const handler = async (args: unknown, ctx: ServerContext) =>
    tool.handler(tool.inputSchema.parse(args), {
      signal: ctx.mcpReq.signal,
      providerAccessToken: dependencies.providerAccessToken,
    });
  const baseConfig = {
    title: tool.title,
    description: tool.description,
    inputSchema: tool.inputSchema as StandardSchemaWithJSON,
    ...(tool.annotations ? { annotations: tool.annotations } : {}),
  };
  if (tool.outputSchema) {
    server.registerTool(
      tool.name,
      {
        ...baseConfig,
        outputSchema: tool.outputSchema as StandardSchemaWithJSON,
      },
      handler,
    );
  } else {
    server.registerTool(tool.name, baseConfig, handler);
  }
}

export function registerTools(
  server: McpServer,
  dependencies: GoogleCalendarToolDependencies,
): void {
  registerTool(server, listCalendarsTool, dependencies);
  registerTool(server, searchEventsTool, dependencies);
  registerTool(server, checkAvailabilityTool, dependencies);
  registerTool(server, createEventTool, dependencies);
  registerTool(server, updateEventTool, dependencies);
  registerTool(server, deleteEventTool, dependencies);
  registerTool(server, respondToEventTool, dependencies);
}

export type { SharedToolDefinition, ToolContext, ToolResult } from './types.js';
export { defineTool } from './types.js';
