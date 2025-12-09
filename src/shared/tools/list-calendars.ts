/**
 * List Calendars tool - discover available calendars and their IDs.
 */

import { z } from 'zod';
import { toolsMetadata } from '../../config/metadata.js';
import { GoogleCalendarClient } from '../../services/google-calendar.js';
import { defineTool, type ToolResult } from './types.js';

const InputSchema = z.object({});

export const listCalendarsTool = defineTool({
  name: toolsMetadata.list_calendars.name,
  title: toolsMetadata.list_calendars.title,
  description: toolsMetadata.list_calendars.description,
  inputSchema: InputSchema,
  annotations: {
    readOnlyHint: true,
    destructiveHint: false,
  },

  handler: async (_args, context): Promise<ToolResult> => {
    const token = context.providerToken;

    if (!token) {
      return {
        isError: true,
        content: [
          {
            type: 'text',
            text: 'Authentication required. Please authenticate with Google Calendar.',
          },
        ],
      };
    }

    const client = new GoogleCalendarClient(token);

    try {
      const result = await client.listCalendars();

      // Format for LLM consumption
      const lines: string[] = [];
      lines.push(`Found ${result.items.length} calendar(s):\n`);

      for (const cal of result.items) {
        const primary = cal.primary ? ' (primary)' : '';
        const access = cal.accessRole ? ` [${cal.accessRole}]` : '';
        lines.push(`- ${cal.summary}${primary}${access}`);
        lines.push(`  id: ${cal.id}`);
        if (cal.timeZone) lines.push(`  timezone: ${cal.timeZone}`);
        if (cal.description) lines.push(`  description: ${cal.description}`);
        lines.push('');
      }

      lines.push("Next: Use calendarId in 'search_events', 'create_event', etc.");

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
        structuredContent: { items: result.items },
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          {
            type: 'text',
            text: `Failed to list calendars: ${(error as Error).message}`,
          },
        ],
      };
    }
  },
});









