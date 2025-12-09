/**
 * Delete Event tool - remove events from calendar.
 */

import { z } from 'zod';
import { toolsMetadata } from '../../config/metadata.js';
import { GoogleCalendarClient } from '../../services/google-calendar.js';
import { defineTool, type ToolResult } from './types.js';

const InputSchema = z.object({
  eventId: z.string().describe('Event ID to delete'),
  calendarId: z.string().optional().describe('Calendar ID (defaults to "primary")'),
  sendUpdates: z
    .enum(['all', 'externalOnly', 'none'])
    .optional()
    .default('none')
    .describe('Notify attendees about the cancellation'),
});

export const deleteEventTool = defineTool({
  name: toolsMetadata.delete_event.name,
  title: toolsMetadata.delete_event.title,
  description: toolsMetadata.delete_event.description,
  inputSchema: InputSchema,
  annotations: {
    readOnlyHint: false,
    destructiveHint: true,
  },

  handler: async (args, context): Promise<ToolResult> => {
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
      await client.deleteEvent({
        eventId: args.eventId,
        calendarId: args.calendarId,
        sendUpdates: args.sendUpdates,
      });

      const calendarId = args.calendarId || 'primary';
      const notified =
        args.sendUpdates === 'all'
          ? 'All attendees were notified.'
          : args.sendUpdates === 'externalOnly'
            ? 'External attendees were notified.'
            : 'No notifications sent.';

      return {
        content: [
          {
            type: 'text',
            text: `✓ Event deleted successfully.\n  eventId: ${args.eventId}\n  calendar: ${calendarId}\n  ${notified}\n\nNext: Use 'search_events' to verify deletion.`,
          },
        ],
        structuredContent: { success: true, eventId: args.eventId, calendarId },
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          { type: 'text', text: `Failed to delete event: ${(error as Error).message}` },
        ],
      };
    }
  },
});









