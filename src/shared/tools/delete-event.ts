/**
 * Delete Event tool - remove events from calendar.
 */

import { z } from 'zod/v4';
import { toolsMetadata } from '../../config/metadata.js';
import { DeleteEventOutputSchema } from '../../schemas/outputs.js';
import {
  CalendarApiError,
  GoogleCalendarClient,
} from '../../services/google-calendar.js';
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
  outputSchema: DeleteEventOutputSchema,
  annotations: {
    readOnlyHint: false,
    destructiveHint: true,
  },

  handler: async (args, context): Promise<ToolResult> => {
    const token = context.providerAccessToken;

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

    const client = new GoogleCalendarClient(token, context.signal);
    const calendarId = args.calendarId || 'primary';

    const attemptDelete = async (effectiveCalendarId: string) => {
      await client.deleteEvent({
        eventId: args.eventId,
        calendarId: effectiveCalendarId,
        sendUpdates: args.sendUpdates,
      });
      return effectiveCalendarId;
    };

    try {
      let usedCalendarId: string;

      try {
        usedCalendarId = await attemptDelete(calendarId);
      } catch (error) {
        // On 404, retry with 'primary' if we used a specific calendarId (email form).
        // Google Calendar API can return events via listEvents using the email alias
        // but require 'primary' for mutations, or vice versa.
        if (
          error instanceof CalendarApiError &&
          error.isNotFound &&
          calendarId !== 'primary'
        ) {
          usedCalendarId = await attemptDelete('primary');
        } else {
          throw error;
        }
      }

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
            text: `✓ Event deleted successfully.\n  eventId: ${args.eventId}\n  calendar: ${usedCalendarId}\n  ${notified}\n\nNext: Use 'search_events' to verify deletion.`,
          },
        ],
        structuredContent: {
          success: true,
          eventId: args.eventId,
          calendarId: usedCalendarId,
        },
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
