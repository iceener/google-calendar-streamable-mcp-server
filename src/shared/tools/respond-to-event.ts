/**
 * Respond to Event tool - accept, decline, or tentatively accept an event invitation.
 */

import { z } from 'zod/v4';
import { toolsMetadata } from '../../config/metadata.js';
import { RespondToEventOutputSchema } from '../../schemas/outputs.js';
import {
  CalendarApiError,
  type CalendarEvent,
  GoogleCalendarClient,
} from '../../services/google-calendar.js';
import { defineTool, type ToolResult } from './types.js';

const InputSchema = z.object({
  eventId: z.string().describe('Event ID to respond to'),
  calendarId: z.string().optional().describe('Calendar ID (defaults to "primary")'),
  response: z
    .enum(['accepted', 'declined', 'tentative'])
    .describe(
      'Your response: "accepted" (yes), "declined" (no), or "tentative" (maybe)',
    ),
  sendUpdates: z.enum(['all', 'externalOnly', 'none']).optional().default('all'),
});

const RESPONSE_LABELS: Record<string, string> = {
  accepted: 'accepted',
  declined: 'declined',
  tentative: 'marked as maybe',
};

function formatResponse(event: CalendarEvent, response: string): string {
  const lines: string[] = [];

  const title = event.summary || '(no title)';
  const start = event.start?.dateTime || event.start?.date || 'no date';
  const responseLabel = RESPONSE_LABELS[response] || response;

  if (event.htmlLink) {
    lines.push(`✓ You ${responseLabel}: [${title}](${event.htmlLink})`);
  } else {
    lines.push(`✓ You ${responseLabel}: ${title}`);
  }

  lines.push(`  when: ${start}`);

  if (event.location) {
    lines.push(`  location: ${event.location}`);
  }

  if (event.hangoutLink) {
    lines.push(`  meet: ${event.hangoutLink}`);
  }

  return lines.join('\n');
}

export const respondToEventTool = defineTool({
  name: toolsMetadata.respond_to_event.name,
  title: toolsMetadata.respond_to_event.title,
  description: toolsMetadata.respond_to_event.description,
  inputSchema: InputSchema,
  outputSchema: RespondToEventOutputSchema,
  annotations: {
    readOnlyHint: false,
    destructiveHint: false,
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

    const attemptRespond = (effectiveCalendarId: string) =>
      client.respondToEvent({
        calendarId: effectiveCalendarId,
        eventId: args.eventId,
        response: args.response,
        sendUpdates: args.sendUpdates,
      });

    try {
      let result: CalendarEvent;

      try {
        result = await attemptRespond(calendarId);
      } catch (error) {
        if (
          error instanceof CalendarApiError &&
          error.isNotFound &&
          calendarId !== 'primary'
        ) {
          result = await attemptRespond('primary');
        } else {
          throw error;
        }
      }

      const text = formatResponse(result, args.response);

      return {
        content: [{ type: 'text', text }],
        structuredContent: {
          ok: true,
          action: 'respond_to_event',
          response: args.response,
          event: result,
        },
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          {
            type: 'text',
            text: `Failed to respond to event: ${(error as Error).message}`,
          },
        ],
      };
    }
  },
});
