/**
 * Create Event tool - create events using natural language or structured input.
 */

import { z } from 'zod/v4';
import { toolsMetadata } from '../../config/metadata.js';
import { CalendarEventOutputSchema } from '../../schemas/outputs.js';
import {
  type CalendarEvent,
  GoogleCalendarClient,
} from '../../services/google-calendar.js';
import { defineTool, type ToolResult } from './types.js';

const ReminderOverrideSchema = z.object({
  method: z.enum(['popup', 'email']),
  minutes: z.number().int().min(0).max(40320),
});

const RemindersSchema = z.object({
  useDefault: z.boolean(),
  overrides: z.array(ReminderOverrideSchema).optional(),
});

const InputSchema = z.object({
  // Natural language mode
  text: z
    .string()
    .optional()
    .describe(
      'Natural language event description (e.g., "Lunch with Anna tomorrow at noon")',
    ),

  // Structured mode
  summary: z.string().optional().describe('Event title'),
  start: z
    .string()
    .optional()
    .describe('Start time (ISO 8601 datetime or YYYY-MM-DD for all-day)'),
  end: z
    .string()
    .optional()
    .describe('End time (ISO 8601 datetime or YYYY-MM-DD for all-day)'),
  description: z.string().optional().describe('Event description'),
  location: z.string().optional().describe('Event location'),
  attendees: z
    .array(z.string().email())
    .optional()
    .describe('List of attendee email addresses'),

  // Shared options
  calendarId: z.string().optional().describe('Calendar ID (defaults to "primary")'),
  addGoogleMeet: z
    .boolean()
    .optional()
    .default(false)
    .describe('Auto-create Google Meet link'),
  recurrence: z
    .array(z.string())
    .optional()
    .describe('RRULE array for recurring events'),
  reminders: RemindersSchema.optional().describe('Reminder settings'),
  visibility: z.enum(['default', 'public', 'private', 'confidential']).optional(),
  colorId: z.string().optional().describe('Color ID (1-11)'),
  sendUpdates: z.enum(['all', 'externalOnly', 'none']).optional().default('none'),
  timeZone: z.string().optional().describe('Time zone for the event'),
});

function isAllDayDate(str: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$/.test(str);
}

function formatCreatedEvent(event: CalendarEvent): string {
  const lines: string[] = [];

  const title = event.summary || '(no title)';
  const start = event.start?.dateTime || event.start?.date || 'no date';

  if (event.htmlLink) {
    lines.push(`✓ Event created: [${title}](${event.htmlLink})`);
  } else {
    lines.push(`✓ Event created: ${title}`);
  }

  lines.push(`  id: ${event.id}`);
  lines.push(`  when: ${start}`);

  if (event.location) {
    lines.push(`  location: ${event.location}`);
  }

  if (event.hangoutLink) {
    lines.push(`  meet: ${event.hangoutLink}`);
  }

  if (event.attendees && event.attendees.length > 0) {
    const attendeeList = event.attendees.map((a) => a.email).join(', ');
    lines.push(`  attendees: ${attendeeList}`);
  }

  return lines.join('\n');
}

export const createEventTool = defineTool({
  name: toolsMetadata.create_event.name,
  title: toolsMetadata.create_event.title,
  description: toolsMetadata.create_event.description,
  inputSchema: InputSchema,
  outputSchema: CalendarEventOutputSchema,
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

    try {
      let result: CalendarEvent;

      // Determine mode: natural language (quickAdd) vs structured
      const useQuickAdd = args.text && !args.summary;

      if (useQuickAdd) {
        // Mode A: Natural language
        result = await client.quickAdd({
          calendarId: args.calendarId,
          text: args.text ?? '',
          sendUpdates: args.sendUpdates,
        });
      } else {
        // Mode B: Structured
        if (!args.summary) {
          return {
            isError: true,
            content: [
              {
                type: 'text',
                text: "Either 'text' (for natural language) or 'summary' (for structured) is required.",
              },
            ],
          };
        }

        if (!args.start || !args.end) {
          return {
            isError: true,
            content: [
              {
                type: 'text',
                text: "'start' and 'end' are required for structured event creation.",
              },
            ],
          };
        }

        // Determine if all-day event
        const isAllDay = isAllDayDate(args.start) && isAllDayDate(args.end);

        result = await client.createEvent({
          calendarId: args.calendarId,
          summary: args.summary,
          description: args.description,
          start: isAllDay
            ? { date: args.start }
            : { dateTime: args.start, timeZone: args.timeZone },
          end: isAllDay
            ? { date: args.end }
            : { dateTime: args.end, timeZone: args.timeZone },
          location: args.location,
          attendees: args.attendees,
          addGoogleMeet: args.addGoogleMeet,
          recurrence: args.recurrence,
          reminders: args.reminders,
          visibility: args.visibility,
          colorId: args.colorId,
          sendUpdates: args.sendUpdates,
        });
      }

      const text = formatCreatedEvent(result);

      return {
        content: [
          {
            type: 'text',
            text:
              text +
              "\n\nNext: Share htmlLink with user. Use 'search_events' to verify.",
          },
        ],
        structuredContent: { ...result },
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          { type: 'text', text: `Failed to create event: ${(error as Error).message}` },
        ],
      };
    }
  },
});
