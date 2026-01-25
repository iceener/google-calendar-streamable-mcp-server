/**
 * Update Event tool - update or move existing events.
 */

import { z } from 'zod';
import { toolsMetadata } from '../../config/metadata.js';
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
  eventId: z.string().describe('Event ID to update'),
  calendarId: z.string().optional().describe('Calendar ID (defaults to "primary")'),
  targetCalendarId: z.string().optional().describe('Move event to this calendar'),

  // Fields to update
  summary: z.string().optional().describe('New event title'),
  start: z.string().optional().describe('New start time (ISO 8601)'),
  end: z.string().optional().describe('New end time (ISO 8601)'),
  description: z.string().optional().describe('New description'),
  location: z.string().optional().describe('New location'),
  attendees: z
    .array(z.string().email())
    .optional()
    .describe('New attendee list (replaces existing)'),
  addGoogleMeet: z.boolean().optional().describe('Add Google Meet link'),
  recurrence: z.array(z.string()).optional().describe('New RRULE array'),
  reminders: RemindersSchema.optional().describe('New reminder settings'),
  visibility: z.enum(['default', 'public', 'private', 'confidential']).optional(),
  colorId: z.string().optional().describe('New color ID (1-11)'),

  // Options
  sendUpdates: z.enum(['all', 'externalOnly', 'none']).optional().default('none'),
  timeZone: z.string().optional().describe('Time zone for datetime values'),
});

function isAllDayDate(str: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$/.test(str);
}

function formatUpdatedEvent(event: CalendarEvent, wasMoved: boolean): string {
  const lines: string[] = [];

  const title = event.summary || '(no title)';
  const start = event.start?.dateTime || event.start?.date || 'no date';

  const action = wasMoved ? 'moved and updated' : 'updated';

  if (event.htmlLink) {
    lines.push(`✓ Event ${action}: [${title}](${event.htmlLink})`);
  } else {
    lines.push(`✓ Event ${action}: ${title}`);
  }

  lines.push(`  id: ${event.id}`);
  lines.push(`  when: ${start}`);

  if (event.location) {
    lines.push(`  location: ${event.location}`);
  }

  if (event.hangoutLink) {
    lines.push(`  meet: ${event.hangoutLink}`);
  }

  return lines.join('\n');
}

export const updateEventTool = defineTool({
  name: toolsMetadata.update_event.name,
  title: toolsMetadata.update_event.title,
  description: toolsMetadata.update_event.description,
  inputSchema: InputSchema,
  annotations: {
    readOnlyHint: false,
    destructiveHint: false,
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
    const calendarId = args.calendarId || 'primary';

    try {
      let result: CalendarEvent;
      let wasMoved = false;

      // Step 1: Move if targetCalendarId is different
      if (args.targetCalendarId && args.targetCalendarId !== calendarId) {
        result = await client.moveEvent({
          calendarId,
          eventId: args.eventId,
          destinationCalendarId: args.targetCalendarId,
          sendUpdates: args.sendUpdates,
        });
        wasMoved = true;
      }

      // Step 2: Patch if any fields to update
      const hasFieldsToUpdate =
        args.summary !== undefined ||
        args.start !== undefined ||
        args.end !== undefined ||
        args.description !== undefined ||
        args.location !== undefined ||
        args.attendees !== undefined ||
        args.addGoogleMeet !== undefined ||
        args.recurrence !== undefined ||
        args.reminders !== undefined ||
        args.visibility !== undefined ||
        args.colorId !== undefined;

      if (hasFieldsToUpdate) {
        // Build start/end objects if provided
        let startObj;
        let endObj;

        if (args.start) {
          startObj = isAllDayDate(args.start)
            ? { date: args.start }
            : { dateTime: args.start, timeZone: args.timeZone };
        }

        if (args.end) {
          endObj = isAllDayDate(args.end)
            ? { date: args.end }
            : { dateTime: args.end, timeZone: args.timeZone };
        }

        result = await client.updateEvent({
          calendarId: wasMoved ? args.targetCalendarId : calendarId,
          eventId: args.eventId,
          summary: args.summary,
          description: args.description,
          start: startObj,
          end: endObj,
          location: args.location,
          attendees: args.attendees,
          addGoogleMeet: args.addGoogleMeet,
          recurrence: args.recurrence,
          reminders: args.reminders,
          visibility: args.visibility,
          colorId: args.colorId,
          sendUpdates: args.sendUpdates,
        });
      } else if (!wasMoved) {
        // Nothing to do
        return {
          isError: true,
          content: [
            {
              type: 'text',
              text: 'No changes specified. Provide at least one field to update or a targetCalendarId to move.',
            },
          ],
        };
      }

      const text = formatUpdatedEvent(result!, wasMoved);

      return {
        content: [
          {
            type: 'text',
            text: text + "\n\nNext: Use 'search_events' to verify changes.",
          },
        ],
        structuredContent: result!,
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          { type: 'text', text: `Failed to update event: ${(error as Error).message}` },
        ],
      };
    }
  },
});































