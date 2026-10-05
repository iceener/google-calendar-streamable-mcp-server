import * as z from 'zod/v4';
import { defineTool, toolError } from '../platform/primitives';
import type { CalendarEvent } from '../services/google-calendar';
import { calendarFailure, calendarFor, signInRequired } from './shared/calendar';
import {
  attendeeEmail,
  CalendarEventOutputSchema,
  isAllDayDate,
  RemindersSchema,
} from './shared/schemas';

export const createEvent = defineTool(
  'create_event',
  {
    title: 'Create Event',
    description: `Create a calendar event using natural language OR structured input. Inputs vary by mode.

MODE A - Natural language (uses Google quickAdd):
- text: string (e.g., "Lunch with Anna tomorrow at noon for 1 hour", "Team standup every Monday 9am")
- calendarId?: string (default: 'primary')
- sendUpdates?: 'all'|'externalOnly'|'none' (default: 'none')
Detection: If 'text' is provided without 'summary', uses quickAdd.

MODE B - Structured:
- summary: string (required, event title)
- start: string (ISO 8601 datetime) or { date: string } for all-day (required)
- end: string (ISO 8601 datetime) or { date: string } for all-day (required)
- calendarId?: string (default: 'primary')
- description?: string
- location?: string
- attendees?: string[] (array of email addresses)
- addGoogleMeet?: boolean (default: false, auto-creates Meet link)
- recurrence?: string[] (RRULE array, e.g., ["RRULE:FREQ=WEEKLY;COUNT=10"])
- reminders?: { useDefault: boolean, overrides?: Array<{ method: 'popup'|'email', minutes: number }> }
- visibility?: 'default'|'public'|'private'|'confidential'
- colorId?: string (1-11)
- sendUpdates?: 'all'|'externalOnly'|'none' (default: 'none')

Returns: Created event object with id, htmlLink, and all fields.
Next: Share htmlLink with user. Use 'search_events' to verify creation.`,
    inputSchema: z.object({
      // Natural language mode
      text: z
        .string()
        .optional()
        .describe('Natural language event description (e.g., "Lunch with Anna tomorrow at noon")'),
      // Structured mode
      summary: z.string().optional().describe('Event title'),
      start: z
        .string()
        .optional()
        .describe('Start time (ISO 8601 datetime or YYYY-MM-DD for all-day)'),
      end: z.string().optional().describe('End time (ISO 8601 datetime or YYYY-MM-DD for all-day)'),
      description: z.string().optional().describe('Event description'),
      location: z.string().optional().describe('Event location'),
      attendees: z.array(attendeeEmail()).optional().describe('List of attendee email addresses'),
      // Shared options
      calendarId: z.string().optional().describe('Calendar ID (defaults to "primary")'),
      addGoogleMeet: z.boolean().optional().default(false).describe('Auto-create Google Meet link'),
      recurrence: z.array(z.string()).optional().describe('RRULE array for recurring events'),
      reminders: RemindersSchema.optional().describe('Reminder settings'),
      visibility: z.enum(['default', 'public', 'private', 'confidential']).optional(),
      colorId: z.string().optional().describe('Color ID (1-11)'),
      sendUpdates: z.enum(['all', 'externalOnly', 'none']).optional().default('none'),
      timeZone: z.string().optional().describe('Time zone for the event'),
    }),
    outputSchema: CalendarEventOutputSchema,
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  async (args, ctx, deps) => {
    const calendar = calendarFor(ctx, deps);
    if (!calendar) return signInRequired();
    try {
      let event: CalendarEvent;
      if (args.text && !args.summary) {
        // Natural language: Google's quick add parses the text.
        event = await calendar.quickAdd({
          calendarId: args.calendarId,
          text: args.text,
          sendUpdates: args.sendUpdates,
        });
      } else {
        if (!args.summary) {
          return toolError(
            "Either 'text' (for natural language) or 'summary' (for structured) is required.",
          );
        }
        if (!args.start || !args.end) {
          return toolError("'start' and 'end' are required for structured event creation.");
        }
        const allDay = isAllDayDate(args.start) && isAllDayDate(args.end);
        event = await calendar.createEvent({
          calendarId: args.calendarId,
          summary: args.summary,
          description: args.description,
          start: allDay ? { date: args.start } : { dateTime: args.start, timeZone: args.timeZone },
          end: allDay ? { date: args.end } : { dateTime: args.end, timeZone: args.timeZone },
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
      return {
        content: [
          {
            type: 'text',
            text: `${describeCreated(event)}\n\nNext: Share htmlLink with user. Use 'search_events' to verify.`,
          },
        ],
        structuredContent: { ...event },
      };
    } catch (error) {
      return calendarFailure('create event', error);
    }
  },
);

function describeCreated(event: CalendarEvent): string {
  const title = event.summary || '(no title)';
  const lines = [
    event.htmlLink ? `✓ Event created: [${title}](${event.htmlLink})` : `✓ Event created: ${title}`,
    `  id: ${event.id}`,
    `  when: ${event.start?.dateTime || event.start?.date || 'no date'}`,
  ];
  if (event.location) lines.push(`  location: ${event.location}`);
  if (event.hangoutLink) lines.push(`  meet: ${event.hangoutLink}`);
  if (event.attendees && event.attendees.length > 0) {
    lines.push(`  attendees: ${event.attendees.map((attendee) => attendee.email).join(', ')}`);
  }
  return lines.join('\n');
}
