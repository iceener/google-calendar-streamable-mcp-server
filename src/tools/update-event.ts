import * as z from 'zod/v4';
import { defineTool, toolError } from '../platform/primitives';
import type { CalendarEvent, EventDateTime } from '../services/google-calendar';
import {
  calendarFailure,
  calendarFor,
  signInRequired,
  withPrimaryFallback,
} from './shared/calendar';
import {
  attendeeEmail,
  CalendarEventOutputSchema,
  isAllDayDate,
  RemindersSchema,
} from './shared/schemas';

export const updateEvent = defineTool(
  'update_event',
  {
    title: 'Update Event',
    description: `Update or move an existing event. Uses PATCH semantics (only provided fields are changed). Inputs: eventId (required), calendarId? (default: 'primary'), targetCalendarId? (moves event if different from calendarId), sendUpdates? ('all'|'externalOnly'|'none', default: 'none'), plus any field to update: summary?, start?, end?, description?, location?, attendees?, addGoogleMeet?, recurrence?, reminders?, visibility?, colorId?.

MOVE BEHAVIOR:
- If targetCalendarId differs from calendarId, performs Move operation first.
- Only 'default' events can be moved (not birthday, focusTime, outOfOffice, workingLocation).

PATCH BEHAVIOR:
- Only sends fields you provide; omitted fields remain unchanged.
- To clear a field, set it to null or empty string where applicable.

Returns: Updated event object.
Next: Use 'search_events' to verify changes. Share updated htmlLink if needed.`,
    inputSchema: z.object({
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
        .array(attendeeEmail())
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
    }),
    outputSchema: CalendarEventOutputSchema,
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  async (args, ctx, deps) => {
    const calendar = calendarFor(ctx, deps);
    if (!calendar) return signInRequired();

    const changesFields = [
      args.summary,
      args.start,
      args.end,
      args.description,
      args.location,
      args.attendees,
      args.addGoogleMeet,
      args.recurrence,
      args.reminders,
      args.visibility,
      args.colorId,
    ].some((value) => value !== undefined);
    if (!changesFields && !args.targetCalendarId) {
      return toolError(
        'No changes specified. Provide at least one field to update or a targetCalendarId to move.',
      );
    }
    const when = (value: string | undefined): EventDateTime | undefined => {
      if (!value) return undefined;
      return isAllDayDate(value) ? { date: value } : { dateTime: value, timeZone: args.timeZone };
    };

    try {
      const { event, moved } = await withPrimaryFallback(
        args.calendarId || 'primary',
        async (calendarId) => {
          let event: CalendarEvent | undefined;
          let moved = false;
          // A different target calendar moves the event first, then patches it there.
          if (args.targetCalendarId && args.targetCalendarId !== calendarId) {
            event = await calendar.moveEvent({
              calendarId,
              eventId: args.eventId,
              destinationCalendarId: args.targetCalendarId,
              sendUpdates: args.sendUpdates,
            });
            moved = true;
          }
          if (changesFields) {
            event = await calendar.updateEvent({
              calendarId: moved ? args.targetCalendarId : calendarId,
              eventId: args.eventId,
              summary: args.summary,
              description: args.description,
              start: when(args.start),
              end: when(args.end),
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
          if (!event) throw new Error('No event update or move was performed');
          return { event, moved };
        },
      );
      return {
        content: [
          {
            type: 'text',
            text: `${describeUpdated(event, moved)}\n\nNext: Use 'search_events' to verify changes.`,
          },
        ],
        structuredContent: { ...event },
      };
    } catch (error) {
      return calendarFailure('update event', error);
    }
  },
);

function describeUpdated(event: CalendarEvent, moved: boolean): string {
  const title = event.summary || '(no title)';
  const action = moved ? 'moved and updated' : 'updated';
  const lines = [
    event.htmlLink
      ? `✓ Event ${action}: [${title}](${event.htmlLink})`
      : `✓ Event ${action}: ${title}`,
    `  id: ${event.id}`,
    `  when: ${event.start?.dateTime || event.start?.date || 'no date'}`,
  ];
  if (event.location) lines.push(`  location: ${event.location}`);
  if (event.hangoutLink) lines.push(`  meet: ${event.hangoutLink}`);
  return lines.join('\n');
}
