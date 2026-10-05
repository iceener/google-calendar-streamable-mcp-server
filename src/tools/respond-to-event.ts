import * as z from 'zod/v4';
import { defineTool } from '../platform/primitives';
import type { CalendarEvent } from '../services/google-calendar';
import {
  calendarFailure,
  calendarFor,
  signInRequired,
  withPrimaryFallback,
} from './shared/calendar';
import { CalendarEventOutputSchema } from './shared/schemas';

const RespondToEventOutputSchema = z.object({
  ok: z.literal(true),
  action: z.literal('respond_to_event'),
  response: z.enum(['accepted', 'declined', 'tentative']),
  event: CalendarEventOutputSchema,
});

const RESPONSE_LABELS = {
  accepted: 'accepted',
  declined: 'declined',
  tentative: 'marked as maybe',
} as const;

export const respondToEvent = defineTool(
  'respond_to_event',
  {
    title: 'Respond to Event',
    description: `Accept, decline, or tentatively accept an event invitation.

Inputs:
- eventId: string (required) — Event ID from search_events
- calendarId?: string (default: 'primary') — Calendar where the event appears
- response: 'accepted' | 'declined' | 'tentative' (required)
  - 'accepted' = Yes, I'll attend
  - 'declined' = No, I won't attend  
  - 'tentative' = Maybe
- sendUpdates?: 'all' | 'externalOnly' | 'none' (default: 'all')

Behavior: Updates YOUR attendance status for the event. You must be an attendee (invited) to respond.

Returns: Updated event object with your new response status.

Note: This only works for events you were invited to. For events you created yourself, you are the organizer, not an attendee.`,
    inputSchema: z.object({
      eventId: z.string().describe('Event ID to respond to'),
      calendarId: z.string().optional().describe('Calendar ID (defaults to "primary")'),
      response: z
        .enum(['accepted', 'declined', 'tentative'])
        .describe('Your response: "accepted" (yes), "declined" (no), or "tentative" (maybe)'),
      sendUpdates: z.enum(['all', 'externalOnly', 'none']).optional().default('all'),
    }),
    outputSchema: RespondToEventOutputSchema,
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  async (args, ctx, deps) => {
    const calendar = calendarFor(ctx, deps);
    if (!calendar) return signInRequired();
    try {
      const event = await withPrimaryFallback(args.calendarId || 'primary', (calendarId) =>
        calendar.respondToEvent({
          calendarId,
          eventId: args.eventId,
          response: args.response,
          sendUpdates: args.sendUpdates,
        }),
      );
      return {
        content: [{ type: 'text', text: describeResponse(event, args.response) }],
        structuredContent: {
          ok: true,
          action: 'respond_to_event',
          response: args.response,
          event,
        },
      };
    } catch (error) {
      return calendarFailure('respond to event', error);
    }
  },
);

function describeResponse(event: CalendarEvent, response: keyof typeof RESPONSE_LABELS): string {
  const title = event.summary || '(no title)';
  const label = RESPONSE_LABELS[response];
  const lines = [
    event.htmlLink ? `✓ You ${label}: [${title}](${event.htmlLink})` : `✓ You ${label}: ${title}`,
    `  when: ${event.start?.dateTime || event.start?.date || 'no date'}`,
  ];
  if (event.location) lines.push(`  location: ${event.location}`);
  if (event.hangoutLink) lines.push(`  meet: ${event.hangoutLink}`);
  return lines.join('\n');
}
