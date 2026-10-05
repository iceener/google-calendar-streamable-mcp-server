import * as z from 'zod/v4';
import { defineTool } from '../platform/primitives';
import {
  calendarFailure,
  calendarFor,
  signInRequired,
  withPrimaryFallback,
} from './shared/calendar';

const DeleteEventOutputSchema = z.object({
  success: z.literal(true),
  eventId: z.string(),
  calendarId: z.string(),
});

const NOTIFIED = {
  all: 'All attendees were notified.',
  externalOnly: 'External attendees were notified.',
  none: 'No notifications sent.',
} as const;

export const deleteEvent = defineTool(
  'delete_event',
  {
    title: 'Delete Event',
    description:
      "Delete an event from a calendar. Inputs: eventId (required), calendarId? (default: 'primary'), sendUpdates? ('all'|'externalOnly'|'none', default: 'none').\nBehavior: Permanently removes the event. If it's a recurring event instance, only that instance is deleted. To delete all instances, delete the parent event (use recurringEventId).\nReturns: { success: true }.\nNext: Use 'search_events' to verify deletion.",
    inputSchema: z.object({
      eventId: z.string().describe('Event ID to delete'),
      calendarId: z.string().optional().describe('Calendar ID (defaults to "primary")'),
      sendUpdates: z
        .enum(['all', 'externalOnly', 'none'])
        .optional()
        .default('none')
        .describe('Notify attendees about the cancellation'),
    }),
    outputSchema: DeleteEventOutputSchema,
    annotations: { readOnlyHint: false, destructiveHint: true },
  },
  async (args, ctx, deps) => {
    const calendar = calendarFor(ctx, deps);
    if (!calendar) return signInRequired();
    try {
      const calendarId = await withPrimaryFallback(args.calendarId || 'primary', async (id) => {
        await calendar.deleteEvent({
          eventId: args.eventId,
          calendarId: id,
          sendUpdates: args.sendUpdates,
        });
        return id;
      });
      return {
        content: [
          {
            type: 'text',
            text: `✓ Event deleted successfully.\n  eventId: ${args.eventId}\n  calendar: ${calendarId}\n  ${NOTIFIED[args.sendUpdates]}\n\nNext: Use 'search_events' to verify deletion.`,
          },
        ],
        structuredContent: { success: true, eventId: args.eventId, calendarId },
      };
    } catch (error) {
      return calendarFailure('delete event', error);
    }
  },
);
