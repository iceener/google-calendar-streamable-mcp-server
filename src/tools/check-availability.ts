import * as z from 'zod/v4';
import { defineTool } from '../platform/primitives';
import { calendarFailure, calendarFor, signInRequired } from './shared/calendar';
import { rfc3339 } from './shared/schemas';

const AvailabilityOutputSchema = z.object({
  timeMin: z.string(),
  timeMax: z.string(),
  calendars: z.record(
    z.string(),
    z.object({
      busy: z.array(z.object({ start: z.string(), end: z.string() })),
      errors: z.array(z.object({ domain: z.string(), reason: z.string() })).optional(),
    }),
  ),
});

export const checkAvailability = defineTool(
  'check_availability',
  {
    title: 'Check Availability',
    description:
      "Check free/busy status for time slots across one or more calendars. Use BEFORE scheduling to find available times. Inputs: timeMin, timeMax (ISO 8601, required), calendarIds? (default: ['primary']).\nReturns: { calendars: { [calendarId]: { busy: Array<{ start, end }> } } }.\nBehavior: Returns only busy time blocks. Free time = gaps between busy blocks.\nNext: Use free slots to suggest meeting times, then 'create_event' to book.",
    inputSchema: z.object({
      timeMin: rfc3339.describe(
        'Start of time range to check (RFC3339 with timezone, e.g., 2025-12-06T09:00:00Z or 2025-12-06T09:00:00+01:00)',
      ),
      timeMax: rfc3339.describe(
        'End of time range to check (RFC3339 with timezone, e.g., 2025-12-06T17:00:00Z or 2025-12-06T17:00:00+01:00)',
      ),
      calendarIds: z
        .array(z.string())
        .optional()
        .default(['primary'])
        .describe('Calendar IDs to check (defaults to ["primary"])'),
      timeZone: z.string().optional().describe('Timezone for the response'),
    }),
    outputSchema: AvailabilityOutputSchema,
    annotations: { readOnlyHint: true, destructiveHint: false },
  },
  async (args, ctx, deps) => {
    const calendar = calendarFor(ctx, deps);
    if (!calendar) return signInRequired();
    try {
      const result = await calendar.getFreeBusy(args);
      const lines = [`Availability check: ${args.timeMin} to ${args.timeMax}\n`];
      let totalBusy = 0;
      for (const [calendarId, data] of Object.entries(result.calendars)) {
        const busyCount = data.busy?.length || 0;
        totalBusy += busyCount;
        if (data.errors && data.errors.length > 0) {
          lines.push(`📅 ${calendarId}: Error - ${data.errors[0]?.reason}`);
          continue;
        }
        if (busyCount === 0) {
          lines.push(`📅 ${calendarId}: Completely free during this period ✓`);
        } else {
          lines.push(`📅 ${calendarId}: ${busyCount} busy slot(s)`);
          for (const slot of data.busy) {
            lines.push(
              `   - ${new Date(slot.start).toLocaleString()} → ${new Date(slot.end).toLocaleString()}`,
            );
          }
        }
        lines.push('');
      }
      if (totalBusy === 0) {
        lines.push('✓ All calendars are free during this time range.');
      } else {
        lines.push(`Total: ${totalBusy} busy slot(s) across all calendars.`);
        lines.push('Free times are the gaps between busy slots.');
      }
      lines.push("\nNext: Use 'create_event' to schedule during free times.");
      return {
        content: [{ type: 'text', text: lines.join('\n') }],
        structuredContent: {
          timeMin: result.timeMin,
          timeMax: result.timeMax,
          calendars: result.calendars,
        },
      };
    } catch (error) {
      return calendarFailure('check availability', error);
    }
  },
);
