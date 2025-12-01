/**
 * Check Availability tool - check free/busy status for time slots.
 */

import { z } from 'zod';
import { toolsMetadata } from '../../config/metadata.js';
import { GoogleCalendarClient } from '../../services/google-calendar.js';
import { defineTool, type ToolResult } from './types.js';

const InputSchema = z.object({
  timeMin: z.string().describe('Start of time range to check (ISO 8601)'),
  timeMax: z.string().describe('End of time range to check (ISO 8601)'),
  calendarIds: z
    .array(z.string())
    .optional()
    .default(['primary'])
    .describe('Calendar IDs to check (defaults to ["primary"])'),
  timeZone: z.string().optional().describe('Timezone for the response'),
});

function formatBusySlot(slot: { start: string; end: string }): string {
  const start = new Date(slot.start).toLocaleString();
  const end = new Date(slot.end).toLocaleString();
  return `${start} → ${end}`;
}

export const checkAvailabilityTool = defineTool({
  name: toolsMetadata.check_availability.name,
  title: toolsMetadata.check_availability.title,
  description: toolsMetadata.check_availability.description,
  inputSchema: InputSchema,
  annotations: {
    readOnlyHint: true,
    destructiveHint: false,
  },

  handler: async (args, context): Promise<ToolResult> => {
    const token = context.providerToken;

    if (!token) {
      return {
        isError: true,
        content: [{ type: 'text', text: 'Authentication required. Please authenticate with Google Calendar.' }],
      };
    }

    const client = new GoogleCalendarClient(token);

    try {
      const result = await client.getFreeBusy({
        timeMin: args.timeMin,
        timeMax: args.timeMax,
        calendarIds: args.calendarIds,
        timeZone: args.timeZone,
      });

      // Format for LLM consumption
      const lines: string[] = [];
      lines.push(`Availability check: ${args.timeMin} to ${args.timeMax}\n`);

      let totalBusySlots = 0;

      for (const [calendarId, data] of Object.entries(result.calendars)) {
        const busyCount = data.busy?.length || 0;
        totalBusySlots += busyCount;

        if (data.errors && data.errors.length > 0) {
          lines.push(`📅 ${calendarId}: Error - ${data.errors[0].reason}`);
          continue;
        }

        if (busyCount === 0) {
          lines.push(`📅 ${calendarId}: Completely free during this period ✓`);
        } else {
          lines.push(`📅 ${calendarId}: ${busyCount} busy slot(s)`);
          for (const slot of data.busy) {
            lines.push(`   - ${formatBusySlot(slot)}`);
          }
        }
        lines.push('');
      }

      if (totalBusySlots === 0) {
        lines.push('✓ All calendars are free during this time range.');
      } else {
        lines.push(`Total: ${totalBusySlots} busy slot(s) across all calendars.`);
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
      return {
        isError: true,
        content: [{ type: 'text', text: `Failed to check availability: ${(error as Error).message}` }],
      };
    }
  },
});



