import * as z from 'zod/v4';
import { defineTool } from '../platform/primitives';
import { calendarFailure, calendarFor, signInRequired } from './shared/calendar';

const ListCalendarsOutputSchema = z.object({
  items: z.array(
    z.object({
      id: z.string(),
      summary: z.string(),
      description: z.string().optional(),
      primary: z.boolean().optional(),
      backgroundColor: z.string().optional(),
      foregroundColor: z.string().optional(),
      accessRole: z.enum(['owner', 'writer', 'reader', 'freeBusyReader']),
      timeZone: z.string().optional(),
    }),
  ),
});

export const listCalendars = defineTool(
  'list_calendars',
  {
    title: 'List Calendars',
    description:
      "List all calendars accessible to the user with their details. Use this FIRST to discover calendar IDs. Inputs: none.\nReturns: { items: Array<{ id, summary, primary?, backgroundColor?, accessRole, timeZone, description? }> }.\nNext: Use calendarId from items in 'search_events', 'create_event', etc. The 'primary' calendar is the user's main calendar.",
    inputSchema: z.object({}),
    outputSchema: ListCalendarsOutputSchema,
    annotations: { readOnlyHint: true, destructiveHint: false },
  },
  async (_args, ctx, deps) => {
    const calendar = calendarFor(ctx, deps);
    if (!calendar) return signInRequired();
    try {
      const { items } = await calendar.listCalendars();
      const lines = [`Found ${items.length} calendar(s):\n`];
      for (const item of items) {
        const primary = item.primary ? ' (primary)' : '';
        const access = item.accessRole ? ` [${item.accessRole}]` : '';
        lines.push(`- ${item.summary}${primary}${access}`);
        lines.push(`  id: ${item.id}`);
        if (item.timeZone) lines.push(`  timezone: ${item.timeZone}`);
        if (item.description) lines.push(`  description: ${item.description}`);
        lines.push('');
      }
      lines.push("Next: Use calendarId in 'search_events', 'create_event', etc.");
      return { content: [{ type: 'text', text: lines.join('\n') }], structuredContent: { items } };
    } catch (error) {
      return calendarFailure('list calendars', error);
    }
  },
);
