import * as z from 'zod/v4';
import { defineTool, toolError } from '../platform/primitives';
import type { CalendarEvent, CalendarListItem } from '../services/google-calendar';
import { calendarFailure, calendarFor, signInRequired } from './shared/calendar';
import { CalendarEventOutputSchema, rfc3339 } from './shared/schemas';

const DEFAULT_FIELDS = [
  'id',
  'summary',
  'start',
  'end',
  'location',
  'htmlLink',
  'status',
  'attendees',
  'organizer',
  'calendarId',
  'calendarName',
];

const SearchEventOutputSchema = CalendarEventOutputSchema.partial().extend({
  calendarId: z.string().optional(),
  calendarName: z.string().optional(),
});

const SearchEventsOutputSchema = z.object({
  items: z.array(SearchEventOutputSchema),
  calendarsSearched: z.array(z.string()),
  nextPageToken: z.string().optional(),
});

type SearchEventItem = z.input<typeof SearchEventOutputSchema>;

interface EventWithCalendar extends CalendarEvent {
  calendarId: string;
  calendarName: string;
}

const RESPONSES: Record<string, string> = {
  accepted: 'you: accepted',
  declined: 'you: declined',
  tentative: 'you: maybe',
  needsAction: 'you: not responded',
};

export const searchEvents = defineTool(
  'search_events',
  {
    title: 'Search Events',
    description: `Search events across ALL calendars by default. Returns merged results sorted by start time.

Inputs: calendarId? (default: 'all' = searches all accessible calendars; can also be a single calendar ID or array of IDs), timeMin?, timeMax? (ISO 8601), query? (searches title, description, location, attendees), maxResults? (default: 50, total across all calendars), eventTypes? (default|birthday|focusTime|outOfOffice|workingLocation), orderBy? (startTime|updated), fields? (array of fields to return).

CALENDAR SEARCH:
- Default ('all'): Searches ALL accessible calendars in parallel.
- Single calendar: calendarId: 'primary' or specific calendar ID.
- Multiple calendars: calendarId: ['primary', 'work@group.calendar.google.com'].

FILTERING BY TIME (important!):
- Today's events: timeMin=start of day, timeMax=end of day in user's timezone.
- This week: timeMin=Monday 00:00, timeMax=Sunday 23:59:59.
- Upcoming: timeMin=now, no timeMax.

FILTERING BY TYPE:
- Regular events only: eventTypes: ['default']
- Focus time: eventTypes: ['focusTime']
- Out of office: eventTypes: ['outOfOffice']

Text search: pass query: "meeting with John" to match title, description, location, or attendee names/emails.

Returns: { items: Array<{ id, summary, start, end, calendarId, calendarName, location?, htmlLink, status, ... }>, calendarsSearched, nextPageToken? }.
IMPORTANT: Each event includes 'calendarId' and 'calendarName' showing which calendar it belongs to.

Next: Use eventId AND calendarId with 'update_event' or 'delete_event'. Pagination only works with single calendar searches.`,
    inputSchema: z.object({
      calendarId: z
        .union([z.literal('all'), z.string(), z.array(z.string())])
        .optional()
        .default('all')
        .describe(
          'Calendar ID(s) to search. Use "all" (default) to search all calendars, a single ID, or array of IDs',
        ),
      timeMin: rfc3339
        .optional()
        .describe(
          'Start of time range (RFC3339 with timezone, e.g., 2025-12-06T19:00:00Z or 2025-12-06T19:00:00+01:00)',
        ),
      timeMax: rfc3339
        .optional()
        .describe(
          'End of time range (RFC3339 with timezone, e.g., 2025-12-06T19:00:00Z or 2025-12-06T19:00:00+01:00)',
        ),
      query: z
        .string()
        .optional()
        .describe('Text search (matches title, description, location, attendees)'),
      maxResults: z
        .number()
        .int()
        .min(1)
        .max(250)
        .optional()
        .default(50)
        .describe('Max events to return (total across all calendars)'),
      eventTypes: z
        .array(z.enum(['default', 'birthday', 'focusTime', 'outOfOffice', 'workingLocation']))
        .optional()
        .describe('Filter by event type'),
      orderBy: z.enum(['startTime', 'updated']).optional().describe('Sort order'),
      pageToken: z
        .string()
        .optional()
        .describe('Token for pagination (only works with single calendar)'),
      fields: z.array(z.string()).optional().describe('Fields to include in response'),
      singleEvents: z
        .boolean()
        .optional()
        .default(true)
        .describe('Expand recurring events into instances'),
    }),
    outputSchema: SearchEventsOutputSchema,
    annotations: { readOnlyHint: true, destructiveHint: false },
  },
  async (args, ctx, deps) => {
    const calendar = calendarFor(ctx, deps);
    if (!calendar) return signInRequired();
    try {
      let calendars: CalendarListItem[];
      if (args.calendarId === 'all') {
        // Every calendar whose events the user can read.
        const { items } = await calendar.listCalendars();
        calendars = items.filter((item) => ['owner', 'writer', 'reader'].includes(item.accessRole));
      } else {
        const ids = Array.isArray(args.calendarId) ? args.calendarId : [args.calendarId];
        calendars = ids.map((id) => ({
          id,
          summary: id === 'primary' ? 'Primary' : id,
          accessRole: 'reader',
        }));
      }
      if (args.pageToken && calendars.length !== 1) {
        return toolError(
          'Pagination (pageToken) only works when searching a single calendar. Specify a calendarId to use pagination.',
        );
      }

      // Google's `q` matches whole words only ("barber" misses "barbershop"), so the query is
      // matched here, as a substring, over a larger page of events.
      const fetchCount = args.query
        ? Math.min(Math.max(args.maxResults * 10, 100), 250)
        : Math.min(Math.max(args.maxResults * 2, 10), 250);

      const results = await Promise.all(
        calendars.map(async (searched) => {
          try {
            const result = await calendar.listEvents({
              calendarId: searched.id,
              timeMin: args.timeMin,
              timeMax: args.timeMax,
              maxResults: fetchCount,
              singleEvents: args.singleEvents,
              orderBy: args.singleEvents ? args.orderBy || 'startTime' : args.orderBy,
              eventTypes: args.eventTypes,
              pageToken: args.pageToken,
            });
            // An event the user organizes is reported under the organizer's address, which
            // update_event and delete_event accept; otherwise under the searched calendar.
            const events: EventWithCalendar[] = result.items.map((event) => ({
              ...event,
              calendarId:
                event.organizer?.self && event.organizer.email
                  ? event.organizer.email
                  : searched.id,
              calendarName: searched.summary,
            }));
            return { calendar: searched, events, nextPageToken: result.nextPageToken };
          } catch (error) {
            // One calendar failing doesn't fail the search.
            deps.logger.warning('Could not search a calendar', { error });
            return { calendar: searched, events: [], error: (error as Error).message };
          }
        }),
      );

      const failed = results.filter((result) => result.error);
      if (results.length > 0 && failed.length === results.length) {
        return toolError(
          `Failed to search all ${results.length} calendar(s): ${failed[0]?.error ?? 'Unknown error'}`,
        );
      }

      let events = results.flatMap((result) => result.events);
      if (args.query) {
        const query = args.query;
        events = events.filter((event) => matchesQuery(event, query));
      }
      if (args.singleEvents !== false && (args.orderBy === 'startTime' || !args.orderBy)) {
        events.sort((a, b) => startTime(a) - startTime(b));
      }
      const hasMore = events.length > args.maxResults;
      events = events.slice(0, args.maxResults);

      const fields = args.fields && args.fields.length > 0 ? args.fields : DEFAULT_FIELDS;
      const items = events.map((event) => pickFields(event, fields));

      const searchedNames = results
        .filter((result) => !result.error)
        .map((r) => r.calendar.summary);
      const failedNames = failed.map((result) => result.calendar.summary);
      const lines: string[] = [];
      if (args.calendarId === 'all' && searchedNames.length > 1) {
        lines.push(`Searched ${searchedNames.length} calendar(s): ${searchedNames.join(', ')}`);
        if (failedNames.length > 0) lines.push(`(Failed to search: ${failedNames.join(', ')})`);
        lines.push('');
      }
      if (events.length === 0) {
        lines.push('No events found matching the criteria.');
      } else {
        lines.push(`Found ${events.length} event(s)${hasMore ? ' (more available)' : ''}:\n`);
        for (const event of events) {
          lines.push(describeEvent(event));
          if (event.location) lines.push(`  location: ${event.location}`);
          if (event.attendees && event.attendees.length > 0) {
            const shown = event.attendees
              .slice(0, 5)
              .map((attendee) => attendee.email)
              .join(', ');
            const more = event.attendees.length > 5 ? ` +${event.attendees.length - 5} more` : '';
            lines.push(`  attendees: ${shown}${more}`);
          }
          if (event.hangoutLink) lines.push(`  meet: ${event.hangoutLink}`);
        }
      }

      // A page token only makes sense for one calendar and without the local query filter.
      const single = calendars.length === 1 ? results[0] : undefined;
      if (events.length > 0 && single?.nextPageToken && !args.query) {
        lines.push(
          `\nMore results available. Pass pageToken: "${single.nextPageToken}" to fetch next page.`,
        );
      } else if (events.length > 0 && hasMore) {
        lines.push('\nMore results available. Increase maxResults or narrow your time range.');
      }
      lines.push(
        "\nNote: Use the calendarId from results when calling 'update_event' or 'delete_event'.",
      );

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
        structuredContent: {
          items,
          calendarsSearched: searchedNames,
          nextPageToken: single?.nextPageToken,
        },
      };
    } catch (error) {
      return calendarFailure('search events', error);
    }
  },
);

function describeEvent(event: EventWithCalendar): string {
  const start = event.start?.dateTime || event.start?.date || 'no date';
  const title = event.summary || '(no title)';
  const calendarName = event.calendarName ? ` (${event.calendarName})` : '';
  const self = event.attendees?.find((attendee) => attendee.self);
  let status = '';
  if (self?.responseStatus) {
    status = ` [${RESPONSES[self.responseStatus] || self.responseStatus}]`;
  } else if (event.status === 'cancelled') {
    status = ' [cancelled]';
  }
  return event.htmlLink
    ? `- [${title}](${event.htmlLink}) — ${start}${calendarName}${status}`
    : `- ${title} — ${start}${calendarName}${status}`;
}

/** Only the requested fields that the event has. */
function pickFields(event: EventWithCalendar, fields: string[]): SearchEventItem {
  const source = event as unknown as Record<string, unknown>;
  return Object.fromEntries(
    fields.filter((field) => field in event).map((field) => [field, source[field]]),
  ) as SearchEventItem;
}

function startTime(event: CalendarEvent): number {
  const value = event.start?.dateTime || event.start?.date;
  return value ? new Date(value).getTime() : 0;
}

function matchesQuery(event: CalendarEvent, query: string): boolean {
  const needle = query.toLowerCase();
  return [
    event.summary,
    event.description,
    event.location,
    ...(event.attendees?.map((attendee) => attendee.email) ?? []),
    ...(event.attendees?.map((attendee) => attendee.displayName) ?? []),
  ].some((field) => field?.toLowerCase().includes(needle));
}
