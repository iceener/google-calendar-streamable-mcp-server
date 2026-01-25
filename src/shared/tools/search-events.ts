/**
 * Search Events tool - search and filter events across all calendars.
 */

import { z } from 'zod';
import { toolsMetadata } from '../../config/metadata.js';
import {
  type CalendarEvent,
  type CalendarListItem,
  GoogleCalendarClient,
} from '../../services/google-calendar.js';
import { defineTool, type ToolResult } from './types.js';

const DEFAULT_FIELDS = [
  'id',
  'summary',
  'start',
  'end',
  'location',
  'htmlLink',
  'status',
  'attendees',
  'calendarId',
  'calendarName',
];

const ALL_FIELDS = [
  'id',
  'summary',
  'description',
  'start',
  'end',
  'location',
  'attendees',
  'organizer',
  'creator',
  'htmlLink',
  'hangoutLink',
  'conferenceData',
  'status',
  'eventType',
  'visibility',
  'colorId',
  'recurringEventId',
  'recurrence',
  'calendarId',
  'calendarName',
];

// Extended event type with calendar info
interface EventWithCalendar extends CalendarEvent {
  calendarId: string;
  calendarName: string;
}

const InputSchema = z.object({
  calendarId: z
    .union([z.literal('all'), z.string(), z.array(z.string())])
    .optional()
    .default('all')
    .describe(
      'Calendar ID(s) to search. Use "all" (default) to search all calendars, a single ID, or array of IDs',
    ),
  timeMin: z.string().optional().describe('Start of time range (RFC3339 with timezone, e.g., 2025-12-06T19:00:00Z or 2025-12-06T19:00:00+01:00)'),
  timeMax: z.string().optional().describe('End of time range (RFC3339 with timezone, e.g., 2025-12-06T19:00:00Z or 2025-12-06T19:00:00+01:00)'),
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
    .array(
      z.enum(['default', 'birthday', 'focusTime', 'outOfOffice', 'workingLocation']),
    )
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
});

function formatEventLine(event: EventWithCalendar): string {
  const start = event.start?.dateTime || event.start?.date || 'no date';
  const title = event.summary || '(no title)';
  const calendar = event.calendarName ? ` (${event.calendarName})` : '';

  // Find user's own response status (self: true in attendees)
  const selfAttendee = event.attendees?.find((a) => a.self);
  let statusInfo = '';

  if (selfAttendee?.responseStatus) {
    // Show user's response status: accepted, declined, tentative, needsAction
    const responseMap: Record<string, string> = {
      accepted: 'you: accepted',
      declined: 'you: declined',
      tentative: 'you: maybe',
      needsAction: 'you: not responded',
    };
    statusInfo = ` [${responseMap[selfAttendee.responseStatus] || selfAttendee.responseStatus}]`;
  } else if (event.status === 'cancelled') {
    statusInfo = ' [cancelled]';
  }

  if (event.htmlLink) {
    return `- [${title}](${event.htmlLink}) — ${start}${calendar}${statusInfo}`;
  }
  return `- ${title} — ${start}${calendar}${statusInfo}`;
}

function pickFields(
  event: EventWithCalendar,
  fields: string[],
): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  for (const field of fields) {
    if (field in event) {
      result[field] = (event as unknown as Record<string, unknown>)[field];
    }
  }
  return result;
}

function getEventStartTime(event: CalendarEvent): number {
  const dateStr = event.start?.dateTime || event.start?.date;
  if (!dateStr) return 0;
  return new Date(dateStr).getTime();
}

/**
 * Client-side substring search to complement Google's exact word matching.
 * Google's API `q` parameter only matches exact words, so "barber" won't find "barbershop".
 * This filter catches those cases with case-insensitive substring matching.
 */
function matchesQuerySubstring(event: CalendarEvent, query: string): boolean {
  const lowerQuery = query.toLowerCase();
  const searchableFields = [
    event.summary,
    event.description,
    event.location,
    ...(event.attendees?.map((a) => a.email) ?? []),
    ...(event.attendees?.map((a) => a.displayName) ?? []),
  ];

  return searchableFields.some(
    (field) => field && field.toLowerCase().includes(lowerQuery),
  );
}

export const searchEventsTool = defineTool({
  name: toolsMetadata.search_events.name,
  title: toolsMetadata.search_events.title,
  description: toolsMetadata.search_events.description,
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
        content: [
          {
            type: 'text',
            text: 'Authentication required. Please authenticate with Google Calendar.',
          },
        ],
      };
    }

    const client = new GoogleCalendarClient(token);

    try {
      // Determine which calendars to search
      let calendarsToSearch: CalendarListItem[] = [];

      if (args.calendarId === 'all') {
        // Fetch all accessible calendars
        const calendarList = await client.listCalendars();
        calendarsToSearch = calendarList.items.filter(
          // Include calendars where user can at least read events
          (cal) => ['owner', 'writer', 'reader'].includes(cal.accessRole),
        );
      } else if (Array.isArray(args.calendarId)) {
        // Use provided calendar IDs
        calendarsToSearch = args.calendarId.map((id) => ({
          id,
          summary: id === 'primary' ? 'Primary' : id,
          accessRole: 'reader' as const,
        }));
      } else {
        // Single calendar ID
        calendarsToSearch = [
          {
            id: args.calendarId,
            summary: args.calendarId === 'primary' ? 'Primary' : args.calendarId,
            accessRole: 'reader' as const,
          },
        ];
      }

      // For pagination with single calendar
      if (args.pageToken && calendarsToSearch.length !== 1) {
        return {
          isError: true,
          content: [
            {
              type: 'text',
              text: 'Pagination (pageToken) only works when searching a single calendar. Specify a calendarId to use pagination.',
            },
          ],
        };
      }

      // Search all calendars in parallel
      // Note: We don't pass `q` to Google API because it only does exact word matching.
      // Instead, we fetch events and filter locally with substring matching.
      // This ensures "barber" will match "barbershop".
      
      // When doing local query filtering, we need to fetch enough events to have
      // a reasonable chance of finding matches. A small maxResults (e.g., 1) with
      // only 2x multiplier means we might miss events that exist further in the list.
      const hasLocalQuery = !!args.query;
      const requestedMax = args.maxResults ?? 50;
      const fetchMultiplier = hasLocalQuery ? 10 : 2; // Fetch more when filtering locally
      const minFetchAmount = hasLocalQuery ? 100 : 10; // Minimum events to fetch for query searches
      
      const searchPromises = calendarsToSearch.map(async (calendar) => {
        try {
          const result = await client.listEvents({
            calendarId: calendar.id,
            timeMin: args.timeMin,
            timeMax: args.timeMax,
            maxResults:
              args.calendarId === 'all'
                ? Math.min(Math.max(requestedMax * fetchMultiplier, minFetchAmount), 250)
                : Math.min(Math.max(requestedMax * fetchMultiplier, minFetchAmount), 250),
            singleEvents: args.singleEvents,
            orderBy: args.singleEvents ? args.orderBy || 'startTime' : args.orderBy,
            // Don't use Google's q parameter - do local substring filtering instead
            eventTypes: args.eventTypes,
            pageToken: args.pageToken,
          });

          // Add calendar info to each event
          const eventsWithCalendar: EventWithCalendar[] = result.items.map((event) => ({
            ...event,
            calendarId: calendar.id,
            calendarName: calendar.summary,
          }));

          return {
            calendar,
            events: eventsWithCalendar,
            nextPageToken: result.nextPageToken,
          };
        } catch (error) {
          // Log error but don't fail the whole search
          console.warn(
            `Failed to search calendar ${calendar.id}: ${(error as Error).message}`,
          );
          return {
            calendar,
            events: [],
            error: (error as Error).message,
          };
        }
      });

      const results = await Promise.all(searchPromises);

      // Merge all events and sort by start time
      let allEvents: EventWithCalendar[] = results.flatMap((r) => r.events);

      // Apply local substring filtering if query is provided
      // This catches partial matches that Google's exact word matching misses
      if (args.query) {
        allEvents = allEvents.filter((event) => matchesQuerySubstring(event, args.query!));
      }

      // Sort by start time if using startTime ordering
      if (
        args.singleEvents !== false &&
        (args.orderBy === 'startTime' || !args.orderBy)
      ) {
        allEvents.sort((a, b) => getEventStartTime(a) - getEventStartTime(b));
      }

      // Apply total maxResults limit
      const maxResults = args.maxResults ?? 50;
      const hasMore = allEvents.length > maxResults;
      allEvents = allEvents.slice(0, maxResults);

      const fields =
        args.fields && args.fields.length > 0 ? args.fields : DEFAULT_FIELDS;
      const filteredItems = allEvents.map((event) => pickFields(event, fields));

      // Format for LLM consumption
      const lines: string[] = [];

      // Show which calendars were searched
      const searchedCalendars = results
        .filter((r) => !r.error)
        .map((r) => r.calendar.summary);
      const failedCalendars = results
        .filter((r) => r.error)
        .map((r) => r.calendar.summary);

      if (args.calendarId === 'all' && searchedCalendars.length > 1) {
        lines.push(
          `Searched ${searchedCalendars.length} calendar(s): ${searchedCalendars.join(', ')}`,
        );
        if (failedCalendars.length > 0) {
          lines.push(`(Failed to search: ${failedCalendars.join(', ')})`);
        }
        lines.push('');
      }

      if (allEvents.length === 0) {
        lines.push('No events found matching the criteria.');
      } else {
        lines.push(
          `Found ${allEvents.length} event(s)${hasMore ? ' (more available)' : ''}:\n`,
        );

        for (const event of allEvents) {
          lines.push(formatEventLine(event));

          if (event.location) {
            lines.push(`  location: ${event.location}`);
          }
          if (event.attendees && event.attendees.length > 0) {
            const attendeeList = event.attendees
              .slice(0, 5)
              .map((a) => a.email)
              .join(', ');
            const more =
              event.attendees.length > 5 ? ` +${event.attendees.length - 5} more` : '';
            lines.push(`  attendees: ${attendeeList}${more}`);
          }
          if (event.hangoutLink) {
            lines.push(`  meet: ${event.hangoutLink}`);
          }
        }
      }

      // Only show nextPageToken for single calendar searches AND when we have results
      // Don't show pagination hints when local filtering returned 0 results - it's misleading
      // because the next page might also not contain matching events
      const singleCalendarResult = calendarsToSearch.length === 1 ? results[0] : null;
      const hasResultsToShow = allEvents.length > 0;
      
      if (hasResultsToShow && singleCalendarResult?.nextPageToken && !args.query) {
        // Only show pageToken when not doing local query filtering
        // (pageToken doesn't account for our substring filter)
        lines.push(
          `\nMore results available. Pass pageToken: "${singleCalendarResult.nextPageToken}" to fetch next page.`,
        );
      } else if (hasResultsToShow && hasMore) {
        lines.push(
          `\nMore results available. Increase maxResults or narrow your time range.`,
        );
      }

      lines.push(
        "\nNote: Use the calendarId from results when calling 'update_event' or 'delete_event'.",
      );

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
        structuredContent: {
          items: filteredItems,
          calendarsSearched: searchedCalendars,
          nextPageToken: singleCalendarResult?.nextPageToken,
        },
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          {
            type: 'text',
            text: `Failed to search events: ${(error as Error).message}`,
          },
        ],
      };
    }
  },
});
