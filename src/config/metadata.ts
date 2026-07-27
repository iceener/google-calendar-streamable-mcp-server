/**
 * Centralized tool metadata for the Google Calendar MCP server.
 */

import type { Icon } from '@modelcontextprotocol/server';
import type { AppConfig } from './env.js';

export interface ToolMetadata {
  name: string;
  title: string;
  description: string;
}

export const serverMetadata = {
  title: 'Google Calendar',
  instructions: `Use these tools to manage Google Calendar events.

Quick start
- Use 'search_events' to find events — it searches ALL calendars by default!
- Use 'create_event' to add events (natural language or structured).
- Use 'update_event' to modify or move events.
- Use 'respond_to_event' to accept, decline, or tentatively accept invitations.
- Use 'check_availability' before scheduling to find free slots.
- Call 'list_calendars' to see all available calendars if needed.

Default behavior
- 'search_events' searches ALL calendars by default and shows which calendar each event belongs to.
- Other tools default to 'primary' calendar if calendarId is omitted.
- 'sendUpdates' defaults to 'none' to avoid spamming attendees during agent operations.
- Date/time values use ISO 8601 format with timezone offset (e.g., 2024-12-01T14:00:00+01:00).
- For all-day events, use date format: 2024-12-01.

Recurring events
- 'search_events' expands recurring events into instances by default (singleEvents: true).
- Each instance has 'recurringEventId' pointing to the parent event.
- To modify a single instance, use its instance 'id'. To modify all, use the parent 'recurringEventId'.

Event types
- 'default': Regular calendar event (can be created, updated, moved, deleted).
- 'birthday', 'focusTime', 'outOfOffice', 'workingLocation': Special types with restrictions.
- 'fromGmail': Auto-created from emails (read-only).
- Only 'default' events can be moved between calendars.

How to chain safely
- calendarId: Get from 'list_calendars', pass to all other tools.
- eventId: Get from 'search_events', pass to 'update_event'/'delete_event'.
- Always verify writes with 'search_events' after modifications.
`,
} as const;

export const toolsMetadata = {
  list_calendars: {
    name: 'list_calendars',
    title: 'List Calendars',
    description:
      "List all calendars accessible to the user with their details. Use this FIRST to discover calendar IDs. Inputs: none.\nReturns: { items: Array<{ id, summary, primary?, backgroundColor?, accessRole, timeZone, description? }> }.\nNext: Use calendarId from items in 'search_events', 'create_event', etc. The 'primary' calendar is the user's main calendar.",
  },

  search_events: {
    name: 'search_events',
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
  },

  check_availability: {
    name: 'check_availability',
    title: 'Check Availability',
    description:
      "Check free/busy status for time slots across one or more calendars. Use BEFORE scheduling to find available times. Inputs: timeMin, timeMax (ISO 8601, required), calendarIds? (default: ['primary']).\nReturns: { calendars: { [calendarId]: { busy: Array<{ start, end }> } } }.\nBehavior: Returns only busy time blocks. Free time = gaps between busy blocks.\nNext: Use free slots to suggest meeting times, then 'create_event' to book.",
  },

  create_event: {
    name: 'create_event',
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
  },

  update_event: {
    name: 'update_event',
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
  },

  delete_event: {
    name: 'delete_event',
    title: 'Delete Event',
    description:
      "Delete an event from a calendar. Inputs: eventId (required), calendarId? (default: 'primary'), sendUpdates? ('all'|'externalOnly'|'none', default: 'none').\nBehavior: Permanently removes the event. If it's a recurring event instance, only that instance is deleted. To delete all instances, delete the parent event (use recurringEventId).\nReturns: { success: true }.\nNext: Use 'search_events' to verify deletion.",
  },

  respond_to_event: {
    name: 'respond_to_event',
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
  },
} as const satisfies Record<string, ToolMetadata>;

/**
 * Type-safe helper to get metadata for a tool.
 */
export function getToolMetadata(toolName: keyof typeof toolsMetadata): ToolMetadata {
  return toolsMetadata[toolName];
}

/**
 * Get all registered tool names.
 */
export function getToolNames(): string[] {
  return Object.keys(toolsMetadata);
}

export const SERVER_ICON_PATH = '/icon.svg';
export const SERVER_ICON_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" role="img" aria-label="Google Calendar"><rect width="64" height="64" rx="12" fill="#fff"/><path d="M14 20h36v32H14z" fill="#4285f4"/><path d="M14 20h36v10H14z" fill="#ea4335"/><path d="M23 12v14m18-14v14" stroke="#34a853" stroke-width="5"/><path d="M24 36h16v10H24z" fill="#fbbc05"/></svg>`;

export function serverIcons(config: AppConfig): Icon[] {
  return [
    {
      src: new URL(SERVER_ICON_PATH, config.MCP_PUBLIC_URL).href,
      mimeType: 'image/svg+xml',
      sizes: ['any'],
      theme: 'light',
    },
  ];
}

export function serverImplementation(config: AppConfig) {
  return {
    name: config.MCP_NAME,
    title: config.MCP_TITLE,
    version: config.MCP_VERSION,
    description: config.MCP_DESCRIPTION,
    icons: serverIcons(config),
    ...(config.MCP_WEBSITE_URL ? { websiteUrl: config.MCP_WEBSITE_URL.href } : {}),
  };
}
