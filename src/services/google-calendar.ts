/**
 * Google Calendar API client.
 */

import { logger } from '../utils/logger.js';

const GOOGLE_CALENDAR_API_BASE = 'https://www.googleapis.com/calendar/v3';

// ============================================================================
// Types
// ============================================================================

export interface CalendarListItem {
  id: string;
  summary: string;
  description?: string;
  primary?: boolean;
  backgroundColor?: string;
  foregroundColor?: string;
  accessRole: 'owner' | 'writer' | 'reader' | 'freeBusyReader';
  timeZone?: string;
}

export interface EventDateTime {
  dateTime?: string;
  date?: string;
  timeZone?: string;
}

export interface EventAttendee {
  email: string;
  displayName?: string;
  responseStatus?: 'needsAction' | 'declined' | 'tentative' | 'accepted';
  optional?: boolean;
  organizer?: boolean;
  self?: boolean;
}

export interface EventReminder {
  method: 'popup' | 'email';
  minutes: number;
}

export interface ConferenceData {
  createRequest?: {
    requestId: string;
    conferenceSolutionKey?: { type: string };
    status?: { statusCode: string };
  };
  entryPoints?: Array<{
    entryPointType: string;
    uri: string;
    label?: string;
  }>;
  conferenceSolution?: {
    key: { type: string };
    name: string;
    iconUri?: string;
  };
  conferenceId?: string;
}

export interface CalendarEvent {
  id: string;
  summary?: string;
  description?: string;
  start?: EventDateTime;
  end?: EventDateTime;
  location?: string;
  status?: 'confirmed' | 'tentative' | 'cancelled';
  htmlLink?: string;
  hangoutLink?: string;
  conferenceData?: ConferenceData;
  attendees?: EventAttendee[];
  organizer?: { email: string; displayName?: string; self?: boolean };
  creator?: { email: string; displayName?: string };
  eventType?:
    | 'default'
    | 'birthday'
    | 'focusTime'
    | 'fromGmail'
    | 'outOfOffice'
    | 'workingLocation';
  visibility?: 'default' | 'public' | 'private' | 'confidential';
  colorId?: string;
  recurringEventId?: string;
  recurrence?: string[];
  reminders?: {
    useDefault: boolean;
    overrides?: EventReminder[];
  };
  created?: string;
  updated?: string;
}

export interface FreeBusyResponse {
  timeMin: string;
  timeMax: string;
  calendars: Record<
    string,
    {
      busy: Array<{ start: string; end: string }>;
      errors?: Array<{ domain: string; reason: string }>;
    }
  >;
}

// ============================================================================
// Request Parameters
// ============================================================================

export interface ListEventsParams {
  calendarId?: string;
  timeMin?: string;
  timeMax?: string;
  maxResults?: number;
  singleEvents?: boolean;
  orderBy?: 'startTime' | 'updated';
  q?: string;
  eventTypes?: string[];
  pageToken?: string;
  showDeleted?: boolean;
}

export interface CreateEventParams {
  calendarId?: string;
  summary: string;
  description?: string;
  start: EventDateTime;
  end: EventDateTime;
  location?: string;
  attendees?: string[];
  addGoogleMeet?: boolean;
  recurrence?: string[];
  reminders?: { useDefault: boolean; overrides?: EventReminder[] };
  visibility?: 'default' | 'public' | 'private' | 'confidential';
  colorId?: string;
  sendUpdates?: 'all' | 'externalOnly' | 'none';
}

export interface QuickAddParams {
  calendarId?: string;
  text: string;
  sendUpdates?: 'all' | 'externalOnly' | 'none';
}

export interface UpdateEventParams {
  calendarId?: string;
  eventId: string;
  summary?: string;
  description?: string;
  start?: EventDateTime;
  end?: EventDateTime;
  location?: string;
  attendees?: string[];
  addGoogleMeet?: boolean;
  recurrence?: string[];
  reminders?: { useDefault: boolean; overrides?: EventReminder[] };
  visibility?: 'default' | 'public' | 'private' | 'confidential';
  colorId?: string;
  sendUpdates?: 'all' | 'externalOnly' | 'none';
}

export interface MoveEventParams {
  calendarId: string;
  eventId: string;
  destinationCalendarId: string;
  sendUpdates?: 'all' | 'externalOnly' | 'none';
}

export interface DeleteEventParams {
  calendarId?: string;
  eventId: string;
  sendUpdates?: 'all' | 'externalOnly' | 'none';
}

export interface RespondToEventParams {
  calendarId?: string;
  eventId: string;
  response: 'accepted' | 'declined' | 'tentative';
  sendUpdates?: 'all' | 'externalOnly' | 'none';
}

export interface FreeBusyParams {
  timeMin: string;
  timeMax: string;
  calendarIds?: string[];
  timeZone?: string;
}

// ============================================================================
// Client
// ============================================================================

export class GoogleCalendarClient {
  private accessToken: string;

  constructor(accessToken: string) {
    this.accessToken = accessToken;
  }

  /**
   * Normalize event ID by removing the calendar email suffix if present.
   * 
   * Google Calendar returns composite IDs like "baseEventId calendarEmail" (base64 encoded)
   * for shared/imported events. The update/delete APIs expect only the base event ID.
   * 
   * Examples of composite IDs (decoded):
   * - "bbmpmmmickpmipbanqoosctntc adam@overment.com" → "bbmpmmmickpmipbanqoosctntc"
   * - "abc123 someone@example.com" → "abc123"
   */
  private normalizeEventId(eventId: string): string {
    // Try to decode as base64 to check for email suffix
    try {
      const decoded = atob(eventId);
      // Check if decoded string contains " email@domain" pattern
      const emailSuffixMatch = decoded.match(/^(.+?)\s+[\w.+-]+@[\w.-]+$/);
      if (emailSuffixMatch && emailSuffixMatch[1]) {
        // Re-encode just the base ID part
        const baseId = emailSuffixMatch[1];
        return btoa(baseId);
      }
    } catch {
      // Not valid base64, use as-is
    }
    return eventId;
  }

  private async request<T>(path: string, options: RequestInit = {}): Promise<T> {
    const url = `${GOOGLE_CALENDAR_API_BASE}${path}`;
    const headers = {
      Authorization: `Bearer ${this.accessToken}`,
      'Content-Type': 'application/json',
      ...options.headers,
    };

    try {
      const response = await fetch(url, { ...options, headers });

      if (!response.ok) {
        let errorMessage = `Google Calendar API error: ${response.status} ${response.statusText}`;
        try {
          const errorData = (await response.json()) as { error?: { message?: string } };
          if (errorData.error?.message) {
            errorMessage += ` - ${errorData.error.message}`;
          }
        } catch {
          // Ignore JSON parse error
        }
        throw new Error(errorMessage);
      }

      // Handle 204 No Content
      if (response.status === 204) {
        return {} as T;
      }

      return (await response.json()) as T;
    } catch (error) {
      logger.error('google-calendar-client', {
        message: 'Request failed',
        url,
        error: (error as Error).message,
      });
      throw error;
    }
  }

  // --------------------------------------------------------------------------
  // Calendars
  // --------------------------------------------------------------------------

  async listCalendars(): Promise<{ items: CalendarListItem[] }> {
    return this.request('/users/me/calendarList');
  }

  // --------------------------------------------------------------------------
  // Events - Get Single
  // --------------------------------------------------------------------------

  async getEvent(calendarId: string, eventId: string): Promise<CalendarEvent> {
    const normalizedId = this.normalizeEventId(eventId);
    const path = `/calendars/${encodeURIComponent(calendarId)}/events/${encodeURIComponent(normalizedId)}`;
    return this.request(path);
  }

  // --------------------------------------------------------------------------
  // Events - List/Search
  // --------------------------------------------------------------------------

  /**
   * Validates that timestamp has a timezone suffix (required by Google Calendar API).
   * Throws if timezone is missing — caller must provide RFC3339 format.
   */
  private requireTimezone(timestamp: string): string {
    // Valid: ends with Z or has +/- offset like +01:00 or -07:00
    if (/Z$|[+-]\d{2}:\d{2}$/.test(timestamp)) {
      return timestamp;
    }
    throw new Error(
      `Invalid timestamp format: "${timestamp}". Must be RFC3339 with timezone (e.g., 2025-12-06T19:00:00Z or 2025-12-06T19:00:00+01:00)`,
    );
  }

  async listEvents(
    params: ListEventsParams,
  ): Promise<{ items: CalendarEvent[]; nextPageToken?: string }> {
    const calendarId = params.calendarId || 'primary';
    const queryParams = new URLSearchParams();

    if (params.timeMin) queryParams.set('timeMin', this.requireTimezone(params.timeMin));
    if (params.timeMax) queryParams.set('timeMax', this.requireTimezone(params.timeMax));
    if (params.maxResults) queryParams.set('maxResults', String(params.maxResults));
    if (params.singleEvents !== undefined)
      queryParams.set('singleEvents', String(params.singleEvents));
    if (params.orderBy) queryParams.set('orderBy', params.orderBy);
    if (params.q) queryParams.set('q', params.q);
    if (params.pageToken) queryParams.set('pageToken', params.pageToken);
    if (params.showDeleted) queryParams.set('showDeleted', String(params.showDeleted));

    // eventTypes can be repeated
    if (params.eventTypes && params.eventTypes.length > 0) {
      for (const et of params.eventTypes) {
        queryParams.append('eventTypes', et);
      }
    }

    const query = queryParams.toString();
    const path = `/calendars/${encodeURIComponent(calendarId)}/events${query ? `?${query}` : ''}`;

    return this.request(path);
  }

  // --------------------------------------------------------------------------
  // Events - Create
  // --------------------------------------------------------------------------

  async createEvent(params: CreateEventParams): Promise<CalendarEvent> {
    const calendarId = params.calendarId || 'primary';
    const queryParams = new URLSearchParams();

    if (params.sendUpdates) queryParams.set('sendUpdates', params.sendUpdates);
    if (params.addGoogleMeet) queryParams.set('conferenceDataVersion', '1');

    const body: Record<string, unknown> = {
      summary: params.summary,
      start: params.start,
      end: params.end,
    };

    if (params.description) body.description = params.description;
    if (params.location) body.location = params.location;
    if (params.visibility) body.visibility = params.visibility;
    if (params.colorId) body.colorId = params.colorId;
    if (params.recurrence) body.recurrence = params.recurrence;
    if (params.reminders) body.reminders = params.reminders;

    if (params.attendees && params.attendees.length > 0) {
      body.attendees = params.attendees.map((email) => ({ email }));
    }

    if (params.addGoogleMeet) {
      body.conferenceData = {
        createRequest: {
          requestId: `meet-${Date.now()}-${Math.random().toString(36).slice(2)}`,
          conferenceSolutionKey: { type: 'hangoutsMeet' },
        },
      };
    }

    const query = queryParams.toString();
    const path = `/calendars/${encodeURIComponent(calendarId)}/events${query ? `?${query}` : ''}`;

    return this.request(path, {
      method: 'POST',
      body: JSON.stringify(body),
    });
  }

  async quickAdd(params: QuickAddParams): Promise<CalendarEvent> {
    const calendarId = params.calendarId || 'primary';
    const queryParams = new URLSearchParams();

    queryParams.set('text', params.text);
    if (params.sendUpdates) queryParams.set('sendUpdates', params.sendUpdates);

    const path = `/calendars/${encodeURIComponent(calendarId)}/events/quickAdd?${queryParams.toString()}`;

    return this.request(path, { method: 'POST' });
  }

  // --------------------------------------------------------------------------
  // Events - Update
  // --------------------------------------------------------------------------

  async updateEvent(params: UpdateEventParams): Promise<CalendarEvent> {
    const calendarId = params.calendarId || 'primary';
    const queryParams = new URLSearchParams();

    if (params.sendUpdates) queryParams.set('sendUpdates', params.sendUpdates);
    if (params.addGoogleMeet) queryParams.set('conferenceDataVersion', '1');

    const body: Record<string, unknown> = {};

    if (params.summary !== undefined) body.summary = params.summary;
    if (params.description !== undefined) body.description = params.description;
    if (params.start !== undefined) body.start = params.start;
    if (params.end !== undefined) body.end = params.end;
    if (params.location !== undefined) body.location = params.location;
    if (params.visibility !== undefined) body.visibility = params.visibility;
    if (params.colorId !== undefined) body.colorId = params.colorId;
    if (params.recurrence !== undefined) body.recurrence = params.recurrence;
    if (params.reminders !== undefined) body.reminders = params.reminders;

    if (params.attendees !== undefined) {
      body.attendees = params.attendees.map((email) => ({ email }));
    }

    if (params.addGoogleMeet) {
      body.conferenceData = {
        createRequest: {
          requestId: `meet-${Date.now()}-${Math.random().toString(36).slice(2)}`,
          conferenceSolutionKey: { type: 'hangoutsMeet' },
        },
      };
    }

    const query = queryParams.toString();
    const normalizedId = this.normalizeEventId(params.eventId);
    const path = `/calendars/${encodeURIComponent(calendarId)}/events/${encodeURIComponent(normalizedId)}${query ? `?${query}` : ''}`;

    return this.request(path, {
      method: 'PATCH',
      body: JSON.stringify(body),
    });
  }

  async moveEvent(params: MoveEventParams): Promise<CalendarEvent> {
    const queryParams = new URLSearchParams();

    queryParams.set('destination', params.destinationCalendarId);
    if (params.sendUpdates) queryParams.set('sendUpdates', params.sendUpdates);

    const normalizedId = this.normalizeEventId(params.eventId);
    const path = `/calendars/${encodeURIComponent(params.calendarId)}/events/${encodeURIComponent(normalizedId)}/move?${queryParams.toString()}`;

    return this.request(path, { method: 'POST' });
  }

  // --------------------------------------------------------------------------
  // Events - Delete
  // --------------------------------------------------------------------------

  async deleteEvent(params: DeleteEventParams): Promise<void> {
    const calendarId = params.calendarId || 'primary';
    const queryParams = new URLSearchParams();

    if (params.sendUpdates) queryParams.set('sendUpdates', params.sendUpdates);

    const query = queryParams.toString();
    const normalizedId = this.normalizeEventId(params.eventId);
    const path = `/calendars/${encodeURIComponent(calendarId)}/events/${encodeURIComponent(normalizedId)}${query ? `?${query}` : ''}`;

    await this.request(path, { method: 'DELETE' });
  }

  // --------------------------------------------------------------------------
  // Events - Respond (Accept/Decline/Tentative)
  // --------------------------------------------------------------------------

  async respondToEvent(params: RespondToEventParams): Promise<CalendarEvent> {
    const calendarId = params.calendarId || 'primary';
    const normalizedId = this.normalizeEventId(params.eventId);

    // First, get the current event to find our attendee entry
    const event = await this.getEvent(calendarId, normalizedId);

    if (!event.attendees || event.attendees.length === 0) {
      throw new Error(
        'This event has no attendees. You can only respond to events you were invited to.',
      );
    }

    // Find the self attendee
    const selfAttendee = event.attendees.find((a) => a.self);
    if (!selfAttendee) {
      throw new Error(
        'You are not an attendee of this event. Cannot update response status.',
      );
    }

    // Update the attendee's response status
    const updatedAttendees = event.attendees.map((a) => {
      if (a.self) {
        return { ...a, responseStatus: params.response };
      }
      return a;
    });

    // PATCH the event with updated attendees
    const queryParams = new URLSearchParams();
    if (params.sendUpdates) queryParams.set('sendUpdates', params.sendUpdates);

    const query = queryParams.toString();
    const path = `/calendars/${encodeURIComponent(calendarId)}/events/${encodeURIComponent(normalizedId)}${query ? `?${query}` : ''}`;

    return this.request(path, {
      method: 'PATCH',
      body: JSON.stringify({ attendees: updatedAttendees }),
    });
  }

  // --------------------------------------------------------------------------
  // Free/Busy
  // --------------------------------------------------------------------------

  async getFreeBusy(params: FreeBusyParams): Promise<FreeBusyResponse> {
    const calendarIds = params.calendarIds || ['primary'];

    const body = {
      timeMin: this.requireTimezone(params.timeMin),
      timeMax: this.requireTimezone(params.timeMax),
      timeZone: params.timeZone,
      items: calendarIds.map((id) => ({ id })),
    };

    return this.request('/freeBusy', {
      method: 'POST',
      body: JSON.stringify(body),
    });
  }
}
