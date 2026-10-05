import type { Logger } from '../platform/logger';

/**
 * Google Calendar API client (https://developers.google.com/workspace/calendar/api/v3/reference).
 * It acts with the user's Google access token, which the OAuth proxy resolved for this request;
 * never with the client's own MCP token. Failures throw `CalendarApiError` with the message
 * `Google Calendar API error: <status> <text> - <message>`; tools show it to the model.
 */
export const CALENDAR_API_BASE = 'https://www.googleapis.com/calendar/v3';

export class CalendarApiError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly statusText: string,
  ) {
    super(message);
    this.name = 'CalendarApiError';
  }

  get isNotFound(): boolean {
    return this.status === 404;
  }
}

export type SendUpdates = 'all' | 'externalOnly' | 'none';
export type Visibility = 'default' | 'public' | 'private' | 'confidential';

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
  dateTime?: string | undefined;
  date?: string | undefined;
  timeZone?: string | undefined;
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

export interface Reminders {
  useDefault: boolean;
  overrides?: EventReminder[] | undefined;
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
  conferenceData?: Record<string, unknown>;
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
  visibility?: Visibility;
  colorId?: string;
  recurringEventId?: string;
  recurrence?: string[];
  reminders?: Reminders;
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

export interface ListEventsParams {
  calendarId?: string | undefined;
  timeMin?: string | undefined;
  timeMax?: string | undefined;
  maxResults?: number | undefined;
  singleEvents?: boolean | undefined;
  orderBy?: 'startTime' | 'updated' | undefined;
  eventTypes?: string[] | undefined;
  pageToken?: string | undefined;
}

/** The fields `createEvent` and `updateEvent` send. `updateEvent` sends only those given. */
export interface EventFields {
  summary?: string | undefined;
  description?: string | undefined;
  start?: EventDateTime | undefined;
  end?: EventDateTime | undefined;
  location?: string | undefined;
  attendees?: string[] | undefined;
  addGoogleMeet?: boolean | undefined;
  recurrence?: string[] | undefined;
  reminders?: Reminders | undefined;
  visibility?: Visibility | undefined;
  colorId?: string | undefined;
}

export interface CreateEventParams extends EventFields {
  calendarId?: string | undefined;
  summary: string;
  start: EventDateTime;
  end: EventDateTime;
  sendUpdates?: SendUpdates | undefined;
}

export interface UpdateEventParams extends EventFields {
  calendarId?: string | undefined;
  eventId: string;
  sendUpdates?: SendUpdates | undefined;
}

export interface CalendarClientOptions {
  logger: Logger;
  /** Injected in tests. Defaults to the runtime's `fetch`. */
  fetch?: typeof fetch;
}

/** A client for one request: the user's access token and the call's cancellation signal. */
export type CalendarClientFactory = (accessToken: string, signal: AbortSignal) => CalendarClient;

export function createCalendarClientFactory(options: CalendarClientOptions): CalendarClientFactory {
  return (accessToken, signal) => new CalendarClient(accessToken, signal, options);
}

const RFC3339_WITH_ZONE = /Z$|[+-]\d{2}:\d{2}$/;

export class CalendarClient {
  constructor(
    private readonly accessToken: string,
    private readonly signal: AbortSignal | undefined,
    private readonly options: CalendarClientOptions,
  ) {}

  listCalendars(): Promise<{ items: CalendarListItem[] }> {
    return this.request('/users/me/calendarList');
  }

  getEvent(calendarId: string, eventId: string): Promise<CalendarEvent> {
    return this.request(eventPath(calendarId, normalizeEventId(eventId)));
  }

  async listEvents(
    params: ListEventsParams,
  ): Promise<{ items: CalendarEvent[]; nextPageToken?: string }> {
    const query = new URLSearchParams();
    if (params.timeMin) query.set('timeMin', requireTimezone(params.timeMin));
    if (params.timeMax) query.set('timeMax', requireTimezone(params.timeMax));
    if (params.maxResults) query.set('maxResults', String(params.maxResults));
    if (params.singleEvents !== undefined) query.set('singleEvents', String(params.singleEvents));
    if (params.orderBy) query.set('orderBy', params.orderBy);
    if (params.pageToken) query.set('pageToken', params.pageToken);
    for (const eventType of params.eventTypes ?? []) query.append('eventTypes', eventType);
    return this.request(`${eventsPath(params.calendarId)}${withQuery(query)}`);
  }

  createEvent(params: CreateEventParams): Promise<CalendarEvent> {
    const query = new URLSearchParams();
    if (params.sendUpdates) query.set('sendUpdates', params.sendUpdates);
    if (params.addGoogleMeet) query.set('conferenceDataVersion', '1');

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
    if (params.addGoogleMeet) body.conferenceData = meetRequest();

    return this.request(`${eventsPath(params.calendarId)}${withQuery(query)}`, {
      method: 'POST',
      body: JSON.stringify(body),
    });
  }

  /** Google parses `text` ("Lunch with Anna tomorrow at noon") into an event. */
  quickAdd(params: {
    calendarId?: string | undefined;
    text: string;
    sendUpdates?: SendUpdates | undefined;
  }): Promise<CalendarEvent> {
    const query = new URLSearchParams({ text: params.text });
    if (params.sendUpdates) query.set('sendUpdates', params.sendUpdates);
    return this.request(`${eventsPath(params.calendarId)}/quickAdd?${query}`, { method: 'POST' });
  }

  /** PATCH: only the fields given change. */
  updateEvent(params: UpdateEventParams): Promise<CalendarEvent> {
    const query = new URLSearchParams();
    if (params.sendUpdates) query.set('sendUpdates', params.sendUpdates);
    if (params.addGoogleMeet) query.set('conferenceDataVersion', '1');

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
    if (params.addGoogleMeet) body.conferenceData = meetRequest();

    const path = eventPath(params.calendarId, normalizeEventId(params.eventId));
    return this.request(`${path}${withQuery(query)}`, {
      method: 'PATCH',
      body: JSON.stringify(body),
    });
  }

  moveEvent(params: {
    calendarId: string;
    eventId: string;
    destinationCalendarId: string;
    sendUpdates?: SendUpdates | undefined;
  }): Promise<CalendarEvent> {
    const query = new URLSearchParams({ destination: params.destinationCalendarId });
    if (params.sendUpdates) query.set('sendUpdates', params.sendUpdates);
    const path = eventPath(params.calendarId, normalizeEventId(params.eventId));
    return this.request(`${path}/move?${query}`, { method: 'POST' });
  }

  async deleteEvent(params: {
    calendarId?: string | undefined;
    eventId: string;
    sendUpdates?: SendUpdates | undefined;
  }): Promise<void> {
    const query = new URLSearchParams();
    if (params.sendUpdates) query.set('sendUpdates', params.sendUpdates);
    const path = eventPath(params.calendarId, normalizeEventId(params.eventId));
    await this.request(`${path}${withQuery(query)}`, { method: 'DELETE' });
  }

  /** Set the signed-in user's own attendance: read the event, then PATCH its attendee list. */
  async respondToEvent(params: {
    calendarId?: string | undefined;
    eventId: string;
    response: 'accepted' | 'declined' | 'tentative';
    sendUpdates?: SendUpdates | undefined;
  }): Promise<CalendarEvent> {
    const calendarId = params.calendarId || 'primary';
    const eventId = normalizeEventId(params.eventId);
    const event = await this.getEvent(calendarId, eventId);

    if (!event.attendees || event.attendees.length === 0) {
      throw new Error(
        'This event has no attendees. You can only respond to events you were invited to.',
      );
    }
    if (!event.attendees.some((attendee) => attendee.self)) {
      throw new Error('You are not an attendee of this event. Cannot update response status.');
    }
    const attendees = event.attendees.map((attendee) =>
      attendee.self ? { ...attendee, responseStatus: params.response } : attendee,
    );

    const query = new URLSearchParams();
    if (params.sendUpdates) query.set('sendUpdates', params.sendUpdates);
    return this.request(`${eventPath(calendarId, eventId)}${withQuery(query)}`, {
      method: 'PATCH',
      body: JSON.stringify({ attendees }),
    });
  }

  async getFreeBusy(params: {
    timeMin: string;
    timeMax: string;
    calendarIds?: string[] | undefined;
    timeZone?: string | undefined;
  }): Promise<FreeBusyResponse> {
    const body = {
      timeMin: requireTimezone(params.timeMin),
      timeMax: requireTimezone(params.timeMax),
      timeZone: params.timeZone,
      items: (params.calendarIds || ['primary']).map((id) => ({ id })),
    };
    return this.request('/freeBusy', { method: 'POST', body: JSON.stringify(body) });
  }

  private async request<T>(path: string, init: RequestInit = {}): Promise<T> {
    const fetchImpl = this.options.fetch ?? fetch;
    try {
      const response = await fetchImpl(`${CALENDAR_API_BASE}${path}`, {
        ...init,
        headers: {
          Authorization: `Bearer ${this.accessToken}`,
          'Content-Type': 'application/json',
        },
        ...(this.signal && { signal: this.signal }),
      });
      if (!response.ok) {
        throw new CalendarApiError(
          await describeFailure(response),
          response.status,
          response.statusText,
        );
      }
      if (response.status === 204) return {} as T;
      return (await response.json()) as T;
    } catch (error) {
      if (!this.signal?.aborted) {
        // The path only: queries carry search terms and quick-add text.
        this.options.logger.warning('Google Calendar request failed', {
          method: init.method ?? 'GET',
          path: path.split('?')[0],
          error,
        });
      }
      throw error;
    }
  }
}

/**
 * Google returns composite IDs for shared and imported events: base64 of
 * "<base id> <calendar email>". Updates and deletes expect the base ID alone, re-encoded.
 */
export function normalizeEventId(eventId: string): string {
  try {
    const base = /^(.+?)\s+[\w.+-]+@[\w.-]+$/.exec(atob(eventId))?.[1];
    if (base) return btoa(base);
  } catch {
    // Not base64: a plain ID.
  }
  return eventId;
}

/** Google needs RFC 3339 timestamps with a zone: `2025-12-06T19:00:00Z` or `…+01:00`. */
function requireTimezone(timestamp: string): string {
  if (RFC3339_WITH_ZONE.test(timestamp)) return timestamp;
  throw new Error(
    `Invalid timestamp format: "${timestamp}". Must be RFC3339 with timezone (e.g., 2025-12-06T19:00:00Z or 2025-12-06T19:00:00+01:00)`,
  );
}

function eventsPath(calendarId: string | undefined): string {
  return `/calendars/${encodeURIComponent(calendarId || 'primary')}/events`;
}

function eventPath(calendarId: string | undefined, eventId: string): string {
  return `${eventsPath(calendarId)}/${encodeURIComponent(eventId)}`;
}

function withQuery(query: URLSearchParams): string {
  const text = query.toString();
  return text ? `?${text}` : '';
}

function meetRequest() {
  return {
    createRequest: {
      requestId: `meet-${crypto.randomUUID()}`,
      conferenceSolutionKey: { type: 'hangoutsMeet' },
    },
  };
}

async function describeFailure(response: Response): Promise<string> {
  let message = `Google Calendar API error: ${response.status} ${response.statusText}`;
  try {
    const data = (await response.json()) as { error?: { message?: string } };
    if (data.error?.message) message += ` - ${data.error.message}`;
  } catch {
    // Not JSON: the status line is all there is.
  }
  return message;
}
