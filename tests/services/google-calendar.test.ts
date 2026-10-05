import { describe, expect, test } from 'bun:test';
import {
  CalendarApiError,
  createCalendarClientFactory,
  normalizeEventId,
} from '../../src/services/google-calendar';
import { type FetchRoute, fakeFetch, type LogEntry, memoryLogger } from '../helpers';

function client(route: FetchRoute, logs: LogEntry[] = [], signal = new AbortController().signal) {
  const fetch = fakeFetch(route);
  const calendar = createCalendarClientFactory({ logger: memoryLogger(logs), fetch })(
    'google-token',
    signal,
  );
  return { calendar, fetch };
}

describe('Google Calendar API client', () => {
  test('every request carries the user’s Google token and JSON', async () => {
    const { calendar, fetch } = client(() => Response.json({ items: [] }));
    await calendar.listCalendars();
    const [request] = fetch.requests;
    expect(request?.url).toBe('https://www.googleapis.com/calendar/v3/users/me/calendarList');
    expect(request?.headers.get('Authorization')).toBe('Bearer google-token');
    expect(request?.headers.get('Content-Type')).toBe('application/json');
  });

  test('list parameters, repeated event types and a 204 answer', async () => {
    const { calendar, fetch } = client((request) =>
      request.method === 'DELETE'
        ? new Response(null, { status: 204 })
        : Response.json({ items: [] }),
    );
    await calendar.listEvents({
      calendarId: 'a@example.com',
      timeMin: '2026-10-05T00:00:00+02:00',
      maxResults: 20,
      singleEvents: false,
      eventTypes: ['default', 'focusTime'],
    });
    expect(await calendar.deleteEvent({ eventId: 'e1' })).toBeUndefined();
    expect(fetch.requests.map((request) => request.url)).toEqual([
      'https://www.googleapis.com/calendar/v3/calendars/a%40example.com/events?timeMin=2026-10-05T00%3A00%3A00%2B02%3A00&maxResults=20&singleEvents=false&eventTypes=default&eventTypes=focusTime',
      'https://www.googleapis.com/calendar/v3/calendars/primary/events/e1',
    ]);
  });

  test('timestamps without a zone are refused before any request', async () => {
    const { calendar, fetch } = client(() => undefined);
    await expect(
      calendar.getFreeBusy({ timeMin: '2026-10-05T09:00:00', timeMax: '2026-10-05T10:00:00Z' }),
    ).rejects.toThrow('Invalid timestamp format: "2026-10-05T09:00:00"');
    expect(fetch.requests).toHaveLength(0);
  });

  test('Google’s error reaches the tool; the query stays out of the log', async () => {
    const logs: LogEntry[] = [];
    const { calendar } = client(
      () =>
        Response.json(
          { error: { message: 'Invalid Credentials' } },
          { status: 401, statusText: 'Unauthorized' },
        ),
      logs,
    );
    const failure = calendar.quickAdd({ text: 'Dinner with a secret person' });
    await expect(failure).rejects.toThrow(
      'Google Calendar API error: 401 Unauthorized - Invalid Credentials',
    );
    await expect(failure).rejects.toBeInstanceOf(CalendarApiError);
    expect(logs).toEqual([
      expect.objectContaining({
        level: 'warning',
        fields: expect.objectContaining({ path: '/calendars/primary/events/quickAdd' }),
      }),
    ]);
    expect(JSON.stringify(logs)).not.toContain('secret');
  });

  test('a 404 is recognisable, for the fallback to primary', async () => {
    const { calendar } = client(() => new Response('gone', { status: 404 }));
    const error = await calendar.getEvent('primary', 'e1').catch((caught: unknown) => caught);
    expect(error).toBeInstanceOf(CalendarApiError);
    expect((error as CalendarApiError).isNotFound).toBe(true);
    expect((error as Error).message).toBe('Google Calendar API error: 404 ');
  });

  test('the call’s signal reaches the request', async () => {
    const controller = new AbortController();
    let seen: AbortSignal | undefined;
    const { calendar } = client(
      (request) => {
        seen = request.signal;
        return Response.json({ items: [] });
      },
      [],
      controller.signal,
    );
    await calendar.listCalendars();
    controller.abort();
    expect(seen?.aborted).toBe(true);
  });
});

describe('event IDs', () => {
  test('a composite ID becomes its base ID; anything else is kept', () => {
    expect(normalizeEventId(btoa('bbmpmmmickpmipbanqoosctntc adam@example.net'))).toBe(
      btoa('bbmpmmmickpmipbanqoosctntc'),
    );
    expect(normalizeEventId('abc123')).toBe('abc123');
    expect(normalizeEventId('not base64!')).toBe('not base64!');
  });
});
