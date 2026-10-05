import { afterEach, describe, expect, test } from 'bun:test';
import { Client, StreamableHTTPClientTransport } from '@modelcontextprotocol/client';
import { createCalendarClientFactory } from '../../src/services/google-calendar';
import { tools } from '../../src/tools';
import {
  cleanup,
  connect,
  type FetchRoute,
  fakeFetch,
  memoryLogger,
  signedIn,
  testDeps,
  textOf,
  track,
} from '../helpers';
import { codeFor, loopback, ORIGIN, proxyFixture, redeem, registerNative } from '../oauth/fixture';

afterEach(cleanup);

/** A client for one Google token, with Google Calendar answered by `route`. */
async function calendarCaller(route: FetchRoute, googleToken = 'google-a') {
  const fetch = fakeFetch(route);
  const deps = testDeps({
    calendar: createCalendarClientFactory({ logger: memoryLogger(), fetch }),
  });
  return { client: await connect({ deps, authInfo: signedIn(googleToken) }), fetch, deps };
}

/** "<METHOD> <path>?<query>" without the API base, for asserting what reached Google. */
function call(request: Request): string {
  const url = new URL(request.url);
  return `${request.method} ${url.pathname.replace('/calendar/v3', '')}${url.search}`;
}

const notFound = () => Response.json({ error: { message: 'Not Found' } }, { status: 404 });

const ARGUMENTS: Record<string, Record<string, unknown>> = {
  list_calendars: {},
  search_events: {},
  check_availability: { timeMin: '2026-10-05T00:00:00Z', timeMax: '2026-10-06T00:00:00Z' },
  create_event: { text: 'Lunch tomorrow' },
  update_event: { eventId: 'e1', summary: 'x' },
  delete_event: { eventId: 'e1' },
  respond_to_event: { eventId: 'e1', response: 'accepted' },
};

describe('every tool', () => {
  test('asks the user to sign in when there is no Google token', async () => {
    const client = await connect();
    for (const tool of tools) {
      const result = await client.callTool({ name: tool.name, arguments: ARGUMENTS[tool.name] });
      expect(result.isError).toBe(true);
      expect(textOf(result)).toBe(
        'Authentication required. Please authenticate with Google Calendar.',
      );
    }
  });
});

describe('through the OAuth proxy, over HTTP', () => {
  async function signIn(f: ReturnType<typeof proxyFixture>): Promise<string> {
    const uri = loopback();
    const client = await registerNative(f, uri);
    const { code } = await codeFor(f, client.client_id, uri);
    return (
      (await (await redeem(f, client.client_id, uri, code)).json()) as { access_token: string }
    ).access_token;
  }

  async function mcpClient(
    f: ReturnType<typeof proxyFixture>,
    token: string,
    era: 'modern' | 'legacy',
  ): Promise<Client> {
    const client = new Client(
      { name: 'test', version: '1.0.0' },
      { versionNegotiation: { mode: era === 'modern' ? 'auto' : 'legacy' } },
    );
    await client.connect(
      new StreamableHTTPClientTransport(new URL(`${ORIGIN}/mcp`), {
        fetch: async (url, init) => {
          const headers = new Headers(init?.headers);
          headers.set('Host', '127.0.0.1:3000');
          headers.set('Authorization', `Bearer ${token}`);
          const response = await f.app.fetch(new Request(String(url), { ...init, headers }));
          expect(response.headers.has('Mcp-Session-Id')).toBe(false);
          return response;
        },
      }),
    );
    return track(client);
  }

  const calendarList = () =>
    Response.json({ items: [{ id: 'primary', summary: 'Primary', accessRole: 'owner' }] });

  for (const era of ['modern', 'legacy'] as const) {
    test(`${era}: Google sees the user's Google token, never the MCP token`, async () => {
      const f = proxyFixture({ bindGrants: false, calendar: calendarList });
      const token = await signIn(f);
      const client = await mcpClient(f, token, era);
      const result = await client.callTool({ name: 'list_calendars', arguments: {} });

      expect(result.structuredContent).toEqual({
        items: [{ id: 'primary', summary: 'Primary', accessRole: 'owner' }],
      });
      expect(f.calendar.requests.map((request) => request.headers.get('Authorization'))).toEqual([
        'Bearer google-access',
      ]);
      expect(JSON.stringify(result)).not.toContain(token);
    });
  }

  test('an expiring Google token is refreshed before the tool runs', async () => {
    const f = proxyFixture({
      bindGrants: false,
      calendar: (request) =>
        request.headers.get('Authorization') === 'Bearer google-fresh'
          ? calendarList()
          : Response.json({}, { status: 401 }),
    });
    const token = await signIn(f);
    const record = await f.tokens.getByRsAccess(token);
    if (!record) throw new Error('expected a record');
    expect(record.oauth).toBeUndefined();
    await f.tokens.updateByRsRefresh(record.rs_refresh_token, {
      ...record.provider,
      expires_at: Date.now() + 10_000,
    });
    const client = await mcpClient(f, token, 'modern');
    const result = await client.callTool({ name: 'list_calendars', arguments: {} });

    expect(result.isError).toBeUndefined();
    expect(f.google.forms.map((form) => form.get('grant_type'))).toEqual([
      'authorization_code',
      'refresh_token',
    ]);
  });
});

describe('list_calendars', () => {
  test('concurrent callers each get their own calendars', async () => {
    const route: FetchRoute = (request) => {
      const who = request.headers.get('Authorization') === 'Bearer google-alice' ? 'Alice' : 'Bob';
      return Response.json({ items: [{ id: who, summary: who, accessRole: 'owner' }] });
    };
    const [alice, bob] = await Promise.all([
      calendarCaller(route, 'google-alice'),
      calendarCaller(route, 'google-bob'),
    ]);
    const [aliceResult, bobResult] = await Promise.all([
      alice.client.callTool({ name: 'list_calendars', arguments: {} }),
      bob.client.callTool({ name: 'list_calendars', arguments: {} }),
    ]);
    expect(aliceResult.structuredContent).toMatchObject({ items: [{ summary: 'Alice' }] });
    expect(bobResult.structuredContent).toMatchObject({ items: [{ summary: 'Bob' }] });
  });

  test('cancelling the call aborts the Google request', async () => {
    let aborted: () => void = () => {};
    const googleAborted = new Promise<void>((resolve) => {
      aborted = resolve;
    });
    const { client } = await calendarCaller(
      (request) =>
        new Promise<Response>((_resolve, reject) => {
          request.signal.addEventListener('abort', () => {
            aborted();
            reject(new DOMException('Aborted', 'AbortError'));
          });
        }),
    );
    const controller = new AbortController();
    const pending = client.callTool(
      { name: 'list_calendars', arguments: {} },
      { signal: controller.signal },
    );
    setTimeout(() => controller.abort(), 20);
    await expect(pending).rejects.toThrow();
    await googleAborted;
  });

  test("a Google failure is told to the model with Google's message", async () => {
    const { client } = await calendarCaller(() =>
      Response.json(
        { error: { message: 'Quota exceeded' } },
        { status: 403, statusText: 'Forbidden' },
      ),
    );
    const result = await client.callTool({ name: 'list_calendars', arguments: {} });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toBe(
      'Failed to list calendars: Google Calendar API error: 403 Forbidden - Quota exceeded',
    );
  });
});

describe('search_events', () => {
  const event = (id: string, summary: string, start: string, extra: object = {}) => ({
    id,
    summary,
    start: { dateTime: start },
    end: { dateTime: start },
    htmlLink: `https://calendar.google.com/${id}`,
    status: 'confirmed',
    ...extra,
  });

  test('searches every readable calendar, matches substrings, and merges by start time', async () => {
    const { client, fetch } = await calendarCaller((request) => {
      const url = new URL(request.url);
      if (url.pathname.endsWith('/calendarList')) {
        return Response.json({
          items: [
            { id: 'me@example.com', summary: 'Me', accessRole: 'owner' },
            { id: 'team@example.com', summary: 'Team', accessRole: 'reader' },
            { id: 'busy@example.com', summary: 'Busy only', accessRole: 'freeBusyReader' },
          ],
        });
      }
      if (url.pathname.includes('me%40example.com')) {
        return Response.json({
          items: [
            event('b', 'Barbershop', '2026-10-05T12:00:00Z', {
              organizer: { email: 'me@example.com', self: true },
            }),
            event('x', 'Dentist', '2026-10-05T08:00:00Z'),
          ],
        });
      }
      return Response.json({
        items: [event('a', 'Team barber day', '2026-10-05T09:00:00Z')],
        nextPageToken: 'ignored',
      });
    });
    const result = await client.callTool({
      name: 'search_events',
      arguments: { query: 'barber', timeMin: '2026-10-05T00:00:00Z' },
    });

    expect(result.structuredContent).toEqual({
      items: [
        {
          id: 'a',
          summary: 'Team barber day',
          start: { dateTime: '2026-10-05T09:00:00Z' },
          end: { dateTime: '2026-10-05T09:00:00Z' },
          htmlLink: 'https://calendar.google.com/a',
          status: 'confirmed',
          calendarId: 'team@example.com',
          calendarName: 'Team',
        },
        {
          id: 'b',
          summary: 'Barbershop',
          start: { dateTime: '2026-10-05T12:00:00Z' },
          end: { dateTime: '2026-10-05T12:00:00Z' },
          htmlLink: 'https://calendar.google.com/b',
          status: 'confirmed',
          organizer: { email: 'me@example.com', self: true },
          calendarId: 'me@example.com',
          calendarName: 'Me',
        },
      ],
      calendarsSearched: ['Me', 'Team'],
    });
    // The query is matched here: Google gets a larger page and no `q`.
    expect(fetch.requests.map(call)).toEqual([
      'GET /users/me/calendarList',
      'GET /calendars/me%40example.com/events?timeMin=2026-10-05T00%3A00%3A00Z&maxResults=250&singleEvents=true&orderBy=startTime',
      'GET /calendars/team%40example.com/events?timeMin=2026-10-05T00%3A00%3A00Z&maxResults=250&singleEvents=true&orderBy=startTime',
    ]);
    expect(textOf(result)).toContain('Searched 2 calendar(s): Me, Team');
  });

  test('one failing calendar is skipped; all failing is an error', async () => {
    const failing = await calendarCaller(() =>
      Response.json({ error: { message: 'Forbidden' } }, { status: 403, statusText: 'Forbidden' }),
    );
    const result = await failing.client.callTool({
      name: 'search_events',
      arguments: { calendarId: ['a@example.com', 'b@example.com'] },
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toBe(
      'Failed to search all 2 calendar(s): Google Calendar API error: 403 Forbidden - Forbidden',
    );
    expect(failing.deps.logs.map((entry) => entry.message)).toContain(
      'Could not search a calendar',
    );

    const partly = await calendarCaller((request) =>
      request.url.includes('a%40example.com')
        ? Response.json({ items: [event('1', 'One', '2026-10-05T09:00:00Z')] })
        : notFound(),
    );
    const partial = await partly.client.callTool({
      name: 'search_events',
      arguments: { calendarId: ['a@example.com', 'b@example.com'], fields: ['id'] },
    });
    expect(partial.structuredContent).toEqual({
      items: [{ id: '1' }],
      calendarsSearched: ['a@example.com'],
    });
  });

  test('a page token needs a single calendar, and the next one is offered', async () => {
    const { client } = await calendarCaller(() =>
      Response.json({
        items: [event('1', 'One', '2026-10-05T09:00:00Z')],
        nextPageToken: 'page-2',
      }),
    );
    const refused = await client.callTool({
      name: 'search_events',
      arguments: { pageToken: 'page-1' },
    });
    expect(textOf(refused)).toStartWith('Pagination (pageToken) only works');

    const single = await client.callTool({
      name: 'search_events',
      arguments: { calendarId: 'primary', pageToken: 'page-1' },
    });
    expect(single.structuredContent).toMatchObject({ nextPageToken: 'page-2' });
    expect(textOf(single)).toContain('Pass pageToken: "page-2"');
  });

  test('a timestamp without a zone is refused before Google is called', async () => {
    const { client, fetch } = await calendarCaller(() => undefined);
    const result = await client.callTool({
      name: 'search_events',
      arguments: { calendarId: 'primary', timeMin: '2026-10-05T09:00:00' },
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain('Timestamp must include timezone');
    expect(fetch.requests).toHaveLength(0);
  });
});

describe('check_availability', () => {
  test('sends the range and calendars, and lists the busy slots', async () => {
    const { client, fetch } = await calendarCaller(async (request) => {
      const body = (await request.json()) as Record<string, unknown>;
      expect(body).toEqual({
        timeMin: '2026-10-05T00:00:00Z',
        timeMax: '2026-10-06T00:00:00Z',
        items: [{ id: 'primary' }, { id: 'team@example.com' }],
      });
      return Response.json({
        timeMin: body.timeMin,
        timeMax: body.timeMax,
        calendars: {
          primary: { busy: [] },
          'team@example.com': { busy: [], errors: [{ domain: 'global', reason: 'notFound' }] },
        },
      });
    });
    const result = await client.callTool({
      name: 'check_availability',
      arguments: {
        timeMin: '2026-10-05T00:00:00Z',
        timeMax: '2026-10-06T00:00:00Z',
        calendarIds: ['primary', 'team@example.com'],
      },
    });
    expect(fetch.requests.map(call)).toEqual(['POST /freeBusy']);
    expect(textOf(result)).toContain('📅 primary: Completely free during this period ✓');
    expect(textOf(result)).toContain('📅 team@example.com: Error - notFound');
    expect(textOf(result)).toContain('✓ All calendars are free during this time range.');
  });
});

describe('create_event', () => {
  const created = () =>
    Response.json({ id: 'e1', summary: 'Planning', htmlLink: 'https://calendar.google.com/e1' });

  test('text without a summary goes to quick add', async () => {
    const { client, fetch } = await calendarCaller(created);
    const result = await client.callTool({
      name: 'create_event',
      arguments: { text: 'Lunch with Anna tomorrow at noon' },
    });
    expect(result.structuredContent).toMatchObject({ id: 'e1' });
    expect(fetch.requests.map(call)).toEqual([
      'POST /calendars/primary/events/quickAdd?text=Lunch+with+Anna+tomorrow+at+noon&sendUpdates=none',
    ]);
  });

  test('a structured all-day event with Meet, attendees and reminders', async () => {
    const { client, fetch } = await calendarCaller(created);
    await client.callTool({
      name: 'create_event',
      arguments: {
        summary: 'Planning',
        start: '2026-10-05',
        end: '2026-10-06',
        attendees: ['anna@example.com'],
        addGoogleMeet: true,
        reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 10 }] },
        calendarId: 'team@example.com',
      },
    });
    const [request] = fetch.requests;
    expect(call(request as Request)).toBe(
      'POST /calendars/team%40example.com/events?sendUpdates=none&conferenceDataVersion=1',
    );
    const body = (await (request as Request).json()) as Record<string, unknown>;
    expect(body).toMatchObject({
      summary: 'Planning',
      start: { date: '2026-10-05' },
      end: { date: '2026-10-06' },
      attendees: [{ email: 'anna@example.com' }],
      reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 10 }] },
      conferenceData: { createRequest: { conferenceSolutionKey: { type: 'hangoutsMeet' } } },
    });
  });

  test('timed events carry the time zone; missing fields are explained', async () => {
    const { client, fetch } = await calendarCaller(created);
    await client.callTool({
      name: 'create_event',
      arguments: {
        summary: 'Planning',
        start: '2026-10-05T09:00:00',
        end: '2026-10-05T10:00:00',
        timeZone: 'Europe/Warsaw',
      },
    });
    expect(await (fetch.requests[0] as Request).json()).toMatchObject({
      start: { dateTime: '2026-10-05T09:00:00', timeZone: 'Europe/Warsaw' },
    });

    const noSummary = await client.callTool({ name: 'create_event', arguments: {} });
    expect(textOf(noSummary)).toBe(
      "Either 'text' (for natural language) or 'summary' (for structured) is required.",
    );
    const noEnd = await client.callTool({
      name: 'create_event',
      arguments: { summary: 'x', start: '2026-10-05' },
    });
    expect(textOf(noEnd)).toBe("'start' and 'end' are required for structured event creation.");
  });

  test('attendee addresses are validated with the published pattern', async () => {
    const { client, fetch } = await calendarCaller(created);
    const result = await client.callTool({
      name: 'create_event',
      arguments: { summary: 'x', start: '2026-10-05', end: '2026-10-06', attendees: ['nope'] },
    });
    expect(result.isError).toBe(true);
    expect(fetch.requests).toHaveLength(0);
  });
});

describe('update_event and delete_event', () => {
  test('a 404 on the owner address is retried on primary', async () => {
    const { client, fetch } = await calendarCaller((request) =>
      request.url.includes('/calendars/primary/')
        ? request.method === 'DELETE'
          ? new Response(null, { status: 204 })
          : Response.json({ id: 'e1', summary: 'Renamed' })
        : notFound(),
    );
    const updated = await client.callTool({
      name: 'update_event',
      arguments: { eventId: 'e1', calendarId: 'me@example.com', summary: 'Renamed' },
    });
    expect(updated.structuredContent).toEqual({ id: 'e1', summary: 'Renamed' });
    const deleted = await client.callTool({
      name: 'delete_event',
      arguments: { eventId: 'e1', calendarId: 'me@example.com' },
    });
    expect(deleted.structuredContent).toEqual({
      success: true,
      eventId: 'e1',
      calendarId: 'primary',
    });
    expect(fetch.requests.map(call)).toEqual([
      'PATCH /calendars/me%40example.com/events/e1?sendUpdates=none',
      'PATCH /calendars/primary/events/e1?sendUpdates=none',
      'DELETE /calendars/me%40example.com/events/e1?sendUpdates=none',
      'DELETE /calendars/primary/events/e1?sendUpdates=none',
    ]);
  });

  test('moving first, then patching on the target calendar', async () => {
    const { client, fetch } = await calendarCaller(() =>
      Response.json({ id: 'e1', summary: 'Moved' }),
    );
    const result = await client.callTool({
      name: 'update_event',
      arguments: { eventId: 'e1', targetCalendarId: 'team@example.com', location: 'Room 2' },
    });
    expect(textOf(result)).toStartWith('✓ Event moved and updated: Moved');
    expect(fetch.requests.map(call)).toEqual([
      'POST /calendars/primary/events/e1/move?destination=team%40example.com&sendUpdates=none',
      'PATCH /calendars/team%40example.com/events/e1?sendUpdates=none',
    ]);
  });

  test('an update without changes is refused before Google is called', async () => {
    const { client, fetch } = await calendarCaller(() => undefined);
    const result = await client.callTool({ name: 'update_event', arguments: { eventId: 'e1' } });
    expect(textOf(result)).toBe(
      'No changes specified. Provide at least one field to update or a targetCalendarId to move.',
    );
    expect(fetch.requests).toHaveLength(0);
  });

  test('composite event IDs are reduced to the base ID', async () => {
    const { client, fetch } = await calendarCaller(() => new Response(null, { status: 204 }));
    const composite = btoa('abc123 someone@example.com');
    await client.callTool({ name: 'delete_event', arguments: { eventId: composite } });
    expect(call(fetch.requests[0] as Request)).toBe(
      `DELETE /calendars/primary/events/${btoa('abc123')}?sendUpdates=none`,
    );
  });
});

describe('respond_to_event', () => {
  test("sets only the user's own response, notifying everyone by default", async () => {
    const attendees = [
      { email: 'me@example.com', self: true, responseStatus: 'needsAction' },
      { email: 'anna@example.com', responseStatus: 'accepted' },
    ];
    const { client, fetch } = await calendarCaller((request) =>
      request.method === 'GET'
        ? Response.json({ id: 'e1', summary: 'Sync', attendees })
        : Response.json({ id: 'e1', summary: 'Sync' }),
    );
    const result = await client.callTool({
      name: 'respond_to_event',
      arguments: { eventId: 'e1', response: 'tentative' },
    });
    expect(result.structuredContent).toEqual({
      ok: true,
      action: 'respond_to_event',
      response: 'tentative',
      event: { id: 'e1', summary: 'Sync' },
    });
    expect(textOf(result)).toStartWith('✓ You marked as maybe: Sync');
    expect(fetch.requests.map(call)).toEqual([
      'GET /calendars/primary/events/e1',
      'PATCH /calendars/primary/events/e1?sendUpdates=all',
    ]);
    expect(await (fetch.requests[1] as Request).json()).toEqual({
      attendees: [{ ...attendees[0], responseStatus: 'tentative' }, attendees[1]],
    });
  });

  test('an event the user was not invited to is explained', async () => {
    const { client } = await calendarCaller(() =>
      Response.json({ id: 'e1', attendees: [{ email: 'anna@example.com' }] }),
    );
    const result = await client.callTool({
      name: 'respond_to_event',
      arguments: { eventId: 'e1', response: 'accepted' },
    });
    expect(textOf(result)).toBe(
      'Failed to respond to event: You are not an attendee of this event. Cannot update response status.',
    );
  });
});
