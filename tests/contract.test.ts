import { afterEach, describe, expect, test } from 'bun:test';
import { createCalendarClientFactory } from '../src/services/google-calendar';
import before from './fixtures/contract-before.json';
import {
  cleanup,
  connect,
  fakeFetch,
  memoryLogger,
  PUBLIC_URL,
  signedIn,
  testDeps,
} from './helpers';

afterEach(cleanup);

/**
 * The MCP contract clients relied on before the template migration, recorded from the deployed
 * code in `fixtures/contract-before.json`. Every tool field must match exactly: names, order,
 * titles, descriptions, annotations, input and output schemas. The few differences elsewhere
 * are deliberate and listed here (and in CHANGELOG.md).
 */

/** Google, answering as the recording's harness did. */
function recordedGoogle() {
  const logger = memoryLogger();
  const fetch = fakeFetch(async (request) => {
    const url = new URL(request.url);
    if (url.pathname === '/calendar/v3/users/me/calendarList') {
      return Response.json({
        items: [
          {
            id: 'someone@example.com',
            summary: 'Someone',
            primary: true,
            accessRole: 'owner',
            timeZone: 'Europe/Warsaw',
            backgroundColor: '#9fe1e7',
          },
        ],
      });
    }
    if (url.pathname === '/calendar/v3/freeBusy') {
      const query = (await request.json()) as { timeMin: string; timeMax: string };
      return Response.json({
        kind: 'calendar#freeBusy',
        timeMin: query.timeMin,
        timeMax: query.timeMax,
        calendars: {
          primary: { busy: [{ start: '2026-10-05T09:00:00Z', end: '2026-10-05T10:00:00Z' }] },
        },
      });
    }
    return undefined;
  });
  return testDeps({ logger, calendar: createCalendarClientFactory({ logger, fetch }) });
}

for (const era of ['modern', 'legacy'] as const) {
  describe(`${era} contract`, () => {
    const recorded = before[era] as (typeof before)['modern'];
    const recordedTools = recorded['tools/list'].tools as unknown as Array<Record<string, unknown>>;

    test('the same tools, in the same order, every field identical', async () => {
      const client = await connect({ era });
      const { tools } = await client.listTools();

      expect(tools.map((tool) => tool.name)).toEqual(
        recordedTools.map((tool) => String(tool.name)),
      );
      for (const [index, tool] of tools.entries()) {
        expect(tool as unknown).toEqual(recordedTools[index]);
      }
    });

    test('identity, capabilities and instructions, with the deliberate differences', async () => {
      const client = await connect({ era });
      const { name, title, description, icons } = recorded.serverInfo;

      expect(client.getServerVersion() as unknown).toEqual({
        name,
        title,
        description,
        // Changed: the version marks this release.
        version: '1.1.0',
        // The same icon, served from the public origin.
        icons: icons.map((icon) => ({ ...icon, src: new URL('/icon.svg', PUBLIC_URL).href })),
      });
      // Changed for 2026-07-28 clients: `listChanged: false`, as 2025-era clients already saw.
      // The list only changes on deploy, and nothing ever sent a notification.
      expect(recorded.capabilities).toEqual({ tools: { listChanged: era === 'modern' } });
      expect(client.getServerCapabilities()).toEqual({ tools: { listChanged: false } });
      expect(client.getInstructions()).toBe(recorded.instructions);
    });

    test('no resources, resource templates or prompts, as before', async () => {
      const client = await connect({ era });

      expect(recorded['resources/list'].resources).toEqual([]);
      expect(recorded['resources/templates/list'].resourceTemplates).toEqual([]);
      expect(recorded['prompts/list'].prompts).toEqual([]);
      expect((await client.listResources()).resources).toEqual([]);
      expect((await client.listResourceTemplates()).resourceTemplates).toEqual([]);
      expect((await client.listPrompts()).prompts).toEqual([]);
    });

    test('tool calls answer as before', async () => {
      const client = await connect({
        era,
        deps: recordedGoogle(),
        authInfo: signedIn('google-access'),
      });
      // `_meta` repeats the server identity; the recording leaves it out.
      const { _meta, ...calendars } = await client.callTool({
        name: 'list_calendars',
        arguments: {},
      });
      expect(calendars as unknown).toEqual(recorded.calls.list_calendars);

      const availability = await client.callTool({
        name: 'check_availability',
        arguments: { timeMin: '2026-10-05T00:00:00Z', timeMax: '2026-10-06T00:00:00Z' },
      });
      const old = recorded.calls.check_availability;
      expect(availability.structuredContent).toEqual(old.structuredContent);
      // The busy slots are written in the server's local time: workerd's is UTC.
      const local = (iso: string) => new Date(iso).toLocaleString();
      const inUtc = (text: string) =>
        text
          .replace(local('2026-10-05T09:00:00Z'), '10/5/2026, 9:00:00 AM')
          .replace(local('2026-10-05T10:00:00Z'), '10/5/2026, 10:00:00 AM');
      expect(inUtc((availability.content as Array<{ text: string }>)[0]?.text ?? '')).toBe(
        old.content[0]?.text as string,
      );
    });
  });
}

test('the tool list is cacheable by every caller: it is the same for everyone', async () => {
  expect(before.modern['tools/list']).toMatchObject({ ttlMs: 60_000, cacheScope: 'private' });
  const client = await connect();
  expect(await client.listTools()).toMatchObject({ ttlMs: 60_000, cacheScope: 'public' });
});
