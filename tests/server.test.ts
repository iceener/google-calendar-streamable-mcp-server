import { afterEach, describe, expect, test } from 'bun:test';
import { prompts } from '../src/prompts';
import { resources } from '../src/resources';
import { serverInfo } from '../src/server';
import { tools } from '../src/tools';
import { cleanup, connect, PUBLIC_URL } from './helpers';

afterEach(cleanup);

/** `src/server.ts`: identity, capabilities, and that every listed definition is served. */
describe('server', () => {
  for (const era of ['modern', 'legacy'] as const) {
    test(`${era}: identifies itself with an icon served from the public origin`, async () => {
      const client = await connect({ era });

      expect(client.getServerVersion()).toMatchObject({
        ...serverInfo,
        icons: [
          {
            src: new URL('/icon.svg', PUBLIC_URL).href,
            mimeType: 'image/svg+xml',
            sizes: ['any'],
            theme: 'light',
          },
        ],
      });
      expect(client.getInstructions()).toBeString();
    });

    test(`${era}: advertises tools only, with no change stream`, async () => {
      const client = await connect({ era });

      // No prompts or resources yet; the capabilities appear when the index lists aren't empty.
      expect(client.getServerCapabilities()).toEqual({ tools: { listChanged: false } });
    });
  }

  test('serves every tool, resource and prompt in the index lists, in order', async () => {
    const client = await connect();

    expect((await client.listTools()).tools.map((tool) => tool.name)).toEqual(
      tools.map((tool) => tool.name),
    );
    const listed = [
      ...(await client.listResources()).resources.map((resource) => resource.name),
      ...(await client.listResourceTemplates()).resourceTemplates.map((template) => template.name),
    ];
    for (const resource of resources) expect(listed).toContain(resource.name);
    expect((await client.listPrompts()).prompts.map((prompt) => prompt.name)).toEqual(
      prompts.map((prompt) => prompt.name),
    );
  });

  test('lists are cacheable by every caller for a minute', async () => {
    const client = await connect();
    const result = await client.listTools();

    expect(result).toMatchObject({ ttlMs: 60_000, cacheScope: 'public' });
  });
});

test('oversized tool arguments are refused before any schema runs', async () => {
  const client = await connect();
  const result = await client.callTool({
    name: 'check_availability',
    arguments: {
      timeMin: '2026-10-05T00:00:00Z',
      timeMax: '2026-10-06T00:00:00Z',
      calendarIds: Array.from({ length: 2_000 }, (_, index) => `c${index}`),
    },
  });

  expect(result.isError).toBe(true);
  expect(JSON.stringify(result.content)).toContain('1000');
});
