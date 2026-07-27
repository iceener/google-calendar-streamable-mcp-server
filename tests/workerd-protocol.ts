import { Client, StreamableHTTPClientTransport } from '@modelcontextprotocol/client';

const port = 8793;
const endpoint = new URL(`http://127.0.0.1:${port}/mcp`);
const worker = Bun.spawn(
  [
    'bun',
    'x',
    'wrangler',
    'dev',
    '--local',
    '--config',
    'wrangler.jsonc',
    '--env-file',
    'tests/workerd.env',
    '--port',
    String(port),
  ],
  { stdout: 'ignore', stderr: 'ignore' },
);

async function waitForWorker(): Promise<void> {
  for (let attempt = 0; attempt < 80; attempt += 1) {
    try {
      const response = await fetch(`http://127.0.0.1:${port}/health`);
      if (response.ok) return;
    } catch {}
    await Bun.sleep(100);
  }
  throw new Error('workerd did not become ready');
}

async function verify(era: 'modern' | 'legacy'): Promise<void> {
  const client = new Client(
    { name: `calendar-workerd-${era}`, version: '1.0.0' },
    era === 'modern'
      ? { versionNegotiation: { mode: { pin: '2026-07-28' } } }
      : undefined,
  );
  try {
    await client.connect(new StreamableHTTPClientTransport(endpoint));
    if (client.getProtocolEra() !== era) throw new Error(`Expected ${era} era`);
    const tools = await client.listTools();
    if (tools.tools.length !== 7) throw new Error('Unexpected Calendar tools');
  } finally {
    await client.close();
  }
}

try {
  await waitForWorker();
  await verify('modern');
  await verify('legacy');
  console.log('workerd modern+legacy protocol checks passed');
} finally {
  worker.kill();
  await worker.exited;
}
