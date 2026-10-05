import { createServer, type IncomingMessage } from 'node:http';
import type { AddressInfo } from 'node:net';

/**
 * Google on loopback, for the smoke tests: the OAuth token endpoint and the one Calendar call
 * the smoke client makes. It runs under Bun and Node. Requests for Google reach it through
 * `toMock`, so nothing leaves the machine.
 */
export const MOCK_CALENDAR = { id: 'smoke@example.com', summary: 'Smoke', accessRole: 'owner' };
export const GOOGLE_HOSTS = ['oauth2.googleapis.com', 'www.googleapis.com'];

export interface MockGoogle {
  origin: string;
  /** Grant types the token endpoint received, in order. */
  grants: string[];
  /** Hold Calendar's answers this long, to test quiet connections. */
  delayMs: number;
  close(): Promise<void>;
}

export async function startMockGoogle(clientId: string, clientSecret: string): Promise<MockGoogle> {
  const mock: MockGoogle = { origin: '', grants: [], delayMs: 0, close: async () => {} };
  const basic = `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString('base64')}`;

  const server = createServer(async (request, response) => {
    const send = (status: number, body: unknown) => {
      response.writeHead(status, { 'Content-Type': 'application/json' });
      response.end(JSON.stringify(body));
    };
    const url = new URL(request.url ?? '/', 'http://mock');
    if (request.method === 'POST' && url.pathname === '/token') {
      const form = new URLSearchParams(await text(request));
      if (request.headers.authorization !== basic) return send(401, { error: 'invalid_client' });
      mock.grants.push(form.get('grant_type') ?? '');
      return send(200, {
        access_token: 'google-access',
        refresh_token: 'google-refresh',
        expires_in: 3600,
        scope:
          'https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.readonly',
        token_type: 'Bearer',
      });
    }
    if (request.method === 'GET' && url.pathname === '/calendar/v3/users/me/calendarList') {
      if (request.headers.authorization !== 'Bearer google-access') return send(401, {});
      await new Promise((resolve) => setTimeout(resolve, mock.delayMs));
      return send(200, { items: [MOCK_CALENDAR] });
    }
    send(404, { error: 'not_found' });
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  mock.origin = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  mock.close = () =>
    new Promise((resolve) => {
      server.closeAllConnections();
      server.close(() => resolve());
    });
  return mock;
}

/** The same request, sent to the mock instead of Google. Other hosts are refused. */
export async function toMock(request: Request, mock: MockGoogle): Promise<Request> {
  const url = new URL(request.url);
  if (!GOOGLE_HOSTS.includes(url.hostname)) {
    throw new Error(`The smoke test blocked a request to ${url.origin}`);
  }
  const body = ['GET', 'HEAD'].includes(request.method) ? undefined : await request.text();
  return new Request(`${mock.origin}${url.pathname}${url.search}`, {
    method: request.method,
    headers: request.headers,
    ...(body !== undefined && { body }),
  });
}

function text(request: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    let body = '';
    request.on('data', (chunk) => {
      body += chunk;
    });
    request.on('end', () => resolve(body));
    request.on('error', reject);
  });
}
