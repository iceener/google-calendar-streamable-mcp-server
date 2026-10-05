# Google Calendar MCP Server

This server is a remote [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server for Google Calendar. A model can use it to list calendars, find events, check free and busy times, create, change, move and delete events, and answer invitations. The server runs on **Cloudflare Workers** and on **Bun**. It uses the [MCP server template](https://github.com/iceener/streamable-mcp-server-template) 2.1 and the official MCP TypeScript SDK 2.3.0.

The server URL is the deployed Worker's `MCP_PUBLIC_URL`, for example `https://google-calendar.<subdomain>.workers.dev/mcp`.

The server uses protocol version `2026-07-28`. It also accepts clients that use the 2025 protocol versions.

> [!WARNING]
> You connect this server to your MCP client at your own risk. A model can make mistakes. Examine what the tools do, and examine the changes in your calendar. The tools can delete events and can send updates to attendees.

## Tools

| Tool | Function |
|---|---|
| `list_calendars` | Lists the calendars of the user, with their IDs, access roles and time zones. |
| `search_events` | Finds events in all readable calendars (the default), in one calendar, or in a list of calendars. It filters by time range, text and event type. The text match is a substring match on the title, description, location and attendees. |
| `check_availability` | Shows the busy times of one or more calendars in a time range. |
| `create_event` | Makes an event from a sentence (Google's quick add) or from fields: title, start, end, attendees, Google Meet, recurrence, reminders. |
| `update_event` | Changes the given fields of an event. It can also move the event to another calendar. |
| `delete_event` | Deletes an event. |
| `respond_to_event` | Sets the user's answer to an invitation: accepted, declined or tentative. |

By default, `create_event`, `update_event` and `delete_event` send no notification to attendees (`sendUpdates: none`). `respond_to_event` notifies everyone (`sendUpdates: all`). If Google cannot find an event under the owner's email address, the server tries again with `primary`.

## Connect a client

Use the server URL and the Streamable HTTP transport. The client signs you in with Google the first time.

| Client | Procedure |
|---|---|
| Claude | In **Settings → Connectors**, add a custom connector with the server URL. |
| Claude Code | Run `claude mcp add --transport http google-calendar <server URL>`. |
| MCP Inspector | Run `bun run inspector`. Select **Streamable HTTP**. Enter the server URL. |

Alice and Wonderlands use the same URL.

## Authentication

The server is its own OAuth authorization server, in front of Google's OAuth. MCP clients do not receive Google tokens.

1. The client registers at `/register` and signs the user in at `/authorize`. The client must use PKCE S256.
2. The server sends the user to Google. Google asks for the two Calendar scopes.
3. Google returns the user to the server's `/oauth/callback`. The server keeps the Google tokens, encrypted.
4. The client gets a code, and exchanges the code at `/token` for an access token and a refresh token. These tokens are opaque and have meaning only for this server.
5. On each MCP request, the server finds the Google token for the access token. If the Google token expires in less than one minute, the server refreshes it first. The tools use the Google token. They do not see the client's token.

For the complete flow, the storage and the redirect rules, refer to [docs/oauth.md](docs/oauth.md).

## Configuration

The deployment settings are in `wrangler.production.jsonc` (Workers; gitignored, with `wrangler.production.example.jsonc` as its shape) or `.env` (Bun). `.env.example` describes all variables.

| Variable | Production value | Function |
|---|---|---|
| `MCP_PUBLIC_URL` | `https://google-calendar.<subdomain>.workers.dev/mcp` | The public URL of the MCP endpoint. It is also the resource that tokens are issued for. |
| `MCP_ALLOWED_HOSTS` | `google-calendar.<subdomain>.workers.dev` | The `Host` headers that the server accepts. |
| `MCP_ALLOWED_ORIGIN_HOSTNAMES` | The server host, `claude.ai`, `claude.com`, and the hosts of the operator's own browser clients | The browser origins that the server accepts. |
| `AUTH_MODE` | `oauth` | The server checks the tokens that it issued. |
| `OAUTH_ISSUER_URL`, `OAUTH_AUTHORIZATION_URL`, `OAUTH_TOKEN_URL`, `OAUTH_REGISTRATION_URL` | This server's origin, `/authorize`, `/token` and `/register` | The authorization server that clients use. The server does not start if they name another server. |
| `OAUTH_SCOPES` | The two Calendar scopes | The scopes that every MCP request must have. |
| `PROXY_REDIRECT_ALLOWLIST` | The callbacks of Alice, Claude and the operator's other clients | The client redirect URIs that the proxy accepts, in addition to native loopback URIs. |
| `MCP_MAX_REQUEST_BYTES` | `1048576` | The largest MCP request body. |
| `MCP_LEGACY_MODE` | `stateless` | The server also accepts 2025-era clients. |

Secrets:

| Secret | Function |
|---|---|
| `PROVIDER_CLIENT_ID` | The server's Google OAuth client ID. |
| `PROVIDER_CLIENT_SECRET` | The server's Google OAuth client secret. |
| `TOKENS_ENC_KEY` | 32 random bytes, base64url. The key encrypts the stored tokens. If you change it, all users must sign in again. |

In production, the server does not start without these three secrets. The Google endpoints, the Calendar scopes and the offline-access parameters are constants in `src/services/google-oauth.ts`. For the setup at Google and at Cloudflare, refer to [docs/deploy.md](docs/deploy.md).

## Development

Requirements: [Bun](https://bun.sh) 1.4 or later, and Node.js 22.18 or later.

1. Install the dependencies:

   ```sh
   bun install
   ```

2. Copy `.env.example` to `.env`. Set `PROVIDER_CLIENT_ID`, `PROVIDER_CLIENT_SECRET` and `TOKENS_ENC_KEY`. At Google, allow the redirect URI `http://127.0.0.1:3000/oauth/callback`.
3. Start the server:

   ```sh
   bun run dev
   ```

   The server URL is `http://127.0.0.1:3000/mcp`. To use the Cloudflare local runtime, run `bun run dev:worker` (port 8787). Put its secrets in `.dev.vars`.

4. Before you commit, run the checks:

   ```sh
   bun run check
   bun run test:smoke
   ```

`bun run check` does the type check, the lint check and the tests. `bun run test:smoke` starts the real server on Bun and on workerd. It signs in through a Google stand-in on loopback and calls the tools. No test uses the network.

| Script | Function |
|---|---|
| `bun run dev` | Starts the server on Bun. |
| `bun run dev:worker` | Starts the server in the Cloudflare local runtime. |
| `bun run check` | Type check, lint check, tests, and the check of the generated Worker types. |
| `bun run test:smoke` | Smoke tests on Bun and on workerd. |
| `bun run deploy` | Deploys with the gitignored `wrangler.production.jsonc` (`wrangler deploy --config wrangler.production.jsonc`). |
| `bun run types:worker` | Makes the Worker types again after a change to `wrangler.jsonc`. |

## Project structure

```
src/
  server.ts       Identity, Runtime, Deps, and the hooks that connect the OAuth proxy
  settings.ts     The Google client, the encryption key, the redirect allowlist
  tools/          One file for each tool; shared/ has schemas and helpers
  services/       The Google Calendar API client, and google-oauth.ts: Google as the OAuth provider
  oauth/          The OAuth proxy. It is the same for every provider (docs/oauth.md)
  platform/       Template code. Do not change it
  bun.ts          Entry point for Bun: file storage
  worker.ts       Entry point for Workers: KV and the NativeOAuthAuthority Durable Object
tests/            Tests for the tools, the proxy and the platform; fixtures from before 1.1.0
scripts/          Smoke tests and the Google stand-in
```

For the template's concepts (the request path, `defineTool`, the error policy, the configuration checks), refer to the [template documentation](https://github.com/iceener/streamable-mcp-server-template#documentation).

## License

[MIT](LICENSE)
