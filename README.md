# Google Calendar MCP Server

Fetch-native MCP server for listing calendars, searching and mutating events, responding to invitations, and checking Google Calendar availability. It runs on Bun and Cloudflare Workers and retains the project-specific Google OAuth proxy.

> This repository targets the prerelease `2026-07-28` MCP candidate with `@modelcontextprotocol/server@2.0.0-beta.5`. It does not claim final conformance until the dated specification and stable packages are published and verified.

## Tools

- `list_calendars`
- `search_events`
- `check_availability`
- `create_event`
- `update_event`
- `delete_event`
- `respond_to_event`

Existing all-calendar search, local substring matching, event-ID normalization, mutation fallback, free/busy behavior, Google Meet creation, and structured/natural-language event creation are preserved. Every tool advertises a complete Zod 4 input and output schema.

## Local setup

1. Enable Google Calendar API in Google Cloud.
2. Create a Web OAuth client.
3. Register `http://127.0.0.1:3001/oauth/callback` as a redirect URI.
4. Configure and start:

```bash
bun install
cp env.example .env
# Set PROVIDER_CLIENT_ID, PROVIDER_CLIENT_SECRET, and RS_TOKENS_ENC_KEY.
bun run dev
```

Local endpoints:

- MCP Resource Server: `http://127.0.0.1:3000/mcp`
- OAuth proxy: `http://127.0.0.1:3001`

## Credential model

Three credentials remain separate:

1. **MCP resource token** — the opaque bearer token accepted by `/mcp`.
2. **Google access token** — resolved from storage, refreshed when necessary, and exposed to a fresh server only as `AuthInfo.extra.providerAccessToken`.
3. **Google refresh token** — remains encrypted in file/KV storage and never enters MCP tool context.

Tools never read or forward `authInfo.token`. Token-separation tests use visibly different MCP and Google tokens and assert that only the Google token reaches the provider mock.

## OAuth proxy

The active proxy retains PKCE authorization, callback exchange, refresh, revocation, registration, redirect allowlists, and RS-to-Google token mappings.

Discovery endpoints:

- `/.well-known/oauth-protected-resource/mcp`
- `/.well-known/oauth-authorization-server`

Proxy endpoints:

- `GET /authorize`
- `GET /oauth/callback`
- `POST /token`
- `POST /revoke`
- `POST /register`

On Bun the proxy runs at `OAUTH_ISSUER_URL` (port 3001 by default). On Workers it is mounted outside `/mcp` on the same isolate.

## Cloudflare Workers

Create KV and update its ID in `wrangler.jsonc`:

```bash
bun run kv:create
bun x wrangler secret put PROVIDER_CLIENT_ID --config wrangler.jsonc
bun x wrangler secret put PROVIDER_CLIENT_SECRET --config wrangler.jsonc
bun x wrangler secret put RS_TOKENS_ENC_KEY --config wrangler.jsonc
bun run deploy
```

Before deployment, set HTTPS values for `MCP_PUBLIC_URL`, `OAUTH_ISSUER_URL`, `OAUTH_REDIRECT_URI`, Host/Origin allowlists, and the Google redirect URI. `RS_TOKENS_ENC_KEY` must be a base64url-encoded 32-byte key.

## Protocol and HTTP behavior

- One deployment-scoped fetch handler and a fresh `McpServer` per request
- Modern protocol pinned to candidate `2026-07-28`
- SDK stateless fallback for legacy `2025-11-25` clients
- `GET /mcp` and `DELETE /mcp` return `405`; no MCP session storage is used
- SDK-owned negotiation, header mismatch, cancellation, and transport errors
- Bounded MCP request bodies, strict Host/Origin checks, and allowlisted CORS headers
- Missing/invalid resource tokens return `401`; missing Google scopes return `403`

## Validate

```bash
bun test
bun run typecheck
bun run lint
bun run format:check
bun run build
bun run build:worker
bun run types:worker:check
```

Provider tests use mocks. A live OAuth flow requires Google credentials, user consent, and redirect URIs configured in Google Cloud.

## License

MIT
