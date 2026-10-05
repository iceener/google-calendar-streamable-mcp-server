# Changelog

## 1.1.0 — 2026-10-05

The server now follows the [MCP server template](https://github.com/iceener/streamable-mcp-server-template) 2.1 and uses `@modelcontextprotocol/server` 2.3.0. The OAuth proxy is the shared one from gmail-mcp (`src/oauth/`). The OAuth behavior, the stored tokens and the tool contract are the same as before. `tests/fixtures/` records the deployed code (commit `180f48b`); `tests/contract.test.ts`, `tests/routes.test.ts` and `tests/oauth/storage.test.ts` compare the server with it.

### Changed for clients

- **Server identity:** version `1.1.0`. The name, title, description, icon and instructions are the same.
- **Capabilities:** `tools.listChanged` is `false` for 2026-07-28 clients too (it was `true`, but nothing ever sent a notification). 2025-era clients already saw `false`.
- **List caching:** `tools/list` and `server/discover` are cacheable by every caller (`cacheScope: public`, was `private`). The lists are the same for every caller.
- **Authorization server metadata:** adds `client_id_metadata_document_supported: false`. Clients treated a missing field the same way: they register at `/register`. All other fields are the same.
- **OAuth endpoints and CORS:** an allowed browser origin (`MCP_ALLOWED_ORIGIN_HOSTNAMES`) can now send a CORS preflight to `/register`, `/token`, `/authorize`, `/revoke` and `/oauth/callback` (was `404`), and can read the answers (`Access-Control-Allow-Origin`). Other origins still get `403`.
- **Unknown paths** answer `404 {"error":"not_found"}` as JSON, not as plain text. `/health` reports `{status, name, version}`.
- **A Worker that is configured wrong** answers `500 {"error":"server_misconfigured"}` to every request (was `503 {"error":"server_error"}`).

Not changed: the tool names, order, titles, descriptions, annotations, input schemas and output schemas; the public URLs; the Google callback `/oauth/callback`; the redirect allowlist; refresh without `client_id`.
- **Token store outage:** when KV can't be reached, an MCP request gets `500` and the client keeps its token, instead of `401`.

### Changed for operators

- `wrangler.jsonc` for development, and the gitignored `wrangler.production.jsonc` for production, with `wrangler.production.example.jsonc` as its committed shape. `bun run deploy` is `wrangler deploy --config wrangler.production.jsonc`. A deploy replaces the vars; it no longer keeps vars that are not in the file (`--keep-vars`).
- Compatibility date `2026-09-08`.
- Bun serves MCP and OAuth on one port. The separate OAuth listener on `PORT + 1` is gone.
- A Google refresh that the verifier stored is used for the request at once. Before, the verifier read the record again and could see an old expiry.
- OAuth failures that clients see as `503 server_error` are logged. Calendar failures are logged without the query string (search text, quick-add text).

### Environment variables

| Before | Now |
|---|---|
| `AUTH_ENABLED=true`, `AUTH_STRATEGY=oauth` | `AUTH_MODE=oauth` |
| `OAUTH_ISSUER_URL` | Same name and value |
| — | `OAUTH_AUTHORIZATION_URL`, `OAUTH_TOKEN_URL`, `OAUTH_REGISTRATION_URL`: this server's `/authorize`, `/token`, `/register` |
| `OAUTH_AUTHORIZATION_URL`, `OAUTH_TOKEN_URL` (Google's) | Constants in `src/services/google-oauth.ts` |
| `OAUTH_EXTRA_AUTH_PARAMS` | Constant `authorizationParams` in `src/services/google-oauth.ts` |
| `OAUTH_SCOPES` | Same name and value. It is the scopes that every request needs, as before. The scopes requested at Google are the constant `CALENDAR_SCOPES`; `tests/production.test.ts` checks that the two are equal. |
| `OAUTH_REDIRECT_ALLOWLIST` | `PROXY_REDIRECT_ALLOWLIST` (same value) |
| `OAUTH_REDIRECT_URI` | Removed. It only added one entry to the allowlist, and its production value is in the allowlist. |
| `OAUTH_REDIRECT_ALLOW_ALL` | Removed. No redirect check read it, and production refused `true`. |
| `OAUTH_REVOCATION_URL`, `PROVIDER_ACCOUNTS_URL`, `PROVIDER_API_URL` | Removed. Google's endpoints were already explicit; `/revoke` never called Google. |
| `MCP_NAME`, `MCP_TITLE`, `MCP_VERSION`, `MCP_DESCRIPTION`, `MCP_INSTRUCTIONS`, `MCP_WEBSITE_URL` | `serverInfo` and `instructions` in `src/server.ts` (same values) |
| `MCP_PROTOCOL_VERSION`, `AUTH_ALLOW_DIRECT_BEARER`, `AUTH_REQUIRE_RS`, `RPS_LIMIT`, `CONCURRENCY_LIMIT` | Removed. They were never read. |
| `RS_TOKENS_ENC_KEY` (preferred over `TOKENS_ENC_KEY`) | Removed. Only `TOKENS_ENC_KEY`, the deployed secret, is read. On Bun, rename it in `.env`. |

Unchanged: the secrets `PROVIDER_CLIENT_ID`, `PROVIDER_CLIENT_SECRET` and `TOKENS_ENC_KEY`, and `MCP_PUBLIC_URL`, `MCP_ALLOWED_HOSTS`, `MCP_ALLOWED_ORIGIN_HOSTNAMES`, `MCP_LEGACY_MODE`, `MCP_MAX_REQUEST_BYTES`, `LOG_LEVEL`, `NODE_ENV`, `HOST`, `PORT` and `RS_TOKENS_FILE` (default `.data/rs_tokens.json`).

The deployed secrets `BEARER_TOKEN` and `RESEND_API_KEY` are not read; `docs/deploy.md` shows how to remove them.

### Removed

- The old adapters, `core/runtime.ts`, and the hand-written HTTP, security, body, logger and configuration modules. The template's `src/platform/` replaces them.
- The copy of the OAuth proxy in `src/shared/`. The shared `src/oauth/` replaces it, with options that keep this server's behavior (`docs/oauth.md`).
- `refreshProviderToken` in `shared/oauth/refresh.ts`, a second implementation of the Google token call. The verifier and `/token` share one.
- The KV transaction and code methods. The authority Durable Object replaced them before this release.
- `docs/calendar-api.md` (a copy of Google's documentation), `docs/tools.md`, `env.example`, `docs/MCP_SDK_COMPATIBILITY.md`, `docs/MCP_2026_DEPLOYMENT.md`, `docs/ALICE_NATIVE_OAUTH.md` and `tests/contract-initial.json`. The facts that still apply are in `README.md`, `docs/oauth.md` and `docs/deploy.md`.
- The `wrangler.production.jsonc`, `wrangler.toml` and `scripts/types.env` files.
- The dependencies `@cloudflare/workers-types`, `@types/node` and `bun-types`. The generated Worker types and `@types/bun` replace them.
