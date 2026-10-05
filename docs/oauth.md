# The OAuth proxy

MCP clients cannot use Google's OAuth directly. A Google token is not issued for this server, and the MCP specification forbids passing a client's token to another API. So this server is the **authorization server** that MCP clients use. Behind it, the server signs the user in with Google and keeps the Google tokens. The template calls this pattern "Be your own authorization server" (template `docs/auth.md`).

The code in `src/oauth/` is the same in every proxy that uses this layout (gmail-mcp first). Only one small module describes the provider: `src/services/google-oauth.ts`.

## Layout

| File | Provider-specific? | Function |
|---|---|---|
| `src/services/google-oauth.ts` | **Yes** | Google's endpoints, the Calendar scopes, the offline-access parameters, and how to read Google's token response. |
| `src/oauth/proxy.ts` | No | `createOAuthProxy`: the verifier, the metadata and the routes. Checks that `OAUTH_*` name this server. Optionally serves the older discovery locations. |
| `src/oauth/routes.ts` | No | The HTTP endpoints, CORS, `no-store`, and the error answers. |
| `src/oauth/flow.ts` | No | Authorize, provider callback, token (code and refresh), registration. |
| `src/oauth/verifier.ts` | No | Checks the issued tokens for MCP requests, refreshes the provider token, and puts it in `authInfo.extra`. |
| `src/oauth/provider.ts` | No | The `OAuthProvider` interface, and the call to the provider's token endpoint. |
| `src/oauth/redirect-policy.ts` | No | Which client redirect URIs are allowed. |
| `src/oauth/authority.ts` | No | The state machine for registrations, transactions and codes. |
| `src/oauth/authority-do.ts` | No | The authority on Workers: the `NativeOAuthAuthority` Durable Object. |
| `src/oauth/authority-file.ts` | No | The authority on Bun: one file for each client. |
| `src/oauth/token-store*.ts` | No | Issued tokens: the interface and memory store, Workers KV, a file on Bun. |
| `src/oauth/crypto.ts`, `encoding.ts`, `errors.ts`, `input.ts` | No | AES-GCM, PKCE, base64, the error convention, request parsing. |

The proxy connects to the template only through the hooks in `src/server.ts`. `src/platform/` does not change.

| Template hook | What the proxy puts there |
|---|---|
| `Runtime` | `{ tokens, authority }`: KV and the Durable Object from `src/worker.ts`, files from `src/bun.ts`. |
| `Deps.oauth` | The proxy, made by `createOAuthProxy` in `createDeps`. |
| `createVerifier` | `deps.oauth.verifier`. |
| `oauthMetadata` | `deps.oauth.metadata(oauth)`: the published RFC 8414 document. |
| `routes` | `deps.oauth.mount(app, oauth, serverInfo.title)`, in OAuth mode only. |

The platform does the rest: it publishes the protected-resource document, checks every token's audience (`expectedResource`) and scopes (`OAUTH_SCOPES`), and removes the client's token before tools run.

## Options

`createOAuthProxy` has two options. The defaults are the stricter behavior. This server sets both, so that it behaves as it did before 1.1.0.

| Option | Default | This server | Effect |
|---|---|---|---|
| `bindGrants` | `true` | `false` | With `true`, each issued token record carries `oauth: { clientId, resource }`, and a refresh needs the same `client_id`. With `false`, records have no binding, and a refresh needs no `client_id`. |
| `discoveryAliases` | `true` | `false` | With `true`, the discovery documents are also served at `/.well-known/oauth-protected-resource`, `/mcp/.well-known/oauth-protected-resource` and `/mcp/.well-known/oauth-authorization-server`, the locations some 2025-era clients try. These are routes, so the Origin check applies. |

These are rules, not options:

- When the token store can't be reached (a KV outage), an MCP request gets `500`, so the client keeps its token and retries. A token that isn't stored, or a record that won't decrypt, gets `401`, and the client refreshes or signs in again.
- KV errors name only the operation and the kind of record (`KV write failed (refresh record)`): the KV keys contain the tokens, and KV's own messages can repeat the key.
- On Bun, records loaded from the token file stay for a week, even when the provider access token has expired, so the verifier can refresh them.

## The flow

1. **Register.** The client sends `POST /register` with its redirect URIs. Only public clients are accepted (`token_endpoint_auth_method: none`). The authority stores the client in a new Durable Object. The answer has a `client_id` of 32 characters. There is no registration access token.
2. **Authorize.** The client sends the user to `GET /authorize` with `client_id`, `redirect_uri`, `code_challenge` (S256 only), and optionally `state` and `resource`. If `resource` is given, it must be `MCP_PUBLIC_URL` exactly. The proxy stores a transaction for 10 minutes and sends the user to Google. Google receives the proxy's client ID, the proxy's callback, the Calendar scopes, `access_type=offline`, `prompt=consent`, and an opaque handle as `state`. Google never sees the client's redirect URI or state.
3. **Callback.** Google sends the user to `GET /oauth/callback?code=…&state=<handle>`. The proxy claims the transaction **before** it calls Google, so a replayed or concurrent callback fails. Then the proxy exchanges the code at Google's token endpoint (HTTP Basic authentication) and encrypts the Google tokens into a code for the client. The code is valid for 2 minutes. The proxy sends the user to the client's redirect URI, exactly as registered, with `code` and the client's `state`.
4. **Token.** The client sends `POST /token` with the code, `client_id`, `redirect_uri` and `code_verifier`. The authority redeems the code once, and only if all four match. The proxy issues two opaque tokens (32 characters each). A failed match leaves the code usable; a redeemed code never comes back, even if a later step fails.
5. **MCP requests.** The client sends `Authorization: Bearer <access token>`. The verifier finds the record. If the Google token expires in less than one minute, the verifier refreshes it at Google and stores it. The tools get the Google token as `authInfo.extra.providerAccessToken`.
6. **Refresh.** The client sends `POST /token` with `grant_type=refresh_token`. `client_id` is optional here: the records have no client binding. A `resource`, if given, must be `MCP_PUBLIC_URL`. The proxy refreshes the Google token if it expires in less than one minute. The access token changes only when Google returns a new refresh token; the refresh token never changes.

The records have no binding, so the verifier reports the caller ID as `rs:<SHA-256 of the token>`, and a record with an empty scope list counts as having the configured scopes.

## Endpoints

| Endpoint | Answers |
|---|---|
| `POST /register` | `201` with the client. `400 invalid_client_metadata` or `invalid_redirect_uri`. |
| `GET /authorize` | `302` to Google. `400` with `unsupported_response_type`, `invalid_client`, `invalid_request` or `invalid_target`. |
| `GET /oauth/callback` | `302` to the client. `400 invalid_callback` or `invalid_grant`. `503 server_error` if Google refuses the code. |
| `POST /token` | `200` with `access_token`, `refresh_token`, `token_type: bearer`, `expires_in`, `scope`. `400` with `invalid_grant`, `invalid_target`, `unsupported_grant_type`, `missing_refresh_token` or `invalid_request`. |
| `POST /revoke` | `200 {"status":"ok"}`. It changes nothing (see "Limits"). |
| `/.well-known/oauth-authorization-server` | The metadata, from any origin (the platform serves it). |
| `/.well-known/oauth-protected-resource/mcp` | The protected-resource metadata, from any origin (the platform serves it). |

Every OAuth answer has `Cache-Control: no-store`. An allowed browser origin (`MCP_ALLOWED_ORIGIN_HOSTNAMES`) can read the answers and send a CORS preflight. Bodies are limited to 16 KiB (`413 request_too_large`). A repeated field is `400 invalid_request`. Any other failure (storage, Google, a bug) is `503 server_error`; the details go to the log, not to the client.

Errors travel as `Error` messages that start with the OAuth error code. Only a plain message survives the Durable Object RPC boundary, so do not replace the convention with error classes.

## Redirect URIs

A client may use a redirect URI only if it registered the URI and one of these rules allows it:

1. **Native loopback.** A client with `application_type: native` may use `http://127.0.0.1:<port>/<path>`. At `/authorize` it may use another port, but the same path. The server refuses `localhost`, IPv6, numeric aliases of 127.0.0.1, user info, queries, fragments and dot segments. The raw spelling is kept, so `:80` stays.
2. **Allowlist.** An entry in `PROXY_REDIRECT_ALLOWLIST`: an exact URI (HTTPS or a custom scheme such as `alice://oauth/callback`), or an HTTPS origin (`https://host/`) that allows every path on that origin. Plain `http` entries never match: loopback uses rule 1 only. So the production entry `http://127.0.0.1:*/oauth/callback` has no effect; it is kept so that the value did not change.

The callback checks the rules again before it calls Google. If you remove an entry, transactions in progress for it fail.

## Storage

Do not change these formats. Records that exist in production use them.

**Issued tokens (Workers KV `TOKENS`).** Each record is written twice, as `rs:access:<access token>` and `rs:refresh:<refresh token>`:

```json
{
  "rs_access_token": "…",
  "rs_refresh_token": "…",
  "provider": { "access_token": "…", "refresh_token": "…", "expires_at": 1790000000000, "scopes": ["…"] },
  "created_at": 1790000000000
}
```

The value is AES-256-GCM with `TOKENS_ENC_KEY`: `base64url(iv[12] || ciphertext || tag[16])`. Without a key, the value is plain JSON; production refuses to start without one. Every write also goes to memory, so a failed KV write still works in the same isolate. KV is eventually consistent, so it never decides whether a code was used: the authority does.

**Authorization state (Durable Object `OAUTH_AUTHORITY`, class `NativeOAuthAuthority`, migration tag `v1-native-oauth-authority`).** There is one object for each client, named by its `client_id`, with one SQLite row (`authority` table, `id = 1`) that holds a JSON document: `{ client, transactions, codes }`. Each command reads, checks and writes in one `transactionSync`. A client has at most 100 open transactions and codes. An alarm removes expired entries. Inside a code, the Google tokens are `enc:` and the KV layout.

**Bun.** The tokens are in one file (`RS_TOKENS_FILE`, default `.data/rs_tokens.json`): `{ version: 1, encrypted, records }`, encrypted whole as `base64url(iv[12] || tag[16] || ciphertext)`. A record stays a week after its last change, also after its Google token expires. The authority keeps one file for each client in `<RS_TOKENS_FILE>.oauth/`. Only one Bun process may use these files.

`tests/fixtures/storage-before.json` holds records, a token file and an authority document that the code before 1.1.0 wrote. `tests/oauth/storage.test.ts` reads and refreshes them.

## Limits

- **Revocation does nothing.** `/revoke` answers `200` and keeps the tokens. Tokens end when the user removes the app at Google, or when the refresh at Google fails.
- **Refresh tokens are not bound to a client and are not rotated.** Anyone who has a refresh token can use it, with any `client_id`, until Google refuses the refresh. Replay of a stolen refresh token is not detected. Turning on `bindGrants` binds new records; check first that every client sends `client_id` on refresh.
- **Client ID metadata documents (CIMD) are not supported.** The metadata says `client_id_metadata_document_supported: false`. Clients must register.
- **No rate limit on `/register`.** Each registration makes a Durable Object.
- **Refresh races.** Two requests at the same moment can both refresh the Google token. Both results are valid.

## Rollback

A Durable Object migration does not reverse when you deploy older code. Keep the `NativeOAuthAuthority` class, the `OAUTH_AUTHORITY` binding and the migration list. Fix forward. Do not restore an authority backup: it can bring back codes that were already used. Do not change `TOKENS_ENC_KEY`: all stored tokens become unreadable, and all users must sign in again.

## Port another provider

Use this procedure to move a sibling proxy (Linear, Spotify) to this layout.

1. Record the "before" state from the checkpoint commit, as this repository does in `tests/fixtures/`: the MCP contract, the route matrix (in workerd, with the deployed vars), and a set of records that the old code wrote (KV, token file, an authority document).
2. Copy `src/oauth/` without changes.
3. Write `src/services/<provider>-oauth.ts`. It exports an `OAuthProvider`: `name`, `authorizationUrl`, `tokenUrl`, `scopes`, `authorizationParams` and `readTokens`. Use `readStandardTokenResponse` if the provider answers as RFC 6749 describes. If the scopes come back in another form (a comma-separated string, an array), write your own `readTokens`.
4. In `src/server.ts`, copy `Runtime`, `Deps.oauth`, the `createOAuthProxy` call in `createDeps`, and the three hooks from this repository. Choose the [options](#options) that match what the old server did: did its records carry a client binding, did it serve the older discovery locations, did its Bun file keep expired provider tokens.
5. In `src/worker.ts`, export `NativeOAuthAuthority` and build `Runtime` from `env.TOKENS` and `env.OAUTH_AUTHORITY`. In `src/bun.ts`, build it from `FileTokenStore` and `FileOAuthAuthority`, with the old file paths.
6. In `src/settings.ts`, declare the secrets **with their deployed names**. Do not rename a deployed secret. This server keeps `TOKENS_ENC_KEY`; gmail-mcp keeps `RS_TOKENS_ENC_KEY`. Pass the key to `createOAuthProxy` and to the stores.
7. In `wrangler.production.jsonc`, map the old vars as `CHANGELOG.md` shows for 1.1.0. Keep every applied migration tag, in order. Keep the KV namespace ID.
8. Copy `tests/oauth/`, `tests/contract.test.ts`, `tests/routes.test.ts`, `tests/production.test.ts` and the smoke scripts. Change the provider stand-in (`tests/oauth/fixture.ts`, `scripts/mock-google.ts`) to your provider's endpoints and answers.

Before you port, read the sibling's deployed bindings and vars (`bunx wrangler versions view <version> --name <worker>`). Look for differences, for example a different encryption-key secret name, more than one applied migration tag, an extra required-scopes setting, or provider settings with their own prefix.
