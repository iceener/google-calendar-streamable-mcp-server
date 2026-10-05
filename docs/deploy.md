# Deploy

The production server is the Cloudflare Worker `google-calendar`, on `workers.dev`. Its URL is `MCP_PUBLIC_URL`, for example `https://google-calendar.<subdomain>.workers.dev/mcp`. For the template's general deployment notes, refer to the template's `docs/deploy.md`.

## Google

The server uses one Google OAuth client of type **Web application**, in a Google Cloud project with the **Google Calendar API** enabled.

- **Authorized redirect URI:** `/oauth/callback` on the origin of `MCP_PUBLIC_URL`. This URI is registered at Google. Do not change the path or the host. For local development, also add `http://127.0.0.1:3000/oauth/callback` (Bun) or `http://127.0.0.1:8787/oauth/callback` (`wrangler dev`).
- **Scopes:** `https://www.googleapis.com/auth/calendar.events` and `https://www.googleapis.com/auth/calendar.readonly`. They are in `src/services/google-oauth.ts`, and in `OAUTH_SCOPES` in `wrangler.production.jsonc`. The two lists must be the same; `tests/production.test.ts` checks this.
- The server asks for `access_type=offline` and `prompt=consent`, so Google always issues a refresh token.

## Cloudflare

`wrangler.jsonc` is for `bun run dev:worker`. It works only on loopback.

`wrangler.production.jsonc` is the production Wrangler config: the Worker name, the account, the variables, the KV namespace and the Durable Object. It is gitignored, because it names the deployed Worker and its storage. `wrangler.production.example.jsonc` has the same shape with stand-in values. To deploy your own copy, copy the example to `wrangler.production.jsonc` and put in your values. `bun run deploy` runs `wrangler deploy --config wrangler.production.jsonc`.

Where the real file exists, `tests/production.test.ts` checks it: it must have the example's shape, serve what the example serves, and match a pinned SHA-256 digest. A deliberate change to the file needs the new digest in that test, in the same commit. Without the file, those tests are skipped.

The production config has these bindings. Do not change them:

| Binding | Value |
|---|---|
| KV namespace `TOKENS` | The namespace that holds the signed-in users' tokens. Its ID is in `wrangler.production.jsonc`. |
| Durable Object `OAUTH_AUTHORITY` | class `NativeOAuthAuthority` |
| Migrations | `v1-native-oauth-authority` (applied). Only add new tags after it. |

The production secrets:

| Secret | Value |
|---|---|
| `PROVIDER_CLIENT_ID` | The Google OAuth client ID. |
| `PROVIDER_CLIENT_SECRET` | The Google OAuth client secret. |
| `TOKENS_ENC_KEY` | 32 random bytes, base64url. Do not change it: all stored tokens use it. |

The Worker does not start in production without all three. To set a secret, run `bunx wrangler secret put NAME --config wrangler.production.jsonc`. To make a key for a new deployment, run `openssl rand -base64 32 | tr -d '=' | tr '+/' '-_'`.

Two more secrets were deployed before 1.1.0, but no code reads them: `BEARER_TOKEN` and `RESEND_API_KEY`. `BEARER_TOKEN` is a name the template uses for `AUTH_MODE=bearer`; in OAuth mode the template ignores it. Remove both:

```sh
bunx wrangler secret delete BEARER_TOKEN --config wrangler.production.jsonc
bunx wrangler secret delete RESEND_API_KEY --config wrangler.production.jsonc
```

## Deploy

1. Run the checks:

   ```sh
   bun run check
   bun run test:smoke
   ```

2. Examine the bindings without a deploy:

   ```sh
   CLOUDFLARE_LOAD_DEV_VARS_FROM_DOT_ENV=false bunx wrangler deploy --dry-run --config wrangler.production.jsonc --outdir /tmp/google-calendar-mcp-dryrun
   ```

   The list must show `env.TOKENS`, with the namespace ID from `wrangler.production.jsonc`, and `env.OAUTH_AUTHORITY (NativeOAuthAuthority)`.

3. Deploy:

   ```sh
   bun run deploy
   ```

   A deploy replaces all vars with the vars in `wrangler.production.jsonc`. Vars that are not in the file are removed. Secrets stay.

4. Examine the deployed server:

   ```sh
   ORIGIN=$(bun -e "console.log(new URL(Bun.JSONC.parse(await Bun.file('wrangler.production.jsonc').text()).vars.MCP_PUBLIC_URL).origin)")
   curl -s "$ORIGIN/.well-known/oauth-authorization-server"
   curl -s -o /dev/null -w '%{http_code}\n' -X POST "$ORIGIN/mcp"
   ```

   The first command shows the issuer, `$ORIGIN`. The second shows `401`.

5. Connect a client and call `list_calendars`.

## Rollback

The version before 1.1.0 is `d3856e67-c312-411d-904a-fef62f6d24bd` (2026-10-02). 1.1.0 adds no migration and keeps the storage formats, so `bunx wrangler rollback` to that version reads the data that 1.1.0 wrote, and the reverse is also true. Do not roll back to a version before the `NativeOAuthAuthority` migration (published 2026-09-08 as `ad7f9e16-2b14-4497-bbc5-53cc96a9bdfa`): Durable Object migrations do not reverse, and the older code cannot read the authority's state. [docs/oauth.md](oauth.md#rollback) has the rules for the authority and the key.

## Logs

The Worker writes JSON logs to Workers Logs. A tool error that the model sees includes a reference; search the logs for it. OAuth failures that clients see as `503 server_error` are logged as `OAuth request failed`, with the path. A failed Google refresh is logged as `Google token refresh failed; using the current token`. A failed Calendar call is logged as `Google Calendar request failed`, with the method and the path, without the query.
