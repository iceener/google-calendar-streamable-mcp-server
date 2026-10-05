import * as z from 'zod/v4';

/**
 * Settings your own code needs: API keys, feature flags, upstream URLs. They are read from
 * the environment and validated at startup together with the platform's configuration, so
 * one misconfigured deploy reports every problem at once. Code reads them as
 * `deps.config.settings`.
 *
 * On Workers, store secrets with `wrangler secret put NAME --config wrangler.production.jsonc`, not in `vars`.
 * docs/oauth.md describes how the OAuth proxy uses each of these. The secret names are the
 * deployed ones: renaming a secret signs every user out.
 */
export const Settings = z
  .object({
    PROVIDER_CLIENT_ID: z
      .string()
      .optional()
      .describe("Secret. The server's Google OAuth client ID."),
    PROVIDER_CLIENT_SECRET: z
      .string()
      .optional()
      .describe("Secret. The server's Google OAuth client secret."),
    TOKENS_ENC_KEY: z
      .string()
      .optional()
      .describe(
        'Secret. 32 random bytes, base64url: encrypts stored tokens. Changing it signs every user out.',
      ),
    PROXY_REDIRECT_ALLOWLIST: z
      .string()
      .optional()
      .transform((value) => [
        ...new Set(
          (value ?? '')
            .split(',')
            .map((entry) => entry.trim())
            .filter(Boolean),
        ),
      ])
      .describe(
        'Comma-separated client redirect URIs the OAuth proxy accepts, besides native loopback.',
      ),
    RS_TOKENS_FILE: z
      .string()
      .default('.data/rs_tokens.json')
      .describe('Bun only: the token file. The authority keeps its files next to it.'),
  })
  .refine((settings) => !settings.PROVIDER_CLIENT_ID === !settings.PROVIDER_CLIENT_SECRET, {
    message: 'PROVIDER_CLIENT_ID and PROVIDER_CLIENT_SECRET must be set together',
  });

export type Settings = z.infer<typeof Settings>;
