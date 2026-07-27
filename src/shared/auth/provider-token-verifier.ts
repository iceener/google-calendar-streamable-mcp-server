import {
  OAuthError,
  OAuthErrorCode,
  type OAuthTokenVerifier,
} from '@modelcontextprotocol/server';
import type { AppConfig } from '../../config/env.js';
import { buildProviderRefreshConfig, ensureFreshToken } from '../oauth/refresh.js';
import type { TokenStore } from '../storage/interface.js';

export const PROVIDER_ACCESS_TOKEN_EXTRA_KEY = 'providerAccessToken';

export function providerAccessTokenFromExtra(
  extra: Record<string, unknown> | undefined,
): string | undefined {
  const value = extra?.[PROVIDER_ACCESS_TOKEN_EXTRA_KEY];
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

export function createProviderTokenVerifier(
  config: AppConfig,
  store: TokenStore,
): OAuthTokenVerifier {
  const resource = new URL(config.MCP_PUBLIC_URL.href);
  resource.hash = '';
  const refreshConfig = buildProviderRefreshConfig(config);

  return {
    async verifyAccessToken(mcpResourceToken) {
      const { accessToken } = await ensureFreshToken(
        mcpResourceToken,
        store,
        refreshConfig,
      );
      const record = await store.getByRsAccess(mcpResourceToken);
      if (!record || !accessToken) {
        throw new OAuthError(
          OAuthErrorCode.InvalidToken,
          'MCP resource token is invalid or expired',
        );
      }
      const expiresAt = record.provider.expires_at
        ? Math.floor(record.provider.expires_at / 1_000)
        : Math.floor(Date.now() / 1_000) + 3_600;
      return {
        token: mcpResourceToken,
        clientId: 'google-calendar-oauth-client',
        scopes: record.provider.scopes ?? [],
        expiresAt,
        resource,
        extra: {
          [PROVIDER_ACCESS_TOKEN_EXTRA_KEY]: accessToken,
        },
      };
    },
  };
}
