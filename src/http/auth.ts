import {
  type AuthInfo,
  type AuthMetadataOptions,
  buildOAuthProtectedResourceMetadata,
  getOAuthProtectedResourceMetadataUrl,
  requireBearerAuth,
} from '@modelcontextprotocol/server';
import type { AppConfig } from '../config/env.js';
import { createProviderTokenVerifier } from '../shared/auth/provider-token-verifier.js';
import type { TokenStore } from '../shared/storage/interface.js';

export interface AuthServices {
  gate: (request: Request) => Promise<AuthInfo | Response>;
  metadata: AuthMetadataOptions;
}

export function createAuthServices(
  config: AppConfig,
  store: TokenStore,
): AuthServices | undefined {
  if (!config.AUTH_ENABLED) return undefined;
  const issuer = config.OAUTH_ISSUER_URL.replace(/\/$/, '');
  const metadata: AuthMetadataOptions = {
    oauthMetadata: {
      issuer,
      authorization_endpoint: `${issuer}/authorize`,
      token_endpoint: `${issuer}/token`,
      revocation_endpoint: `${issuer}/revoke`,
      registration_endpoint: `${issuer}/register`,
      response_types_supported: ['code'],
      grant_types_supported: ['authorization_code', 'refresh_token'],
      code_challenge_methods_supported: ['S256'],
      token_endpoint_auth_methods_supported: ['none'],
      scopes_supported: config.OAUTH_REQUIRED_SCOPES,
    },
    resourceServerUrl: config.MCP_PUBLIC_URL,
    scopesSupported: config.OAUTH_REQUIRED_SCOPES,
    resourceName: config.MCP_TITLE,
    ...(config.MCP_WEBSITE_URL
      ? { serviceDocumentationUrl: config.MCP_WEBSITE_URL }
      : {}),
    dangerouslyAllowInsecureIssuerUrl: config.NODE_ENV !== 'production',
  };
  buildOAuthProtectedResourceMetadata(metadata);
  return {
    metadata,
    gate: requireBearerAuth({
      verifier: createProviderTokenVerifier(config, store),
      requiredScopes: config.OAUTH_REQUIRED_SCOPES,
      resourceMetadataUrl: getOAuthProtectedResourceMetadataUrl(config.MCP_PUBLIC_URL),
    }),
  };
}
