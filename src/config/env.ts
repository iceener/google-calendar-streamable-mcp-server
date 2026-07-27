export type RuntimeEnvironment = 'development' | 'production' | 'test';
export type LegacyMode = 'stateless' | 'reject';
export type LogLevel = 'debug' | 'info' | 'warning' | 'error';

export interface AppConfig {
  HOST: string;
  PORT: number;
  NODE_ENV: RuntimeEnvironment;
  LOG_LEVEL: LogLevel;

  MCP_NAME: string;
  MCP_TITLE: string;
  MCP_VERSION: string;
  MCP_DESCRIPTION: string;
  MCP_INSTRUCTIONS: string;
  MCP_PUBLIC_URL: URL;
  MCP_WEBSITE_URL?: URL;
  MCP_ALLOWED_HOSTS: string[];
  MCP_ALLOWED_ORIGIN_HOSTNAMES: string[];
  MCP_LEGACY_MODE: LegacyMode;
  MCP_MAX_REQUEST_BYTES: number;

  AUTH_ENABLED: boolean;
  OAUTH_ISSUER_URL: string;
  OAUTH_SCOPES: string;
  OAUTH_REQUIRED_SCOPES: string[];
  OAUTH_REDIRECT_URI: string;
  OAUTH_REDIRECT_ALLOWLIST: string[];
  OAUTH_REDIRECT_ALLOW_ALL: boolean;
  OAUTH_EXTRA_AUTH_PARAMS?: string;
  OAUTH_AUTHORIZATION_URL: string;
  OAUTH_TOKEN_URL: string;
  OAUTH_REVOCATION_URL: string;

  PROVIDER_CLIENT_ID?: string;
  PROVIDER_CLIENT_SECRET?: string;
  PROVIDER_ACCOUNTS_URL: string;
  PROVIDER_API_URL: string;

  RS_TOKENS_FILE: string;
  RS_TOKENS_ENC_KEY?: string;
  RPS_LIMIT: number;
  CONCURRENCY_LIMIT: number;
}

function stringValue(env: Record<string, unknown>, key: string, fallback = ''): string {
  const value = env[key];
  return value === undefined || value === null || value === ''
    ? fallback
    : String(value).trim();
}

function booleanValue(env: Record<string, unknown>, key: string): boolean {
  const value = stringValue(env, key).toLowerCase();
  if (!value || ['0', 'false', 'no', 'off'].includes(value)) return false;
  if (['1', 'true', 'yes', 'on'].includes(value)) return true;
  throw new Error(`${key} must be true or false`);
}

function integerValue(
  env: Record<string, unknown>,
  key: string,
  fallback: number,
  minimum: number,
  maximum: number,
): number {
  const value = Number(stringValue(env, key, String(fallback)));
  if (!Number.isInteger(value) || value < minimum || value > maximum) {
    throw new Error(`${key} must be an integer between ${minimum} and ${maximum}`);
  }
  return value;
}

function listValue(
  env: Record<string, unknown>,
  key: string,
  fallback: string[] = [],
): string[] {
  const value = stringValue(env, key);
  if (!value) return [...fallback];
  const separator = value.includes(',') ? /\s*,\s*/ : /\s+/;
  return [
    ...new Set(
      value
        .split(separator)
        .map((item) => item.trim())
        .filter(Boolean),
    ),
  ];
}

function enumValue<T extends string>(
  env: Record<string, unknown>,
  key: string,
  values: readonly T[],
  fallback: T,
): T {
  const value = stringValue(env, key, fallback) as T;
  if (!values.includes(value)) {
    throw new Error(`${key} must be one of: ${values.join(', ')}`);
  }
  return value;
}

function urlValue(
  env: Record<string, unknown>,
  key: string,
  fallback?: string,
): URL | undefined {
  const value = stringValue(env, key, fallback);
  if (!value) return undefined;
  try {
    return new URL(value);
  } catch {
    throw new Error(`${key} must be an absolute URL`);
  }
}

function isLoopback(hostname: string): boolean {
  return ['localhost', '127.0.0.1', '[::1]', '::1'].includes(hostname);
}

function validateSecureUrl(
  url: URL,
  key: string,
  environment: RuntimeEnvironment,
): void {
  if (
    environment === 'production' &&
    url.protocol !== 'https:' &&
    !isLoopback(url.hostname)
  ) {
    throw new Error(`${key} must use HTTPS in production`);
  }
}

export function parseConfig(env: Record<string, unknown>): AppConfig {
  const port = integerValue(env, 'PORT', 3000, 1, 65_535);
  const environment = enumValue(
    env,
    'NODE_ENV',
    ['development', 'production', 'test'] as const,
    'development',
  );
  const configuredPublicUrl = stringValue(env, 'MCP_PUBLIC_URL');
  if (environment === 'production' && !configuredPublicUrl) {
    throw new Error('MCP_PUBLIC_URL is required in production');
  }
  const publicUrl = urlValue(
    env,
    'MCP_PUBLIC_URL',
    `http://localhost:${port}/mcp`,
  ) as URL;
  if (publicUrl.search || publicUrl.hash) {
    throw new Error('MCP_PUBLIC_URL must not include a query string or fragment');
  }
  validateSecureUrl(publicUrl, 'MCP_PUBLIC_URL', environment);

  const redirectUri = urlValue(
    env,
    'OAUTH_REDIRECT_URI',
    `http://127.0.0.1:${port + 1}/oauth/callback`,
  ) as URL;
  const issuerUrl = urlValue(env, 'OAUTH_ISSUER_URL', redirectUri.origin) as URL;
  issuerUrl.pathname = issuerUrl.pathname.replace(/\/$/, '');
  issuerUrl.search = '';
  issuerUrl.hash = '';

  const authorizationUrl = urlValue(
    env,
    'OAUTH_AUTHORIZATION_URL',
    'https://accounts.google.com/o/oauth2/v2/auth',
  ) as URL;
  const tokenUrl = urlValue(
    env,
    'OAUTH_TOKEN_URL',
    'https://oauth2.googleapis.com/token',
  ) as URL;
  const revocationUrl = urlValue(
    env,
    'OAUTH_REVOCATION_URL',
    'https://oauth2.googleapis.com/revoke',
  ) as URL;
  for (const [key, url] of [
    ['OAUTH_ISSUER_URL', issuerUrl],
    ['OAUTH_REDIRECT_URI', redirectUri],
    ['OAUTH_AUTHORIZATION_URL', authorizationUrl],
    ['OAUTH_TOKEN_URL', tokenUrl],
    ['OAUTH_REVOCATION_URL', revocationUrl],
  ] as Array<[string, URL]>) {
    validateSecureUrl(url, key, environment);
  }

  const defaultHosts = [publicUrl.hostname];
  if (environment !== 'production') {
    defaultHosts.push('localhost', '127.0.0.1', '[::1]');
  }
  const allowedHosts = listValue(env, 'MCP_ALLOWED_HOSTS', defaultHosts);
  const allowedOrigins = listValue(env, 'MCP_ALLOWED_ORIGIN_HOSTNAMES', defaultHosts);
  if (allowedHosts.length === 0 || allowedOrigins.length === 0) {
    throw new Error('MCP Host and Origin allowlists must not be empty');
  }

  const authEnabled = booleanValue(env, 'AUTH_ENABLED');
  const providerClientId = stringValue(env, 'PROVIDER_CLIENT_ID') || undefined;
  const providerClientSecret = stringValue(env, 'PROVIDER_CLIENT_SECRET') || undefined;
  if (
    authEnabled &&
    environment === 'production' &&
    (!providerClientId || !providerClientSecret)
  ) {
    throw new Error(
      'PROVIDER_CLIENT_ID and PROVIDER_CLIENT_SECRET are required when AUTH_ENABLED=true in production',
    );
  }

  const scopes = listValue(env, 'OAUTH_SCOPES', [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/calendar.readonly',
  ]);

  return {
    HOST: stringValue(env, 'HOST', '127.0.0.1'),
    PORT: port,
    NODE_ENV: environment,
    LOG_LEVEL: enumValue(
      env,
      'LOG_LEVEL',
      ['debug', 'info', 'warning', 'error'] as const,
      'info',
    ),
    MCP_NAME: stringValue(env, 'MCP_NAME', 'google-calendar-mcp'),
    MCP_TITLE: stringValue(env, 'MCP_TITLE', 'Google Calendar'),
    MCP_VERSION: stringValue(env, 'MCP_VERSION', '1.0.0'),
    MCP_DESCRIPTION: stringValue(
      env,
      'MCP_DESCRIPTION',
      'Manage Google Calendar events, calendars, invitations, and availability.',
    ),
    MCP_INSTRUCTIONS: stringValue(
      env,
      'MCP_INSTRUCTIONS',
      'Manage Google Calendar events. Search before mutating and verify changes afterward.',
    ),
    MCP_PUBLIC_URL: publicUrl,
    MCP_WEBSITE_URL: urlValue(env, 'MCP_WEBSITE_URL'),
    MCP_ALLOWED_HOSTS: allowedHosts,
    MCP_ALLOWED_ORIGIN_HOSTNAMES: allowedOrigins,
    MCP_LEGACY_MODE: enumValue(
      env,
      'MCP_LEGACY_MODE',
      ['stateless', 'reject'] as const,
      'stateless',
    ),
    MCP_MAX_REQUEST_BYTES: integerValue(
      env,
      'MCP_MAX_REQUEST_BYTES',
      1_048_576,
      1_024,
      10_485_760,
    ),
    AUTH_ENABLED: authEnabled,
    OAUTH_ISSUER_URL: issuerUrl.href.replace(/\/$/, ''),
    OAUTH_SCOPES: scopes.join(' '),
    OAUTH_REQUIRED_SCOPES: scopes,
    OAUTH_REDIRECT_URI: redirectUri.href,
    OAUTH_REDIRECT_ALLOWLIST: listValue(env, 'OAUTH_REDIRECT_ALLOWLIST'),
    OAUTH_REDIRECT_ALLOW_ALL: booleanValue(env, 'OAUTH_REDIRECT_ALLOW_ALL'),
    OAUTH_EXTRA_AUTH_PARAMS: stringValue(env, 'OAUTH_EXTRA_AUTH_PARAMS') || undefined,
    OAUTH_AUTHORIZATION_URL: authorizationUrl.href,
    OAUTH_TOKEN_URL: tokenUrl.href,
    OAUTH_REVOCATION_URL: revocationUrl.href,
    PROVIDER_CLIENT_ID: providerClientId,
    PROVIDER_CLIENT_SECRET: providerClientSecret,
    PROVIDER_ACCOUNTS_URL: stringValue(
      env,
      'PROVIDER_ACCOUNTS_URL',
      'https://accounts.google.com',
    ),
    PROVIDER_API_URL: stringValue(
      env,
      'PROVIDER_API_URL',
      'https://www.googleapis.com',
    ),
    RS_TOKENS_FILE: stringValue(env, 'RS_TOKENS_FILE', '.data/rs_tokens.json'),
    RS_TOKENS_ENC_KEY:
      stringValue(env, 'RS_TOKENS_ENC_KEY') ||
      stringValue(env, 'TOKENS_ENC_KEY') ||
      undefined,
    RPS_LIMIT: integerValue(env, 'RPS_LIMIT', 10, 1, 1_000),
    CONCURRENCY_LIMIT: integerValue(env, 'CONCURRENCY_LIMIT', 5, 1, 1_000),
  };
}

export type UnifiedConfig = AppConfig;
