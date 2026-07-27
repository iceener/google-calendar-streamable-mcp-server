import { parseConfig } from './config/env.js';
import { buildHttpApp } from './http/app.js';
import { buildOAuthApp } from './http/oauth-app.js';
import { FileTokenStore } from './shared/storage/file.js';
import { sharedLogger as logger } from './shared/utils/logger.js';

const config = parseConfig(process.env);
const tokenStore = new FileTokenStore(config.RS_TOKENS_FILE, config.RS_TOKENS_ENC_KEY);
const runtime = buildHttpApp(config, {
  runtimeName: 'bun',
  tokenStore,
});
const mcpServer = Bun.serve({
  hostname: config.HOST,
  port: config.PORT,
  fetch: runtime.fetch,
});

const issuerUrl = new URL(config.OAUTH_ISSUER_URL);
const oauthServer = config.AUTH_ENABLED
  ? Bun.serve({
      hostname: config.HOST,
      port: issuerUrl.port ? Number(issuerUrl.port) : config.PORT + 1,
      fetch: buildOAuthApp(config, tokenStore).fetch,
    })
  : undefined;

logger.info('server', {
  message: 'Google Calendar MCP server started',
  url: config.MCP_PUBLIC_URL.href,
  oauthIssuer: config.AUTH_ENABLED ? config.OAUTH_ISSUER_URL : undefined,
  protocol: '2026-07-28',
  legacyMode: config.MCP_LEGACY_MODE,
  authEnabled: config.AUTH_ENABLED,
});

let shuttingDown = false;
async function shutdown(signal: string): Promise<void> {
  if (shuttingDown) return;
  shuttingDown = true;
  logger.info('server', { message: 'Shutting down', signal });
  const stops = [mcpServer.stop(false), oauthServer?.stop(false)].filter(
    (stop): stop is Promise<void> => Boolean(stop),
  );
  await runtime.close();
  tokenStore.flush();
  tokenStore.stopCleanup();
  let timeout: ReturnType<typeof setTimeout> | undefined;
  const stopped = await Promise.race([
    Promise.all(stops).then(() => true),
    new Promise<false>((resolve) => {
      timeout = setTimeout(() => resolve(false), 5_000);
    }),
  ]);
  if (timeout) clearTimeout(timeout);
  if (!stopped) {
    await mcpServer.stop(true);
    if (oauthServer) await oauthServer.stop(true);
  }
}

process.once('SIGINT', () => void shutdown('SIGINT'));
process.once('SIGTERM', () => void shutdown('SIGTERM'));
