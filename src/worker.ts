import { preloadSchemas } from '@modelcontextprotocol/server';
import { parseConfig } from './config/env.js';
import { buildHttpApp, type HttpRuntime } from './http/app.js';
import { createEncryptor } from './shared/crypto/aes-gcm.js';
import type { TokenStore } from './shared/storage/interface.js';
import { KvTokenStore } from './shared/storage/kv.js';
import { MemoryTokenStore } from './shared/storage/memory.js';

preloadSchemas();

function createWorkerTokenStore(env: Env): TokenStore {
  const fallback = new MemoryTokenStore();
  const encryptionKey = env.RS_TOKENS_ENC_KEY;
  const encryptor =
    typeof encryptionKey === 'string' && encryptionKey
      ? createEncryptor(encryptionKey)
      : undefined;
  return new KvTokenStore(env.TOKENS, {
    fallback,
    ...(encryptor ? { encrypt: encryptor.encrypt, decrypt: encryptor.decrypt } : {}),
  });
}

export function createWorkerRuntime(env: Env): HttpRuntime {
  const config = parseConfig({ ...env });
  const tokenStore = createWorkerTokenStore(env);
  return buildHttpApp(config, {
    runtimeName: 'cloudflare-workers',
    tokenStore,
    mountOAuthProxy: true,
  });
}

let runtime: HttpRuntime | undefined;

export default {
  fetch(request, env) {
    runtime ??= createWorkerRuntime(env);
    return runtime.fetch(request);
  },
} satisfies ExportedHandler<Env>;
