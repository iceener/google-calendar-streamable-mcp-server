import { afterEach, describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createEncryptor } from '../src/shared/crypto/aes-gcm.js';
import { FileTokenStore } from '../src/shared/storage/file.js';
import { KvTokenStore } from '../src/shared/storage/kv.js';

const temporaryDirectories: string[] = [];

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) {
    rmSync(directory, { recursive: true, force: true });
  }
});

function encryptionKey(): string {
  return Buffer.from(crypto.getRandomValues(new Uint8Array(32))).toString('base64url');
}

describe('provider token storage and encryption', () => {
  test('file storage keeps refresh tokens encrypted after access-token expiry', async () => {
    const directory = mkdtempSync(join(tmpdir(), 'google-calendar-mcp-'));
    temporaryDirectories.push(directory);
    const path = join(directory, 'tokens.json');
    const key = encryptionKey();
    const first = new FileTokenStore(path, key);
    await first.storeRsMapping(
      'mcp-resource-token',
      {
        access_token: 'expired-google-access',
        refresh_token: 'persistent-google-refresh',
        expires_at: Date.now() - 1_000,
        scopes: ['calendar'],
      },
      'mcp-refresh-token',
    );
    first.flush();
    first.stopCleanup();

    const ciphertext = readFileSync(path, 'utf8');
    expect(ciphertext).not.toContain('expired-google-access');
    expect(ciphertext).not.toContain('persistent-google-refresh');

    const restored = new FileTokenStore(path, key);
    const record = await restored.getByRsAccess('mcp-resource-token');
    expect(record?.provider.refresh_token).toBe('persistent-google-refresh');
    restored.stopCleanup();
  });

  test('KV storage encrypts RS-to-provider records and reads them back', async () => {
    const values = new Map<string, string>();
    const kv = {
      async get(key: string): Promise<string | null> {
        return values.get(key) ?? null;
      },
      async put(key: string, value: string): Promise<void> {
        values.set(key, value);
      },
      async delete(key: string): Promise<void> {
        values.delete(key);
      },
    };
    const encryptor = createEncryptor(encryptionKey());
    const first = new KvTokenStore(kv, encryptor);
    await first.storeRsMapping(
      'mcp-token',
      {
        access_token: 'google-token',
        refresh_token: 'google-refresh',
        expires_at: Date.now() + 60_000,
      },
      'mcp-refresh',
    );
    expect(values.get('rs:access:mcp-token')).not.toContain('google-refresh');

    const second = new KvTokenStore(kv, encryptor);
    const restored = await second.getByRsAccess('mcp-token');
    expect(restored?.provider.access_token).toBe('google-token');
    expect(restored?.provider.refresh_token).toBe('google-refresh');
  });
});
