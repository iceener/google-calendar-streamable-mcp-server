/**
 * AES-256-GCM encryption/decryption using Web Crypto API.
 * Works in both Cloudflare Workers and Node.js 18+.
 */

import { base64UrlDecode, base64UrlEncode } from '../utils/base64.js';

const ALGORITHM = 'AES-GCM';
const KEY_LENGTH = 256;
const IV_LENGTH = 12; // 96 bits recommended for GCM
const TAG_LENGTH = 128; // bits

/**
 * Derive a CryptoKey from a base64url-encoded secret.
 */
async function deriveKey(secret: string): Promise<CryptoKey> {
  const keyBytes = base64UrlDecode(secret);

  if (keyBytes.length !== 32) {
    throw new Error(`Invalid key length: expected 32 bytes, got ${keyBytes.length}`);
  }

  const rawKey = new Uint8Array(keyBytes.byteLength);
  rawKey.set(keyBytes);
  return crypto.subtle.importKey(
    'raw',
    rawKey,
    { name: ALGORITHM, length: KEY_LENGTH },
    false,
    ['encrypt', 'decrypt'],
  );
}

/**
 * Encrypt plaintext string using AES-256-GCM.
 *
 * @param plaintext - String to encrypt
 * @param secret - Base64url-encoded 32-byte secret key
 * @returns Base64url-encoded ciphertext (IV prepended)
 */
export async function encrypt(plaintext: string, secret: string): Promise<string> {
  const key = await deriveKey(secret);
  const iv = crypto.getRandomValues(new Uint8Array(IV_LENGTH));
  const plaintextBytes = new TextEncoder().encode(plaintext);

  const ciphertext = await crypto.subtle.encrypt(
    { name: ALGORITHM, iv, tagLength: TAG_LENGTH },
    key,
    plaintextBytes,
  );

  const combined = new Uint8Array(iv.length + ciphertext.byteLength);
  combined.set(iv, 0);
  combined.set(new Uint8Array(ciphertext), iv.length);

  return base64UrlEncode(combined);
}

/**
 * Decrypt ciphertext string using AES-256-GCM.
 *
 * @param ciphertext - Base64url-encoded ciphertext (IV prepended)
 * @param secret - Base64url-encoded 32-byte secret key
 * @returns Decrypted plaintext string
 */
export async function decrypt(ciphertext: string, secret: string): Promise<string> {
  const key = await deriveKey(secret);
  const combined = base64UrlDecode(ciphertext);

  if (combined.length < IV_LENGTH + 16) {
    throw new Error('Invalid ciphertext: too short');
  }

  const iv = combined.slice(0, IV_LENGTH);
  const encrypted = combined.slice(IV_LENGTH);

  const plaintextBytes = await crypto.subtle.decrypt(
    { name: ALGORITHM, iv, tagLength: TAG_LENGTH },
    key,
    encrypted,
  );

  return new TextDecoder().decode(plaintextBytes);
}

/**
 * Generate a random 32-byte (256-bit) key suitable for AES-256.
 * Returns base64url-encoded string.
 */
export function generateKey(): string {
  const bytes = crypto.getRandomValues(new Uint8Array(32));
  return base64UrlEncode(bytes);
}

/**
 * Create encryption/decryption functions bound to a specific key.
 * Useful for initializing KV stores.
 */
export function createEncryptor(secret: string): {
  encrypt: (plaintext: string) => Promise<string>;
  decrypt: (ciphertext: string) => Promise<string>;
} {
  return {
    encrypt: (plaintext: string) => encrypt(plaintext, secret),
    decrypt: (ciphertext: string) => decrypt(ciphertext, secret),
  };
}
