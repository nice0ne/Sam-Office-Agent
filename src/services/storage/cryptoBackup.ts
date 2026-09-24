/**
 * Web Crypto AES-256-GCM and PBKDF2 backup engine.
 * Provides client-side encrypted backup & restore for Sam Office Agent.
 */

export interface SamBackupEnvelope {
  format: 'sam-office-agent-backup';
  version: 1;
  encrypted: boolean;
  exportedAt: number;
  salt?: string;
  iv?: string;
  payload: string;
  checksum: string;
}

export interface SamBackupPayload {
  version: 1;
  exportedAt: number;
  environment: {
    appVersion: string;
    userAgent?: string;
  };
  data: {
    providers?: any[];
    activeProvider?: string;
    searchSettings?: any;
    soulConfig?: any;
    customTools?: any[];
    crossAppSnapshots?: any[];
  };
}

export interface RestoreResult {
  success: boolean;
  restoredAt: number;
  cleanRestore: boolean;
  summary: {
    providersCount: number;
    searchConfigured: boolean;
    soulRestored: boolean;
    customToolsCount: number;
    snapshotsCount: number;
  };
  error?: string;
}

/**
 * Access Web Crypto SubtleCrypto safely across browser and Node environments.
 */
function getSubtleCrypto(): SubtleCrypto {
  const subtle = globalThis.crypto?.subtle || (globalThis as any).crypto?.webcrypto?.subtle;
  if (!subtle) {
    throw new Error('Web Crypto API (SubtleCrypto) is not supported in this environment.');
  }
  return subtle;
}

/**
 * Access getRandomValues safely across browser and Node environments.
 */
function getRandomValues(array: Uint8Array): Uint8Array {
  const cryptoObj = globalThis.crypto || (globalThis as any).crypto?.webcrypto;
  if (!cryptoObj?.getRandomValues) {
    throw new Error('Web Crypto getRandomValues is not supported in this environment.');
  }
  return cryptoObj.getRandomValues(array);
}

/**
 * Helper to convert Uint8Array / ArrayBuffer to Base64 string.
 */
function arrayBufferToBase64(buffer: ArrayBuffer | Uint8Array): string {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  if (typeof btoa !== 'undefined') {
    let binary = '';
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  }
  const NodeBuffer = (globalThis as any).Buffer;
  if (NodeBuffer) {
    return NodeBuffer.from(bytes).toString('base64');
  }
  throw new Error('No base64 encoding mechanism available.');
}

/**
 * Helper to convert Base64 string to Uint8Array.
 */
function base64ToArrayBuffer(base64: string): Uint8Array {
  if (typeof atob !== 'undefined') {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }
  const NodeBuffer = (globalThis as any).Buffer;
  if (NodeBuffer) {
    return new Uint8Array(NodeBuffer.from(base64, 'base64'));
  }
  throw new Error('No base64 decoding mechanism available.');
}

/**
 * Calculates SHA-256 hash of a string and returns a 64-character lowercase hex string.
 */
export async function calculateSha256(data: string): Promise<string> {
  const subtle = getSubtleCrypto();
  const encoder = new TextEncoder();
  const dataBuffer = encoder.encode(data);
  const hashBuffer = await subtle.digest('SHA-256', dataBuffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
}

/**
 * Derives an AES-GCM 256-bit CryptoKey from a user password and salt using PBKDF2 with 100,000 iterations.
 */
export async function deriveKey(password: string, salt: Uint8Array): Promise<CryptoKey> {
  const subtle = getSubtleCrypto();
  const encoder = new TextEncoder();
  const passwordKey = await subtle.importKey(
    'raw',
    encoder.encode(password),
    'PBKDF2',
    false,
    ['deriveKey']
  );

  return subtle.deriveKey(
    {
      name: 'PBKDF2',
      salt: salt as unknown as BufferSource,
      iterations: 100000,
      hash: 'SHA-256',
    },
    passwordKey,
    {
      name: 'AES-GCM',
      length: 256,
    },
    false,
    ['encrypt', 'decrypt']
  );
}

/**
 * Encrypts a JSON payload string with a password using AES-256-GCM.
 * Generates a fresh 16-byte salt and 12-byte IV.
 * Returns base64 encoded ciphertext, salt, iv, and the SHA-256 checksum of the original unencrypted payload.
 */
export async function encryptPayload(
  payloadJson: string,
  password: string
): Promise<{ ciphertext: string; salt: string; iv: string; checksum: string }> {
  const subtle = getSubtleCrypto();
  const checksum = await calculateSha256(payloadJson);

  const salt = getRandomValues(new Uint8Array(16));
  const iv = getRandomValues(new Uint8Array(12));

  const key = await deriveKey(password, salt);
  const encoder = new TextEncoder();
  const encodedPayload = encoder.encode(payloadJson);

  const cipherBuffer = await subtle.encrypt(
    {
      name: 'AES-GCM',
      iv: iv as unknown as BufferSource,
    },
    key,
    encodedPayload
  );

  return {
    ciphertext: arrayBufferToBase64(cipherBuffer),
    salt: arrayBufferToBase64(salt),
    iv: arrayBufferToBase64(iv),
    checksum,
  };
}

/**
 * Decrypts an AES-256-GCM ciphertext using the password, salt, and iv.
 * If expectedChecksum is provided, validates that the decrypted payload matches the checksum.
 */
export async function decryptPayload(
  ciphertext: string,
  password: string,
  salt: string,
  iv: string,
  expectedChecksum?: string
): Promise<string> {
  const subtle = getSubtleCrypto();
  const saltBytes = base64ToArrayBuffer(salt);
  const ivBytes = base64ToArrayBuffer(iv);
  const cipherBytes = base64ToArrayBuffer(ciphertext);

  const key = await deriveKey(password, saltBytes);

  let decryptedBuffer: ArrayBuffer;
  try {
    decryptedBuffer = await subtle.decrypt(
      {
        name: 'AES-GCM',
        iv: ivBytes as unknown as BufferSource,
      },
      key,
      cipherBytes as unknown as BufferSource
    );
  } catch (_err) {
    throw new Error('Decryption failed: invalid password or corrupted data');
  }

  const decoder = new TextDecoder();
  const plainText = decoder.decode(decryptedBuffer);

  if (expectedChecksum) {
    const actualChecksum = await calculateSha256(plainText);
    if (actualChecksum !== expectedChecksum) {
      throw new Error(`Integrity check failed: expected checksum ${expectedChecksum} but got ${actualChecksum}`);
    }
  }

  return plainText;
}
