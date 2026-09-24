import { describe, it, expect } from 'vitest';
import {
  encryptPayload,
  decryptPayload,
  calculateSha256,
} from '../src/services/storage/cryptoBackup';

describe('cryptoBackup Web Crypto Core', () => {
  it('encrypts and decrypts payload correctly with password', async () => {
    const rawData = JSON.stringify({ message: 'Rahasia Korporat Sam Agent', timestamp: 123456789 });
    const password = 'KunciKuatRahasia123!';

    const { ciphertext, salt, iv, checksum } = await encryptPayload(rawData, password);

    expect(ciphertext).toBeDefined();
    expect(salt).toBeDefined();
    expect(iv).toBeDefined();
    expect(checksum).toBeDefined();

    const decrypted = await decryptPayload(ciphertext, password, salt, iv, checksum);
    expect(decrypted).toBe(rawData);
  });

  it('fails decryption when provided with wrong password', async () => {
    const rawData = 'Data penting';
    const password = 'PasswordBenar123';
    const { ciphertext, salt, iv, checksum } = await encryptPayload(rawData, password);

    await expect(
      decryptPayload(ciphertext, 'PasswordSalah!', salt, iv, checksum)
    ).rejects.toThrow();
  });

  it('detects corrupted or tampered ciphertext', async () => {
    const rawData = 'Data asli';
    const password = 'Password123';
    const { ciphertext, salt, iv, checksum } = await encryptPayload(rawData, password);

    const tamperedCiphertext = ciphertext.substring(0, ciphertext.length - 4) + 'AAAA';

    await expect(
      decryptPayload(tamperedCiphertext, password, salt, iv, checksum)
    ).rejects.toThrow();
  });

  it('calculates deterministic SHA-256 hash', async () => {
    const text = 'Hello Sam Office Agent';
    const hash1 = await calculateSha256(text);
    const hash2 = await calculateSha256(text);

    expect(hash1).toBe(hash2);
    expect(hash1.length).toBe(64); // 256 bits = 64 hex characters
  });
});
