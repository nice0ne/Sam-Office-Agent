# Encrypted Backup & Restore Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement an enterprise-grade, client-side encrypted backup and restore engine (AES-256-GCM + PBKDF2) to safely export and restore all LLM configs, search settings, `SOUL.md` directives, custom tools, and cross-app snapshots.

**Architecture:** A zero-dependency cryptographic module (`src/services/storage/cryptoBackup.ts`) utilizing the standard Web Crypto API (`window.crypto.subtle`) for PBKDF2 key derivation (100,000 iterations), AES-256-GCM encryption/decryption, and SHA-256 checksums. A backup manager (`src/services/storage/backupManager.ts`) gathers data from all storage modules and handles smart merge or clean restore. An intuitive management panel in `SettingsModal.tsx` provides export/restore actions with password prompts.

**Tech Stack:** TypeScript, Web Crypto API (SubtleCrypto), React, LocalStorage, Vitest.

**Spec:** `docs/superpowers/specs/2026-09-24-encrypted-backup-restore-design.md`

## Global Constraints
- Zero external npm cryptographic dependencies (uses browser-native SubtleCrypto).
- Must function identically in both Node.js (test environment with globalThis.crypto) and browser/Office Add-in WebView2.
- Support both encrypted (`.sam-backup`) and plain JSON export modes.
- Full integrity verification using AES-GCM authentication tags and SHA-256 checksums.
- 100% test pass rate on `npm run build` and `npx vitest run`.

---

### Task 1: Web Crypto Cryptographic Core (`src/services/storage/cryptoBackup.ts`)

**Files:**
- Create: `src/services/storage/cryptoBackup.ts`
- Create: `tests/cryptoBackup.test.ts`

**Interfaces:**
- Produces:
  - `SamBackupEnvelope`
  - `SamBackupPayload`
  - `RestoreResult`
  - `encryptPayload(payloadJson: string, password: string): Promise<{ ciphertext: string; salt: string; iv: string; checksum: string }>`
  - `decryptPayload(ciphertext: string, password: string, salt: string, iv: string, expectedChecksum?: string): Promise<string>`
  - `calculateSha256(data: string): Promise<string>`

- [ ] **Step 1: Write the failing test**

Create `tests/cryptoBackup.test.ts`:

```typescript
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/cryptoBackup.test.ts`
Expected: FAIL with "Cannot find module '../src/services/storage/cryptoBackup'".

- [ ] **Step 3: Write minimal implementation**

Create `src/services/storage/cryptoBackup.ts`:
- Define interfaces: `SamBackupEnvelope`, `SamBackupPayload`, `RestoreResult`.
- Helper functions: `arrayBufferToBase64`, `base64ToArrayBuffer`.
- Helper to get crypto subtle: `const subtle = (globalThis.crypto?.subtle || (globalThis as any).crypto?.webcrypto?.subtle);`
- `calculateSha256(data: string): Promise<string>`:
  - Uses `subtle.digest('SHA-256', ...)`.
  - Converts ArrayBuffer to hex string.
- `deriveKey(password: string, salt: Uint8Array): Promise<CryptoKey>`:
  - `subtle.importKey('raw', ...)`
  - `subtle.deriveKey({ name: 'PBKDF2', salt, iterations: 100000, hash: 'SHA-256' }, ...)`
- `encryptPayload(payloadJson: string, password: string)`:
  - Generates 16-byte salt, 12-byte IV.
  - Derives key, runs `subtle.encrypt({ name: 'AES-GCM', iv }, key, encodedText)`.
  - Returns `{ ciphertext: base64, salt: base64, iv: base64, checksum }`.
- `decryptPayload(ciphertext: string, password: string, salt: string, iv: string, expectedChecksum?: string)`:
  - Derives key from password & decoded salt.
  - Runs `subtle.decrypt({ name: 'AES-GCM', iv: decodedIv }, key, decodedCiphertext)`.
  - Decodes UTF-8 string.
  - If `expectedChecksum`, computes SHA-256 and validates match.
  - Returns decrypted plain string.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/cryptoBackup.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/storage/cryptoBackup.ts tests/cryptoBackup.test.ts
git commit -m "feat(crypto): implement Web Crypto AES-256-GCM and PBKDF2 backup engine"
```

---

### Task 2: Backup & Restore Manager (`src/services/storage/backupManager.ts`)

**Files:**
- Create: `src/services/storage/backupManager.ts`
- Create: `tests/backupManager.test.ts`

**Interfaces:**
- Consumes:
  - `settingsStorage` (`getProviders`, `saveProviders`, `getActiveProvider`, `saveActiveProvider`, `getSearchSettings`, `setSearchSettings`)
  - `soulStorage` (`getSoulConfig`, `saveSoulConfig`, `clearSoulConfig`)
  - `customToolStorage` (`getCustomTools`, `saveCustomTool`, `clearCustomTools`)
  - `crossAppBridge` (`listCrossAppSnapshots`, `saveCrossAppSnapshot`, `clearCrossAppSnapshots`)
  - `cryptoBackup` (`encryptPayload`, `decryptPayload`, `calculateSha256`, `SamBackupEnvelope`, `SamBackupPayload`, `RestoreResult`)
- Produces:
  - `createBackup(options?: { password?: string; appVersion?: string }): Promise<string>`
  - `restoreBackup(envelopeJson: string, options?: { password?: string; cleanRestore?: boolean }): Promise<RestoreResult>`

- [x] **Step 1: Write the failing test**

Create `tests/backupManager.test.ts`:

```typescript
import { describe, it, expect, beforeEach } from 'vitest';
import { createBackup, restoreBackup } from '../src/services/storage/backupManager';
import { saveProviders, saveActiveProvider, getProviders, getActiveProvider } from '../src/services/storage/settingsStorage';
import { saveSoulConfig, getSoulConfig, clearSoulConfig } from '../src/services/storage/soulStorage';
import { saveCustomTool, getCustomTools, clearCustomTools } from '../src/services/storage/customToolStorage';
import { saveCrossAppSnapshot, listCrossAppSnapshots, clearCrossAppSnapshots } from '../src/services/storage/crossAppBridge';

describe('backupManager', () => {
  beforeEach(() => {
    clearSoulConfig();
    clearCustomTools();
    clearCrossAppSnapshots();
  });

  it('creates and restores unencrypted backup cleanly', async () => {
    saveSoulConfig({ corporateName: 'PT Nusantara Jaya', brandVoicePreset: 'financial_compliance' });
    saveCustomTool({ name: 'calc_discount', description: 'Hitung diskon', hostType: 'Excel', code: 'return 10;' });
    saveCrossAppSnapshot({ sourceHost: 'Excel', title: 'Snapshot Keuangan Q3', artifactType: 'table_data' });

    const backupJson = await createBackup();
    const parsed = JSON.parse(backupJson);
    expect(parsed.format).toBe('sam-office-agent-backup');
    expect(parsed.encrypted).toBe(false);

    // Clear local state
    clearSoulConfig();
    clearCustomTools();
    clearCrossAppSnapshots();
    expect(getSoulConfig().corporateName).not.toBe('PT Nusantara Jaya');

    // Restore
    const res = await restoreBackup(backupJson);
    expect(res.success).toBe(true);
    expect(getSoulConfig().corporateName).toBe('PT Nusantara Jaya');
    expect(getCustomTools().some(t => t.name === 'calc_discount')).toBe(true);
    expect(listCrossAppSnapshots().length).toBeGreaterThan(0);
  });

  it('creates and restores password-encrypted backup', async () => {
    saveSoulConfig({ corporateName: 'PT Enkripsi Indonesia' });
    const password = 'SuperSecretPassword2026';

    const backupJson = await createBackup({ password });
    const parsed = JSON.parse(backupJson);
    expect(parsed.encrypted).toBe(true);
    expect(parsed.salt).toBeDefined();
    expect(parsed.iv).toBeDefined();

    clearSoulConfig();

    // Wrong password fails
    await expect(restoreBackup(backupJson, { password: 'WrongPassword' })).rejects.toThrow();

    // Correct password succeeds
    const res = await restoreBackup(backupJson, { password });
    expect(res.success).toBe(true);
    expect(getSoulConfig().corporateName).toBe('PT Enkripsi Indonesia');
  });

  it('performs clean restore when cleanRestore option is true', async () => {
    saveSoulConfig({ corporateName: 'PT Awal' });
    const backupJson = await createBackup();

    saveCustomTool({ name: 'temporary_tool', description: 'Temp', hostType: 'Word', code: 'return 1;' });
    expect(getCustomTools().length).toBe(1);

    await restoreBackup(backupJson, { cleanRestore: true });
    // Temporary tool added after backup should be wiped by cleanRestore
    expect(getCustomTools().length).toBe(0);
  });
});
```

- [x] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/backupManager.test.ts`
Expected: FAIL with "Cannot find module '../src/services/storage/backupManager'".

- [x] **Step 3: Write minimal implementation**

Create `src/services/storage/backupManager.ts`:
- Implement `createBackup(options?: { password?: string; appVersion?: string }): Promise<string>`:
  - Collects `providers`, `activeProvider`, `searchSettings`, `soulConfig`, `customTools`, `crossAppSnapshots`.
  - Builds payload object and JSON string.
  - Computes `checksum` using `calculateSha256`.
  - If password is provided: runs `encryptPayload()`, returns envelope with `encrypted: true`.
  - Else: returns envelope with `encrypted: false`, `payload: payloadJson`.
- Implement `restoreBackup(envelopeJson: string, options?: { password?: string; cleanRestore?: boolean }): Promise<RestoreResult>`:
  - Validates `envelope.format === 'sam-office-agent-backup'`.
  - Decrypts or parses payload.
  - Verifies checksum.
  - If `options.cleanRestore`: calls `clearCustomTools()`, `clearCrossAppSnapshots()`, `clearSoulConfig()`.
  - Restores each slice into its respective storage.
  - Returns `RestoreResult`.

- [x] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/backupManager.test.ts`
Expected: PASS.

- [x] **Step 5: Commit**

```bash
git add src/services/storage/backupManager.ts tests/backupManager.test.ts
git commit -m "feat(storage): implement comprehensive backup and restore manager"
```

---

### Task 3: UI Settings Management Panel (`SettingsModal.tsx`)

**Files:**
- Modify: `src/components/Settings/SettingsModal.tsx`
- Test: `src/components/__tests__/headerAndSettings.test.tsx`

**Interfaces:**
- Consumes: `createBackup`, `restoreBackup` from `../../services/storage/backupManager`
- Produces: "Cadangan & Pemulihan (Backup & Restore)" panel in SettingsModal

- [ ] **Step 1: Write the failing test**

In `src/components/__tests__/headerAndSettings.test.tsx`:
Add a test verifying backup & restore UI elements render in `SettingsModal`:

```typescript
it('renders Backup & Restore section with export and restore controls in SettingsModal', () => {
  render(
    <SettingsModal
      isOpen={true}
      onClose={vi.fn()}
      providers={mockProviders}
      activeProvider="gemini"
      onUpdateProviders={vi.fn()}
      onSelectProvider={vi.fn()}
    />
  );

  expect(screen.getByText(/Cadangan & Pemulihan/i)).toBeInTheDocument();
  expect(screen.getByText(/Unduh Cadangan/i)).toBeInTheDocument();
  expect(screen.getByText(/Pilih Berkas Cadangan/i)).toBeInTheDocument();
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/components/__tests__/headerAndSettings.test.tsx`
Expected: FAIL with "Unable to find an element with text: /Cadangan & Pemulihan/i".

- [ ] **Step 3: Write minimal implementation**

In `src/components/Settings/SettingsModal.tsx`:
- Import `createBackup` and `restoreBackup` from `../../services/storage/backupManager`.
- Add "Cadangan & Pemulihan Data (Backup & Restore)" section:
  - **Export Card**:
    - Password input (type text/password with show/hide toggle).
    - Checkbox: "Ekspor tanpa enkripsi (Plain JSON)".
    - Button: "📥 Unduh Cadangan (.sam-backup)":
      - Calls `createBackup({ password })`.
      - Triggers download of Blob with MIME `application/json` as `sam-office-backup-YYYY-MM-DD.sam-backup`.
  - **Restore Card**:
    - File input for `.sam-backup` / `.json`.
    - Password prompt input if encrypted file is selected.
    - Checkbox: "Bersihkan data lokal saat ini sebelum restore (Clean Restore)".
    - Button: "Pulihkan Data (Restore)".
    - Inline alert banner displaying success or error messages.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/components/__tests__/headerAndSettings.test.tsx`
Expected: PASS.

- [ ] **Step 5: Full verification & regression suite**

Run:
```bash
npm run build
npx vitest run
```
Expected: All build checks pass with exit code 0, and all test files pass 100% green.

- [ ] **Step 6: Commit**

```bash
git add src/components/Settings/SettingsModal.tsx src/components/__tests__/headerAndSettings.test.tsx
git commit -m "feat(ui): add encrypted backup & restore section to SettingsModal"
```
