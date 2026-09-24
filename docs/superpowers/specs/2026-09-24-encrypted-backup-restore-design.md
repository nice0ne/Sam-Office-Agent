# Encrypted Backup & Restore (AES-256-GCM / JSON) Design Spec

## 1. Overview & Business Value

As users configure multiple LLM providers, private API keys, search engine credentials, brand voice rules (`SOUL.md`), learned corporate preferences, custom dynamic tools, and cross-application snapshots, the risk of data loss due to browser storage purges, Office Add-in cache invalidation, or machine migration increases significantly.

### Key Pain Points
1. **Accidental Data Loss**: Browser `localStorage` can be wiped during Office Add-in re-installation, WebView2 cache cleaning, or corporate IT workstation maintenance.
2. **Setup Friction Across Devices**: Manually re-entering API keys, provider parameters, and custom prompt directives across multiple desktop installations is tedious.
3. **Security Risks in Export**: Exporting private API keys, client credentials, and corporate guidelines in plain unencrypted JSON files exposes organizations to leakage risks.

### The Solution: Zero-Dependency Encrypted Backup & Restore Engine
Equip **Sam Office Agent** with an enterprise-grade, 100% client-side backup and restore engine leveraging standard **Web Crypto API** (AES-256-GCM + PBKDF2):
- **Military-Grade Encryption (AES-256-GCM)**: All credentials, settings, and memory directives can be encrypted with a user-chosen passphrase using 100,000 PBKDF2 iterations, SHA-256, and random Salt/IV.
- **Data Integrity & Tamper Protection**: AES-GCM 128-bit authentication tags and SHA-256 payload checksums ensure corrupted or modified backup files cannot be restored.
- **Unencrypted Fallback Option**: An optional plain JSON mode for developer inspection or corporate audit environments.
- **Comprehensive Ecosystem Coverage**: Backs up and restores:
  - LLM Provider configs & API keys (`settingsStorage.ts`)
  - Web Search settings & Tavily credentials (`settingsStorage.ts`)
  - Corporate brand voice & `SOUL.md` learned rules (`soulStorage.ts`)
  - Dynamic meta-tools & custom scripts (`customToolStorage.ts`)
  - Universal Hub cross-app snapshots (`crossAppBridge.ts`)
- **Smart Merge & Clean Restore**: Allows seamless merging with existing data or complete clean resets on new machines.

---

## 2. Architecture & Data Contracts

### 2.1 Types (`src/services/storage/cryptoBackup.ts` & `src/types/index.ts`)

```typescript
export interface SamBackupEnvelope {
  format: 'sam-office-agent-backup';
  version: 1;
  encrypted: boolean;
  exportedAt: number; // UNIX timestamp
  salt?: string;      // Base64-encoded PBKDF2 salt (16 bytes)
  iv?: string;        // Base64-encoded AES-GCM IV (12 bytes)
  payload: string;    // Base64 ciphertext (if encrypted) or raw JSON string (if plain)
  checksum: string;   // SHA-256 hex digest of unencrypted payload JSON
}

export interface SamBackupPayload {
  version: 1;
  exportedAt: number;
  environment: {
    appVersion: string;
    userAgent?: string;
  };
  data: {
    providers?: ProviderConfig[];
    activeProvider?: string;
    searchSettings?: SearchSettings;
    soulConfig?: SoulConfig;
    customTools?: CustomTool[];
    crossAppSnapshots?: CrossAppSnapshot[];
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
```

---

## 3. Component Design & Implementation

### 3.1 Web Crypto Cryptographic Core (`src/services/storage/cryptoBackup.ts`)

- **Key Derivation (PBKDF2)**:
  - Imports user password as raw key material via `crypto.subtle.importKey('raw', ...)`.
  - Derives AES-GCM key with `PBKDF2`, `SHA-256`, 100,000 iterations, and a randomly generated 16-byte salt (`crypto.getRandomValues`).
- **AES-GCM Encryption**:
  - Generates 12-byte initialization vector (IV).
  - Encrypts raw UTF-8 string using `AES-GCM` (256-bit key length).
  - Encodes ciphertext, salt, and IV to Base64 strings.
- **AES-GCM Decryption**:
  - Reconstructs PBKDF2 key from password and salt.
  - Decrypts Base64 ciphertext using IV.
  - Automatically throws error if authentication tag mismatches or password is invalid.
- **Checksum Calculation**:
  - Computes SHA-256 hex digest of string data using `crypto.subtle.digest('SHA-256', ...)`.

### 3.2 Backup & Restore Manager (`src/services/storage/backupManager.ts`)

- **`createBackup(options?: { password?: string; appVersion?: string }): Promise<string>`**:
  - Gathers state from `getProviders()`, `getActiveProvider()`, `getSearchSettings()`, `getSoulConfig()`, `getCustomTools()`, and `listCrossAppSnapshots()`.
  - Compiles `SamBackupPayload`.
  - Computes SHA-256 checksum.
  - If password is provided and non-empty: encrypts payload and returns serialized `SamBackupEnvelope` with `encrypted: true`.
  - If no password: returns serialized envelope with `encrypted: false` and stringified payload.
- **`restoreBackup(envelopeJson: string, options?: { password?: string; cleanRestore?: boolean }): Promise<RestoreResult>`**:
  - Parses envelope and verifies `format === 'sam-office-agent-backup'`.
  - If encrypted, requires `options.password` and decrypts payload via `decryptPayload()`.
  - Verifies payload SHA-256 checksum against `envelope.checksum`.
  - If `options.cleanRestore`: clears existing providers, tools, soul config, and snapshots.
  - Restores each subsystem:
    - Providers & active provider (`saveProviders`, `saveActiveProvider`)
    - Search settings (`setSearchSettings`)
    - Soul config (`saveSoulConfig`)
    - Custom tools (`saveCustomTool`)
    - Cross-app snapshots (`saveCrossAppSnapshot`)
  - Returns `RestoreResult` with detailed summary counts.

### 3.3 UI Integration (`src/components/Settings/SettingsModal.tsx`)

- **"Cadangan & Pemulihan Data (Backup & Restore)" Section**:
  - **Export Card**:
    - Optional password input field (masked with toggle visibility).
    - Checkbox: "Ekspor tanpa enkripsi (Plain JSON)".
    - Button: "📥 Unduh Cadangan (.sam-backup)" (triggers browser file download with timestamped filename).
  - **Restore Card**:
    - File input button: "📤 Pilih Berkas Cadangan".
    - Password prompt appears if loaded envelope is encrypted.
    - Checkbox: "Bersihkan data lokal saat ini sebelum restore (Clean Restore)".
    - Action button: "Pulihkan Data (Restore)".
    - Inline status banner showing success or error details.

---

## 4. Verification & Testing Strategy

1. **Unit Tests for Web Crypto Engine (`tests/cryptoBackup.test.ts`)**:
   - Verify encryption and decryption roundtrip for plain text.
   - Verify decryption fails with incorrect password.
   - Verify SHA-256 checksum calculation and tamper detection.
2. **Integration Tests for Backup Manager (`tests/backupManager.test.ts`)**:
   - Verify `createBackup()` aggregates all subsystems in encrypted and plain modes.
   - Verify `restoreBackup()` correctly restores all configurations into their respective storages.
   - Verify `cleanRestore` option wipes previous data prior to applying the backup.
   - Verify invalid backup format or corrupted payload throws descriptive error.
3. **UI Settings Tests (`src/components/__tests__/headerAndSettings.test.tsx`)**:
   - Verify backup and restore controls render in `SettingsModal`.
4. **Full Regression Suite**:
   - Verify all 320+ existing tests pass with 0 regressions.
