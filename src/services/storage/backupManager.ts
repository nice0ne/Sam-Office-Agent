/**
 * Backup & Restore Manager
 * Collects and restores configuration across settings, soul directives, custom tools, and snapshots.
 */

import {
  getProviders,
  saveProviders,
  getActiveProvider,
  saveActiveProvider,
  getSearchSettings,
  setSearchSettings,
} from './settingsStorage';
import { getSoulConfig, saveSoulConfig, clearSoulConfig } from './soulStorage';
import { getCustomTools, saveCustomTool, clearCustomTools } from './customToolStorage';
import { listCrossAppSnapshots, saveCrossAppSnapshot, clearCrossAppSnapshots } from './crossAppBridge';
import {
  encryptPayload,
  decryptPayload,
  calculateSha256,
  SamBackupEnvelope,
  SamBackupPayload,
  RestoreResult,
} from './cryptoBackup';

/**
 * Creates an encrypted or plain JSON backup envelope of all application state.
 */
export async function createBackup(options?: {
  password?: string;
  appVersion?: string;
}): Promise<string> {
  const providers = getProviders();
  const activeProvider = getActiveProvider()?.id || 'gemini';
  const searchSettings = getSearchSettings();
  const soulConfig = getSoulConfig();
  const customTools = getCustomTools();
  const crossAppSnapshots = listCrossAppSnapshots(Number.MAX_SAFE_INTEGER);

  const payload: SamBackupPayload = {
    version: 1,
    exportedAt: Date.now(),
    environment: {
      appVersion: options?.appVersion || '1.0.0',
      userAgent: typeof navigator !== 'undefined' ? navigator.userAgent : 'Node.js',
    },
    data: {
      providers,
      activeProvider,
      searchSettings,
      soulConfig,
      customTools,
      crossAppSnapshots,
    },
  };

  const payloadJson = JSON.stringify(payload);
  const checksum = await calculateSha256(payloadJson);

  const hasPassword = Boolean(options?.password && options.password.trim().length > 0);

  if (hasPassword && options?.password) {
    const { ciphertext, salt, iv } = await encryptPayload(payloadJson, options.password);
    const envelope: SamBackupEnvelope = {
      format: 'sam-office-agent-backup',
      version: 1,
      encrypted: true,
      exportedAt: payload.exportedAt,
      salt,
      iv,
      payload: ciphertext,
      checksum,
    };
    return JSON.stringify(envelope, null, 2);
  }

  const envelope: SamBackupEnvelope = {
    format: 'sam-office-agent-backup',
    version: 1,
    encrypted: false,
    exportedAt: payload.exportedAt,
    payload: payloadJson,
    checksum,
  };
  return JSON.stringify(envelope, null, 2);
}

/**
 * Restores all application state from a serialized backup envelope.
 */
export async function restoreBackup(
  envelopeJson: string,
  options?: { password?: string; cleanRestore?: boolean }
): Promise<RestoreResult> {
  let envelope: SamBackupEnvelope;
  try {
    envelope = JSON.parse(envelopeJson);
  } catch (err: any) {
    throw new Error(`Invalid backup file: not valid JSON (${err?.message || err})`);
  }

  if (!envelope || envelope.format !== 'sam-office-agent-backup') {
    throw new Error('Invalid backup format: expected "sam-office-agent-backup"');
  }

  let payloadJson: string;

  if (envelope.encrypted) {
    if (!options?.password) {
      throw new Error('This backup is encrypted with a password. Please provide a password.');
    }
    if (!envelope.salt || !envelope.iv) {
      throw new Error('Corrupted backup: missing salt or iv in encrypted envelope.');
    }
    payloadJson = await decryptPayload(
      envelope.payload,
      options.password,
      envelope.salt,
      envelope.iv,
      envelope.checksum
    );
  } else {
    payloadJson = envelope.payload;
    if (envelope.checksum) {
      const actualChecksum = await calculateSha256(payloadJson);
      if (actualChecksum !== envelope.checksum) {
        throw new Error(
          `Checksum integrity check failed: expected ${envelope.checksum} but got ${actualChecksum}`
        );
      }
    }
  }

  let payload: SamBackupPayload;
  try {
    payload = typeof payloadJson === 'string' ? JSON.parse(payloadJson) : payloadJson;
  } catch (err: any) {
    throw new Error(`Invalid backup payload: not valid JSON (${err?.message || err})`);
  }

  if (options?.cleanRestore) {
    clearSoulConfig();
    clearCustomTools();
    clearCrossAppSnapshots();
  }

  let providersCount = 0;
  if (payload.data?.providers) {
    saveProviders(payload.data.providers);
    providersCount = Array.isArray(payload.data.providers)
      ? payload.data.providers.length
      : Object.keys(payload.data.providers).length;
  }

  if (payload.data?.activeProvider) {
    saveActiveProvider(payload.data.activeProvider);
  }

  let searchConfigured = false;
  if (payload.data?.searchSettings) {
    setSearchSettings(payload.data.searchSettings);
    searchConfigured = true;
  }

  let soulRestored = false;
  if (payload.data?.soulConfig) {
    saveSoulConfig(payload.data.soulConfig);
    soulRestored = true;
  }

  let customToolsCount = 0;
  if (Array.isArray(payload.data?.customTools)) {
    for (const tool of payload.data.customTools) {
      saveCustomTool(tool);
      customToolsCount++;
    }
  }

  let snapshotsCount = 0;
  if (Array.isArray(payload.data?.crossAppSnapshots)) {
    const snapshots = [...payload.data.crossAppSnapshots].reverse();
    for (const snap of snapshots) {
      saveCrossAppSnapshot(snap);
      snapshotsCount++;
    }
  }

  return {
    success: true,
    restoredAt: Date.now(),
    cleanRestore: Boolean(options?.cleanRestore),
    summary: {
      providersCount,
      searchConfigured,
      soulRestored,
      customToolsCount,
      snapshotsCount,
    },
  };
}
