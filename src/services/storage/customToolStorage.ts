import { HostType, ToolDefinition } from '../../types';

export interface CustomTool {
  id: string;
  name: string;
  description: string;
  hostType: HostType | 'Universal';
  parameters: ToolDefinition['parameters'];
  code: string;
  createdAt: number;
  usageCount: number;
}

const STORAGE_KEY = 'sam_learned_tools_v1';

export function getCustomTools(host?: HostType): CustomTool[] {
  if (typeof localStorage === 'undefined') return [];
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const list: CustomTool[] = JSON.parse(raw);
    if (!Array.isArray(list)) return [];
    if (!host) return list;
    return list.filter(t => t.hostType === host || t.hostType === 'Universal');
  } catch (e) {
    console.warn('Gagal membaca custom tools:', e);
    return [];
  }
}

export function saveCustomTool(
  tool: Omit<CustomTool, 'id' | 'createdAt' | 'usageCount' | 'parameters'> & {
    parameters?: ToolDefinition['parameters'];
  }
): CustomTool {
  const existing = getCustomTools();
  const normalizedName = tool.name.trim().toLowerCase().replace(/[^a-z0-9_]/g, '_');

  const existingIdx = existing.findIndex(t => t.name.toLowerCase() === normalizedName);
  const now = Date.now();

  const newTool: CustomTool = {
    id: existingIdx >= 0 ? existing[existingIdx].id : `tool_${normalizedName}_${now}`,
    name: normalizedName,
    description: tool.description,
    hostType: tool.hostType || 'Universal',
    parameters: tool.parameters || { type: 'object', properties: {} },
    code: tool.code,
    createdAt: existingIdx >= 0 ? existing[existingIdx].createdAt : now,
    usageCount: existingIdx >= 0 ? existing[existingIdx].usageCount : 0,
  };

  if (existingIdx >= 0) {
    existing[existingIdx] = newTool;
  } else {
    existing.push(newTool);
  }

  if (typeof localStorage !== 'undefined') {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(existing));
  }

  return newTool;
}

export function incrementToolUsage(nameOrId: string): void {
  const list = getCustomTools();
  const found = list.find(t => t.id === nameOrId || t.name === nameOrId);
  if (found) {
    found.usageCount = (found.usageCount || 0) + 1;
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
    }
  }
}

export function deleteCustomTool(nameOrId: string): boolean {
  const list = getCustomTools();
  const filtered = list.filter(t => t.id !== nameOrId && t.name !== nameOrId);
  if (filtered.length !== list.length) {
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(filtered));
    }
    return true;
  }
  return false;
}

export function clearCustomTools(): void {
  if (typeof localStorage !== 'undefined') {
    localStorage.removeItem(STORAGE_KEY);
  }
}
