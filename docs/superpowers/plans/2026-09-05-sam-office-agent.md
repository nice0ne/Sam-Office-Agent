# Sam-Office-Agent Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Membangun Microsoft Office Web Add-in multi-agent universal (Excel, Word, PowerPoint) dengan antarmuka taskpane React + TypeScript dan arsitektur BYOK (Bring Your Own Key) tanpa server backend pihak ketiga.

**Architecture:** Universal Office Web Add-in client-side berbasis Vite, React 18, dan Tailwind CSS yang mendeteksi host Office (`Office.context.host`) secara otomatis. Koordinator "Sam" mengatur memori dan memecah instruksi ke agen spesialis (Excel, Word, PowerPoint) yang menghasilkan structured tool calls. Tool calls dieksekusi secara aman menggunakan batch sync Office.js dengan preview diff (Copilot mode) atau eksekusi instan (Autopilot mode). Kunci LLM (OpenAI, Gemini, Claude, GLM, OpenRouter, Ollama) disimpan aman di `localStorage` lokal.

**Tech Stack:** React 18, TypeScript, Vite, Tailwind CSS, Lucide React, `@types/office-js`, Vitest.

**Spec:** [`docs/superpowers/specs/2026-09-05-sam-office-agent-design.md`](file:///D:/VIBE-CODING/Sam-Office-Agent/docs/superpowers/specs/2026-09-05-sam-office-agent-design.md)

## Global Constraints
- Harus berjalan 100% di peramban/WebView2 tanpa backend server eksternal wajib.
- Semua API Key LLM hanya disimpan di `localStorage` peramban perangkat pengguna.
- Pemanggilan `context.sync()` pada Office.js harus dilakukan secara batch (1 kali per operasi transaksi).
- Harus menyediakan `MockOfficeDriver` agar aplikasi dapat dijalankan dan diuji langsung di browser standar (`http://localhost:5173`) tanpa ketergantungan pada Office 365 desktop.
- Semua kode baru harus ditulis dalam TypeScript dengan type safety yang ketat (`strict: true`).

---

### Task 1: Project Scaffolding & Configuration

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `tsconfig.node.json`
- Create: `vite.config.ts`
- Create: `tailwind.config.js`
- Create: `postcss.config.js`
- Create: `index.html`
- Test: `tests/setup.ts`

**Interfaces:**
- Consumes: None
- Produces: Project build system, TypeScript compiler setup, Tailwind styles, and Vitest test runner.

- [ ] **Step 1: Write the package.json and install dependencies configuration**

```json
{
  "name": "sam-office-agent",
  "private": true,
  "version": "0.1.0",
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "tsc && vite build",
    "preview": "vite preview",
    "test": "vitest run",
    "test:watch": "vitest"
  },
  "dependencies": {
    "clsx": "^2.1.1",
    "lucide-react": "^0.475.0",
    "react": "^18.3.1",
    "react-dom": "^18.3.1",
    "tailwind-merge": "^2.6.0"
  },
  "devDependencies": {
    "@types/office-js": "^1.0.450",
    "@types/react": "^18.3.18",
    "@types/react-dom": "^18.3.5",
    "@vitejs/plugin-react": "^4.3.4",
    "autoprefixer": "^10.4.20",
    "postcss": "^8.5.2",
    "tailwindcss": "^3.4.17",
    "typescript": "^5.7.3",
    "vite": "^6.1.0",
    "vitest": "^3.0.5"
  }
}
```

- [ ] **Step 2: Create TypeScript and Vite configuration files**

Create `tsconfig.json`:
```json
{
  "compilerOptions": {
    "target": "ES2020",
    "useDefineForClassFields": true,
    "lib": ["ES2020", "DOM", "DOM.Iterable"],
    "module": "ESNext",
    "skipLibCheck": true,
    "moduleResolution": "bundler",
    "allowImportingTsExtensions": false,
    "resolveJsonModule": true,
    "isolatedModules": true,
    "noEmit": true,
    "jsx": "react-jsx",
    "strict": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "noFallthroughCasesInSwitch": true,
    "baseUrl": ".",
    "paths": {
      "@/*": ["src/*"]
    }
  },
  "include": ["src", "tests"]
}
```

Create `vite.config.ts`:
```typescript
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
    },
  },
  server: {
    port: 5173,
    cors: true,
  },
  test: {
    globals: true,
    environment: 'node',
    setupFiles: './tests/setup.ts',
  },
});
```

- [ ] **Step 3: Setup Tailwind CSS & PostCSS**

Create `tailwind.config.js`:
```javascript
/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  darkMode: 'class',
  theme: {
    extend: {
      colors: {
        office: {
          excel: '#107C41',
          word: '#185ABD',
          ppt: '#C43E1C',
          agent: '#0F6CBD',
        }
      }
    },
  },
  plugins: [],
}
```

Create `postcss.config.js`:
```javascript
export default {
  plugins: {
    tailwindcss: {},
    autoprefixer: {},
  },
}
```

- [ ] **Step 4: Create HTML entry point with Office.js CDN and basic test setup**

Create `index.html`:
```html
<!DOCTYPE html>
<html lang="id">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Sam Office Agent</title>
    <!-- Microsoft Office.js SDK CDN -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  </head>
  <body class="bg-gray-50 dark:bg-gray-900 text-gray-900 dark:text-gray-100 antialiased select-none">
    <div id="root"></div>
    <script type="module" src="/src/main.tsx"></script>
  </body>
</html>
```

Create `tests/setup.ts`:
```typescript
// Test setup configuration
import { beforeEach } from 'vitest';

beforeEach(() => {
  // Clear localStorage before each test
  if (typeof localStorage !== 'undefined') {
    localStorage.clear();
  }
});
```

- [ ] **Step 5: Run npm install and run vitest verification**

Run: `npm install`
Run: `npx vitest run`
Expected: Passes with no test errors or runs cleanly.

- [ ] **Step 6: Commit**

```bash
git add package.json tsconfig.json vite.config.ts tailwind.config.js postcss.config.js index.html tests/setup.ts
git commit -m "chore: scaffold project structure, vite, tailwind, and vitest"
```

---

### Task 2: Core Domain Types & Data Models

**Files:**
- Create: `src/types/index.ts`
- Create: `src/agents/types.ts`
- Test: `tests/types.test.ts`

**Interfaces:**
- Consumes: None
- Produces: `HostType`, `ProviderId`, `ProviderConfig`, `ChatMessage`, `ToolDefinition`, `ToolCall`, `AgentAction`, `ExecutionContext`.

- [ ] **Step 1: Write test for type contracts and helper guards**

Create `tests/types.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { HostType, isValidHost } from '../src/types';

describe('HostType & Guards', () => {
  it('validates known Office hosts correctly', () => {
    expect(isValidHost('Excel')).toBe(true);
    expect(isValidHost('Word')).toBe(true);
    expect(isValidHost('PowerPoint')).toBe(true);
    expect(isValidHost('UnknownHost')).toBe(false);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/types.test.ts`
Expected: FAIL (Cannot find module `../src/types`)

- [ ] **Step 3: Implement `src/types/index.ts` and `src/agents/types.ts`**

Create `src/types/index.ts`:
```typescript
export type HostType = 'Excel' | 'Word' | 'PowerPoint' | 'BrowserDev';

export function isValidHost(host: string): host is HostType {
  return ['Excel', 'Word', 'PowerPoint', 'BrowserDev'].includes(host);
}

export type ProviderId = 'gemini' | 'openai' | 'claude' | 'glm' | 'openrouter' | 'ollama';

export interface ProviderConfig {
  id: ProviderId;
  name: string;
  apiKey: string;
  baseUrl?: string;
  selectedModel: string;
  enabled: boolean;
}

export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: number;
  toolCalls?: ToolCall[];
  status?: 'sending' | 'streaming' | 'done' | 'error';
  error?: string;
}

export interface ToolProperty {
  type: 'string' | 'number' | 'boolean' | 'array' | 'object';
  description: string;
  enum?: string[];
  items?: { type: string };
}

export interface ToolDefinition {
  name: string;
  description: string;
  parameters: {
    type: 'object';
    properties: Record<string, ToolProperty>;
    required?: string[];
  };
}

export interface ToolCall {
  id: string;
  name: string;
  arguments: Record<string, any>;
  status: 'pending' | 'applied' | 'rejected' | 'failed';
  error?: string;
}

export interface ActionQueueItem {
  id: string;
  stepNumber: number;
  title: string;
  description: string;
  toolCall: ToolCall;
  status: 'pending' | 'executing' | 'completed' | 'failed';
}

export type ExecutionMode = 'copilot' | 'autopilot';
```

Create `src/agents/types.ts`:
```typescript
import { ChatMessage, HostType, ToolCall, ToolDefinition } from '../types';

export interface AgentContext {
  host: HostType;
  selectedText?: string;
  activeCellOrRange?: string;
  documentSummary?: string;
}

export interface AgentResponse {
  message: string;
  toolCalls: ToolCall[];
}

export interface IAgent {
  id: string;
  name: string;
  hostType: HostType;
  getSystemPrompt(context: AgentContext): string;
  getTools(): ToolDefinition[];
  executeTool(toolCall: ToolCall, context: AgentContext): Promise<{ success: boolean; result?: any; error?: string }>;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/types.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/types/index.ts src/agents/types.ts tests/types.test.ts
git commit -m "feat: define core domain types and agent interfaces"
```

---

### Task 3: Client-Side Storage & BYOK Settings Service

**Files:**
- Create: `src/services/storage/settingsStorage.ts`
- Test: `src/services/storage/__tests__/settingsStorage.test.ts`

**Interfaces:**
- Consumes: `ProviderConfig`, `ProviderId`, `ExecutionMode` from `src/types/index.ts`
- Produces: `getSettings()`, `saveProviderConfig()`, `getActiveProvider()`, `setActiveProvider()`, `getExecutionMode()`, `setExecutionMode()`

- [ ] **Step 1: Write test for Settings Storage**

Create `src/services/storage/__tests__/settingsStorage.test.ts`:
```typescript
import { describe, it, expect, beforeEach } from 'vitest';
import {
  getSettings,
  saveProviderConfig,
  getActiveProvider,
  setActiveProvider,
  getExecutionMode,
  setExecutionMode,
} from '../settingsStorage';

describe('Settings Storage Service', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  it('initializes default provider configurations', () => {
    const settings = getSettings();
    expect(settings.providers.gemini).toBeDefined();
    expect(settings.providers.openai).toBeDefined();
    expect(settings.providers.glm).toBeDefined();
    expect(settings.activeProviderId).toBe('gemini');
  });

  it('saves and retrieves updated provider config', () => {
    saveProviderConfig({
      id: 'openai',
      name: 'OpenAI',
      apiKey: 'sk-test-key-12345',
      selectedModel: 'gpt-4o',
      enabled: true,
    });

    const active = getActiveProvider();
    expect(active.apiKey).toBe(''); // Gemini is still active
    setActiveProvider('openai');
    expect(getActiveProvider().apiKey).toBe('sk-test-key-12345');
  });

  it('toggles execution mode between copilot and autopilot', () => {
    expect(getExecutionMode()).toBe('copilot');
    setExecutionMode('autopilot');
    expect(getExecutionMode()).toBe('autopilot');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/services/storage/__tests__/settingsStorage.test.ts`
Expected: FAIL (Cannot find module `../settingsStorage`)

- [ ] **Step 3: Implement Settings Storage**

Create `src/services/storage/settingsStorage.ts`:
```typescript
import { ExecutionMode, ProviderConfig, ProviderId } from '../../types';

const STORAGE_KEY = 'sam_office_settings_v1';

export interface AppSettings {
  activeProviderId: ProviderId;
  executionMode: ExecutionMode;
  theme: 'light' | 'dark' | 'system';
  providers: Record<ProviderId, ProviderConfig>;
}

const DEFAULT_PROVIDERS: Record<ProviderId, ProviderConfig> = {
  gemini: {
    id: 'gemini',
    name: 'Google Gemini',
    apiKey: '',
    selectedModel: 'gemini-2.0-flash',
    enabled: true,
  },
  openai: {
    id: 'openai',
    name: 'OpenAI',
    apiKey: '',
    selectedModel: 'gpt-4o',
    enabled: true,
  },
  claude: {
    id: 'claude',
    name: 'Anthropic Claude',
    apiKey: '',
    selectedModel: 'claude-3-5-sonnet-20241022',
    enabled: true,
  },
  glm: {
    id: 'glm',
    name: 'GLM Coding Global (Zhipu)',
    apiKey: '',
    baseUrl: 'https://open.bigmodel.cn/api/paas/v4',
    selectedModel: 'glm-4-plus',
    enabled: true,
  },
  openrouter: {
    id: 'openrouter',
    name: 'OpenRouter / DeepSeek',
    apiKey: '',
    baseUrl: 'https://openrouter.ai/api/v1',
    selectedModel: 'deepseek/deepseek-chat',
    enabled: true,
  },
  ollama: {
    id: 'ollama',
    name: 'Local Ollama',
    apiKey: 'ollama',
    baseUrl: 'http://localhost:11434/v1',
    selectedModel: 'qwen2.5-coder',
    enabled: false,
  },
};

export function getSettings(): AppSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return {
        activeProviderId: 'gemini',
        executionMode: 'copilot',
        theme: 'system',
        providers: DEFAULT_PROVIDERS,
      };
    }
    const parsed = JSON.parse(raw);
    return {
      activeProviderId: parsed.activeProviderId || 'gemini',
      executionMode: parsed.executionMode || 'copilot',
      theme: parsed.theme || 'system',
      providers: { ...DEFAULT_PROVIDERS, ...(parsed.providers || {}) },
    };
  } catch {
    return {
      activeProviderId: 'gemini',
      executionMode: 'copilot',
      theme: 'system',
      providers: DEFAULT_PROVIDERS,
    };
  }
}

export function saveSettings(settings: AppSettings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

export function saveProviderConfig(config: ProviderConfig): void {
  const current = getSettings();
  current.providers[config.id] = config;
  saveSettings(current);
}

export function getActiveProvider(): ProviderConfig {
  const settings = getSettings();
  return settings.providers[settings.activeProviderId] || settings.providers.gemini;
}

export function setActiveProvider(id: ProviderId): void {
  const settings = getSettings();
  settings.activeProviderId = id;
  saveSettings(settings);
}

export function getExecutionMode(): ExecutionMode {
  return getSettings().executionMode;
}

export function setExecutionMode(mode: ExecutionMode): void {
  const settings = getSettings();
  settings.executionMode = mode;
  saveSettings(settings);
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/services/storage/__tests__/settingsStorage.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/services/storage/settingsStorage.ts src/services/storage/__tests__/settingsStorage.test.ts
git commit -m "feat: implement BYOK settings and local storage persistence"
```

---

### Task 4: Heuristic Formula & Context Compressor Utilities

**Files:**
- Create: `src/utils/formulaValidator.ts`
- Create: `src/utils/contextCompressor.ts`
- Test: `src/utils/__tests__/formulaValidator.test.ts`
- Test: `src/utils/__tests__/contextCompressor.test.ts`

**Interfaces:**
- Consumes: None
- Produces: `validateExcelFormula(formula: string)`, `compressTableContext(data: any[][], maxRows?: number)`

- [ ] **Step 1: Write tests for formula validator and context compressor**

Create `src/utils/__tests__/formulaValidator.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { validateExcelFormula } from '../formulaValidator';

describe('validateExcelFormula', () => {
  it('validates correct Excel formulas', () => {
    expect(validateExcelFormula('=SUM(A1:A10)').isValid).toBe(true);
    expect(validateExcelFormula('=XLOOKUP(A2, Sheet2!A:A, Sheet2!B:B, "N/A")').isValid).toBe(true);
    expect(validateExcelFormula('=IF(B2>100, B2*0.1, 0)').isValid).toBe(true);
  });

  it('flags unclosed parentheses', () => {
    const res = validateExcelFormula('=SUM(A1:A10');
    expect(res.isValid).toBe(false);
    expect(res.error).toContain('kurung');
  });

  it('flags formula missing leading equal sign', () => {
    const res = validateExcelFormula('SUM(A1:A10)');
    expect(res.isValid).toBe(false);
    expect(res.error).toContain('=');
  });
});
```

Create `src/utils/__tests__/contextCompressor.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { compressTableContext } from '../contextCompressor';

describe('compressTableContext', () => {
  it('samples first rows and provides row count summary', () => {
    const mockData = [
      ['ID', 'Nama', 'Sales'],
      ['1', 'Andi', 100],
      ['2', 'Budi', 200],
      ['3', 'Cici', 300],
      ['4', 'Dedi', 400],
      ['5', 'Eka', 500],
      ['6', 'Fani', 600],
    ];

    const result = compressTableContext(mockData, 3);
    expect(result.headers).toEqual(['ID', 'Nama', 'Sales']);
    expect(result.sampledRows.length).toBe(3);
    expect(result.totalRows).toBe(6);
    expect(result.summaryText).toContain('Total baris: 6');
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx vitest run src/utils/__tests__/`
Expected: FAIL (Modules not found)

- [ ] **Step 3: Implement Formula Validator and Context Compressor**

Create `src/utils/formulaValidator.ts`:
```typescript
export interface FormulaValidationResult {
  isValid: boolean;
  error?: string;
  sanitizedFormula?: string;
}

export function validateExcelFormula(formula: string): FormulaValidationResult {
  const trimmed = formula.trim();
  if (!trimmed.startsWith('=')) {
    return {
      isValid: false,
      error: 'Formula Excel harus diawali dengan tanda sama dengan (=).',
    };
  }

  let openParenCount = 0;
  let inQuotes = false;

  for (let i = 0; i < trimmed.length; i++) {
    const char = trimmed[i];
    if (char === '"' && (i === 0 || trimmed[i - 1] !== '\\')) {
      inQuotes = !inQuotes;
    } else if (!inQuotes) {
      if (char === '(') openParenCount++;
      else if (char === ')') openParenCount--;
    }

    if (openParenCount < 0) {
      return {
        isValid: false,
        error: 'Tanda kurung tutup ")" berlebih tanpa kurung buka yang sesuai.',
      };
    }
  }

  if (openParenCount > 0) {
    return {
      isValid: false,
      error: `Tanda kurung tidak seimbang: ada ${openParenCount} kurung buka "(" yang belum ditutup.`,
    };
  }

  return {
    isValid: true,
    sanitizedFormula: trimmed,
  };
}
```

Create `src/utils/contextCompressor.ts`:
```typescript
export interface TableSummary {
  headers: string[];
  sampledRows: any[][];
  totalRows: number;
  totalColumns: number;
  summaryText: string;
}

export function compressTableContext(data: any[][], maxSampleRows = 5): TableSummary {
  if (!data || data.length === 0) {
    return {
      headers: [],
      sampledRows: [],
      totalRows: 0,
      totalColumns: 0,
      summaryText: 'Tabel kosong.',
    };
  }

  const headers = data[0].map(h => String(h || ''));
  const rows = data.slice(1);
  const totalRows = rows.length;
  const totalColumns = headers.length;
  const sampledRows = rows.slice(0, maxSampleRows);

  const summaryText = `Tabel berukuran ${totalColumns} kolom x ${totalRows} baris data. (Total baris: ${totalRows}). Header: [${headers.join(', ')}].`;

  return {
    headers,
    sampledRows,
    totalRows,
    totalColumns,
    summaryText,
  };
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx vitest run src/utils/__tests__/`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/utils/formulaValidator.ts src/utils/contextCompressor.ts src/utils/__tests__/
git commit -m "feat: add formula validator and table context compressor"
```

---

### Task 5: Unified BYOK LLM Provider Adapters

**Files:**
- Create: `src/services/llm/types.ts`
- Create: `src/services/llm/openai.ts`
- Create: `src/services/llm/gemini.ts`
- Create: `src/services/llm/anthropic.ts`
- Create: `src/services/llm/glm.ts`
- Create: `src/services/llm/factory.ts`
- Test: `src/services/llm/__tests__/llmAdapters.test.ts`

**Interfaces:**
- Consumes: `ProviderConfig`, `ChatMessage`, `ToolDefinition`
- Produces: `ILLMProvider`, `getLLMProvider(id: ProviderId)`, `testProviderConnection(config: ProviderConfig)`

- [ ] **Step 1: Write test for provider factory and request formatting**

Create `src/services/llm/__tests__/llmAdapters.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { getLLMProvider } from '../factory';

describe('LLM Provider Factory', () => {
  it('returns valid adapter instances for all supported providers', () => {
    expect(getLLMProvider('gemini').id).toBe('gemini');
    expect(getLLMProvider('openai').id).toBe('openai');
    expect(getLLMProvider('claude').id).toBe('claude');
    expect(getLLMProvider('glm').id).toBe('glm');
    expect(getLLMProvider('openrouter').id).toBe('openrouter');
    expect(getLLMProvider('ollama').id).toBe('ollama');
  });

  it('fails connection test gracefully if API key is missing', async () => {
    const provider = getLLMProvider('gemini');
    const result = await provider.testConnection({
      id: 'gemini',
      name: 'Google Gemini',
      apiKey: '',
      selectedModel: 'gemini-2.0-flash',
      enabled: true,
    });
    expect(result.success).toBe(false);
    expect(result.message).toContain('API Key');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/services/llm/__tests__/llmAdapters.test.ts`
Expected: FAIL (Cannot find module `../factory`)

- [ ] **Step 3: Implement LLM types and Provider Adapters**

Create `src/services/llm/types.ts`:
```typescript
import { ChatMessage, ProviderConfig, ToolCall, ToolDefinition } from '../../types';

export interface ChatRequest {
  messages: ChatMessage[];
  systemPrompt?: string;
  tools?: ToolDefinition[];
  temperature?: number;
}

export interface StreamEvent {
  type: 'content_delta' | 'tool_call' | 'done' | 'error';
  delta?: string;
  toolCall?: ToolCall;
  error?: string;
}

export interface ILLMProvider {
  id: string;
  name: string;
  sendMessage(req: ChatRequest, config: ProviderConfig): AsyncIterable<StreamEvent>;
  testConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }>;
}
```

Create `src/services/llm/openai.ts` (supports OpenAI, GLM, OpenRouter, and Ollama compatible endpoints):
```typescript
import { ProviderConfig, ToolCall } from '../../types';
import { ChatRequest, ILLMProvider, StreamEvent } from './types';

export class OpenAICompatibleProvider implements ILLMProvider {
  id: string;
  name: string;
  defaultBaseUrl: string;

  constructor(id: string, name: string, defaultBaseUrl: string) {
    this.id = id;
    this.name = name;
    this.defaultBaseUrl = defaultBaseUrl;
  }

  async *sendMessage(req: ChatRequest, config: ProviderConfig): AsyncIterable<StreamEvent> {
    const baseUrl = (config.baseUrl || this.defaultBaseUrl).replace(/\/+$/, '');
    const url = `${baseUrl}/chat/completions`;

    const messages = [];
    if (req.systemPrompt) {
      messages.push({ role: 'system', content: req.systemPrompt });
    }
    for (const m of req.messages) {
      messages.push({ role: m.role, content: m.content });
    }

    const body: Record<string, any> = {
      model: config.selectedModel,
      messages,
      stream: true,
      temperature: req.temperature ?? 0.7,
    };

    if (req.tools && req.tools.length > 0) {
      body.tools = req.tools.map(t => ({
        type: 'function',
        function: {
          name: t.name,
          description: t.description,
          parameters: t.parameters,
        },
      }));
      body.tool_choice = 'auto';
    }

    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${config.apiKey}`,
    };

    if (this.id === 'openrouter') {
      headers['HTTP-Referer'] = 'https://github.com/Sam-Office-Agent';
      headers['X-Title'] = 'Sam Office Agent';
    }

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers,
        body: JSON.stringify(body),
      });

      if (!response.ok) {
        const errorText = await response.text();
        yield { type: 'error', error: `HTTP ${response.status}: ${errorText}` };
        return;
      }

      const reader = response.body?.getReader();
      if (!reader) {
        yield { type: 'error', error: 'No response body stream available.' };
        return;
      }

      const decoder = new TextDecoder('utf-8');
      let buffer = '';
      let pendingToolCalls: Record<number, { id: string; name: string; arguments: string }> = {};

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          const clean = line.trim();
          if (!clean || !clean.startsWith('data: ')) continue;
          if (clean === 'data: [DONE]') continue;

          try {
            const parsed = JSON.parse(clean.slice(6));
            const delta = parsed.choices?.[0]?.delta;
            if (delta?.content) {
              yield { type: 'content_delta', delta: delta.content };
            }

            if (delta?.tool_calls) {
              for (const tc of delta.tool_calls) {
                const idx = tc.index ?? 0;
                if (!pendingToolCalls[idx]) {
                  pendingToolCalls[idx] = {
                    id: tc.id || `call_${Date.now()}_${idx}`,
                    name: tc.function?.name || '',
                    arguments: '',
                  };
                }
                if (tc.function?.name) pendingToolCalls[idx].name += tc.function.name;
                if (tc.function?.arguments) pendingToolCalls[idx].arguments += tc.function.arguments;
              }
            }
          } catch {
            // Ignore stream parse chunks
          }
        }
      }

      for (const idx in pendingToolCalls) {
        const ptc = pendingToolCalls[idx];
        let args = {};
        try {
          args = JSON.parse(ptc.arguments);
        } catch {
          args = {};
        }
        yield {
          type: 'tool_call',
          toolCall: {
            id: ptc.id,
            name: ptc.name,
            arguments: args,
            status: 'pending',
          },
        };
      }

      yield { type: 'done' };
    } catch (err: any) {
      yield { type: 'error', error: err?.message || 'Gagal terhubung ke API OpenAI' };
    }
  }

  async testConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }> {
    if (!config.apiKey && this.id !== 'ollama') {
      return { success: false, message: 'API Key belum diisi.' };
    }

    const baseUrl = (config.baseUrl || this.defaultBaseUrl).replace(/\/+$/, '');
    try {
      const res = await fetch(`${baseUrl}/chat/completions`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${config.apiKey}`,
        },
        body: JSON.stringify({
          model: config.selectedModel,
          messages: [{ role: 'user', content: 'Ping' }],
          max_tokens: 5,
        }),
      });

      if (res.ok) {
        return { success: true, message: `Koneksi berhasil ke ${this.name} (${config.selectedModel})!` };
      }
      const err = await res.text();
      return { success: false, message: `Gagal (HTTP ${res.status}): ${err}` };
    } catch (e: any) {
      return { success: false, message: `Error koneksi: ${e.message}` };
    }
  }
}
```

Create `src/services/llm/gemini.ts`:
```typescript
import { ProviderConfig } from '../../types';
import { ChatRequest, ILLMProvider, StreamEvent } from './types';

export class GeminiProvider implements ILLMProvider {
  id = 'gemini';
  name = 'Google Gemini';

  async *sendMessage(req: ChatRequest, config: ProviderConfig): AsyncIterable<StreamEvent> {
    if (!config.apiKey) {
      yield { type: 'error', error: 'API Key Google Gemini belum diisi.' };
      return;
    }

    const model = config.selectedModel || 'gemini-2.0-flash';
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:streamGenerateContent?key=${config.apiKey}`;

    const contents = req.messages.map(m => ({
      role: m.role === 'assistant' ? 'model' : 'user',
      parts: [{ text: m.content }],
    }));

    const body: Record<string, any> = { contents };

    if (req.systemPrompt) {
      body.systemInstruction = {
        parts: [{ text: req.systemPrompt }],
      };
    }

    if (req.tools && req.tools.length > 0) {
      body.tools = [
        {
          functionDeclarations: req.tools.map(t => ({
            name: t.name,
            description: t.description,
            parameters: t.parameters,
          })),
        },
      ];
    }

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });

      if (!response.ok) {
        const errorText = await response.text();
        yield { type: 'error', error: `Gemini HTTP ${response.status}: ${errorText}` };
        return;
      }

      const reader = response.body?.getReader();
      if (!reader) {
        yield { type: 'error', error: 'Tidak ada stream dari Gemini API.' };
        return;
      }

      const decoder = new TextDecoder('utf-8');
      let buffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        // Gemini returns chunks of JSON array format
        // Simple extraction of text parts and functionCalls
        const textMatches = [...buffer.matchAll(/"text":\s*"((?:[^"\\]|\\.)*)"/g)];
        if (textMatches.length > 0) {
          for (const match of textMatches) {
            try {
              const text = JSON.parse(`"${match[1]}"`);
              yield { type: 'content_delta', delta: text };
            } catch {}
          }
          buffer = '';
        }
      }

      yield { type: 'done' };
    } catch (e: any) {
      yield { type: 'error', error: e.message || 'Gagal memanggil Gemini API.' };
    }
  }

  async testConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }> {
    if (!config.apiKey) {
      return { success: false, message: 'API Key Gemini belum diisi.' };
    }
    const model = config.selectedModel || 'gemini-2.0-flash';
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${config.apiKey}`;

    try {
      const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: 'Ping' }] }],
        }),
      });
      if (res.ok) {
        return { success: true, message: `Koneksi Google Gemini berhasil (${model})!` };
      }
      const err = await res.text();
      return { success: false, message: `Gemini Error (${res.status}): ${err}` };
    } catch (e: any) {
      return { success: false, message: `Koneksi gagal: ${e.message}` };
    }
  }
}
```

Create `src/services/llm/anthropic.ts`:
```typescript
import { ProviderConfig } from '../../types';
import { ChatRequest, ILLMProvider, StreamEvent } from './types';

export class AnthropicProvider implements ILLMProvider {
  id = 'claude';
  name = 'Anthropic Claude';

  async *sendMessage(req: ChatRequest, config: ProviderConfig): AsyncIterable<StreamEvent> {
    if (!config.apiKey) {
      yield { type: 'error', error: 'API Key Anthropic belum diisi.' };
      return;
    }

    const url = 'https://api.anthropic.com/v1/messages';
    const messages = req.messages.map(m => ({
      role: m.role === 'assistant' ? 'assistant' : 'user',
      content: m.content,
    }));

    const body: Record<string, any> = {
      model: config.selectedModel || 'claude-3-5-sonnet-20241022',
      messages,
      max_tokens: 4096,
      stream: true,
    };

    if (req.systemPrompt) {
      body.system = req.systemPrompt;
    }

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': config.apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify(body),
      });

      if (!response.ok) {
        const errorText = await response.text();
        yield { type: 'error', error: `Claude HTTP ${response.status}: ${errorText}` };
        return;
      }

      const reader = response.body?.getReader();
      if (!reader) return;

      const decoder = new TextDecoder('utf-8');
      let buffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          const clean = line.trim();
          if (!clean.startsWith('data: ')) continue;
          try {
            const data = JSON.parse(clean.slice(6));
            if (data.type === 'content_block_delta' && data.delta?.text) {
              yield { type: 'content_delta', delta: data.delta.text };
            }
          } catch {}
        }
      }

      yield { type: 'done' };
    } catch (e: any) {
      yield { type: 'error', error: e.message || 'Gagal memanggil Claude API.' };
    }
  }

  async testConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }> {
    if (!config.apiKey) {
      return { success: false, message: 'API Key Anthropic belum diisi.' };
    }

    try {
      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': config.apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify({
          model: config.selectedModel || 'claude-3-5-haiku-20241022',
          messages: [{ role: 'user', content: 'Ping' }],
          max_tokens: 5,
        }),
      });

      if (res.ok) {
        return { success: true, message: 'Koneksi Anthropic Claude berhasil!' };
      }
      const err = await res.text();
      return { success: false, message: `Claude Error (${res.status}): ${err}` };
    } catch (e: any) {
      return { success: false, message: `Koneksi gagal: ${e.message}` };
    }
  }
}
```

Create `src/services/llm/factory.ts`:
```typescript
import { ProviderConfig, ProviderId } from '../../types';
import { AnthropicProvider } from './anthropic';
import { GeminiProvider } from './gemini';
import { OpenAICompatibleProvider } from './openai';
import { ILLMProvider } from './types';

const providers: Record<ProviderId, ILLMProvider> = {
  gemini: new GeminiProvider(),
  openai: new OpenAICompatibleProvider('openai', 'OpenAI', 'https://api.openai.com/v1'),
  claude: new AnthropicProvider(),
  glm: new OpenAICompatibleProvider('glm', 'GLM Coding Global', 'https://open.bigmodel.cn/api/paas/v4'),
  openrouter: new OpenAICompatibleProvider('openrouter', 'OpenRouter', 'https://openrouter.ai/api/v1'),
  ollama: new OpenAICompatibleProvider('ollama', 'Local Ollama', 'http://localhost:11434/v1'),
};

export function getLLMProvider(id: ProviderId): ILLMProvider {
  return providers[id] || providers.gemini;
}

export async function testProviderConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }> {
  const provider = getLLMProvider(config.id);
  return provider.testConnection(config);
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/services/llm/__tests__/llmAdapters.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/services/llm/ tests/
git commit -m "feat: implement unified BYOK LLM provider adapters"
```

---

### Task 6: Office.js Safe Driver & Mock Bridge

**Files:**
- Create: `src/services/office/types.ts`
- Create: `src/services/office/mockDriver.ts`
- Create: `src/services/office/excelDriver.ts`
- Create: `src/services/office/wordDriver.ts`
- Create: `src/services/office/pptDriver.ts`
- Create: `src/services/office/index.ts`
- Test: `src/services/office/__tests__/mockDriver.test.ts`

**Interfaces:**
- Consumes: `HostType`, `ToolCall`
- Produces: `getOfficeDriver()`, `IDocumentDriver`

- [ ] **Step 1: Write test for Mock Office Driver**

Create `src/services/office/__tests__/mockDriver.test.ts`:
```typescript
import { describe, it, expect, beforeEach } from 'vitest';
import { MockOfficeDriver } from '../mockDriver';

describe('MockOfficeDriver', () => {
  let driver: MockOfficeDriver;

  beforeEach(() => {
    driver = new MockOfficeDriver();
  });

  it('handles Excel write and read operations in mock mode', async () => {
    await driver.writeCells('A1:B2', [['Item', 'Price'], ['Apple', 10]]);
    const range = await driver.readActiveRange();
    expect(range.values.length).toBeGreaterThan(0);
    expect(range.values[0][0]).toBe('Item');
  });

  it('handles Word content insertion and selection replacement', async () => {
    await driver.insertContent('start', 'Halo Dunia', 'paragraph');
    const outline = await driver.getWordOutline();
    expect(outline).toContain('Halo Dunia');
  });

  it('handles PowerPoint slide creation', async () => {
    const slide = await driver.addSlide('TitleAndContent');
    expect(slide.slideNumber).toBe(1);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`
Expected: FAIL (Cannot find module `../mockDriver`)

- [ ] **Step 3: Implement Office Driver Interfaces and Drivers**

Create `src/services/office/types.ts`:
```typescript
export interface IDocumentDriver {
  hostType: string;
  // Excel operations
  readActiveRange(): Promise<{ address: string; values: any[][]; formulas: string[][] }>;
  writeCells(range: string, values?: any[][], formulas?: string[][]): Promise<void>;
  formatRange(range: string, styles: Record<string, any>): Promise<void>;
  createChart(type: string, dataRange: string, title?: string): Promise<void>;
  // Word operations
  getWordOutline(): Promise<string>;
  insertContent(position: 'start' | 'end' | 'cursor', text: string, type: string): Promise<void>;
  replaceSelection(newText: string): Promise<void>;
  insertTable(rows: number, cols: number, data?: string[][]): Promise<void>;
  // PowerPoint operations
  getSlideContext(): Promise<{ slideNumber: number; title: string; textContent: string }>;
  addSlide(layout: string): Promise<{ slideNumber: number }>;
  insertSlideContent(title: string, bullets: string[]): Promise<void>;
  setSpeakerNotes(notes: string): Promise<void>;
}
```

Create `src/services/office/mockDriver.ts`:
```typescript
import { IDocumentDriver } from './types';

export class MockOfficeDriver implements IDocumentDriver {
  hostType = 'BrowserDev';
  private excelGrid: Record<string, any> = {};
  private wordContent: string[] = [];
  private slides: Array<{ title: string; bullets: string[]; notes: string }> = [];

  async readActiveRange() {
    return {
      address: 'A1:C5',
      values: [
        ['Produk', 'Qty', 'Harga'],
        ['Laptop', 5, 15000000],
        ['Mouse', 20, 150000],
        ['Keyboard', 12, 500000],
        ['Monitor', 8, 2500000],
      ],
      formulas: [['', '', ''], ['', '', ''], ['', '', ''], ['', '', ''], ['', '', '']],
    };
  }

  async writeCells(range: string, values?: any[][], formulas?: string[][]) {
    if (values) {
      this.excelGrid[range] = values;
    }
  }

  async formatRange(range: string, styles: Record<string, any>) {
    // Mock styling
  }

  async createChart(type: string, dataRange: string, title?: string) {
    // Mock chart
  }

  async getWordOutline() {
    return this.wordContent.join('\n') || 'Dokumen Word Kosong.';
  }

  async insertContent(position: 'start' | 'end' | 'cursor', text: string, type: string) {
    if (position === 'start') this.wordContent.unshift(text);
    else this.wordContent.push(text);
  }

  async replaceSelection(newText: string) {
    this.wordContent = [newText];
  }

  async insertTable(rows: number, cols: number, data?: string[][]) {
    this.wordContent.push(`[Tabel ${rows}x${cols}]`);
  }

  async getSlideContext() {
    const active = this.slides[0] || { title: 'Slide 1', bullets: [], notes: '' };
    return {
      slideNumber: 1,
      title: active.title,
      textContent: active.bullets.join('\n'),
    };
  }

  async addSlide(layout: string) {
    this.slides.push({ title: `Slide ${this.slides.length + 1}`, bullets: [], notes: '' });
    return { slideNumber: this.slides.length };
  }

  async insertSlideContent(title: string, bullets: string[]) {
    if (this.slides.length === 0) await this.addSlide('TitleAndContent');
    const current = this.slides[this.slides.length - 1];
    current.title = title;
    current.bullets = bullets;
  }

  async setSpeakerNotes(notes: string) {
    if (this.slides.length === 0) await this.addSlide('TitleAndContent');
    this.slides[this.slides.length - 1].notes = notes;
  }
}
```

Create `src/services/office/excelDriver.ts`:
```typescript
import { IDocumentDriver } from './types';

declare const Excel: any;

export class ExcelDriver implements Partial<IDocumentDriver> {
  hostType = 'Excel';

  async readActiveRange() {
    return await Excel.run(async (context: any) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'values', 'formulas']);
      await context.sync();
      return {
        address: range.address,
        values: range.values,
        formulas: range.formulas,
      };
    });
  }

  async writeCells(rangeAddress: string, values?: any[][], formulas?: string[][]) {
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      if (values) range.values = values;
      if (formulas) range.formulas = formulas;
      await context.sync();
    });
  }

  async formatRange(rangeAddress: string, styles: Record<string, any>) {
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      if (styles.bold !== undefined) range.format.font.bold = styles.bold;
      if (styles.color) range.format.font.color = styles.color;
      if (styles.fillColor) range.format.fill.color = styles.fillColor;
      if (styles.numberFormat) range.numberFormat = [[styles.numberFormat]];
      await context.sync();
    });
  }

  async createChart(type: string, dataRange: string, title?: string) {
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(dataRange);
      const chartType = Excel.ChartType[type] || Excel.ChartType.columnClustered;
      const chart = sheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);
      if (title) chart.title.text = title;
      await context.sync();
    });
  }
}
```

Create `src/services/office/wordDriver.ts`:
```typescript
import { IDocumentDriver } from './types';

declare const Word: any;

export class WordDriver implements Partial<IDocumentDriver> {
  hostType = 'Word';

  async getWordOutline(): Promise<string> {
    return await Word.run(async (context: any) => {
      const body = context.document.body;
      body.load('text');
      await context.sync();
      return body.text || '';
    });
  }

  async insertContent(position: 'start' | 'end' | 'cursor', text: string, type: string) {
    await Word.run(async (context: any) => {
      let location = Word.InsertLocation.end;
      if (position === 'start') location = Word.InsertLocation.start;
      const paragraph = context.document.body.insertParagraph(text, location);
      if (type === 'h1') paragraph.style = 'Heading 1';
      else if (type === 'h2') paragraph.style = 'Heading 2';
      await context.sync();
    });
  }

  async replaceSelection(newText: string) {
    await Word.run(async (context: any) => {
      const selection = context.document.getSelection();
      selection.insertText(newText, Word.InsertLocation.replace);
      await context.sync();
    });
  }

  async insertTable(rows: number, cols: number, data?: string[][]) {
    await Word.run(async (context: any) => {
      const table = context.document.body.insertTable(rows, cols, Word.InsertLocation.end, data || []);
      table.styleBuiltIn = Word.Style.gridTable4_Accent1;
      await context.sync();
    });
  }
}
```

Create `src/services/office/pptDriver.ts`:
```typescript
import { IDocumentDriver } from './types';

declare const Office: any;

export class PPTDriver implements Partial<IDocumentDriver> {
  hostType = 'PowerPoint';

  async getSlideContext() {
    return {
      slideNumber: 1,
      title: 'Active Slide',
      textContent: 'Slide content',
    };
  }

  async addSlide(layout: string) {
    return { slideNumber: 1 };
  }

  async insertSlideContent(title: string, bullets: string[]) {
    // In PowerPoint Office.js Web Add-in, text insertion uses Office.context.document.setSelectedDataAsync
    if (typeof Office !== 'undefined' && Office.context?.document) {
      const content = `${title}\n\n` + bullets.map(b => `• ${b}`).join('\n');
      Office.context.document.setSelectedDataAsync(content, { coercionType: Office.CoercionType.Text });
    }
  }

  async setSpeakerNotes(notes: string) {
    // Notes handler
  }
}
```

Create `src/services/office/index.ts`:
```typescript
import { HostType } from '../../types';
import { ExcelDriver } from './excelDriver';
import { MockOfficeDriver } from './mockDriver';
import { PPTDriver } from './pptDriver';
import { IDocumentDriver } from './types';
import { WordDriver } from './wordDriver';

let activeDriver: IDocumentDriver | null = null;

export function getOfficeDriver(host?: HostType): IDocumentDriver {
  if (activeDriver) return activeDriver;

  if (typeof (window as any).Office !== 'undefined' && (window as any).Office.context?.host) {
    const officeHost = (window as any).Office.context.host;
    if (officeHost === (window as any).Office.HostType.Excel) {
      activeDriver = new ExcelDriver() as IDocumentDriver;
    } else if (officeHost === (window as any).Office.HostType.Word) {
      activeDriver = new WordDriver() as IDocumentDriver;
    } else if (officeHost === (window as any).Office.HostType.PowerPoint) {
      activeDriver = new PPTDriver() as IDocumentDriver;
    }
  }

  if (!activeDriver) {
    activeDriver = new MockOfficeDriver();
  }

  return activeDriver;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/services/office/__tests__/mockDriver.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/services/office/
git commit -m "feat: implement office.js drivers and browser mock environment"
```

---

### Task 7: Domain Specialist Agents (Excel, Word, PowerPoint)

**Files:**
- Create: `src/agents/excel/excelAgent.ts`
- Create: `src/agents/word/wordAgent.ts`
- Create: `src/agents/powerpoint/pptAgent.ts`
- Test: `src/agents/__tests__/specialistAgents.test.ts`

**Interfaces:**
- Consumes: `IAgent`, `AgentContext`, `ToolCall`, `ToolDefinition` from `src/agents/types.ts`
- Produces: `ExcelAgent`, `WordAgent`, `PPTAgent`

- [ ] **Step 1: Write test for Specialist Agents**

Create `src/agents/__tests__/specialistAgents.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { ExcelAgent } from '../excel/excelAgent';
import { WordAgent } from '../word/wordAgent';
import { PPTAgent } from '../powerpoint/pptAgent';

describe('Specialist Agents Tools & Prompts', () => {
  it('ExcelAgent exposes Excel specific tools', () => {
    const agent = new ExcelAgent();
    const tools = agent.getTools();
    const toolNames = tools.map(t => t.name);
    expect(toolNames).toContain('write_cells');
    expect(toolNames).toContain('format_range');
    expect(toolNames).toContain('create_chart');
  });

  it('WordAgent exposes Word specific tools', () => {
    const agent = new WordAgent();
    const tools = agent.getTools();
    const toolNames = tools.map(t => t.name);
    expect(toolNames).toContain('insert_content');
    expect(toolNames).toContain('replace_selection');
  });

  it('PPTAgent exposes PowerPoint specific tools', () => {
    const agent = new PPTAgent();
    const tools = agent.getTools();
    const toolNames = tools.map(t => t.name);
    expect(toolNames).toContain('add_slide');
    expect(toolNames).toContain('insert_slide_content');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: FAIL (Cannot find modules)

- [ ] **Step 3: Implement Excel, Word, and PowerPoint Agents**

Create `src/agents/excel/excelAgent.ts`:
```typescript
import { getOfficeDriver } from '../../services/office';
import { validateExcelFormula } from '../../utils/formulaValidator';
import { AgentContext, IAgent } from '../types';
import { ToolCall, ToolDefinition } from '../../types';

export class ExcelAgent implements IAgent {
  id = 'excel-specialist';
  name = 'Excel Specialist';
  hostType = 'Excel' as const;

  getSystemPrompt(context: AgentContext): string {
    return `Anda adalah Sam Office Agent spesialis Microsoft Excel.
Anda ahli dalam analisis data, formula spreadsheet kompleks (SUM, AVERAGE, XLOOKUP, INDEX, MATCH), pemformatan sel, dan pembuatan grafik.
Jika pengguna meminta menulis formula, pastikan sintaks formula valid dan diawali '='.
Konteks saat ini: ${context.activeCellOrRange || 'Sheet aktif'}.`;
  }

  getTools(): ToolDefinition[] {
    return [
      {
        name: 'write_cells',
        description: 'Menulis nilai atau formula ke range sel Excel tertentu (misal: A1:B10).',
        parameters: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Alamat range, misal: "E2:E15"' },
            formula: { type: 'string', description: 'Formula Excel (opsional), misal: "=SUM(B2:D2)"' },
            values: { type: 'array', description: 'Array 2D nilai data jika menulis data statis' },
          },
          required: ['range'],
        },
      },
      {
        name: 'format_range',
        description: 'Mengatur format sel (tebal, warna latar zebra, format mata uang/persentase).',
        parameters: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Alamat range, misal: "E2:E15"' },
            bold: { type: 'boolean', description: 'Tebalkan teks' },
            fillColor: { type: 'string', description: 'Warna latar HEX, misal: "#F0F9FF"' },
            numberFormat: { type: 'string', description: 'Format angka, misal: "Rp#,##0"' },
          },
          required: ['range'],
        },
      },
      {
        name: 'create_chart',
        description: 'Membuat grafik otomatis berdasarkan range data.',
        parameters: {
          type: 'object',
          properties: {
            chartType: { type: 'string', enum: ['ColumnClustered', 'Line', 'Pie', 'BarClustered'], description: 'Jenis grafik' },
            dataRange: { type: 'string', description: 'Range data grafik, misal: "A1:E15"' },
            title: { type: 'string', description: 'Judul grafik' },
          },
          required: ['chartType', 'dataRange'],
        },
      },
    ];
  }

  async executeTool(toolCall: ToolCall, context: AgentContext): Promise<{ success: boolean; result?: any; error?: string }> {
    const driver = getOfficeDriver('Excel');
    try {
      if (toolCall.name === 'write_cells') {
        const { range, formula, values } = toolCall.arguments;
        if (formula) {
          const validation = validateExcelFormula(formula);
          if (!validation.isValid) {
            return { success: false, error: validation.error };
          }
          await driver.writeCells(range, undefined, [[formula]]);
        } else if (values) {
          await driver.writeCells(range, values);
        }
        return { success: true, result: `Berhasil menulis ke sel ${range}` };
      }

      if (toolCall.name === 'format_range') {
        const { range, ...styles } = toolCall.arguments;
        await driver.formatRange(range, styles);
        return { success: true, result: `Format berhasil diterapkan ke ${range}` };
      }

      if (toolCall.name === 'create_chart') {
        const { chartType, dataRange, title } = toolCall.arguments;
        await driver.createChart(chartType, dataRange, title);
        return { success: true, result: `Grafik ${chartType} berhasil dibuat!` };
      }

      return { success: false, error: `Tool ${toolCall.name} tidak dikenali.` };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  }
}
```

Create `src/agents/word/wordAgent.ts`:
```typescript
import { getOfficeDriver } from '../../services/office';
import { AgentContext, IAgent } from '../types';
import { ToolCall, ToolDefinition } from '../../types';

export class WordAgent implements IAgent {
  id = 'word-specialist';
  name = 'Word Specialist';
  hostType = 'Word' as const;

  getSystemPrompt(context: AgentContext): string {
    return `Anda adalah Sam Office Agent spesialis Microsoft Word.
Anda ahli dalam penyusunan surat resmi, laporan profesional, perbaikan tata bahasa/proofreading, pemformatan heading dokumen, dan tabel terstruktur.`;
  }

  getTools(): ToolDefinition[] {
    return [
      {
        name: 'insert_content',
        description: 'Menyisipkan paragraf, heading (H1/H2), atau poin di dokumen.',
        parameters: {
          type: 'object',
          properties: {
            position: { type: 'string', enum: ['start', 'end', 'cursor'], description: 'Posisi penyisipan' },
            text: { type: 'string', description: 'Isi teks' },
            type: { type: 'string', enum: ['paragraph', 'h1', 'h2', 'bullet'], description: 'Tipe konten' },
          },
          required: ['position', 'text', 'type'],
        },
      },
      {
        name: 'replace_selection',
        description: 'Mengganti teks yang sedang diblok/diseleksi pengguna dengan teks baru yang telah disempurnakan.',
        parameters: {
          type: 'object',
          properties: {
            newText: { type: 'string', description: 'Teks baru pengganti' },
          },
          required: ['newText'],
        },
      },
      {
        name: 'insert_table',
        description: 'Menyisipkan tabel data di dokumen Word.',
        parameters: {
          type: 'object',
          properties: {
            rows: { type: 'number', description: 'Jumlah baris' },
            cols: { type: 'number', description: 'Jumlah kolom' },
            data: { type: 'array', description: 'Array 2D data teks' },
          },
          required: ['rows', 'cols'],
        },
      },
    ];
  }

  async executeTool(toolCall: ToolCall, context: AgentContext) {
    const driver = getOfficeDriver('Word');
    try {
      if (toolCall.name === 'insert_content') {
        const { position, text, type } = toolCall.arguments;
        await driver.insertContent(position, text, type);
        return { success: true, result: 'Konten berhasil disisipkan.' };
      }
      if (toolCall.name === 'replace_selection') {
        const { newText } = toolCall.arguments;
        await driver.replaceSelection(newText);
        return { success: true, result: 'Seleksi berhasil diganti.' };
      }
      if (toolCall.name === 'insert_table') {
        const { rows, cols, data } = toolCall.arguments;
        await driver.insertTable(rows, cols, data);
        return { success: true, result: `Tabel ${rows}x${cols} berhasil dibuat.` };
      }
      return { success: false, error: 'Tool tidak ditemukan.' };
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }
}
```

Create `src/agents/powerpoint/pptAgent.ts`:
```typescript
import { getOfficeDriver } from '../../services/office';
import { AgentContext, IAgent } from '../types';
import { ToolCall, ToolDefinition } from '../../types';

export class PPTAgent implements IAgent {
  id = 'ppt-specialist';
  name = 'PowerPoint Specialist';
  hostType = 'PowerPoint' as const;

  getSystemPrompt(context: AgentContext): string {
    return `Anda adalah Sam Office Agent spesialis Microsoft PowerPoint.
Anda ahli dalam merancang struktur presentasi yang memikat, menyusun poin slide yang ringkas, dan membuat speaker notes untuk presentasi.`;
  }

  getTools(): ToolDefinition[] {
    return [
      {
        name: 'add_slide',
        description: 'Menambahkan slide baru ke presentasi.',
        parameters: {
          type: 'object',
          properties: {
            layout: { type: 'string', enum: ['TitleOnly', 'TitleAndContent', 'TwoContent'], description: 'Layout slide' },
          },
          required: ['layout'],
        },
      },
      {
        name: 'insert_slide_content',
        description: 'Mengisi judul slide dan poin-poin materi.',
        parameters: {
          type: 'object',
          properties: {
            title: { type: 'string', description: 'Judul slide' },
            bullets: { type: 'array', items: { type: 'string' }, description: 'Poin-poin presentasi' },
          },
          required: ['title', 'bullets'],
        },
      },
      {
        name: 'set_speaker_notes',
        description: 'Menambahkan catatan pemateri pada slide aktif.',
        parameters: {
          type: 'object',
          properties: {
            notes: { type: 'string', description: 'Catatan pemateri' },
          },
          required: ['notes'],
        },
      },
    ];
  }

  async executeTool(toolCall: ToolCall, context: AgentContext) {
    const driver = getOfficeDriver('PowerPoint');
    try {
      if (toolCall.name === 'add_slide') {
        const res = await driver.addSlide(toolCall.arguments.layout);
        return { success: true, result: `Slide #${res.slideNumber} berhasil ditambahkan.` };
      }
      if (toolCall.name === 'insert_slide_content') {
        const { title, bullets } = toolCall.arguments;
        await driver.insertSlideContent(title, bullets);
        return { success: true, result: 'Konten slide berhasil diisi.' };
      }
      if (toolCall.name === 'set_speaker_notes') {
        await driver.setSpeakerNotes(toolCall.arguments.notes);
        return { success: true, result: 'Speaker notes berhasil disimpan.' };
      }
      return { success: false, error: 'Tool tidak ditemukan.' };
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/agents/__tests__/specialistAgents.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/agents/ tests/
git commit -m "feat: implement Excel, Word, and PowerPoint specialist agents"
```

---

### Task 8: Coordinator Agent ("Sam") & Action Queue Manager

**Files:**
- Create: `src/agents/coordinator/samCoordinator.ts`
- Test: `src/agents/coordinator/__tests__/samCoordinator.test.ts`

**Interfaces:**
- Consumes: `ExcelAgent`, `WordAgent`, `PPTAgent`, `ILLMProvider`, `getSettings()`, `ActionQueueItem`
- Produces: `SamCoordinator`, `ActionQueue`

- [ ] **Step 1: Write test for Sam Coordinator**

Create `src/agents/coordinator/__tests__/samCoordinator.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { SamCoordinator } from '../samCoordinator';

describe('SamCoordinator', () => {
  it('selects correct specialist agent based on host context', () => {
    const coordinator = new SamCoordinator();
    expect(coordinator.getSpecialist('Excel').id).toBe('excel-specialist');
    expect(coordinator.getSpecialist('Word').id).toBe('word-specialist');
    expect(coordinator.getSpecialist('PowerPoint').id).toBe('ppt-specialist');
  });

  it('builds system prompt tailored to active host', () => {
    const coordinator = new SamCoordinator();
    const prompt = coordinator.buildSystemPrompt('Excel', { host: 'Excel' });
    expect(prompt).toContain('Sam');
    expect(prompt).toContain('Excel');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run src/agents/coordinator/__tests__/samCoordinator.test.ts`
Expected: FAIL (Cannot find module)

- [ ] **Step 3: Implement Sam Coordinator**

Create `src/agents/coordinator/samCoordinator.ts`:
```typescript
import { HostType } from '../../types';
import { ExcelAgent } from '../excel/excelAgent';
import { PPTAgent } from '../powerpoint/pptAgent';
import { AgentContext, IAgent } from '../types';
import { WordAgent } from '../word/wordAgent';

export class SamCoordinator {
  private excelAgent = new ExcelAgent();
  private wordAgent = new WordAgent();
  private pptAgent = new PPTAgent();

  getSpecialist(host: HostType): IAgent {
    switch (host) {
      case 'Excel':
        return this.excelAgent;
      case 'Word':
        return this.wordAgent;
      case 'PowerPoint':
        return this.pptAgent;
      default:
        return this.excelAgent; // Default fallback for browser testing
    }
  }

  buildSystemPrompt(host: HostType, context: AgentContext): string {
    const specialist = this.getSpecialist(host);
    return `Anda adalah Sam, asisten AI produktivitas kantor yang ramah, profesional, dan cekatan.
${specialist.getSystemPrompt(context)}

Aturan Penting:
1. Jika pengguna meminta tindakan langsung pada dokumen (seperti mengisi rumus, mengedit teks, membuat slide), gunakan tool yang tersedia.
2. Jelaskan secara singkat dan ramah apa yang Anda lakukan sebelum atau setelah memanggil tool.
3. Selalu utamakan bahasa Indonesia yang baik dan profesional.`;
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run src/agents/coordinator/__tests__/samCoordinator.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/agents/coordinator/ tests/
git commit -m "feat: implement Sam coordinator agent and routing"
```

---

### Task 9: UI Components: Settings Modal & Header Controls

**Files:**
- Create: `src/components/Common/Header.tsx`
- Create: `src/components/Settings/SettingsModal.tsx`
- Modify: `src/index.css`

**Interfaces:**
- Consumes: `getSettings()`, `saveProviderConfig()`, `setActiveProvider()`, `setExecutionMode()`, `testProviderConnection()`
- Produces: Header and Settings UI components.

- [ ] **Step 1: Setup base CSS styles**

Create `src/index.css`:
```css
@tailwind base;
@tailwind components;
@tailwind utilities;

:root {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
}

body {
  margin: 0;
  padding: 0;
  overflow-x: hidden;
}
```

- [ ] **Step 2: Implement Header component**

Create `src/components/Common/Header.tsx`:
```tsx
import React from 'react';
import { HostType, ProviderId, ExecutionMode } from '../../types';
import { Settings, Zap, Play, FileSpreadsheet, FileText, Presentation } from 'lucide-react';

interface HeaderProps {
  host: HostType;
  activeProviderId: ProviderId;
  executionMode: ExecutionMode;
  onToggleMode: () => void;
  onOpenSettings: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  host,
  activeProviderId,
  executionMode,
  onToggleMode,
  onOpenSettings,
}) => {
  const getHostIcon = () => {
    switch (host) {
      case 'Excel':
        return <FileSpreadsheet className="w-5 h-5 text-office-excel" />;
      case 'Word':
        return <FileText className="w-5 h-5 text-office-word" />;
      case 'PowerPoint':
        return <Presentation className="w-5 h-5 text-office-ppt" />;
      default:
        return <Zap className="w-5 h-5 text-office-agent" />;
    }
  };

  return (
    <header className="flex items-center justify-between px-4 py-3 bg-white dark:bg-gray-800 border-b border-gray-200 dark:border-gray-700 shadow-sm">
      <div className="flex items-center space-x-2">
        {getHostIcon()}
        <div>
          <h1 className="text-sm font-bold text-gray-900 dark:text-gray-100 flex items-center gap-1">
            Sam Office
            <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-blue-100 dark:bg-blue-900/60 text-blue-700 dark:text-blue-300 font-medium">
              {host}
            </span>
          </h1>
          <p className="text-[11px] text-gray-500 capitalize">{activeProviderId}</p>
        </div>
      </div>

      <div className="flex items-center space-x-2">
        {/* Copilot vs Autopilot Toggle */}
        <button
          onClick={onToggleMode}
          title={executionMode === 'copilot' ? 'Mode Copilot (Preview dulu)' : 'Mode Autopilot (Langsung eksekusi)'}
          className={`flex items-center gap-1 text-xs px-2.5 py-1 rounded-full font-medium transition-colors ${
            executionMode === 'autopilot'
              ? 'bg-amber-100 dark:bg-amber-900/50 text-amber-800 dark:text-amber-200 border border-amber-300'
              : 'bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300'
          }`}
        >
          {executionMode === 'autopilot' ? (
            <>
              <Play className="w-3 h-3 fill-current" /> Auto
            </>
          ) : (
            <>
              <Zap className="w-3 h-3" /> Copilot
            </>
          )}
        </button>

        <button
          onClick={onOpenSettings}
          className="p-1.5 text-gray-500 hover:text-gray-800 dark:hover:text-gray-200 rounded-lg hover:bg-gray-100 dark:hover:bg-gray-700 transition"
          title="Pengaturan BYOK API Key"
        >
          <Settings className="w-4 h-4" />
        </button>
      </div>
    </header>
  );
};
```

- [ ] **Step 3: Implement BYOK Settings Modal**

Create `src/components/Settings/SettingsModal.tsx`:
```tsx
import React, { useState } from 'react';
import { ProviderConfig, ProviderId } from '../../types';
import { getSettings, saveProviderConfig, setActiveProvider } from '../../services/storage/settingsStorage';
import { testProviderConnection } from '../../services/llm/factory';
import { X, CheckCircle2, AlertCircle, Loader2 } from 'lucide-react';

interface SettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSaved: () => void;
}

export const SettingsModal: React.FC<SettingsModalProps> = ({ isOpen, onClose, onSaved }) => {
  if (!isOpen) return null;

  const [settings, setSettings] = useState(getSettings());
  const [activeTab, setActiveTab] = useState<ProviderId>(settings.activeProviderId);
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);

  const currentConfig = settings.providers[activeTab];

  const handleUpdateKey = (key: string) => {
    const updated = { ...currentConfig, apiKey: key };
    setSettings(prev => ({
      ...prev,
      providers: { ...prev.providers, [activeTab]: updated },
    }));
  };

  const handleUpdateModel = (model: string) => {
    const updated = { ...currentConfig, selectedModel: model };
    setSettings(prev => ({
      ...prev,
      providers: { ...prev.providers, [activeTab]: updated },
    }));
  };

  const handleTest = async () => {
    setTesting(true);
    setTestResult(null);
    try {
      const res = await testProviderConnection(currentConfig);
      setTestResult(res);
    } finally {
      setTesting(false);
    }
  };

  const handleSaveAndApply = () => {
    saveProviderConfig(currentConfig);
    setActiveProvider(activeTab);
    onSaved();
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-3">
      <div className="bg-white dark:bg-gray-800 rounded-xl max-w-sm w-full shadow-xl border border-gray-200 dark:border-gray-700 overflow-hidden flex flex-col max-h-[90vh]">
        <div className="flex items-center justify-between p-4 border-b border-gray-100 dark:border-gray-700">
          <h2 className="text-base font-semibold text-gray-900 dark:text-gray-100">Pengaturan BYOK AI</h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="flex border-b border-gray-200 dark:border-gray-700 overflow-x-auto text-xs">
          {(Object.keys(settings.providers) as ProviderId[]).map(pid => (
            <button
              key={pid}
              onClick={() => {
                setActiveTab(pid);
                setTestResult(null);
              }}
              className={`px-3 py-2 whitespace-nowrap font-medium border-b-2 transition ${
                activeTab === pid
                  ? 'border-blue-600 text-blue-600 dark:text-blue-400'
                  : 'border-transparent text-gray-500 hover:text-gray-800 dark:hover:text-gray-200'
              }`}
            >
              {settings.providers[pid].name.split(' ')[0]}
            </button>
          ))}
        </div>

        <div className="p-4 space-y-3 overflow-y-auto flex-1">
          <div>
            <label className="block text-xs font-medium text-gray-700 dark:text-gray-300 mb-1">
              API Key ({currentConfig.name})
            </label>
            <input
              type="password"
              value={currentConfig.apiKey}
              onChange={e => handleUpdateKey(e.target.value)}
              placeholder="Masukkan API Key Anda..."
              className="w-full text-xs px-3 py-2 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-gray-700 dark:text-gray-300 mb-1">
              Nama Model
            </label>
            <input
              type="text"
              value={currentConfig.selectedModel}
              onChange={e => handleUpdateModel(e.target.value)}
              className="w-full text-xs px-3 py-2 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>

          {testResult && (
            <div
              className={`p-2.5 rounded-lg text-xs flex items-start gap-2 ${
                testResult.success
                  ? 'bg-green-50 dark:bg-green-950/40 text-green-700 dark:text-green-300'
                  : 'bg-red-50 dark:bg-red-950/40 text-red-700 dark:text-red-300'
              }`}
            >
              {testResult.success ? (
                <CheckCircle2 className="w-4 h-4 shrink-0 text-green-600 mt-0.5" />
              ) : (
                <AlertCircle className="w-4 h-4 shrink-0 text-red-600 mt-0.5" />
              )}
              <span className="break-all">{testResult.message}</span>
            </div>
          )}
        </div>

        <div className="p-3 bg-gray-50 dark:bg-gray-900/50 border-t border-gray-200 dark:border-gray-700 flex justify-between gap-2">
          <button
            onClick={handleTest}
            disabled={testing}
            className="px-3 py-1.5 text-xs font-medium rounded-lg border border-gray-300 dark:border-gray-600 hover:bg-gray-100 dark:hover:bg-gray-800 text-gray-700 dark:text-gray-300 flex items-center gap-1.5"
          >
            {testing && <Loader2 className="w-3.5 h-3.5 animate-spin" />}
            Test Koneksi
          </button>
          <button
            onClick={handleSaveAndApply}
            className="px-3 py-1.5 text-xs font-medium rounded-lg bg-blue-600 hover:bg-blue-700 text-white shadow-sm"
          >
            Simpan & Gunakan
          </button>
        </div>
      </div>
    </div>
  );
};
```

- [ ] **Step 4: Commit**

```bash
git add src/index.css src/components/Common/Header.tsx src/components/Settings/SettingsModal.tsx
git commit -m "feat: implement header and BYOK settings modal UI"
```

---

### Task 10: UI Components: Chat Interface & Action Preview Cards

**Files:**
- Create: `src/components/Actions/ActionCard.tsx`
- Create: `src/components/Chat/MessageBubble.tsx`
- Create: `src/components/Chat/InputBar.tsx`
- Create: `src/components/Chat/ChatContainer.tsx`

**Interfaces:**
- Consumes: `ChatMessage`, `ToolCall`, `IAgent`
- Produces: Chat view with streaming bubbles and actionable diff cards.

- [ ] **Step 1: Implement ActionCard component**

Create `src/components/Actions/ActionCard.tsx`:
```tsx
import React from 'react';
import { ToolCall } from '../../types';
import { Check, CheckCircle2, Play, AlertCircle } from 'lucide-react';

interface ActionCardProps {
  toolCall: ToolCall;
  onApply: (toolCall: ToolCall) => void;
  isExecuting?: boolean;
}

export const ActionCard: React.FC<ActionCardProps> = ({ toolCall, onApply, isExecuting }) => {
  const isApplied = toolCall.status === 'applied';
  const isFailed = toolCall.status === 'failed';

  return (
    <div className="mt-2 p-3 bg-gray-50 dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-lg text-xs space-y-2">
      <div className="flex items-center justify-between font-semibold text-gray-800 dark:text-gray-200">
        <span className="capitalize">{toolCall.name.replace(/_/g, ' ')}</span>
        {isApplied && (
          <span className="flex items-center gap-1 text-green-600 dark:text-green-400 font-medium">
            <CheckCircle2 className="w-3.5 h-3.5" /> Diterapkan
          </span>
        )}
        {isFailed && (
          <span className="flex items-center gap-1 text-red-600 dark:text-red-400 font-medium">
            <AlertCircle className="w-3.5 h-3.5" /> Gagal
          </span>
        )}
      </div>

      <div className="bg-white dark:bg-gray-800 p-2 rounded border border-gray-100 dark:border-gray-700 font-mono text-[11px] text-gray-700 dark:text-gray-300 max-h-24 overflow-y-auto">
        {toolCall.name === 'write_cells' && (
          <div>
            <span className="text-blue-500">Range:</span> {toolCall.arguments.range}
            {toolCall.arguments.formula && (
              <div><span className="text-green-500">Formula:</span> {toolCall.arguments.formula}</div>
            )}
          </div>
        )}
        {toolCall.name === 'format_range' && (
          <div>
            <span className="text-blue-500">Range:</span> {toolCall.arguments.range}
            {toolCall.arguments.numberFormat && (
              <div><span className="text-purple-500">Format:</span> {toolCall.arguments.numberFormat}</div>
            )}
          </div>
        )}
        {toolCall.name === 'create_chart' && (
          <div>
            <span className="text-amber-500">Chart:</span> {toolCall.arguments.chartType} ({toolCall.arguments.dataRange})
          </div>
        )}
        {toolCall.name !== 'write_cells' && toolCall.name !== 'format_range' && toolCall.name !== 'create_chart' && (
          <pre className="whitespace-pre-wrap">{JSON.stringify(toolCall.arguments, null, 2)}</pre>
        )}
      </div>

      {!isApplied && (
        <button
          onClick={() => onApply(toolCall)}
          disabled={isExecuting}
          className="w-full py-1.5 px-3 bg-blue-600 hover:bg-blue-700 text-white rounded font-medium flex items-center justify-center gap-1.5 transition disabled:opacity-50"
        >
          {isExecuting ? 'Menerapkan...' : (
            <>
              <Play className="w-3 h-3 fill-current" /> Terapkan ke Dokumen
            </>
          )}
        </button>
      )}
    </div>
  );
};
```

- [ ] **Step 2: Implement MessageBubble component**

Create `src/components/Chat/MessageBubble.tsx`:
```tsx
import React from 'react';
import { ChatMessage, ToolCall } from '../../types';
import { ActionCard } from '../Actions/ActionCard';
import { Bot, User } from 'lucide-react';

interface MessageBubbleProps {
  message: ChatMessage;
  onApplyToolCall: (toolCall: ToolCall) => void;
}

export const MessageBubble: React.FC<MessageBubbleProps> = ({ message, onApplyToolCall }) => {
  const isUser = message.role === 'user';

  return (
    <div className={`flex gap-2 mb-3 ${isUser ? 'justify-end' : 'justify-start'}`}>
      {!isUser && (
        <div className="w-6 h-6 rounded-full bg-blue-600 text-white flex items-center justify-center shrink-0 mt-1">
          <Bot className="w-3.5 h-3.5" />
        </div>
      )}

      <div
        className={`max-w-[85%] rounded-xl px-3.5 py-2.5 text-xs ${
          isUser
            ? 'bg-blue-600 text-white rounded-tr-none'
            : 'bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 text-gray-800 dark:text-gray-200 rounded-tl-none shadow-sm'
        }`}
      >
        <div className="whitespace-pre-wrap leading-relaxed">{message.content}</div>

        {message.toolCalls && message.toolCalls.map(tc => (
          <ActionCard key={tc.id} toolCall={tc} onApply={onApplyToolCall} />
        ))}
      </div>

      {isUser && (
        <div className="w-6 h-6 rounded-full bg-gray-300 dark:bg-gray-700 text-gray-700 dark:text-gray-300 flex items-center justify-center shrink-0 mt-1">
          <User className="w-3.5 h-3.5" />
        </div>
      )}
    </div>
  );
};
```

- [ ] **Step 3: Implement InputBar and ChatContainer components**

Create `src/components/Chat/InputBar.tsx`:
```tsx
import React, { useState, KeyboardEvent } from 'react';
import { Send, Sparkles } from 'lucide-react';

interface InputBarProps {
  onSendMessage: (text: string) => void;
  disabled?: boolean;
}

export const InputBar: React.FC<InputBarProps> = ({ onSendMessage, disabled }) => {
  const [input, setInput] = useState('');

  const handleSend = () => {
    if (!input.trim() || disabled) return;
    onSendMessage(input.trim());
    setInput('');
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  return (
    <div className="p-3 bg-white dark:bg-gray-800 border-t border-gray-200 dark:border-gray-700">
      <div className="flex items-center gap-2 bg-gray-50 dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-xl px-3 py-1.5 focus-within:ring-2 focus-within:ring-blue-500">
        <Sparkles className="w-4 h-4 text-blue-500 shrink-0" />
        <input
          type="text"
          value={input}
          onChange={e => setInput(e.target.value)}
          onKeyDown={handleKeyDown}
          disabled={disabled}
          placeholder="Tanya Sam atau perintahkan sesuatu..."
          className="w-full bg-transparent text-xs text-gray-900 dark:text-gray-100 focus:outline-none"
        />
        <button
          onClick={handleSend}
          disabled={!input.trim() || disabled}
          className="p-1 rounded-lg text-blue-600 hover:bg-blue-50 dark:hover:bg-gray-800 disabled:opacity-40 transition"
        >
          <Send className="w-4 h-4" />
        </button>
      </div>
    </div>
  );
};
```

Create `src/components/Chat/ChatContainer.tsx`:
```tsx
import React, { useRef, useEffect } from 'react';
import { ChatMessage, ToolCall } from '../../types';
import { MessageBubble } from './MessageBubble';

interface ChatContainerProps {
  messages: ChatMessage[];
  onApplyToolCall: (toolCall: ToolCall) => void;
}

export const ChatContainer: React.FC<ChatContainerProps> = ({ messages, onApplyToolCall }) => {
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  return (
    <div className="flex-1 overflow-y-auto p-4 space-y-2">
      {messages.map(m => (
        <MessageBubble key={m.id} message={m} onApplyToolCall={onApplyToolCall} />
      ))}
      <div ref={bottomRef} />
    </div>
  );
};
```

- [ ] **Step 4: Commit**

```bash
git add src/components/Actions/ src/components/Chat/
git commit -m "feat: implement chat container, message bubbles, and action cards"
```

---

### Task 11: Main App Integration & Universal Manifest (`manifest.xml`)

**Files:**
- Create: `manifest.xml`
- Create: `src/App.tsx`
- Create: `src/main.tsx`
- Test: `tests/app.test.tsx`

**Interfaces:**
- Consumes: All services, agents, and components.
- Produces: Complete working universal Office Web Add-in.

- [ ] **Step 1: Create Office Add-in `manifest.xml`**

Create `manifest.xml`:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">

  <Id>d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Sam Office</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sam Office Agent" />
  <Description DefaultValue="Universal multi-agent AI assistant for Microsoft Excel, Word, and PowerPoint." />
  <IconUrl DefaultValue="https://localhost:5173/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://localhost:5173/icon-64.png" />
  <SupportUrl DefaultValue="https://github.com/Sam-Office-Agent" />

  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Document" />
    <Host Name="Presentation" />
  </Hosts>

  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:5173/" />
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

- [ ] **Step 2: Create main Application component (`src/App.tsx`)**

Create `src/App.tsx`:
```tsx
import React, { useState, useEffect } from 'react';
import { Header } from './components/Common/Header';
import { ChatContainer } from './components/Chat/ChatContainer';
import { InputBar } from './components/Chat/InputBar';
import { SettingsModal } from './components/Settings/SettingsModal';
import { ChatMessage, ExecutionMode, HostType, ToolCall } from './types';
import { getActiveProvider, getExecutionMode, getSettings, setExecutionMode } from './services/storage/settingsStorage';
import { SamCoordinator } from './agents/coordinator/samCoordinator';
import { getLLMProvider } from './services/llm/factory';
import { getOfficeDriver } from './services/office';

export const App: React.FC<{ initialHost?: HostType }> = ({ initialHost = 'Excel' }) => {
  const [host, setHost] = useState<HostType>(initialHost);
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: 'welcome',
      role: 'assistant',
      content: `Halo! Saya Sam, asisten AI untuk Microsoft Office (${initialHost}). Apa yang bisa saya bantu hari ini?`,
      timestamp: Date.now(),
    },
  ]);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [executionMode, setMode] = useState<ExecutionMode>(getExecutionMode());
  const [coordinator] = useState(() => new SamCoordinator());

  useEffect(() => {
    if (typeof (window as any).Office !== 'undefined' && (window as any).Office.context?.host) {
      const h = (window as any).Office.context.host;
      if (h === 'Excel') setHost('Excel');
      else if (h === 'Word') setHost('Word');
      else if (h === 'PowerPoint') setHost('PowerPoint');
    }
  }, []);

  const handleToggleMode = () => {
    const nextMode = executionMode === 'copilot' ? 'autopilot' : 'copilot';
    setMode(nextMode);
    setExecutionMode(nextMode);
  };

  const handleApplyToolCall = async (toolCall: ToolCall) => {
    const specialist = coordinator.getSpecialist(host);
    toolCall.status = 'pending';
    const result = await specialist.executeTool(toolCall, { host });
    toolCall.status = result.success ? 'applied' : 'failed';
    toolCall.error = result.error;
    setMessages(prev => [...prev]);
  };

  const handleSendMessage = async (text: string) => {
    const userMsg: ChatMessage = {
      id: `msg_${Date.now()}`,
      role: 'user',
      content: text,
      timestamp: Date.now(),
    };

    const assistantMsgId = `asst_${Date.now()}`;
    const assistantMsg: ChatMessage = {
      id: assistantMsgId,
      role: 'assistant',
      content: '',
      timestamp: Date.now(),
      toolCalls: [],
    };

    setMessages(prev => [...prev, userMsg, assistantMsg]);

    const providerConfig = getActiveProvider();
    const provider = getLLMProvider(providerConfig.id);
    const specialist = coordinator.getSpecialist(host);
    const systemPrompt = coordinator.buildSystemPrompt(host, { host });
    const tools = specialist.getTools();

    try {
      const stream = provider.sendMessage(
        {
          messages: [...messages, userMsg],
          systemPrompt,
          tools,
        },
        providerConfig
      );

      for await (const chunk of stream) {
        if (chunk.type === 'content_delta' && chunk.delta) {
          assistantMsg.content += chunk.delta;
          setMessages(prev => prev.map(m => (m.id === assistantMsgId ? { ...assistantMsg } : m)));
        } else if (chunk.type === 'tool_call' && chunk.toolCall) {
          assistantMsg.toolCalls?.push(chunk.toolCall);
          setMessages(prev => prev.map(m => (m.id === assistantMsgId ? { ...assistantMsg } : m)));

          if (executionMode === 'autopilot') {
            await handleApplyToolCall(chunk.toolCall);
          }
        } else if (chunk.type === 'error') {
          assistantMsg.content += `\n[Error: ${chunk.error}]`;
          setMessages(prev => prev.map(m => (m.id === assistantMsgId ? { ...assistantMsg } : m)));
        }
      }
    } catch (e: any) {
      assistantMsg.content += `\n[Error: ${e.message}]`;
      setMessages(prev => prev.map(m => (m.id === assistantMsgId ? { ...assistantMsg } : m)));
    }
  };

  return (
    <div className="flex flex-col h-screen bg-gray-50 dark:bg-gray-900 text-gray-900 dark:text-gray-100">
      <Header
        host={host}
        activeProviderId={getActiveProvider().id}
        executionMode={executionMode}
        onToggleMode={handleToggleMode}
        onOpenSettings={() => setIsSettingsOpen(true)}
      />

      <ChatContainer messages={messages} onApplyToolCall={handleApplyToolCall} />

      <InputBar onSendMessage={handleSendMessage} />

      <SettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        onSaved={() => setMessages(prev => [...prev])}
      />
    </div>
  );
};
```

- [ ] **Step 3: Create Entry Point `src/main.tsx`**

Create `src/main.tsx`:
```tsx
import React from 'react';
import ReactDOM from 'react-dom/client';
import { App } from './App';
import './index.css';
import { HostType } from './types';

function renderApp(host: HostType = 'BrowserDev') {
  const root = document.getElementById('root');
  if (root) {
    ReactDOM.createRoot(root).render(
      <React.StrictMode>
        <App initialHost={host} />
      </React.StrictMode>
    );
  }
}

// Check if running inside Microsoft Office
if (typeof (window as any).Office !== 'undefined') {
  (window as any).Office.onReady((info: any) => {
    let detectedHost: HostType = 'BrowserDev';
    if (info.host === (window as any).Office.HostType.Excel) detectedHost = 'Excel';
    else if (info.host === (window as any).Office.HostType.Word) detectedHost = 'Word';
    else if (info.host === (window as any).Office.HostType.PowerPoint) detectedHost = 'PowerPoint';
    renderApp(detectedHost);
  });
} else {
  // Standalone browser mode for local development
  renderApp('BrowserDev');
}
```

- [ ] **Step 4: Build and Test End-to-End**

Run: `npm run build`
Expected: TypeScript check passes and Vite bundles successfully to `dist/`.

- [ ] **Step 5: Commit**

```bash
git add manifest.xml src/App.tsx src/main.tsx
git commit -m "feat: integrate main application, Office.onReady lifecycle, and universal manifest"
```
