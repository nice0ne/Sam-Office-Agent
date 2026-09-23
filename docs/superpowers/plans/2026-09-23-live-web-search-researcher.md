# Live Web Search & Autonomous Research Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the Autonomous Live Web Search Service for Sam Office Agent with zero-config DuckDuckGo and optional Tavily AI search, enabling Excel, Word, and PowerPoint agents to research real-time facts, currency rates, and market statistics and inject them directly into active documents with transparent citations.

**Architecture:** Build a modular search service (`src/services/search/`) with a DuckDuckGo HTML scraper and Tavily API client with auto-fallback. Wire up a universal `web_search` tool across `ExcelAgent`, `WordAgent`, and `PPTAgent`. Expose search engine preferences in `SettingsModal.tsx` and 1-click Quick Action Presets in `QuickActionPresets.tsx`.

**Tech Stack:** TypeScript, React 18, Vitest, Tailwind CSS, Lucide icons.

**Spec:** [docs/superpowers/specs/2026-09-23-live-web-search-researcher-design.md](file:///E:/VIBE-CODE-WS/Sam-Office-Agent/docs/superpowers/specs/2026-09-23-live-web-search-researcher-design.md)

## Global Constraints

- 100% Client-Side BYOK: Search requests originate directly from the user's browser/WebView2 context without third-party proxy relays.
- Zero-config default: DuckDuckGo requires no API keys or setup.
- TDD workflow: Write tests before implementation for every task.
- Backward compatibility: All 294 existing tests must remain passing.

---

### Task 1: Search Service Engine & Settings Storage

**Files:**
- Create: `src/services/search/types.ts`
- Create: `src/services/search/duckduckgo.ts`
- Create: `src/services/search/tavily.ts`
- Create: `src/services/search/index.ts`
- Modify: `src/services/storage/settingsStorage.ts`
- Create: `tests/webSearch.test.ts`

**Interfaces:**
- Produces:
  - Types: `SearchResultItem`, `SearchOptions`, `SearchResponse`, `SearchSettings`.
  - Functions: `searchWeb(query: string, options?: SearchOptions): Promise<SearchResponse>`, `getSearchSettings(): SearchSettings`, `setSearchSettings(s: SearchSettings): void`.

- [ ] **Step 1: Write the failing test**

In `tests/webSearch.test.ts`:
```typescript
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { searchWeb } from '../src/services/search';
import { getSearchSettings, setSearchSettings } from '../src/services/storage/settingsStorage';

describe('Web Search Service', () => {
  beforeEach(() => {
    localStorage.clear();
    vi.restoreAllMocks();
  });

  it('manages search settings with DuckDuckGo as default', () => {
    const settings = getSearchSettings();
    expect(settings.searchProvider).toBe('duckduckgo');

    setSearchSettings({ searchProvider: 'tavily', tavilyApiKey: 'tvly-test-123' });
    const updated = getSearchSettings();
    expect(updated.searchProvider).toBe('tavily');
    expect(updated.tavilyApiKey).toBe('tvly-test-123');
  });

  it('parses DuckDuckGo HTML results and formats clean citations', async () => {
    const mockHtml = `
      <div class="result web-result">
        <a class="result__url" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fwww.bi.go.id%2Fkurs">Bank Indonesia</a>
        <h2 class="result__title"><a class="result__a" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fwww.bi.go.id%2Fkurs">Informasi Kurs BI</a></h2>
        <a class="result__snippet">Kurs transaksi USD ke IDR hari ini adalah Rp16.200 per dolar.</a>
      </div>
    `;

    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      text: async () => mockHtml,
    } as any);

    const res = await searchWeb('kurs USD ke IDR');
    expect(res.provider).toBe('duckduckgo');
    expect(res.results.length).toBe(1);
    expect(res.results[0].title).toBe('Informasi Kurs BI');
    expect(res.results[0].url).toBe('https://www.bi.go.id/kurs');
    expect(res.results[0].snippet).toContain('Rp16.200');
    expect(res.citationsFormatted).toContain('https://www.bi.go.id/kurs');
  });

  it('falls back to DuckDuckGo when Tavily API key is missing or fails', async () => {
    setSearchSettings({ searchProvider: 'tavily', tavilyApiKey: '' });

    const mockHtml = `
      <div class="result">
        <h2 class="result__title"><a class="result__a" href="https://example.com">Fallback Result</a></h2>
        <a class="result__snippet">Contoh snippet fallback.</a>
      </div>
    `;

    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      text: async () => mockHtml,
    } as any);

    const res = await searchWeb('test query');
    expect(res.provider).toBe('duckduckgo');
    expect(res.results.length).toBeGreaterThan(0);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/webSearch.test.ts`
Expected: FAIL with module not found.

- [ ] **Step 3: Implement search service & settings**

- In `src/services/storage/settingsStorage.ts`: add `getSearchSettings` and `setSearchSettings`.
- In `src/services/search/types.ts`: define contracts.
- In `src/services/search/duckduckgo.ts`: implement HTML parser with `DOMParser` and `cleanDdgUrl`.
- In `src/services/search/tavily.ts`: implement JSON client.
- In `src/services/search/index.ts`: implement `searchWeb` with auto-fallback and citation formatting.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/webSearch.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/search/ src/services/storage/settingsStorage.ts tests/webSearch.test.ts
git commit -m "feat(search): implement modular web search service and settings"
```

---

### Task 2: Universal `web_search` Tool in Agent Specialists

**Files:**
- Modify: `src/agents/excel/excelAgent.ts`
- Modify: `src/agents/word/wordAgent.ts`
- Modify: `src/agents/powerpoint/pptAgent.ts`
- Modify: `src/agents/coordinator/samCoordinator.ts`
- Create: `tests/webSearchAgents.test.ts`

**Interfaces:**
- Consumes: `searchWeb` from `src/services/search`.
- Produces: `web_search` tool declared in `getTools()` and handled in `executeTool()` for all specialists.

- [ ] **Step 1: Write the failing test**

In `tests/webSearchAgents.test.ts`:
```typescript
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ExcelAgent } from '../src/agents/excel/excelAgent';
import { WordAgent } from '../src/agents/word/wordAgent';
import { PPTAgent } from '../src/agents/powerpoint/pptAgent';
import { MockOfficeDriver } from '../src/services/office/mockDriver';

describe('Universal web_search tool across specialists', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    const mockHtml = `
      <div class="result">
        <h2 class="result__title"><a class="result__a" href="https://bi.go.id">Bank Indonesia Kurs</a></h2>
        <a class="result__snippet">Kurs USD adalah Rp16.250.</a>
      </div>
    `;
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      text: async () => mockHtml,
    } as any);
  });

  it('provides web_search in ExcelAgent and returns search results with citations', async () => {
    const driver = new MockOfficeDriver();
    const agent = new ExcelAgent(driver);
    expect(agent.getTools().some(t => t.name === 'web_search')).toBe(true);

    const res = await agent.executeTool(
      { id: 'w1', name: 'web_search', arguments: { query: 'kurs dollar hari ini' }, status: 'pending' },
      { host: 'Excel' }
    );
    expect(res.success).toBe(true);
    expect(res.result).toContain('Bank Indonesia Kurs');
    expect(res.result).toContain('Rp16.250');
    expect(res.result).toContain('Sumber Referensi');
  });

  it('provides web_search in WordAgent', async () => {
    const driver = new MockOfficeDriver();
    const agent = new WordAgent(driver);
    expect(agent.getTools().some(t => t.name === 'web_search')).toBe(true);

    const res = await agent.executeTool(
      { id: 'w2', name: 'web_search', arguments: { query: 'tren HR 2026' }, status: 'pending' },
      { host: 'Word' }
    );
    expect(res.success).toBe(true);
    expect(res.result).toContain('Sumber Referensi');
  });

  it('provides web_search in PPTAgent', async () => {
    const driver = new MockOfficeDriver();
    const agent = new PPTAgent(driver);
    expect(agent.getTools().some(t => t.name === 'web_search')).toBe(true);

    const res = await agent.executeTool(
      { id: 'w3', name: 'web_search', arguments: { query: 'statistik e-commerce' }, status: 'pending' },
      { host: 'PowerPoint' }
    );
    expect(res.success).toBe(true);
    expect(res.result).toContain('Sumber Referensi');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/webSearchAgents.test.ts`
Expected: FAIL.

- [ ] **Step 3: Register and handle `web_search` in `ExcelAgent`, `WordAgent`, and `PPTAgent`**

- In `ExcelAgent`, `WordAgent`, and `PPTAgent`:
  - Add `web_search` definition to `getTools()`.
  - Add execution block in `executeTool()`, calling `searchWeb(query, { maxResults })` and returning formatted results with citations.
- In `SamCoordinator.buildSystemPrompt`:
  - Add guideline on using `web_search` when external real-time data or research is requested.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/webSearchAgents.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/agents/ tests/webSearchAgents.test.ts
git commit -m "feat(agents): register and handle universal web_search tool across specialists"
```

---

### Task 3: Settings UI, Quick Action Presets & Full Verification

**Files:**
- Modify: `src/components/Settings/SettingsModal.tsx`
- Modify: `src/components/Chat/QuickActionPresets.tsx`
- Modify: `src/components/__tests__/headerAndSettings.test.tsx`
- Modify: `src/components/__tests__/chatAndActions.test.tsx`

**Interfaces:**
- Exposes web search engine controls in SettingsModal and adds 1-click web research presets to all hosts.

- [ ] **Step 1: Write the failing tests**

In `src/components/__tests__/headerAndSettings.test.tsx`:
```typescript
it('renders web search provider settings in SettingsModal', () => {
  render(<SettingsModal isOpen={true} onClose={vi.fn()} />);
  expect(screen.getByText(/Pencarian Web/i)).toBeInTheDocument();
  expect(screen.getByText(/DuckDuckGo/i)).toBeInTheDocument();
});
```

In `src/components/__tests__/chatAndActions.test.tsx`:
```typescript
it('renders web research presets for Excel, Word, and PowerPoint hosts', () => {
  const { rerender } = render(<QuickActionPresets host="Excel" onSelectPreset={vi.fn()} />);
  expect(screen.getByText(/Riset Web/i)).toBeInTheDocument();

  rerender(<QuickActionPresets host="Word" onSelectPreset={vi.fn()} />);
  expect(screen.getByText(/Riset Web/i)).toBeInTheDocument();

  rerender(<QuickActionPresets host="PowerPoint" onSelectPreset={vi.fn()} />);
  expect(screen.getByText(/Riset Data & Buat Slide/i)).toBeInTheDocument();
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx vitest run src/components/__tests__/headerAndSettings.test.tsx src/components/__tests__/chatAndActions.test.tsx`
Expected: FAIL.

- [ ] **Step 3: Update `SettingsModal.tsx` and `QuickActionPresets.tsx`**

- Add Search Settings card in `SettingsModal.tsx` with provider selector and Tavily API key field.
- Add search presets to `PRESETS_BY_HOST` in `QuickActionPresets.tsx`.

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx vitest run src/components/__tests__/headerAndSettings.test.tsx src/components/__tests__/chatAndActions.test.tsx`
Expected: PASS.

- [ ] **Step 5: Run full test suite to ensure 0 regressions**

Run: `npx vitest run`
Expected: All 20+ test files and 294+ tests pass.

- [ ] **Step 6: Commit**

```bash
git add src/components/Settings/SettingsModal.tsx src/components/Chat/QuickActionPresets.tsx src/components/__tests__/
git commit -m "feat(ui): add web search settings and quick action presets"
```
