# Live Web Search & Autonomous Research Design Spec

## 1. Overview & Business Value

In corporate daily office work, knowledge workers regularly require live, up-to-date external facts, market statistics, currency exchange rates, commodity prices, industry trends, and competitor intelligence that cannot be found in their current local documents or within static LLM training cutoffs.

### Key Pain Points
1. **Static LLM Knowledge Cutoff**: Standard offline LLMs cannot answer questions about real-time currency rates, today's news, or current commodity prices, often leading to hallucinations or refusal.
2. **Context Switching Friction**: Users must leave Microsoft Office, open a web browser, search for statistics, copy tables/snippets, return to Office, and format the data manually.
3. **Missing Citations**: Data pasted from manual web searches often lacks transparent sources, making corporate audits difficult.

### The Solution: Autonomous Web Search Service
Inspired by `sam-agent`'s zero-config web search capability, this feature equips **Sam Office Agent** with an autonomous, provider-agnostic, zero-config web search engine:
- **Zero-Config Default (DuckDuckGo)**: Built-in free web search parsing `https://html.duckduckgo.com/html/` without requiring any API keys or registrations.
- **Optional Premium Provider (Tavily AI)**: Configurable in Settings for structured AI-grounded deep research.
- **Universal ReAct Tool (`web_search`)**: Registered across all Office specialists (`ExcelAgent`, `WordAgent`, and `PPTAgent`).
- **Direct Document Action Pipeline**: The agent autonomously executes web searches and immediately applies the results into the active document (e.g. writing table data to Excel cells, formatting reports in Word, or building presentation slides in PowerPoint).
- **Transparent Citations**: Clean, structured reference links appended to the chat message.

---

## 2. Architecture & Data Contracts

### 2.1 Types (`src/services/search/types.ts`)

```typescript
export interface SearchResultItem {
  title: string;
  url: string;
  snippet: string;
}

export interface SearchOptions {
  maxResults?: number;
  provider?: 'duckduckgo' | 'tavily';
  tavilyApiKey?: string;
}

export interface SearchResponse {
  query: string;
  provider: 'duckduckgo' | 'tavily';
  results: SearchResultItem[];
  citationsFormatted: string;
  error?: string;
}
```

### 2.2 Storage & Settings Extensions (`src/services/storage/settingsStorage.ts`)

```typescript
export interface SearchSettings {
  searchProvider: 'duckduckgo' | 'tavily';
  tavilyApiKey?: string;
}

export function getSearchSettings(): SearchSettings;
export function setSearchSettings(settings: SearchSettings): void;
```

---

## 3. Component Design & Implementation

### 3.1 Search Service (`src/services/search/`)

- **`duckduckgo.ts`**:
  - Sends a POST request to `https://html.duckduckgo.com/html/` with form-encoded `q={query}&b=`.
  - Parses `.result, .web-result` nodes using browser-standard `DOMParser`.
  - Extracts title text, raw URL, and text snippet.
  - Decodes `uddg` redirected query parameter to return clean target links.
- **`tavily.ts`**:
  - Sends a JSON POST request to `https://api.tavily.com/search`.
  - Extracts `title`, `url`, and `content` as snippets.
- **`index.ts` (`searchWeb`)**:
  - Resolves active search settings.
  - If provider is `tavily` and `tavilyApiKey` is configured, calls `searchWithTavily()`.
  - Falls back to `searchWithDuckDuckGo()` if Tavily fails or key is empty.
  - Formats clean Markdown citation badges:
    ```markdown
    🌐 **Sumber Referensi:**
    1. [Judul 1](url1)
    2. [Judul 2](url2)
    ```

### 3.2 Universal Agent ReAct Tool (`web_search`)

Registered in `ExcelAgent`, `WordAgent`, and `PPTAgent`:
- **Tool Schema**:
  - `name`: `web_search`
  - `description`: `"Mencari informasi, data terkini, statistik, kurs valuta asing, harga komoditas, atau fakta terbaru dari internet untuk disintesis dan dimasukkan ke dalam dokumen."`
  - `parameters`: `{ query: string, maxResults?: number }`
- **Execution**:
  - Calls `searchWeb(query, { maxResults })`.
  - Injects search results into LLM ReAct context.
  - LLM subsequently invokes host writing tools (`write_cells`, `insert_table`, `create_presentation_deck`) and presents findings with citations in chat.

### 3.3 UI Integration

- **Settings Card (`SettingsModal.tsx`)**:
  - Provider selector: `DuckDuckGo (Gratis Zero-Config)` vs `Tavily AI Search`.
  - Password-masked input field for Tavily API Key with save/persist logic.
- **Quick Action Presets (`QuickActionPresets.tsx`)**:
  - Excel: `🌐 Riset Web & Buat Tabel`
  - Word: `🌐 Riset Web & Tulis Laporan`
  - PowerPoint: `🌐 Riset Data & Buat Slide`

---

## 4. Verification & Testing Strategy

1. **Unit Tests for Search Service (`tests/webSearch.test.ts`)**:
   - Verify DuckDuckGo HTML parser extracting results, decoding uddg URLs, and handling empty/error responses.
   - Verify Tavily client and automatic fallback to DuckDuckGo when API key is missing or request fails.
   - Verify citation link formatting.
2. **Specialist Agents Tests (`src/agents/__tests__/specialistAgents.test.ts`)**:
   - Verify `web_search` tool exists and executes in `ExcelAgent`, `WordAgent`, and `PPTAgent`.
3. **Settings & Preset UI Tests (`src/components/__tests__/headerAndSettings.test.tsx`, `chatAndActions.test.tsx`)**:
   - Verify search provider toggle and preset button rendering.
4. **Full Test Suite Run**:
   - Verify that all existing 294+ tests continue to pass with 0 regressions.
