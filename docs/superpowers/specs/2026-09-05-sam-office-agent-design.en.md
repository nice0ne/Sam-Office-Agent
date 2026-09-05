# Design Specification: Sam-Office-Agent

- **Date**: 2026-09-05
- **Topic**: Universal Microsoft Office Multi-Agent Assistant (Excel, Word, PowerPoint) with BYOK
- **Status**: Approved
- **Repository**: `D:/VIBE-CODING/Sam-Office-Agent`

English | [Bahasa Indonesia](2026-09-05-sam-office-agent-design.md)

---

## 1. Executive Summary & Objective

**Sam-Office-Agent** is an office productivity AI assistant based on the **Microsoft Office Web Add-in** architecture that runs in the side panel (*Taskpane*) of Microsoft Office applications (**Excel, Word, PowerPoint**), supporting both Desktop Office 365 and Office Online (Web) versions.

The system adopts the **BYOK (Bring Your Own Key)** principle, allowing users to input their own LLM API keys (OpenAI, Google Gemini, Anthropic Claude, GLM Coding Global / Zhipu AI, OpenRouter/DeepSeek, and Local Ollama) without any third-party intermediary servers. All processing runs directly on the browser/taskpane side (*client-side sandbox*), making it exceptionally lightweight (50–120 MB RAM), secure, fast, and isolated.

---

## 2. Core Architecture & Tech Stack

### 2.1 Technology Stack
* **Framework**: React 18+ with TypeScript.
* **Build Tool**: Vite (supports local HTTPS dev server for Office Add-ins).
* **Styling**: Tailwind CSS + Shadcn UI / Radix UI + Lucide Icons (supports modern Light and Dark themes).
* **Office Integration SDK**: Official `@types/office-js` and `Office.js` from Microsoft CDN.
* **Testing**: Vitest for unit tests (LLM adapters, logic validators) and Mock Office Driver for testing in standard browsers.

### 2.2 Host Auto-Detection
The Add-in uses a single universal `manifest.xml` file registered with Word, Excel, and PowerPoint. When the taskpane opens:
1. `Office.onReady()` checks `Office.context.host`.
2. The UI and agent modules automatically adapt:
   * If in **Excel**: Activates `ExcelAgent`, green Excel badge, and formula/table toolbar.
   * If in **Word**: Activates `WordAgent`, blue Word badge, and text/document toolbar.
   * If in **PowerPoint**: Activates `PPTAgent`, orange PPT badge, and slide/outline toolbar.
   * If in **Standard Browser (Dev)**: Activates `MockOfficeDriver` for local simulation without Office.

### 2.3 File & Directory Structure
```text
Sam-Office-Agent/
├── manifest.xml                 # Office Add-in manifest mendaftarkan taskpane untuk Excel, Word, PPT
├── package.json                 # Project dependencies & scripts
├── vite.config.ts               # Vite configuration (HTTPS certs, path aliases)
├── tsconfig.json                # TypeScript strict configuration
├── docs/
│   └── superpowers/
│       └── specs/
│           └── 2026-09-05-sam-office-agent-design.md
├── src/
│   ├── index.html               # Entry point HTML yang memuat script Office.js
│   ├── main.tsx                 # Office.onReady() bootstrap & React root
│   ├── components/
│   │   ├── Chat/                # Chat container, MessageBubble, StreamingCursor, InputBar
│   │   ├── Actions/             # ActionCard (Diff preview, step status, Apply Button)
│   │   ├── Settings/            # Modal konfigurasi BYOK, Test Connection, Model selector
│   │   └── Common/              # Header, HostBadge, ThemeToggle, ModeToggle (Copilot/Autopilot)
│   ├── agents/
│   │   ├── types.ts             # Definisi Agent, ToolDefinition, ToolCall, ExecutionContext
│   │   ├── coordinator/         # Sam Coordinator (Intent parser, task decomposition, memory)
│   │   ├── excel/               # Excel Specialist Agent & tool implementations
│   │   ├── word/                # Word Specialist Agent & tool implementations
│   │   └── powerpoint/          # PowerPoint Specialist Agent & tool implementations
│   ├── services/
│   │   ├── llm/                 # Unified BYOK Adapter
│   │   │   ├── types.ts         # ILLMProvider interface
│   │   │   ├── openai.ts        # OpenAI adapter (GPT-4o, mini, o3-mini)
│   │   │   ├── gemini.ts        # Gemini adapter (Gemini 2.0 Flash, 1.5 Pro)
│   │   │   ├── anthropic.ts     # Claude adapter (Claude 3.5 Sonnet, Haiku)
│   │   │   ├── glm.ts           # GLM Coding Global (Zhipu BigModel API)
│   │   │   ├── openrouter.ts    # OpenRouter & DeepSeek API
│   │   │   └── ollama.ts        # Local Ollama / Custom Base URL adapter
│   │   ├── office/              # Office.js Safe Driver (Batch sync executor)
│   │   │   ├── excelDriver.ts   # Excel.run batching
│   │   │   ├── wordDriver.ts    # Word.run batching
│   │   │   ├── pptDriver.ts     # PowerPoint.run batching
│   │   │   └── mockDriver.ts    # Browser standalone mock driver
│   │   └── storage/             # localStorage helper untuk API Keys & chat history
│   └── utils/
│       ├── formulaValidator.ts  # Heuristic formula syntax checker
│       └── contextCompressor.ts # Sampling data tabel/dokumen untuk efisiensi token
```

---

## 3. Multi-Agent Specialization & Tool Calling

### 3.1 Coordinator Agent ("Sam")
* **Focus**: User-friendly interface, intent recognition, task decomposition, and chat memory.
* **Flow**:
  1. Receives user instructions.
  2. Identifies whether the instruction is a consultation question or a document manipulation command.
  3. Identifies task complexity:
     * **Single-step**: Directly invokes the relevant specialist agent.
     * **Multi-step**: Builds an `Action Queue` (a chained task list with statuses: *Pending*, *In Progress*, *Completed*).

### 3.2 Specialist Agents & Tool Definitions

#### A. Excel Specialist Agent
* `get_active_range()`: Reads the active cell address, table dimensions, column header names, and samples the first 5 rows of data.
* `write_cells(range: string, values?: any[][], formulas?: string[][])`: Writes data or Excel formulas (e.g., `=SUM(B2:D2)`, `=XLOOKUP(...)`).
* `format_range(range: string, styles: CellStyle)`: Configures cell styling (zebra banding), bold fonts, borders, and number formatting (IDR/USD/Percentage).
* `create_chart(chartType: string, dataRange: string, title?: string)`: Creates bar, line, pie, or column charts automatically.
* `manage_sheet(action: 'add' | 'rename' | 'activate', name: string)`: Manages worksheets.

#### B. Word Specialist Agent
* `get_selection_or_outline()`: Reads currently selected text or the document structure outline (H1/H2 headings).
* `insert_content(position: 'start' | 'end' | 'cursor', text: string, type: 'paragraph' | 'h1' | 'h2' | 'bullet')`: Inserts text with structured formatting.
* `replace_selection(newText: string)`: Replaces text selection with a revised/translated/polished version.
* `insert_table(rows: number, cols: number, data: string[][])`: Creates and populates formal tables in Word.
* `format_text(styles: TextStyle)`: Modifies text formatting (bold, italic, font size, color).

#### C. PowerPoint Specialist Agent
* `get_slide_context()`: Reads the title, active slide number, and text content within the currently open slide.
* `add_slide(layout: 'TitleOnly' | 'TitleAndContent' | 'TwoContent')`: Adds a new presentation slide.
* `insert_slide_content(title: string, bullets: string[])`: Populates presentation titles and bullet points.
* `set_speaker_notes(notes: string)`: Adds speaker notes for presentation delivery.

---

## 4. BYOK Provider Engine & Adapter

### 4.1 Unified Interface (`ILLMProvider`)
All providers implement the interface:
```typescript
export interface ChatRequest {
  messages: Array<{ role: 'system' | 'user' | 'assistant'; content: string }>;
  tools?: Array<ToolDefinition>;
  temperature?: number;
}

export interface StreamEvent {
  type: 'content_delta' | 'tool_call' | 'done' | 'error';
  delta?: string;
  toolCall?: { id: string; name: string; arguments: Record<string, any> };
  error?: string;
}

export interface ILLMProvider {
  id: string;
  name: string;
  sendMessage(req: ChatRequest, config: ProviderConfig): AsyncIterable<StreamEvent>;
  testConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }>;
}
```

### 4.2 Supported Providers
1. **Google Gemini**: Models `gemini-2.0-flash`, `gemini-1.5-pro` (Endpoint: `generativelanguage.googleapis.com`).
2. **OpenAI**: Models `gpt-4o`, `gpt-4o-mini`, `o3-mini` (Endpoint: `api.openai.com/v1`).
3. **Anthropic Claude**: Models `claude-3-5-sonnet`, `claude-3-5-haiku` (Header: `anthropic-dangerous-direct-browser-access: true`).
4. **GLM Coding Global (Zhipu AI)**: Models `glm-4-plus`, `glm-4-air`, `codegeex-4` (Endpoint: `open.bigmodel.cn/api/paas/v4`).
5. **OpenRouter / DeepSeek**: Models `deepseek/deepseek-chat`, `deepseek/deepseek-r1` (Endpoint: `openrouter.ai/api/v1`).
6. **Local LLM / Custom**: Supports Ollama (`localhost:11434`), LM Studio, and Custom Base URL / Reverse Proxy.

---

## 5. Performance, Execution, & Security

### 5.1 Execution Modes (Copilot vs Autopilot)
* **Copilot Mode (Default)**: The agent displays an action preview card (*Action Card*) in the chat panel containing the change description, diff snippet, and an **"Apply to Document"** button.
* **Autopilot Mode**: The agent immediately executes actions to the document without waiting for manual confirmation (ideal for formatting automation or bulk formula population).
* Users can instantly switch modes via the toggle at the top of the taskpane.

### 5.2 Speed Optimization (Batch `context.sync()`)
To prevent lag or freezing in Office applications:
* Cell/paragraph manipulation operations are grouped into a single memory transaction.
* Calling `await context.sync()` is performed **only once per execution batch**.

### 5.3 Security & Privacy
* **Zero Middleman**: All LLM requests are sent directly from the WebView2 browser to the provider via HTTPS / TLS 1.3.
* **Local Storage Only**: API keys are stored exclusively in the local computer's browser `localStorage`.
* **Office Sandbox**: The add-in is fully isolated within the Microsoft Office sandbox, with no permissions to access the Windows OS disk file system.
* **Native Undo**: Every batch action registers an undo state, allowing users to press `Ctrl+Z` at any time.

---

## 6. Testing & Quality Assurance Plan

1. **Vitest Unit Tests**:
   * Request serialization & response parser testing for each LLM provider.
   * Heuristic formula validator testing (detecting formulas with syntax errors or unbalanced parentheses).
   * Host context router testing (Word, Excel, PowerPoint).
2. **Mock Office Environment**:
   * Simulation of Excel worksheet state (2D cell array), Word document state (paragraph array), and PPT slide state in a standard browser for rapid daily development.
3. **Office Manifest Validation**:
   * XML manifest validation via Microsoft CLI (`npx office-addin-manifest validate manifest.xml`) to ensure complete compatibility before installation/sideloading.
