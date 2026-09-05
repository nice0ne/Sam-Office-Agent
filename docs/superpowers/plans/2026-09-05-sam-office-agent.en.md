# Sam-Office-Agent Implementation Plan (Summary)

English | [Bahasa Indonesia](2026-09-05-sam-office-agent.md)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a universal multi-agent Microsoft Office Web Add-in (Excel, Word, PowerPoint) featuring a React + TypeScript taskpane interface and a Bring Your Own Key (BYOK) architecture with no mandatory third-party backend servers.

**Architecture:** Client-side universal Office Web Add-in powered by Vite, React 18, and Tailwind CSS that automatically detects the active Office host (`Office.context.host`). The "Sam" coordinator agent manages memory and routes instructions to domain specialist agents (Excel, Word, PowerPoint) that produce structured tool calls. Tool calls are safely executed via Office.js batch sync with diff preview (Copilot mode) or direct execution (Autopilot mode). LLM credentials (OpenAI, Gemini, Claude, GLM, OpenRouter, Ollama) are securely persisted in local `localStorage`.

**Tech Stack:** React 18, TypeScript, Vite, Tailwind CSS, Lucide React, `@types/office-js`, Vitest.

**Spec:** [`docs/superpowers/specs/2026-09-05-sam-office-agent-design.md`](file:///D:/VIBE-CODING/Sam-Office-Agent/docs/superpowers/specs/2026-09-05-sam-office-agent-design.md)

---

## Global Constraints
- Must run 100% in the browser / WebView2 without requiring any external backend server.
- All LLM API keys must only be stored in the user's browser `localStorage`.
- Office.js `context.sync()` calls must be batched (1 invocation per transactional operation).
- Must provide a `MockOfficeDriver` so the application can be run and tested directly in standard browsers (`http://localhost:5173`) without dependencies on desktop Office 365.
- All new code must be written in TypeScript with strict type safety (`strict: true`).

---

## Tasks Overview

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

**Summary:**
Sets up the project directory structure, dependencies (React 18, Tailwind CSS, Lucide icons, Vitest, and `@types/office-js`), and Vite build configuration for an Office Web Add-in client. It establishes strict TypeScript checking, path aliases (`@/*`), and a base testing setup. This ensures a clean development workflow and reproducible build output.

---

### Task 2: Core Domain Types & Data Models

**Files:**
- Create: `src/types/index.ts`
- Create: `src/agents/types.ts`
- Test: `tests/types.test.ts`

**Interfaces:**
- Consumes: None
- Produces: `HostType`, `ProviderId`, `ProviderConfig`, `ChatMessage`, `ToolDefinition`, `ToolCall`, `AgentAction`, `ExecutionContext`.

**Summary:**
Defines core TypeScript interfaces and type guards for host detection, BYOK provider configurations, chat messages, and agent execution. It formalizes structured tool call data structures and runtime verification helpers used across all domain agents. This guarantees strict end-to-end type safety across the entire application.

---

### Task 3: Client-Side Storage & BYOK Settings Service

**Files:**
- Create: `src/services/storage/settingsStorage.ts`
- Test: `src/services/storage/__tests__/settingsStorage.test.ts`

**Interfaces:**
- Consumes: `ProviderConfig`, `ProviderId`, `ExecutionMode` from `src/types/index.ts`
- Produces: `getSettings()`, `saveProviderConfig()`, `getActiveProvider()`, `setActiveProvider()`, `getExecutionMode()`, `setExecutionMode()`

**Summary:**
Implements browser `localStorage`-based persistence for API keys and BYOK provider settings without any external backend. It provides secure, local-only storage helpers to retrieve, save, and update LLM credentials and user execution preferences (Copilot vs. Autopilot). Comprehensive unit tests verify fallback defaults, serialization, and secret isolation.

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

**Summary:**
Provides utility functions to heuristically validate Excel formula syntax (checking matched parentheses, valid formula names, and quotes) before execution. It also provides table context compression to summarize large spreadsheet grids and stay within LLM token limits. These utilities prevent syntax runtime errors and optimize prompt size when querying models.

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

**Summary:**
Implements unified client-side adapters for BYOK LLM providers including OpenAI, Google Gemini, Anthropic Claude, Zhipu GLM, OpenRouter, and local Ollama. Each adapter normalizes provider-specific API formats, streaming/chat completions, tool definitions, and tool call extraction directly via browser `fetch`. Connection testing functions allow instant credential verification in the UI.

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

**Summary:**
Builds safe, abstracted drivers for interacting with Microsoft Office applications via Office.js batch synchronization (`context.sync()`). It includes a full in-memory `MockOfficeDriver` enabling fast local browser development and testing at `localhost:5173` without desktop Office 365. The driver layer isolates Office.js lifecycle intricacies and executes transactional document modifications reliably.

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

**Summary:**
Creates specialized domain agents for Excel, Word, and PowerPoint that define application-specific tool schemas and system prompts. Each agent interprets user intent in its respective document context and formulates concrete tool calls such as cell writing, formula calculation, document formatting, and slide generation. Unit tests verify tool parameter schemas, validation logic, and prompt construction.

---

### Task 8: Coordinator Agent ("Sam") & Action Queue Manager

**Files:**
- Create: `src/agents/coordinator/samCoordinator.ts`
- Test: `src/agents/coordinator/__tests__/samCoordinator.test.ts`

**Interfaces:**
- Consumes: `ExcelAgent`, `WordAgent`, `PPTAgent`, `ILLMProvider`, `getSettings()`, `ActionQueueItem`
- Produces: `SamCoordinator`, `ActionQueue`

**Summary:**
Implements the central "Sam" coordinator agent that orchestrates multi-turn conversations and routes requests to the appropriate host specialist agent. It manages conversational context, action queues, diff previews for Copilot mode, and direct tool execution for Autopilot mode. The coordinator ensures smooth execution flow and error recovery across all user interactions.

---

### Task 9: UI Components: Settings Modal & Header Controls

**Files:**
- Create: `src/components/Common/Header.tsx`
- Create: `src/components/Settings/SettingsModal.tsx`
- Modify: `src/index.css`

**Interfaces:**
- Consumes: `getSettings()`, `saveProviderConfig()`, `setActiveProvider()`, `setExecutionMode()`, `testProviderConnection()`
- Produces: Header and Settings UI components.

**Summary:**
Builds the top navigation header and BYOK settings modal in the taskpane UI. The header allows users to toggle between Copilot and Autopilot execution modes and monitor active host status. The settings modal enables configuring API keys, custom base URLs, selecting default models, and testing provider connectivity.

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

**Summary:**
Develops the responsive chat interface comprising message bubbles, quick input prompts, and interactive Action Cards. In Copilot mode, Action Cards present a clear visual diff of proposed changes with an "Apply to Document" button before applying changes to Office. In Autopilot mode, cards display live execution status and immediate completion feedback.

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

**Summary:**
Integrates all components into the main application container and hooks into the `Office.onReady` lifecycle. It generates the universal `manifest.xml` targeting Excel, Word, and PowerPoint, and validates the entire build and end-to-end user experience. Both mock browser mode and production Office Web Add-in environments are fully supported.

---

## Self-Review Checklist

- [ ] **Architecture & BYOK Security:**
  - [ ] Operates 100% client-side without mandatory external backend servers.
  - [ ] All LLM API keys and provider endpoints are stored exclusively in browser `localStorage`.
  - [ ] Supported BYOK providers (OpenAI, Gemini, Anthropic, GLM, OpenRouter, Ollama) connect directly via client `fetch`.
- [ ] **Office.js & Host Isolation:**
  - [ ] `context.sync()` calls are strictly batched (1 sync per transactional operation).
  - [ ] `MockOfficeDriver` enables full functional testing and development at `http://localhost:5173` without desktop Office 365.
  - [ ] Universal manifest (`manifest.xml`) correctly targets Excel (`Workbook`), Word (`Document`), and PowerPoint (`Presentation`).
- [ ] **Agent Workflow & Execution Safety:**
  - [ ] Host auto-detection routes user prompts to the correct specialist agent (Excel, Word, PowerPoint).
  - [ ] Copilot mode presents Action Preview Cards with visual diffs and manual "Apply to Document" approval.
  - [ ] Autopilot mode executes validated tool calls immediately with status indicators.
  - [ ] Excel formula validator flags unclosed parentheses, invalid formula names, and malformed strings.
  - [ ] Table context compressor summarizes large spreadsheets to prevent LLM token overflows.
- [ ] **Code Quality & Verification:**
  - [ ] All code written in strict TypeScript (`strict: true`).
  - [ ] Automated unit test suites pass (`npm run test`).
  - [ ] Production build succeeds without TypeScript errors (`npm run build`).
