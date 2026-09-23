# Self-Learning Corporate Brand Voice & SOUL Directives Design Spec

## 1. Overview & Business Value

Every enterprise and organization maintains distinct brand guidelines, communication tones, compliance boundaries, and formatting standards (e.g. currency formatting, date standards, forbidden terms, official terminology). Without persistent directives, generic LLMs revert to generic, bland, or inconsistent responses across Office sessions.

### Key Pain Points
1. **Inconsistent Brand Voice**: Different documents produced by the agent reflect inconsistent tones—sometimes overly informal, sometimes overly academic.
2. **Repetitive Prompting**: Users must repeatedly remind the agent about internal rules (e.g. "always write IDR as Rp with periods", "don't use the word 'skema'", "we are in the logistics division").
3. **Static & Fragile Customization**: Most customization systems require manual prompt engineering rather than organically learning from ongoing user corrections and feedback.
4. **Lack of Transparency**: Enterprise users need to see, audit, export, and edit what the AI has learned and what directives it follows.

### The Solution: Self-Learning AI & Corporate Directives Engine (`SOUL.md`)
Inspired by `sam-agent`'s persistent personality and autonomous memory, this feature equips **Sam Office Agent** with a persistent, self-learning corporate memory and brand voice engine:
- **Autonomous Feedback Learning (`learn_corporate_directive`)**: An autonomous ReAct tool that extracts user corrections, formatting requirements, and terminology preferences directly from conversation and persists them.
- **Structured `SOUL.md` Format**: A clean, human-readable and machine-parseable Markdown format defining corporate identity, tone of voice, formatting standards, negative directives (forbidden words), and learned preferences.
- **Dynamic ReAct Prompt Injection**: Automatically injects active `SOUL.md` directives and auto-learned preferences into `SamCoordinator` and specialist agents (`ExcelAgent`, `WordAgent`, `PPTAgent`).
- **Comprehensive Settings UI**: A dedicated "Brand Voice & SOUL.md" panel in `SettingsModal` offering pre-built corporate presets (*Formal Executive*, *Modern Tech*, *Financial & Compliance*), a live Markdown editor, and individual badge controls for learned preferences.
- **Local File Import/Export**: Users can export their `SOUL.md` to local disk or import company-wide `SOUL.md` guidelines.

---

## 2. Architecture & Data Contracts

### 2.1 Types (`src/types/index.ts` / `src/services/storage/soulStorage.ts`)

```typescript
export type DirectiveCategory = 'tone' | 'terminology' | 'formatting' | 'constraint' | 'general';

export interface LearnedDirective {
  id: string;
  category: DirectiveCategory;
  rule: string;
  learnedAt: number;
  source?: string;
  explanation?: string;
}

export type BrandVoicePreset = 'formal_executive' | 'modern_professional' | 'financial_compliance' | 'custom';

export interface SoulConfig {
  enabled: boolean;
  corporateName: string;
  brandVoicePreset: BrandVoicePreset;
  rawSoulMarkdown: string;
  learnedDirectives: LearnedDirective[];
  lastUpdated: number;
}
```

### 2.2 Storage Engine (`src/services/storage/soulStorage.ts`)

- **`getSoulConfig(): SoulConfig`**:
  Retrieves active configuration from `localStorage` (`sam_office_soul_config`). If not present, initializes with sensible corporate default:
  - Preset: `formal_executive`
  - Enabled: `true`
  - Corporate Name: `"Organisasi Korporat"`
  - Default Markdown covering identity, executive tone, Indonesian currency/date formatting, and prohibited ambiguous words.
- **`saveSoulConfig(config: Partial<SoulConfig>): SoulConfig`**:
  Merges and persists configuration updates to `localStorage`.
- **`addLearnedDirective(rule: string, category?: DirectiveCategory, explanation?: string, source?: string): LearnedDirective`**:
  Adds a new learned directive, appends it to `learnedDirectives`, updates the `rawSoulMarkdown` learned section, and saves.
- **`removeLearnedDirective(id: string): void`**:
  Removes a learned directive by ID and updates storage.
- **`exportSoulToMarkdown(): string`**:
  Compiles the full Markdown representation ready for file download.
- **`importSoulFromMarkdown(markdownText: string): SoulConfig`**:
  Parses an imported Markdown file, updates `rawSoulMarkdown`, extracts any learned bullets, and saves.
- **`clearSoulConfig(): void`**:
  Resets `SOUL.md` to factory default.

---

## 3. Component Design & Implementation

### 3.1 Autonomous ReAct Tool: `learn_corporate_directive`

Registered in `ExcelAgent`, `WordAgent`, `PPTAgent`, and coordinator meta-tools:
- **Tool Schema**:
  ```typescript
  {
    name: 'learn_corporate_directive',
    description: 'Mencatat dan mempelajari aturan gaya baru, preferensi format, koreksi istilah, atau direktif penulisan dari pengguna ke dalam memori persisten korporat (SOUL.md) agar selalu dipatuhi di masa mendatang.',
    parameters: {
      type: 'object',
      properties: {
        rule: {
          type: 'string',
          description: 'Aturan atau preferensi spesifik yang harus dipelajari (misal: "Selalu gunakan mata uang Rp dengan pemisah titik", "Gunakan istilah inisiatif bukan skema").',
        },
        category: {
          type: 'string',
          enum: ['tone', 'terminology', 'formatting', 'constraint', 'general'],
          description: 'Kategori direktif yang dipelajari.',
        },
        explanation: {
          type: 'string',
          description: 'Alasan atau konteks mengapa aturan ini dipelajari dari percakapan.',
        },
      },
      required: ['rule'],
    },
  }
  ```
- **Execution**:
  - Calls `addLearnedDirective(rule, category, explanation, 'react_tool')`.
  - Returns a user-friendly response:
    `"🧠 Berhasil mempelajari direktif korporat: '[rule]' (Kategori: [category]). Aturan ini telah dicatat ke dalam SOUL.md dan akan otomatis diterapkan pada seluruh respons dan dokumen mendatang."`

### 3.2 Dynamic Prompt Injection (`samCoordinator.ts`)

In `SamCoordinator.buildSystemPrompt()`:
- Calls `getSoulConfig()`.
- If `config.enabled` is `true`:
  - Formats an authoritative section:
    ```markdown
    === CORPORATE BRAND VOICE & SOUL DIRECTIVES (SOUL.md) ===
    Nama Entitas: ${config.corporateName}
    Preset Karakter: ${config.brandVoicePreset}

    ${config.rawSoulMarkdown}

    Aturan yang Dipelajari Secara Mandiri (Self-Learned Directives):
    ${config.learnedDirectives.map(d => `- [${d.category.toUpperCase()}]: ${d.rule}`).join('\n')}

    PANDUAN KEPATUHAN: Anda WAJIB mematuhi seluruh direktif SOUL di atas dalam setiap penyusunan naskah, tabel angka, dan slide presentasi.
    === END CORPORATE DIRECTIVES ===
    ```
- Agents automatically reflect these rules during synthesis, writing, and auditing.

### 3.3 UI Integration (`SettingsModal.tsx` & `QuickActionPresets.tsx`)

- **SettingsModal Section ("Brand Voice & SOUL.md")**:
  - Enable/Disable Toggle.
  - Preset Selector (`Formal Executive`, `Modern Professional`, `Financial Compliance`, `Custom`).
  - Input field for `corporateName`.
  - Monospace Markdown textarea for `rawSoulMarkdown`.
  - Badges list of `learnedDirectives` with delete button for easy pruning.
  - Action buttons:
    - `📥 Impor SOUL.md` (via HTML file input)
    - `📤 Ekspor SOUL.md` (triggers browser `.md` file download)
    - `🔄 Reset ke Standar`
- **Quick Action Preset (`QuickActionPresets.tsx`)**:
  - `id: 'universal_learn_soul'`
  - Label: `🧠 Pelajari Preferensi Ini`
  - Prompt: `Catat gaya bahasa, aturan format angka, dan preferensi istilah dari dokumen ini ke dalam direktif korporat (SOUL.md) agar selalu diterapkan pada pekerjaan berikutnya.`

---

## 4. Verification & Testing Strategy

1. **Storage Unit Tests (`tests/soulStorage.test.ts`)**:
   - Verify `getSoulConfig()` defaults and `saveSoulConfig()` updates.
   - Verify `addLearnedDirective()` and `removeLearnedDirective()` behavior.
   - Verify Markdown export and import fidelity.
2. **Autonomous Tool & Prompt Integration Tests (`tests/selfLearningSoul.test.ts`)**:
   - Verify `learn_corporate_directive` executes successfully and persists to storage.
   - Verify `buildSystemPrompt()` in `SamCoordinator` injects the `SOUL.md` block when enabled, and omits when disabled.
   - Verify specialist agents (`WordAgent`, `PPTAgent`, `ExcelAgent`) have `learn_corporate_directive` available.
3. **UI Settings & Presets Tests (`src/components/__tests__/headerAndSettings.test.tsx`, `chatAndActions.test.tsx`)**:
   - Verify Brand Voice section and editor controls in `SettingsModal`.
   - Verify `🧠 Pelajari Preferensi Ini` quick action button renders across hosts.
4. **Full Regression Test Suite**:
   - Verify all existing 311+ tests continue to pass with 0 regressions.
