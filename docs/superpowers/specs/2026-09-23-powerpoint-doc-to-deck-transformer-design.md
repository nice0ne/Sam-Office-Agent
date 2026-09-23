# PowerPoint Doc-to-Deck Transformer Design Spec

## 1. Overview & Business Value

In corporate daily office work, professionals frequently encounter long reports, project charters, minutes of meetings (MoM), standard operating procedures (SOPs), and tabular spreadsheets that need to be transformed into executive presentation slide decks.

Creating presentation decks manually from documents suffers from several pain points:
1. **High Cognitive Friction**: Reading long prose documents and condensing them into punchy bullet points without losing critical information takes significant time.
2. **Text-Heavy / Cluttered Slides**: Untrained users often paste entire paragraphs onto slides, creating unreadable walls of text.
3. **Missing Presenter Scripts**: Crafting what to say (spoken narration, transitions between topics, opening hooks) is usually left unprepared until the meeting begins.
4. **Office.js Notes API Limitation**: The official Office.js PowerPoint API does not expose native Speaker Notes panel writing, making it difficult for add-ins to record spoken notes directly into the PPTX slide notes pane.

The **Doc-to-Deck Transformer** solves these challenges by:
1. Transforming raw text or document outlines into a structured **Executive Storyline Arc** (Cover, Context/Problem, Strategy/Solution, Key Metrics, Roadmap/Next Steps).
2. Generating concise, high-impact bullet points (max 3-5 bullets per slide) styled according to professional corporate themes (`corporate_blue`, `emerald_executive`, `modern_dark`, `minimalist_clean`).
3. Generating a full, structured **Presenter Script** (Opening Hook, Key Talking Points, Smooth Verbal Transition) presented directly in the chat taskpane for easy rehearsal and delivery.
4. Providing a 1-click **Quick Action Preset** in the chat UI for instant access in PowerPoint.

---

## 2. Architecture & Data Contracts

### 2.1 Type Definitions (`src/services/office/types.ts`)

```typescript
export interface DocToDeckSlide {
  title: string;
  category: 'cover' | 'context' | 'strategy' | 'metrics' | 'roadmap' | 'general';
  bullets: string[];
  metrics?: Array<{ label: string; value: string; trend?: string }>;
  speakerScript: {
    hook: string;
    keyTalkingPoints: string[];
    transition: string;
  };
}

export interface DocToDeckOptions {
  documentText?: string;
  targetSlideCount?: number;
  theme?: 'corporate_blue' | 'emerald_executive' | 'modern_dark' | 'minimalist_clean';
  presentationTitle?: string;
  targetAudience?: 'executive' | 'technical' | 'team_all_hands';
}

export interface DocToDeckResult {
  deckTitle: string;
  appliedTheme: string;
  totalSlidesCreated: number;
  slides: DocToDeckSlide[];
  summaryMessage: string;
}
```

### 2.2 Driver Interface Extensions (`IDocumentDriver`)

```typescript
export interface IDocumentDriver extends IOfficeDriver {
  // Existing methods...
  transformDocToDeck?(options?: DocToDeckOptions): Promise<DocToDeckResult>;
}
```

---

## 3. Component Design & Implementation

### 3.1 Transformation Engine (`src/services/office/docToDeckTransformer.ts`)

The transformation engine operates deterministically with intelligent heuristics to structure arbitrary document text:
- **Title Extraction**:
  - Identifies the document's main subject from leading headers or generates an executive title based on detected keywords.
- **Narrative Arc Partitioning**:
  1. **Cover Slide (`cover`)**: High-impact title and subtitle representing the strategic briefing.
  2. **Context & Problem Statement (`context`)**: Extracts background constraints, market challenges, or problem statements into 3-4 clean bullet points.
  3. **Core Strategy & Solution (`strategy`)**: Identifies proposed interventions, architecture, or solutions as 3-4 strategic pillars.
  4. **Key Metrics & Impact (`metrics`)**: Extracts quantitative numbers (percentages, currency, SLA numbers, dates) and formats them into metric highlights.
  5. **Roadmap & Next Steps (`roadmap`)**: Extracts deadlines, milestones, and immediate action items.
- **Presenter Script Synthesis**:
  - For each slide, synthesizes an executive Indonesian spoken script:
    - `hook`: Engaging rhetorical question or scene-setting opening.
    - `keyTalkingPoints`: 2-3 bulleted conversational explanations.
    - `transition`: Bridge sentence linking current slide to the next.

### 3.2 Office Drivers (`mockDriver.ts` & `pptDriver.ts`)

- **`MockOfficeDriver`**:
  - Invokes `synthesizeDocToDeck(options.documentText, options)`.
  - Appends created slides to `this.slides` array.
  - Returns `DocToDeckResult`.
- **`PPTDriver` (Office.js Execution)**:
  - Invokes `synthesizeDocToDeck(options.documentText, options)`.
  - Wraps PowerPoint shape creation in `PowerPoint.run(async (context) => { ... })`:
    - Iterates over slides, calling `context.presentation.slides.add()`.
    - Creates title text box (font size 28, bold, theme primary color).
    - Creates content box with bullets (font size 18, line spacing 1.25, theme dark slate color).
    - Batches changes with a single `await context.sync()`.
  - Returns `DocToDeckResult`.

### 3.3 PowerPoint Specialist Agent (`src/agents/powerpoint/pptAgent.ts`)

- Registers `transform_doc_to_deck` in `getTools()`:
  - Declares parameters `documentText`, `targetSlideCount`, `theme`, `presentationTitle`, `targetAudience`.
- Implements execution in `executeTool()`:
  - Calls `driver.transformDocToDeck()`.
  - Formats output into a rich Markdown presentation dashboard:
    - Header with theme and total slides created.
    - Précis of each slide's content.
    - Clear `🎙️ Naskah Presenter / Catatan Pemateri` block per slide.
- Updates `getSystemPrompt()` to guide the agent on when to recommend or execute doc-to-deck transformation.

### 3.4 Quick Action Presets (`src/components/Chat/QuickActionPresets.tsx`)

Adds a new preset to `PRESETS_BY_HOST.PowerPoint`:
```typescript
{
  id: 'ppt_doc_to_deck',
  label: '📊 Dokumen ke Slide (Doc-to-Deck)',
  prompt: 'Ubah teks/dokumen laporan ini menjadi 5 slide presentasi eksekutif terstruktur lengkap dengan naskah pembicara (speaker notes).',
}
```

---

## 4. Verification & Testing Plan

1. **Unit Tests for Transformation Engine (`tests/docToDeckTransformer.test.ts`)**:
   - Verify parsing of unstructured prose into 5-slide Executive Storyline Arc.
   - Verify slide counts and default theme assignment (`corporate_blue`).
   - Verify presence of `speakerScript` with non-empty `hook`, `keyTalkingPoints`, and `transition`.
   - Verify handling of empty or brief input text with graceful fallback defaults.
2. **Driver Tests (`src/services/office/__tests__/mockDriver.test.ts`)**:
   - Verify `MockOfficeDriver.transformDocToDeck` creates slides in `driver.slides` and returns structured `DocToDeckResult`.
3. **Agent Integration Tests (`src/agents/__tests__/specialistAgents.test.ts`)**:
   - Verify `PPTAgent.getTools()` contains `transform_doc_to_deck`.
   - Verify `PPTAgent.executeTool` handles `transform_doc_to_deck` and formats Markdown output with presenter script badges.
4. **Full Test Suite Run**:
   - Run `npx vitest run` to ensure all existing and new tests pass cleanly.
