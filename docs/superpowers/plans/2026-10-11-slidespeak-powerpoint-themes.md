# Implementation Plan: SlideSpeak-Inspired PowerPoint Themes & Blueprints

**Branch:** `feat/slidespeak-powerpoint-themes`  
**Spec:** `docs/superpowers/specs/2026-10-11-slidespeak-powerpoint-themes-design.md`  

---

### Task Breakdown

- [ ] **Task 1: Slide Blueprint Engine & Types (`slideBlueprintEngine.ts`)**
  - Definisikan `PresentationTheme` dan update `ThemedDeckOptions`, `DocToDeckOptions` di `src/services/office/types.ts`.
  - Buat `src/services/office/slideBlueprintEngine.ts`:
    - Palet warna lengkap (10 tema termasuk McKinsey, Demo Day, Linear, Stripe, TED, Sequoia).
    - Helper `formatActionTitle(topic: string, takeaway?: string): string`.
    - Generator `buildSequoiaPitchDeck(options: SequoiaPitchDeckOptions): DocToDeckSlide[]`.
  - Export dari `src/services/office/index.ts`.
  - Tulis test unit `src/services/office/__tests__/slideBlueprintEngine.test.ts` (TDD).
  - Commit.

- [ ] **Task 2: Driver Theme Palette Upgrades (`pptDriver.ts` & `mockDriver.ts`)**
  - Integrasikan palet 10 tema ke `themePalette` di `PPTDriver` dan `MockOfficeDriver`.
  - Update `docToDeckTransformer.ts` untuk mendukung tema baru.
  - Update test `mockDriver.test.ts` & commit.

- [ ] **Task 3: Agent Tooling (`pptAgent.ts`)**
  - Update enum `theme` di tool `generate_themed_deck` dan `transform_doc_to_deck`.
  - Tambahkan tool `create_sequoia_pitch_deck` pada `pptAgent.ts`.
  - Perbarui prompt sistem `PPTAgent`.
  - Update integration test di `specialistAgents.test.ts` & commit.

- [ ] **Task 4: UI Presets (`QuickActionPresets.tsx`)**
  - Tambahkan tombol preset:
    - 💼 *Deck ala McKinsey*
    - 🚀 *Pitch Deck ala Sequoia (10 Slide)*
    - ⚡ *Tema Linear Dark*
    - 🎤 *Slide TED Keynote + Speaker Notes*
  - Update test `chatAndActions.test.tsx` & commit.

- [ ] **Task 5: Full Regression Suite & Merge ke Main**
  - Jalankan `npx vitest run` (pastikan 388+ tests lulus).
  - Merge `feat/slidespeak-powerpoint-themes` ke `main`.
