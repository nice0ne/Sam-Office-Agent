# Design Specification: 1-Click Executive Reporting Pipeline (Excel ➔ Word ➔ PPT)

**Date:** 2026-10-11  
**Status:** Approved  
**Author:** AI Coding Assistant & Team  
**Scope:** `src/services/office/executivePipeline.ts`, `src/services/office/types.ts`, `src/agents/coordinator/`, UI Presets

---

## 1. Executive Summary & Goals

Dalam tata kelola administrasi korporat dan instansi, penyusunan laporan berkala (mingguan/bulanan/triwulanan) biasanya melibatkan tiga tahap manual yang berulang:
1. Rekapitulasi angka kinerja / transaksi di **Microsoft Excel**.
2. Penyusunan naskah narasi laporan eksekutif (Nota Dinas / Executive Brief) di **Microsoft Word**.
3. Pembuatan slide deck presentasi ringkas untuk rapat pimpinan / BoD di **Microsoft PowerPoint**.

**Fase 4: 1-Click Executive Reporting Pipeline** menyatukan seluruh alur ini menjadi satu orkestrasi otomatis terpadu:
- Menganalisis tabel data / metrik dari lembar kerja Excel atau Universal Hub.
- Menghasilkan naskah laporan Word berstandar formal korporat Indonesia lengkap dengan tabel metrik dan temuan utama.
- Mengubah narasi laporan tersebut secara otomatis menjadi slide deck PowerPoint eksekutif lengkap dengan poin kunci dan naskah pembicara (*speaker script*).
- Menyimpan snapshot ke Sam Universal Hub sehingga sinkron di seluruh aplikasi.

---

## 2. Technical Architecture

```
                    ┌───────────────────────────┐
                    │  Data Sumber (Excel Hub)  │
                    │   - Headers, Rows, Metrik │
                    └─────────────┬─────────────┘
                                  │
                                  ▼
               ┌──────────────────────────────────────┐
               │     Executive Reporting Pipeline     │
               │  (executivePipeline.ts)              │
               └──────────┬────────────────┬──────────┘
                          │                │
            ┌─────────────┴─────┐    ┌─────┴────────────────┐
            ▼                   │    │                      ▼
┌───────────────────────┐       │    │       ┌────────────────────────┐
│   Word Document       │       │    │       │   PowerPoint Deck      │
│   - Executive Memo    │       │    │       │   - 3-5 Executive      │
│   - Table & Findings  │       │    │       │     Slides             │
│   - Sign-off Block    │       │    │       │   - Speaker Scripts    │
└───────────────────────┘       │    │       └────────────────────────┘
                                ▼    ▼
                    ┌───────────────────────────┐
                    │  Sam Universal Hub (Sync) │
                    │   - Cross-App Snapshots   │
                    └───────────────────────────┘
```

---

## 3. Data Structures & Interfaces

```typescript
export interface ExecutiveReportPipelineOptions {
  title: string;
  topic?: string;
  tableData?: {
    headers: string[];
    rows: (string | number)[][];
  };
  metrics?: Array<{ label: string; value: string | number; trend?: 'up' | 'down' | 'neutral' }>;
  targetSlideCount?: number;
  theme?: 'corporate_blue' | 'emerald_executive' | 'modern_dark' | 'minimalist_clean';
  executiveSummary?: string;
}

export interface ExecutiveReportPipelineResult {
  title: string;
  wordDocument: {
    title: string;
    sectionsCount: number;
    formattedNarrative: string;
  };
  powerPointDeck: {
    deckTitle: string;
    totalSlides: number;
    slides: DocToDeckSlide[];
  };
  hubSnapshotId: string;
  summaryMessage: string;
}
```

---

## 4. Verification & Testing Strategy
- Unit test: `src/services/office/__tests__/executivePipeline.test.ts`
- Coordinator & specialist integration test.
- UI Quick Action Preset test di `chatAndActions.test.tsx`.
- Full regression suite `npx vitest run` (385+ tests).
