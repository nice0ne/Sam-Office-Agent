# Design Specification: SlideSpeak-Inspired PowerPoint Themes & Blueprints

**Date:** 2026-10-11  
**Status:** Approved  
**Author:** AI Coding Assistant & Team  
**Scope:** `src/services/office/types.ts`, `src/services/office/pptDriver.ts`, `src/services/office/slideBlueprintEngine.ts`, `src/agents/powerpoint/pptAgent.ts`, `src/components/Chat/QuickActionPresets.tsx`

---

## 1. Overview & Objectives

Berdasarkan acuan industri dari [SlideSpeak Slide Design Prompts](https://slidespeak.co/slide-design-prompts), fitur PowerPoint diperkaya dengan:
1. **Palet Desain & Tema Terkurasi:**
   - `mckinsey_consulting`: White background, deep navy `#051C2C`, serif action titles, slate body `#334155`.
   - `demo_day`: White background, near-black `#111111`, bold geometric sans, single high-contrast orange accent `#FF6B00`.
   - `linear_dark`: Dark canvas `#08090A`, card background `#12151A`, subtle borders `#1F242C`, violet accent `#5E6AD2`.
   - `stripe_saas`: Modern fintech aesthetic, surface `#F6F9FC`, borders `#E3E8EE`, navy `#0A2540`, vibrant indigo `#635BFF`.
   - `ted_stage`: Near-black `#0A0A0A`, dark cards `#161616`, punchy TED coral `#E62B1E`.
   - `sequoia_pitch`: High-trust clean white canvas, deep evergreen `#0F2D21`, vibrant green `#00875A`.
2. **Slide Blueprint Generator:**
   - `generateSequoiaPitchBlueprint(companyName, problem, solution, ...)`: Menghasilkan struktur 10-slide pitch deck standar Sequoia Capital lengkap dengan naskah pembicara (*Speaker Script*).
   - *Action Title Reformatter*: Mengubah judul topik pasif menjadi *Action Titles* (kalimat kesimpulan bernilai bisnis tinggi ala McKinsey).
3. **Agent & UI Integration:**
   - Penambahan opsi tema baru pada `generate_themed_deck` dan `transform_doc_to_deck`.
   - Tombol Quick Action Presets baru di antarmuka PowerPoint untuk akses 1-klik.

---

## 2. Palette Architecture

```typescript
export type PresentationTheme =
  | 'corporate_blue'
  | 'emerald_executive'
  | 'modern_dark'
  | 'minimalist_clean'
  | 'mckinsey_consulting'
  | 'demo_day'
  | 'linear_dark'
  | 'stripe_saas'
  | 'ted_stage'
  | 'sequoia_pitch';
```

---

## 3. Testing Strategy
- Unit test palet dan blueprint generator di `src/services/office/__tests__/slideBlueprintEngine.test.ts`.
- Pengujian `PPTDriver` dan `mockDriver.test.ts`.
- Pengujian tool di `specialistAgents.test.ts`.
- Pengujian UI presets di `chatAndActions.test.tsx`.
- Full regression suite `npx vitest run`.
