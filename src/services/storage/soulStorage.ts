import { BrandVoicePreset, DirectiveCategory, LearnedDirective, SoulConfig } from '../../types';
export type { BrandVoicePreset };

export const STORAGE_KEY = 'sam_office_soul_config';

export const SOUL_PRESETS: Record<BrandVoicePreset, string> = {
  formal_executive: `# SOUL & CORPORATE BRAND DIRECTIVES

## Identitas & Nada Bicara (Tone of Voice)
- Gaya Komunikasi: Formal, terstruktur, berbasis data, berwibawa, dan lugas.
- Perspektif: Berorientasi pada keputusan eksekutif dan dampak strategis bisnis.
- Hindari bahasa kasual, singkatan informal, atau gaya bahasa emosional.

## Standar Bahasa & Format
- Bahasa resmi: Bahasa Indonesia baku sesuai PUEBI / EYD.
- Angka & Finansial: Gunakan simbol "Rp" dengan pemisah ribuan titik (contoh: Rp1.500.000).
- Format tanggal: DD MMMM YYYY (contoh: 23 September 2026).
- Judul & Subjudul: Gunakan Title Case yang ringkas dan informatif.

## Batasan & Larangan (Negative Directives)
- Jangan memberikan asumsi atau klaim tanpa dukungan angka atau fakta dokumen.
- Jangan gunakan istilah ambigu atau jargon yang belum terstandardisasi.`,

  modern_professional: `# SOUL & CORPORATE BRAND DIRECTIVES

## Identitas & Nada Bicara (Tone of Voice)
- Gaya Komunikasi: Modern, gesit (agile), kolaboratif, profesional, dan to-the-point.
- Perspektif: Mengutamakan efisiensi tindakan, solusi cepat, dan kejelasan alur kerja.
- Kalimat aktif dan ringkas lebih disukai daripada kalimat pasif yang panjang.

## Standar Bahasa & Format
- Bahasa: Campuran profesional Indonesia-Inggris bisnis yang wajar (tech-savvy).
- Bullet points dan pemformatan visual ringkas diutamakan.
- Ringkas tabel dan data menjadi takeaway yang tajam.

## Batasan & Larangan (Negative Directives)
- Hindari birokrasi kata-kata yang bertele-tele.
- Hindari istilah arkais atau terlalu kaku.`,

  financial_compliance: `# SOUL & CORPORATE BRAND DIRECTIVES

## Identitas & Nada Bicara (Tone of Voice)
- Gaya Komunikasi: Sangat teliti, presisi tinggi, berorientasi audit, dan objektif.
- Perspektif: Kepatuhan regulasi, mitigasi risiko keuangan, dan transparansi kalkulasi.
- Setiap angka harus dapat diverifikasi dan memiliki rujukan sel/tabel yang jelas.

## Standar Bahasa & Format
- Penulisan mata uang dan desimal: Standar akuntansi formal (contoh: Rp 1.500.000,00).
- Wajib menyertakan disclaimer kepatuhan jika memuat proyeksi keuangan.
- Standar penamaan pos anggaran dan kode akun harus konsisten.

## Batasan & Larangan (Negative Directives)
- Dilarang membulatkan angka desimal tanpa instruksi spesifik.
- Dilarang menghilangkan jejak formula kalkulasi dalam lembar kerja.`,

  custom: `# SOUL & CORPORATE BRAND DIRECTIVES

## Panduan Kustom Organisasi
- Tuliskan direktif, panduan gaya bahasa, format data, dan batasan korporat Anda di sini.`,
};

export const DEFAULT_SOUL_CONFIG: SoulConfig = {
  enabled: true,
  corporateName: 'Organisasi Korporat',
  brandVoicePreset: 'formal_executive',
  rawSoulMarkdown: SOUL_PRESETS.formal_executive,
  learnedDirectives: [],
  lastUpdated: Date.now(),
};

let memoryFallbackConfig: SoulConfig | null = null;

function getSafeLocalStorage(): Storage | null {
  try {
    if (typeof localStorage !== 'undefined' && typeof localStorage.getItem === 'function') {
      return localStorage;
    }
  } catch {
    // Non-browser, SSR, or storage disabled environment
  }
  return null;
}

export function getSoulConfig(): SoulConfig {
  const storage = getSafeLocalStorage();
  if (!storage) {
    if (!memoryFallbackConfig) {
      memoryFallbackConfig = structuredClone(DEFAULT_SOUL_CONFIG);
    }
    return structuredClone(memoryFallbackConfig);
  }

  try {
    const raw = storage.getItem(STORAGE_KEY);
    if (!raw) {
      return structuredClone(DEFAULT_SOUL_CONFIG);
    }
    const parsed = JSON.parse(raw);
    return {
      enabled: typeof parsed.enabled === 'boolean' ? parsed.enabled : DEFAULT_SOUL_CONFIG.enabled,
      corporateName: parsed.corporateName || DEFAULT_SOUL_CONFIG.corporateName,
      brandVoicePreset: parsed.brandVoicePreset || DEFAULT_SOUL_CONFIG.brandVoicePreset,
      rawSoulMarkdown: typeof parsed.rawSoulMarkdown === 'string' ? parsed.rawSoulMarkdown : DEFAULT_SOUL_CONFIG.rawSoulMarkdown,
      learnedDirectives: Array.isArray(parsed.learnedDirectives) ? parsed.learnedDirectives : [],
      lastUpdated: parsed.lastUpdated || Date.now(),
    };
  } catch {
    return structuredClone(DEFAULT_SOUL_CONFIG);
  }
}

export function saveSoulConfig(partial: Partial<SoulConfig>): SoulConfig {
  const current = getSoulConfig();
  const updated: SoulConfig = {
    ...current,
    ...partial,
    lastUpdated: Date.now(),
  };

  memoryFallbackConfig = structuredClone(updated);

  const storage = getSafeLocalStorage();
  if (storage) {
    try {
      storage.setItem(STORAGE_KEY, JSON.stringify(updated));
    } catch (e) {
      console.warn('Gagal menyimpan SOUL config ke localStorage:', e);
    }
  }

  return updated;
}

export function addLearnedDirective(
  rule: string,
  category: DirectiveCategory = 'general',
  explanation?: string,
  source?: string
): LearnedDirective {
  const current = getSoulConfig();
  const id = typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
    ? crypto.randomUUID()
    : `dir_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;

  const newDirective: LearnedDirective = {
    id,
    category,
    rule,
    learnedAt: Date.now(),
    source,
    explanation,
  };

  const updatedDirectives = [...current.learnedDirectives, newDirective];
  saveSoulConfig({
    learnedDirectives: updatedDirectives,
  });

  return newDirective;
}

export function removeLearnedDirective(id: string): void {
  const current = getSoulConfig();
  const updatedDirectives = current.learnedDirectives.filter((d) => d.id !== id);
  saveSoulConfig({
    learnedDirectives: updatedDirectives,
  });
}

export function exportSoulToMarkdown(): string {
  const config = getSoulConfig();
  const parts: string[] = [];

  let content = config.rawSoulMarkdown.trim();
  if (!content.includes('SOUL & CORPORATE BRAND DIRECTIVES')) {
    content = `# SOUL & CORPORATE BRAND DIRECTIVES\n\n${content}`;
  }
  parts.push(content);

  if (config.learnedDirectives && config.learnedDirectives.length > 0) {
    parts.push('\n## Direktif yang Dipelajari (Self-Learned Directives)');
    for (const directive of config.learnedDirectives) {
      const cat = directive.category ? `[${directive.category.toUpperCase()}] ` : '';
      const expl = directive.explanation ? ` (Alasan: ${directive.explanation})` : '';
      parts.push(`- ${cat}${directive.rule}${expl}`);
    }
  }

  return parts.join('\n');
}

export function importSoulFromMarkdown(markdownText: string): SoulConfig {
  const current = getSoulConfig();

  // Extract learned section if present
  const learnedRegex = /##\s+(?:Direktif yang Dipelajari|Self-Learned Directives)[^\n]*\n([\s\S]*)$/i;
  const match = markdownText.match(learnedRegex);

  let baseMarkdown = markdownText;
  const parsedDirectives: LearnedDirective[] = [];

  if (match) {
    const fullMatch = match[0];
    const bulletText = match[1];
    baseMarkdown = markdownText.replace(fullMatch, '').trim();

    const lines = bulletText.split('\n');
    for (const rawLine of lines) {
      const line = rawLine.trim();
      if (!line.startsWith('- ') && !line.startsWith('* ')) continue;

      const content = line.substring(2).trim();
      if (!content || content.toLowerCase().includes('belum ada direktif')) continue;

      const categoryMatch = content.match(/^\[(TONE|TERMINOLOGY|FORMATTING|CONSTRAINT|GENERAL)\]\s*(.*)$/i);
      let category: DirectiveCategory = 'general';
      let ruleWithExpl = content;

      if (categoryMatch) {
        category = categoryMatch[1].toLowerCase() as DirectiveCategory;
        ruleWithExpl = categoryMatch[2].trim();
      }

      const explMatch = ruleWithExpl.match(/^(.*?)(?:\s*\((?:Konteks|Alasan):\s*(.*?)\))?$/);
      const rule = explMatch && explMatch[1] ? explMatch[1].trim() : ruleWithExpl;
      const explanation = explMatch && explMatch[2] ? explMatch[2].trim() : undefined;

      parsedDirectives.push({
        id: typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
          ? crypto.randomUUID()
          : `dir_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`,
        category,
        rule,
        learnedAt: Date.now(),
        explanation,
        source: 'imported',
      });
    }
  }

  const finalMarkdown = baseMarkdown || markdownText;

  return saveSoulConfig({
    rawSoulMarkdown: finalMarkdown,
    learnedDirectives: match ? parsedDirectives : current.learnedDirectives,
  });
}

export function clearSoulConfig(): void {
  const storage = getSafeLocalStorage();
  if (storage) {
    try {
      storage.removeItem(STORAGE_KEY);
    } catch {
      // ignore
    }
  }
  memoryFallbackConfig = null;
}
