import { DocToDeckOptions, DocToDeckResult, DocToDeckSlide } from './types';

interface SectionStore {
  context: string[];
  strategy: string[];
  metrics: string[];
  roadmap: string[];
  general: string[];
}

const CONTEXT_PREFIX_REGEX = /^(latar\s*belakang|background|tantangan|masalah|problem|isu|konteks|urgensi|challenges?|context)\s*[:\-–]/i;
const STRATEGY_PREFIX_REGEX = /^(solusi|strategi|inisiatif|pendekatan|pillars?|solution|strategy|initiatives?|usulan)\s*[:\-–]/i;
const METRICS_PREFIX_REGEX = /^(capaian\s*(&|dan)?\s*metrik|metrik|target|kpi|hasil|dampak|metrics?|impact|results?|outcomes?)\s*[:\-–]/i;
const ROADMAP_PREFIX_REGEX = /^(rencana\s*aksi|roadmap|langkah\s*berikutnya|tindak\s*lanjut|action\s*plan|next\s*steps?|timeline|linimasa|jadwal)\s*[:\-–]/i;

function cleanBulletText(text: string): string {
  return text.trim().replace(/^[-*•\d+.)\s]+/, '').replace(/\.$/, '').trim();
}

function splitIntoBullets(rawText: string): string[] {
  const lines = rawText
    .split(/\r?\n/)
    .map(cleanBulletText)
    .filter((s) => s.length > 0);

  const bullets: string[] = [];

  for (const line of lines) {
    if (line.includes(';') || (line.includes(',') && line.length > 40)) {
      const parts = line
        .split(/(?:;|\s*,\s*(?=(?:peningkatan|akurasi|kepuasan|reduksi|penurunan|efisiensi|fase|tahap|sosialisasi|evaluasi|[a-zA-Z0-9])))/i)
        .map(cleanBulletText)
        .filter((p) => p.length > 0);

      if (parts.length > 1 && parts.every((p) => p.length > 10)) {
        for (const p of parts) {
          bullets.push(p.charAt(0).toUpperCase() + p.slice(1));
        }
        continue;
      }
    }

    bullets.push(line.charAt(0).toUpperCase() + line.slice(1));
  }

  return bullets;
}

function extractMetrics(bullets: string[]): Array<{ label: string; value: string; trend?: string }> {
  const metrics: Array<{ label: string; value: string; trend?: string }> = [];

  for (const b of bullets) {
    const pctMatch = b.match(/(\d+(?:\.\d+)?%)/);
    const timeMatch = b.match(/(\d+(?:\.\d+)?\s*(?:hari|jam|bulan|x|kali))/i);
    const currMatch = b.match(/((?:Rp\.?|\$)\s*[\d.,]+)/i);

    const valMatch = pctMatch || timeMatch || currMatch;
    if (valMatch) {
      const value = valMatch[1].trim();

      let label = 'Metrik Kunci';
      if (/efisiensi/i.test(b)) label = 'Efisiensi Waktu';
      else if (/akurasi/i.test(b)) label = 'Akurasi Data';
      else if (/kepuasan/i.test(b)) label = 'Kepuasan Pengguna';
      else if (/sla/i.test(b)) label = 'SLA';
      else if (/uptime/i.test(b)) label = 'Uptime';
      else {
        const cleaned = b
          .replace(valMatch[0], '')
          .replace(/^(peningkatan|penurunan|reduksi|mencapai|hingga|target|sebesar)\s*/i, '')
          .replace(/[:\-–]/g, '')
          .trim();
        if (cleaned.length >= 3) {
          label = cleaned.charAt(0).toUpperCase() + cleaned.slice(1);
        }
      }

      const isPositive = /(peningkatan|efisiensi|kepuasan|akurasi|uptime|tumbuh|naik)/i.test(b);
      const isReduction = /(reduksi|penurunan|turun|berkurang|hemat)/i.test(b);

      metrics.push({
        label,
        value,
        trend: isPositive ? 'up' : isReduction ? 'down' : 'neutral',
      });
    }
  }

  return metrics;
}

export function synthesizeDocToDeck(rawText?: string, options?: DocToDeckOptions): DocToDeckResult {
  const text = (rawText || options?.documentText || '').trim();
  const theme = options?.theme || 'corporate_blue';

  // Normalize period-delimited inline sections into newlines
  const normalizedText = text.replace(
    /([.!?])\s+(?=(?:latar\s*belakang|background|tantangan|masalah|problem|solusi|strategi|inisiatif|capaian|metrik|target|kpi|rencana\s*aksi|roadmap|langkah)\s*[:\-–])/gi,
    '$1\n'
  );

  const rawLines = normalizedText
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  let extractedTitle: string | undefined;
  const sections: SectionStore = {
    context: [],
    strategy: [],
    metrics: [],
    roadmap: [],
    general: [],
  };

  let currentCategory: keyof SectionStore | null = null;

  for (let i = 0; i < rawLines.length; i++) {
    const line = rawLines[i];

    if (CONTEXT_PREFIX_REGEX.test(line)) {
      currentCategory = 'context';
      const content = line.replace(CONTEXT_PREFIX_REGEX, '').trim();
      if (content) sections.context.push(...splitIntoBullets(content));
    } else if (STRATEGY_PREFIX_REGEX.test(line)) {
      currentCategory = 'strategy';
      const content = line.replace(STRATEGY_PREFIX_REGEX, '').trim();
      if (content) sections.strategy.push(...splitIntoBullets(content));
    } else if (METRICS_PREFIX_REGEX.test(line)) {
      currentCategory = 'metrics';
      const content = line.replace(METRICS_PREFIX_REGEX, '').trim();
      if (content) sections.metrics.push(...splitIntoBullets(content));
    } else if (ROADMAP_PREFIX_REGEX.test(line)) {
      currentCategory = 'roadmap';
      const content = line.replace(ROADMAP_PREFIX_REGEX, '').trim();
      if (content) sections.roadmap.push(...splitIntoBullets(content));
    } else {
      if (i === 0 && !extractedTitle && !options?.presentationTitle) {
        // The first line without explicit category prefix is the title
        extractedTitle = line.replace(/^#+\s*/, '').replace(/\*+/g, '').replace(/[.:]+$/, '').trim();
      } else if (currentCategory) {
        sections[currentCategory].push(...splitIntoBullets(line));
      } else {
        // Classify by keywords if no current category
        if (/latar\s*belakang|masalah|problem|tantangan|kendala|kesalahan|manual|risiko|isu|rugi/i.test(line)) {
          sections.context.push(...splitIntoBullets(line));
        } else if (/solusi|strategi|solutif|penerapan|mengembangkan|fitur|arsitektur|inisiatif/i.test(line)) {
          sections.strategy.push(...splitIntoBullets(line));
        } else if (/metrik|capaian|efisiensi|persen|%|\$|rp\b|target|sla|akurasi|kpi|roi/i.test(line)) {
          sections.metrics.push(...splitIntoBullets(line));
        } else if (/rencana|roadmap|jadwal|tahapan|fase|oktober|november|desember|kuartal|timeline/i.test(line)) {
          sections.roadmap.push(...splitIntoBullets(line));
        } else {
          sections.general.push(line);
        }
      }
    }
  }

  const deckTitle = options?.presentationTitle || extractedTitle || 'Briefing Eksekutif';

  // 1. Cover Slide
  const coverSlide: DocToDeckSlide = {
    title: deckTitle,
    category: 'cover',
    bullets: [
      'Paparan Strategis & Kerangka Solusi',
      'Disiapkan untuk Tim Eksekutif & Stakeholder Terkait',
    ],
    speakerScript: {
      hook: `Selamat pagi Bapak/Ibu sekalian, terima kasih telah meluangkan waktu untuk hadir. Hari ini saya memaparkan inisiatif strategis "${deckTitle}".`,
      keyTalkingPoints: [
        'Paparan ini menyajikan sintesis komprehensif mulai dari pemetaan tantangan, arsitektur solusi, hingga dampak kuantitatif.',
        'Fokus utama sesi ini adalah menyelaraskan pemahaman dan menyepakati langkah prioritas pelaksanaan.',
      ],
      transition: 'Mari kita awali dengan meninjau latar belakang dan konteks tantangan operasional yang dihadapi.',
    },
  };

  // 2. Context Slide
  const contextBullets = sections.context.length > 0
    ? sections.context.slice(0, 4)
    : [
        'Tantangan efisiensi dan alur kerja manual yang memakan waktu',
        'Tingginya potensi risiko kesalahan manusia dan keterlambatan rekonsiliasi',
        'Kebutuhan mendesak untuk modernisasi dan standardisasi alur kerja',
      ];

  const contextSlide: DocToDeckSlide = {
    title: 'Latar Belakang & Analisis Masalah',
    category: 'context',
    bullets: contextBullets,
    speakerScript: {
      hook: 'Sebelum kita membahas solusi, mari kita telaah realitas operasional dan kendala utama yang sedang kita hadapi.',
      keyTalkingPoints: contextBullets.slice(0, 3).map((b) => `Urgensi lapangan: ${b}.`),
      transition: 'Menanggapi tantangan tersebut, kami telah menyusun pilar inisiatif strategis yang terstruktur.',
    },
  };

  // 3. Strategy Slide
  const strategyBullets = sections.strategy.length > 0
    ? sections.strategy.slice(0, 4)
    : [
        'Penerapan otomatisasi sistem cerdas untuk memangkas proses manual',
        'Standardisasi alur kerja dokumen terintegrasi lintas unit',
        'Penguatan tata kelola operasional dan akuntabilitas kerja',
      ];

  const strategySlide: DocToDeckSlide = {
    title: 'Pilar Strategi & Solusi Utama',
    category: 'strategy',
    bullets: strategyBullets,
    speakerScript: {
      hook: 'Solusi yang efektif harus menyasar langsung akar masalah dengan pendekatan yang terukur dan mudah diadopsi.',
      keyTalkingPoints: strategyBullets.slice(0, 3).map((b) => `Pilar solusi: ${b}.`),
      transition: 'Selanjutnya, mari kita cermati bagaimana efektivitas solusi ini diukur melalui indikator kuantitatif.',
    },
  };

  // 4. Metrics Slide
  const metricsBullets = sections.metrics.length > 0
    ? sections.metrics.slice(0, 4)
    : [
        'Peningkatan efisiensi waktu pemrosesan hingga 75%',
        'Tingkat akurasi data operasional mencapai 99.8%',
        'Peningkatan kepatuhan SLA dan kepuasan pemangku kepentingan hingga 95%',
      ];

  const detectedMetrics = extractMetrics(metricsBullets);
  const slideMetrics = detectedMetrics.length > 0
    ? detectedMetrics
    : [
        { label: 'Efisiensi Waktu', value: '+75%', trend: 'up' },
        { label: 'Akurasi Data', value: '99.8%', trend: 'up' },
        { label: 'Kepuasan Pengguna', value: '95%', trend: 'up' },
      ];

  const metricsSlide: DocToDeckSlide = {
    title: 'Target Capaian & Metrik Dampak',
    category: 'metrics',
    bullets: metricsBullets,
    metrics: slideMetrics,
    speakerScript: {
      hook: 'Ukuran keberhasilan inisiatif ini harus dapat diverifikasi secara objektif melalui data dan indikator terukur.',
      keyTalkingPoints: metricsBullets.slice(0, 3).map((b) => `Target capaian: ${b}.`),
      transition: 'Untuk merealisasikan seluruh target ini, kami telah menyusun peta jalan dan rencana aksi yang komprehensif.',
    },
  };

  // 5. Roadmap Slide
  const roadmapBullets = sections.roadmap.length > 0
    ? sections.roadmap.slice(0, 4)
    : [
        'Fase 1: Sosialisasi dan pelatihan menyeluruh bagi tim pengguna',
        'Fase 2: Implementasi bertahap dengan pendampingan intensif',
        'Fase 3: Evaluasi dampak berkala dan penyempurnaan berkelanjutan',
      ];

  const roadmapSlide: DocToDeckSlide = {
    title: 'Rencana Aksi & Langkah Berikutnya',
    category: 'roadmap',
    bullets: roadmapBullets,
    speakerScript: {
      hook: 'Strategi yang hebat hanya akan membuahkan hasil nyata bila diiringi disiplin eksekusi dan linimasa yang jelas.',
      keyTalkingPoints: roadmapBullets.slice(0, 3).map((b) => `Tahap eksekusi: ${b}.`),
      transition: 'Demikian rencana strategis ini kami sampaikan, kami membuka ruang diskusi untuk masukan dan arahan Bapak/Ibu.',
    },
  };

  const slides: DocToDeckSlide[] = [coverSlide, contextSlide, strategySlide, metricsSlide, roadmapSlide];

  return {
    deckTitle,
    appliedTheme: theme,
    totalSlidesCreated: slides.length,
    slides,
    summaryMessage: `Berhasil mensintesis ${slides.length} slide Executive Storyline Arc untuk "${deckTitle}".`,
  };
}
