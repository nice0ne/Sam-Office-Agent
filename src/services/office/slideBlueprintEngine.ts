import { DocToDeckSlide, PresentationTheme, SequoiaPitchOptions } from './types';

export interface SlideThemeDefinition {
  primary: string;
  secondary: string;
  bg: string;
  text: string;
  accent: string;
  cardBg?: string;
  fontFamily?: string;
  description: string;
}

/**
 * Curated presentation design themes inspired by SlideSpeak prompt architecture.
 */
export const SLIDESPEAK_THEME_PALETTES: Record<PresentationTheme, SlideThemeDefinition> = {
  corporate_blue: {
    primary: '#1E3A8A',
    secondary: '#3B82F6',
    bg: '#F8FAFC',
    text: '#1E293B',
    accent: '#2563EB',
    cardBg: '#FFFFFF',
    fontFamily: 'Segoe UI, Arial, sans-serif',
    description: 'Corporate Executive Navy: Formal, dependable, and authoritative.',
  },
  emerald_executive: {
    primary: '#065F46',
    secondary: '#10B981',
    bg: '#F0FDF4',
    text: '#0F172A',
    accent: '#059669',
    cardBg: '#FFFFFF',
    fontFamily: 'Calibri, Aptos, sans-serif',
    description: 'Emerald Executive: Premium green palette ideal for ESG, sustainability, and growth.',
  },
  modern_dark: {
    primary: '#38BDF8',
    secondary: '#818CF8',
    bg: '#0F172A',
    text: '#F8FAFC',
    accent: '#38BDF8',
    cardBg: '#1E293B',
    fontFamily: 'Inter, Segoe UI, sans-serif',
    description: 'Modern Dark: Sleek contrast dark theme for technology and digital products.',
  },
  minimalist_clean: {
    primary: '#18181B',
    secondary: '#71717A',
    bg: '#FFFFFF',
    text: '#27272A',
    accent: '#52525B',
    cardBg: '#FAFAFA',
    fontFamily: 'Arial, Helvetica, sans-serif',
    description: 'Minimalist Clean: Monochromatic Swiss design, typography-first, zero noise.',
  },
  mckinsey_consulting: {
    primary: '#051C2C',
    secondary: '#006699',
    bg: '#FFFFFF',
    text: '#334155',
    accent: '#006699',
    cardBg: '#F8FAFC',
    fontFamily: 'Georgia, Times New Roman, serif',
    description: 'McKinsey Strategy: Deep navy, serif action titles, and sharp data-first hierarchy.',
  },
  demo_day: {
    primary: '#111111',
    secondary: '#FF6B00',
    bg: '#FFFFFF',
    text: '#222222',
    accent: '#FF6B00',
    cardBg: '#FFF7ED',
    fontFamily: 'Poppins, Arial, sans-serif',
    description: 'YC Demo Day: 3-minute accelerator clock, massive numbers, single bold orange accent.',
  },
  linear_dark: {
    primary: '#F8FAFC',
    secondary: '#5E6AD2',
    bg: '#08090A',
    text: '#94A3B8',
    accent: '#5E6AD2',
    cardBg: '#12151A',
    fontFamily: 'Inter, monospace, sans-serif',
    description: 'Linear Aesthetic: Dark issue-tracker canvas with purple glow and mono accents.',
  },
  stripe_saas: {
    primary: '#0A2540',
    secondary: '#635BFF',
    bg: '#F6F9FC',
    text: '#425466',
    accent: '#635BFF',
    cardBg: '#FFFFFF',
    fontFamily: 'Segoe UI, Inter, sans-serif',
    description: 'Stripe SaaS: Fintech clean white, slate borders, and vibrant electric indigo.',
  },
  ted_stage: {
    primary: '#FFFFFF',
    secondary: '#E62B1E',
    bg: '#0A0A0A',
    text: '#E5E7EB',
    accent: '#E62B1E',
    cardBg: '#161616',
    fontFamily: 'Inter, Arial, sans-serif',
    description: 'TED Keynote: Dark auditorium stage with punchy coral red and emotional focus.',
  },
  sequoia_pitch: {
    primary: '#0F2D21',
    secondary: '#00875A',
    bg: '#FFFFFF',
    text: '#334155',
    accent: '#00875A',
    cardBg: '#F4FBF7',
    fontFamily: 'Segoe UI, Aptos, sans-serif',
    description: 'Sequoia Venture: Canonical 10-slide framework with trustworthy evergreen palette.',
  },
};

/**
 * Converts a topic and takeaway into a McKinsey-style Action Title.
 */
export function formatActionTitle(topic: string, takeaway?: string): string {
  if (!takeaway || takeaway.trim().length === 0) {
    return topic.trim();
  }

  const cleanTopic = topic.trim();
  const cleanTakeaway = takeaway.trim();

  // Capitalize words for clear executive impact
  const capitalizedTakeaway = cleanTakeaway
    .split(' ')
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');

  return `${cleanTopic}: ${capitalizedTakeaway}`;
}

/**
 * Builds the canonical 10-slide Venture Capital Pitch Deck blueprint modeled after Sequoia Capital.
 */
export function buildSequoiaPitchDeck(options: SequoiaPitchOptions): DocToDeckSlide[] {
  const company = options.companyName || 'Startup Indonesia';
  const oneLiner = options.oneLiner || 'Solusi masa depan untuk efisiensi industri nasional';
  const problem = options.problem || 'Proses operasional masih manual, memakan waktu lama dan rentan error';
  const solution = options.solution || 'Platform terotomasi cerdas dengan integrasi instan dan keamanan data';
  const market = options.targetMarket || 'Potensi pasar terjangkau (TAM) sebesar Rp 25 Triliun di Asia Tenggara';

  return [
    {
      title: `${company} — Tujuan & Visi Perusahaan (Company Purpose)`,
      category: 'cover',
      bullets: [oneLiner, 'Deck Presentasi Pendanaan & Rencana Pertumbuhan Strategis (Sequoia Blueprint)'],
      speakerScript: {
        hook: `Selamat pagi/siang Dewan Investor. Kami memperkenalkan ${company}: ${oneLiner}.`,
        keyTalkingPoints: [`Misi kami adalah mentransformasi industri melalui ${solution}.`],
        transition: 'Mari kita mulai dari tantangan paling krusial yang dihadapi industri saat ini.',
      },
    },
    {
      title: '1. Masalah Pasar (The Problem)',
      category: 'context',
      bullets: [
        problem,
        'Solusi warisan (legacy) lambat beradaptasi dengan kebutuhan modern',
        'Pelanggan kehilangan waktu dan efisiensi biaya secara signifikan setiap bulannya',
      ],
      speakerScript: {
        hook: 'Rasa sakit yang dirasakan pelanggan saat ini nyata dan berulang setiap hari.',
        keyTalkingPoints: [problem, 'Pelanggan saat ini frustrasi dengan alat lama yang mahal dan lambat.'],
        transition: 'Inilah alasan kuat mengapa kami membangun solusi kami.',
      },
    },
    {
      title: '2. Solusi Kami (The Solution)',
      category: 'strategy',
      bullets: [
        solution,
        'Memberikan nilai tambah langsung sejak hari pertama implementasi',
        'Otomatisasi menyeluruh yang memangkas beban kerja hingga 70%',
      ],
      speakerScript: {
        hook: 'Solusi kami bukan perbaikan inkremental, melainkan lompatan efisiensi 10x lipat.',
        keyTalkingPoints: [solution, 'Adopsi cepat dengan antarmuka yang intuitif bagi pengguna bisnis.'],
        transition: 'Pertanyaannya: mengapa waktu yang paling tepat untuk meluncurkannya adalah sekarang?',
      },
    },
    {
      title: '3. Mengapa Sekarang? (Why Now)',
      category: 'context',
      bullets: [
        'Konvergensi teknologi AI dan komputasi awan yang siap produksi',
        'Dorongan regulasi dan transformasi digital instansi yang semakin cepat',
        'Kebutuhan efisiensi biaya menjadi prioritas utama para pengambil keputusan',
      ],
      speakerScript: {
        hook: 'Jendela peluang ini baru terbuka karena perubahan teknologi dan regulasi terkini.',
        keyTalkingPoints: ['Tiga tahun lalu ini belum memungkinkan, tiga tahun ke depan sudah terlambat.'],
        transition: 'Mari kita lihat seberapa besar skala peluang pasar yang kami sasar.',
      },
    },
    {
      title: '4. Skala Pasar (Market Potential & TAM)',
      category: 'metrics',
      bullets: [
        market,
        'SAM (Serviceable Addressable Market): Rp 8.5 Triliun di tier 1 & 2',
        'SOM (Serviceable Obtainable Market): Target awal 5% pangsa pasar dalam 24 bulan',
      ],
      metrics: [
        { label: 'TAM Nasional', value: 'Rp 25 T', trend: 'up' },
        { label: 'Pertumbuhan YoY', value: '28%', trend: 'up' },
      ],
      speakerScript: {
        hook: 'Kami bermain di pasar yang masif dan tumbuh secara eksponensial setiap tahunnya.',
        keyTalkingPoints: [market, 'Fokus awal kami pada segmen dengan siklus penjualan tercepat.'],
        transition: 'Bagaimana peta persaingan dan apa keunggulan kompetitif (moat) kami?',
      },
    },
    {
      title: '5. Lanskap Persaingan & Parit Pertahanan (Competition & Moat)',
      category: 'strategy',
      bullets: [
        options.competition || 'Kompetitor tradisional lambat dan vendor asing kurang adaptif terhadap kebutuhan lokal',
        'Keunggulan kami: Biaya 50% lebih hemat, integrasi lokal mulus, dan kepatuhan data',
        'Efek jaringan data dan retensi pengguna yang tinggi sebagai pertahanan berkelanjutan',
      ],
      speakerScript: {
        hook: 'Kami tidak bersaing dalam perang fitur, melainkan pada kemudahan adopsi dan kepatuhan lokal.',
        keyTalkingPoints: ['Solusi kami dirancang khusus menjawab kelemahan kompetitor lama.'],
        transition: 'Inilah wujud produk dan cara kerja platform kami.',
      },
    },
    {
      title: '6. Arsitektur Produk & Alur Kerja (Product Workflow)',
      category: 'general',
      bullets: [
        'Alur kerja 1-klik yang mengotomatisasi input data ke output laporan eksekutif',
        'Konektor universal ke spreadsheet, dokumen, dan presentasi perkantoran',
        'Standar keamanan enkripsi tingkat perbankan dan pencatatan audit lengkap',
      ],
      speakerScript: {
        hook: 'Desain produk kami memprioritaskan kesederhanaan dan reliabilitas tanpa friksi.',
        keyTalkingPoints: ['Pengguna dapat mulai bekerja dalam hitungan menit tanpa pelatihan rumit.'],
        transition: 'Selanjutnya, mari kita bedah model monetisasi dan pendapatan kami.',
      },
    },
    {
      title: '7. Model Bisnis (Business Model & Unit Economics)',
      category: 'general',
      bullets: [
        options.businessModel || 'Model SaaS B2B Berlangganan (Tiered Subscription) + Add-on Lisensi Enterprise',
        'LTV to CAC Ratio mencapai 4.2x dengan Net Revenue Retention (NRR) di atas 120%',
        'Siklus penjualan rata-rata 30-45 hari dengan konversi pilot ke kontrak 80%',
      ],
      speakerScript: {
        hook: 'Model ekonomi unit kami telah terbukti menghasilkan arus kas yang sehat dan terprediksi.',
        keyTalkingPoints: ['Tingkat retensi tinggi memastikan pertumbuhan pendapatan berulang (ARR).'],
        transition: 'Siapa tim penggerak di balik eksekusi strategi ini?',
      },
    },
    {
      title: '8. Tim Pelaksana (Team & Leadership)',
      category: 'general',
      bullets: [
        options.teamHighlight || 'Kombinasi pendiri dengan pengalaman 10+ tahun di bidang teknologi, korporasi, dan AI',
        'Rekam jejak kepemimpinan teknik di perusahaan teknologi terkemuka dan industri B2B',
        'Didukung oleh dewan penasihat industri yang berpengaruh',
      ],
      speakerScript: {
        hook: 'Kami adalah tim yang tepat dengan keahlian komplementer untuk memenangkan pasar ini.',
        keyTalkingPoints: ['Keahlian mendalam dalam teknik, operasi produk, dan penetrasi pasar B2B.'],
        transition: 'Terakhir, proyeksi keuangan dan rencana pendanaan kami.',
      },
    },
    {
      title: '9. Proyeksi Keuangan & Alokasi Dana (Financials & The Ask)',
      category: 'roadmap',
      bullets: [
        options.financialMilestone || 'Target ARR Rp 15 Miliar dalam 18 bulan ke depan',
        options.fundraiseAsk || 'Alokasi pendanaan: 50% ekspansi produk & R&D, 35% pemasaran/penjualan, 15% operasional',
        'Pencapaian titik impas (breakeven) diproyeksikan pada kuartal ke-6 pasca pendanaan',
      ],
      speakerScript: {
        hook: 'Pendanaan ini akan mengakselerasi kepemimpinan pasar dan memperluas jaringan pelanggan.',
        keyTalkingPoints: ['Penggunaan dana fokus pada pertumbuhan terukur dan profitabilitas berkelanjutan.'],
        transition: 'Terima kasih atas perhatian Dewan Investor, kami siap untuk sesi tanya jawab dan kolaborasi.',
      },
    },
  ];
}
