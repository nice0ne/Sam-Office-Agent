import { StructuredDocOptions, StructuredDocType } from './types';

export type AdministrativeDocType =
  | 'OfficialLetter'
  | 'BAST'
  | 'ExecutiveMoM'
  | 'NotaDinas';

export interface OfficialLetterPayload {
  institutionName: string;
  institutionAddress?: string;
  letterNumber: string;
  date?: string | Date;
  attachments?: string;
  subject: string;
  recipient: {
    title: string;
    address?: string;
  };
  openingText?: string;
  bodyParagraphs: string[];
  closingText?: string;
  signatory: {
    name: string;
    title: string;
    nipOrNik?: string;
  };
  copiesTo?: string[];
}

export interface BastPayload {
  documentNumber: string;
  date?: string | Date;
  firstParty: {
    name: string;
    title: string;
    company: string;
  };
  secondParty: {
    name: string;
    title: string;
    company: string;
  };
  handoverItems: {
    headers: string[];
    rows: string[][];
  };
  clauses: string[];
}

export interface ExecutiveMomPayload {
  meetingTitle: string;
  date?: string | Date;
  timeAndPlace?: string;
  attendees?: string[];
  agenda: string[];
  discussionSummary: string[];
  actionItems: Array<{
    task: string;
    pic: string;
    deadline: string;
    status?: string;
  }>;
}

export interface InternalMemoPayload {
  memoNumber: string;
  from: string;
  to: string;
  date?: string | Date;
  subject: string;
  bodyParagraphs: string[];
  signatory: {
    name: string;
    title: string;
  };
}

const INDONESIAN_MONTHS = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember',
];

const ROMAN_MONTHS = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];

/**
 * Format string tanggal atau Date menjadi format resmi bahasa Indonesia (contoh: "20 Oktober 2026")
 */
export function formatIndonesianDate(inputDate?: string | Date): string {
  if (!inputDate) {
    const now = new Date();
    return `${now.getDate()} ${INDONESIAN_MONTHS[now.getMonth()]} ${now.getFullYear()}`;
  }

  if (inputDate instanceof Date) {
    return `${inputDate.getDate()} ${INDONESIAN_MONTHS[inputDate.getMonth()]} ${inputDate.getFullYear()}`;
  }

  const trimmed = inputDate.trim();
  // Check if already in Indonesian word format: e.g. "20 Oktober 2026"
  const indonesianMonthMatch = INDONESIAN_MONTHS.some(m => new RegExp(`\\b${m}\\b`, 'i').test(trimmed));
  if (indonesianMonthMatch && /\d{4}/.test(trimmed)) {
    return trimmed;
  }

  // Check ISO format YYYY-MM-DD
  const isoMatch = trimmed.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    const year = parseInt(isoMatch[1], 10);
    const month = parseInt(isoMatch[2], 10) - 1;
    const day = parseInt(isoMatch[3], 10);
    if (month >= 0 && month < 12) {
      return `${day} ${INDONESIAN_MONTHS[month]} ${year}`;
    }
  }

  // Check DD-MM-YYYY or DD/MM/YYYY
  const dmyMatch = trimmed.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10) - 1;
    const year = parseInt(dmyMatch[3], 10);
    if (month >= 0 && month < 12) {
      return `${day} ${INDONESIAN_MONTHS[month]} ${year}`;
    }
  }

  // Fallback Date.parse
  const parsed = new Date(trimmed);
  if (!isNaN(parsed.getTime())) {
    return `${parsed.getDate()} ${INDONESIAN_MONTHS[parsed.getMonth()]} ${parsed.getFullYear()}`;
  }

  return trimmed;
}

/**
 * Mengubah angka bulan 1-12 menjadi angka Romawi (I s.d. XII)
 */
export function toRomanMonth(monthNumber: number): string {
  if (monthNumber >= 1 && monthNumber <= 12) {
    return ROMAN_MONTHS[monthNumber - 1];
  }
  return 'I';
}

/**
 * Menyusun nomor surat resmi terstruktur
 */
export function formatOfficialNumber(num: string, dept: string, monthRoman: string, year: number): string {
  return `${num}/${dept}/${monthRoman}/${year}`;
}

/**
 * Builder utama untuk mengonversi naskah dinas ke format StructuredDocOptions
 */
export function buildAdministrativeDoc(type: AdministrativeDocType, payload: any): StructuredDocOptions {
  const dateFormatted = formatIndonesianDate(payload.date);

  switch (type) {
    case 'OfficialLetter': {
      const p = payload as OfficialLetterPayload;
      const sections = [
        {
          heading: 'KOP SURAT & IDENTITAS LEMBAGA',
          content: `${p.institutionName}\n${p.institutionAddress || ''}`.trim(),
        },
        {
          heading: 'NOMOR & PERIHAL SURAT',
          content: [
            `Nomor     : ${p.letterNumber}`,
            `Lampiran  : ${p.attachments || '-'}`,
            `Perihal   : ${p.subject}`,
            `Tanggal   : ${dateFormatted}`,
            '',
            'Kepada Yth.',
            p.recipient?.title || '',
            p.recipient?.address || 'di Tempat',
          ].join('\n'),
        },
        {
          heading: 'ISI SURAT',
          content: [
            p.openingText || 'Dengan hormat,',
            '',
            ...(p.bodyParagraphs || []),
            '',
            p.closingText || 'Demikian surat ini kami sampaikan. Atas perhatian dan kerja sama Bapak/Ibu, kami ucapkan terima kasih.',
          ].join('\n'),
        },
        {
          heading: 'TANDA TANGAN & PENGESAHAN',
          content: [
            dateFormatted,
            p.signatory?.title || '',
            '',
            '',
            p.signatory?.name || '',
            p.signatory?.nipOrNik ? `NIP/NIK: ${p.signatory.nipOrNik}` : '',
          ].filter(Boolean).join('\n'),
        },
      ];

      if (p.copiesTo && p.copiesTo.length > 0) {
        sections.push({
          heading: 'TEMBUSAN',
          bullets: p.copiesTo,
        } as any);
      }

      return {
        templateType: 'OfficialLetter' as StructuredDocType,
        title: p.institutionName ? `SURAT DINAS - ${p.institutionName}` : 'SURAT KELUAR DINAS RESMI',
        author: p.signatory?.name,
        sections,
      };
    }

    case 'BAST': {
      const p = payload as BastPayload;
      return {
        templateType: 'BAST' as StructuredDocType,
        title: `BERITA ACARA SERAH TERIMA (${p.documentNumber})`,
        sections: [
          {
            heading: 'NOMOR & WAKTU PELAKSANAAN',
            content: `Nomor: ${p.documentNumber}\nTanggal: ${dateFormatted}\n\nPada hari ini, ${dateFormatted}, telah disepakati serah terima barang/pekerjaan antara pihak-pihak di bawah ini:`,
          },
          {
            heading: 'IDENTITAS PARA PIHAK',
            content: [
              '1. PIHAK PERTAMA (Yang Menyerahkan):',
              `   Nama    : ${p.firstParty?.name || '-'}`,
              `   Jabatan : ${p.firstParty?.title || '-'}`,
              `   Instansi: ${p.firstParty?.company || '-'}`,
              '',
              '2. PIHAK KEDUA (Yang Menerima):',
              `   Nama    : ${p.secondParty?.name || '-'}`,
              `   Jabatan : ${p.secondParty?.title || '-'}`,
              `   Instansi: ${p.secondParty?.company || '-'}`,
            ].join('\n'),
          },
          {
            heading: 'OBJEK SERAH TERIMA',
            content: 'Pihak Pertama telah menyerahkan kepada Pihak Kedua rincian barang/pekerjaan sebagai berikut:',
            table: p.handoverItems,
          },
          {
            heading: 'KLAUSUL & KETENTUAN PENGALIHAN',
            bullets: p.clauses || [
              'Pihak Pertama telah menyerahkan seluruh barang/pekerjaan dalam kondisi baik.',
              'Pihak Kedua telah memeriksa dan menerima barang/pekerjaan dengan baik.',
            ],
          },
          {
            heading: 'PENANDATANGAN PARA PIHAK',
            table: {
              headers: ['PIHAK PERTAMA (Menyerahkan)', 'PIHAK KEDUA (Menerima)'],
              rows: [
                [
                  `${p.firstParty?.title || 'Pihak Pertama'}\n\n\n\n(${p.firstParty?.name || '-'})`,
                  `${p.secondParty?.title || 'Pihak Kedua'}\n\n\n\n(${p.secondParty?.name || '-'})`,
                ],
              ],
              borderless: true,
            },
          },
        ],
      };
    }

    case 'ExecutiveMoM': {
      const p = payload as ExecutiveMomPayload;
      return {
        templateType: 'ExecutiveMoM' as StructuredDocType,
        title: p.meetingTitle || 'NOTULEN RAPAT EKSEKUTIF (MINUTES OF MEETING)',
        sections: [
          {
            heading: 'INFORMASI PELAKSANAAN RAPAT',
            content: `Tanggal        : ${dateFormatted}\nWaktu & Tempat : ${p.timeAndPlace || '-'}`,
            bullets: p.attendees ? p.attendees.map(a => `Peserta: ${a}`) : [],
          },
          {
            heading: 'AGENDA RAPAT',
            bullets: p.agenda || [],
          },
          {
            heading: 'POKOK PEMBAHASAN & KEPUTUSAN',
            bullets: p.discussionSummary || [],
          },
          {
            heading: 'TINDAK LANJUT & ACTION ITEMS',
            table: {
              headers: ['No', 'Tugas / Tindak Lanjut', 'PIC', 'Deadline', 'Status'],
              rows: (p.actionItems || []).map((item, idx) => [
                String(idx + 1),
                item.task,
                item.pic,
                item.deadline,
                item.status || 'Pending',
              ]),
            },
          },
        ],
      };
    }

    case 'NotaDinas': {
      const p = payload as InternalMemoPayload;
      return {
        templateType: 'NotaDinas' as StructuredDocType,
        title: `NOTA DINAS: ${p.subject}`,
        author: p.signatory?.name,
        sections: [
          {
            heading: 'LEMBAR NOTA DINAS',
            content: [
              `Nomor    : ${p.memoNumber}`,
              `Dari     : ${p.from}`,
              `Kepada   : ${p.to}`,
              `Tanggal  : ${dateFormatted}`,
              `Perihal  : ${p.subject}`,
            ].join('\n'),
          },
          {
            heading: 'POKOK PESAN & ARAHAN',
            content: (p.bodyParagraphs || []).join('\n\n'),
          },
          {
            heading: 'PENANDATANGAN',
            content: [
              dateFormatted,
              p.signatory?.title || '',
              '',
              '',
              p.signatory?.name || '',
            ].join('\n'),
          },
        ],
      };
    }

    default:
      throw new Error(`Unsupported administrative document type: ${type}`);
  }
}
