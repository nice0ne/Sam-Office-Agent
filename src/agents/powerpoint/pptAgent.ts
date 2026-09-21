import { getOfficeDriver } from '../../services/office';
import { AgentContext, IAgent } from '../types';
import { ToolCall, ToolDefinition } from '../../types';
import { executeMetaOrCustomTool, getLearnedAndMetaTools } from '../metaTools';

export class PPTAgent implements IAgent {
  id = 'ppt-specialist';
  name = 'PowerPoint Specialist';
  hostType = 'PowerPoint' as const;

  getSystemPrompt(context: AgentContext): string {
    const summaryHeader = context.documentSummary
      ? '\n--- SLIDE / MATERI SAAT INI ---\n' + context.documentSummary + '\n----------------------------\n'
      : '';

    return 'Anda adalah Sam Office Agent spesialis Microsoft PowerPoint tingkat Expert.\n' +
      'Anda ahli dalam merancang struktur presentasi yang memikat, menyusun poin slide yang ringkas dan padat, membuat slide baru secara instan, dan membuat speaker notes untuk pemateri.\n' +
      summaryHeader +
      '\nPANDUAN ALUR KERJA:\n' +
      '1. MEMBACA & MERINGKAS SLIDE PRESENTASI:\n' +
      '   - Anda MEMILIKI AKSES PENUH ke teks materi pada slide aktif dan seluruh slide dalam presentasi yang tertera di bagian "--- SLIDE / MATERI SAAT INI ---" di atas.\n' +
      '   - Selain itu, Anda memiliki tool "read_slides" untuk membaca isi slide aktif maupun seluruh slide secara langsung dari PowerPoint.\n' +
      '   - Jika pengguna meminta "baca slide ini buat summary", "ringkas slide aktif", atau menganalisis materi slide:\n' +
      '     * Jika konteks slide di atas sudah memiliki isi materi yang jelas, Anda bisa langsung menyajikan ringkasannya secara komprehensif, padat, dan profesional di chat.\n' +
      '     * Jika teks materi di konteks belum terbaca lengkap atau pengguna meminta membaca ulang/slide lain, PANGGIL tool "read_slides" untuk mengambil isi slide terbaru dari PowerPoint!\n' +
      '     * JANGAN PERNAH berasumsi slide kosong tanpa memanggil "read_slides" jika pengguna menyatakan ada materi di slidenya!\n' +
      '2. MEMBUAT SLIDE PRESENTASI:\n' +
      '   - Jika pengguna meminta membuat slide presentasi (misal: "buatkan slide tentang...", "buat 3 slide presentasi", "jadikan materi presentasi dari chat"):\n' +
      '     * Gunakan tool create_presentation_deck jika pengguna meminta membuat beberapa slide sekaligus.\n' +
      '     * Gunakan tool create_slide jika pengguna meminta membuat 1 slide baru.\n' +
      '     * Pastikan setiap slide memiliki judul yang kuat (title) dan 3-5 butir materi (bullets) yang ringkas, profesional, dan to-the-point. Sertakan notes (catatan pemateri) yang bermanfaat.\n' +
      '3. PANDUAN SPEAKER NOTES (CATATAN PEMATERI):\n' +
      '   - BATASAN TEKNIS OFFICE.JS: Microsoft PowerPoint JavaScript API (Office.js) secara resmi TIDAK menyediakan API untuk menulis langsung ke panel Notes bawaan PowerPoint (properti slide.notesPage tidak ada/tidak diekspos oleh Microsoft).\n' +
      '   - JANGAN PERNAH membuat dynamic office script yang memanggil "slide.notesPage" karena akan error runtime.\n' +
      '   - Jika pengguna meminta speaker notes untuk satu atau seluruh slide: Sajikan naskah speaker notes yang terstruktur, elegan, dan siap pakai langsung di pesan chat dengan format per slide (Intonasi, Naskah Penyampaian, Poin Nilai Juri, Transisi) agar pemateri dapat langsung menyalinnya ke panel Notes PowerPoint atau menjadikannya lembar panduan saat presentasi.\n' +
      '4. Berikan penjelasan ringkas, terstruktur, dan ramah di chat mengenai topik slide yang telah Anda ringkas atau buatkan ke dalam presentasi.';
  }

  getTools(): ToolDefinition[] {
    return [
      {
        name: 'read_slides',
        description: 'Membaca teks, judul, poin-poin materi, dan isi slide presentasi yang sedang dibuka di PowerPoint.',
        parameters: {
          type: 'object',
          properties: {
            slideNumber: {
              type: 'number',
              description: 'Nomor slide spesifik yang ingin dibaca (opsional, contoh: 1). Jika diabaikan, akan membaca seluruh slide atau slide aktif.',
            },
            allSlides: {
              type: 'boolean',
              description: 'Set true untuk membaca seluruh slide yang ada di presentasi.',
            },
          },
        },
      },
      {
        name: 'create_slide',
        description: 'Membuat slide baru lengkap dengan judul, poin-poin materi, dan catatan pemateri (opsional) di PowerPoint.',
        parameters: {
          type: 'object',
          properties: {
            title: { type: 'string', description: 'Judul slide' },
            bullets: { type: 'array', items: { type: 'string' }, description: 'Daftar butir poin materi slide' },
            notes: { type: 'string', description: 'Catatan pemateri / speaker notes (opsional)' },
            layout: { type: 'string', enum: ['TitleOnly', 'TitleAndContent', 'TwoContent'], description: 'Layout slide (opsional)' },
          },
          required: ['title', 'bullets'],
        },
      },
      {
        name: 'create_presentation_deck',
        description: 'Membuat beberapa slide presentasi sekaligus secara berurutan ke dalam PowerPoint.',
        parameters: {
          type: 'object',
          properties: {
            slides: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  title: { type: 'string', description: 'Judul slide' },
                  bullets: { type: 'array', items: { type: 'string' }, description: 'Poin-poin materi' },
                  notes: { type: 'string', description: 'Catatan pemateri (opsional)' },
                  layout: { type: 'string', enum: ['TitleOnly', 'TitleAndContent', 'TwoContent'], description: 'Layout slide (opsional)' },
                },
                required: ['title', 'bullets'],
              },
              description: 'Daftar slide yang akan dibuat',
            },
          },
          required: ['slides'],
        },
      },
      {
        name: 'add_slide',
        description: 'Menambahkan slide baru kosong ke presentasi.',
        parameters: {
          type: 'object',
          properties: {
            layout: { type: 'string', enum: ['TitleOnly', 'TitleAndContent', 'TwoContent'], description: 'Layout slide' },
          },
          required: ['layout'],
        },
      },
      {
        name: 'insert_slide_content',
        description: 'Mengisi judul slide dan poin-poin materi pada slide aktif.',
        parameters: {
          type: 'object',
          properties: {
            title: { type: 'string', description: 'Judul slide' },
            bullets: { type: 'array', items: { type: 'string' }, description: 'Poin-poin presentasi' },
          },
          required: ['title', 'bullets'],
        },
      },
      {
        name: 'set_speaker_notes',
        description: 'Menambahkan catatan pemateri pada slide aktif.',
        parameters: {
          type: 'object',
          properties: {
            notes: { type: 'string', description: 'Catatan pemateri' },
          },
          required: ['notes'],
        },
      },
      {
        name: 'generate_themed_deck',
        description: 'Membuat presentasi lengkap dengan tema desain visual profesional (corporate_blue, emerald_executive, modern_dark, minimalist_clean) dan variasi layout slide.',
        parameters: {
          type: 'object',
          properties: {
            topic: { type: 'string', description: 'Topik utama presentasi' },
            theme: {
              type: 'string',
              enum: ['corporate_blue', 'emerald_executive', 'modern_dark', 'minimalist_clean'],
              description: 'Tema warna visual presentasi',
            },
            slides: {
              type: 'array',
              description: 'Daftar slide berformat tema dan layout terstruktur',
              items: {
                type: 'object',
                properties: {
                  title: { type: 'string', description: 'Judul slide' },
                  layout: {
                    type: 'string',
                    enum: ['title_cover', 'split_comparison', 'bullet_points', 'metric_highlights'],
                    description: 'Layout slide',
                  },
                  content: { type: 'array', items: { type: 'string' }, description: 'Poin-poin isi materi slide' },
                  metrics: {
                    type: 'array',
                    items: {
                      type: 'object',
                      properties: {
                        label: { type: 'string' },
                        value: { type: 'string' },
                      },
                      required: ['label', 'value'],
                    },
                    description: 'Data metrik untuk layout metric_highlights',
                  },
                  notes: { type: 'string', description: 'Catatan pemateri' },
                },
                required: ['title'],
              },
            },
          },
          required: ['topic', 'slides'],
        },
      },
      ...getLearnedAndMetaTools(this.hostType),
    ];
  }

  async executeTool(toolCall: ToolCall, _context: AgentContext): Promise<{ success: boolean; result?: any; error?: string }> {
    const metaCheck = await executeMetaOrCustomTool(this.hostType, toolCall, _context);
    if (metaCheck.handled) {
      return metaCheck.result!;
    }

    const driver = getOfficeDriver('PowerPoint');
    try {
      if (toolCall.name === 'read_slides') {
        const { slideNumber, allSlides } = toolCall.arguments || {};
        if (driver.readSlideData) {
          const res = await driver.readSlideData(slideNumber, allSlides ?? (slideNumber === undefined));
          const slideLines = res.slides
            .map((s) => {
              let line = `[Slide #${s.slideIndex}]: "${s.title}"\n${s.textContent}`;
              if (s.notes) line += `\nSpeaker Notes: ${s.notes}`;
              return line;
            })
            .join('\n\n');

          return {
            success: true,
            result: `📋 Hasil Pembacaan Slide PowerPoint:\nTotal Slide: ${res.totalSlides} | Slide Aktif: #${res.activeSlideIndex}\n\n${slideLines}`,
          };
        }

        const ctx = await driver.getSlideContext();
        return {
          success: true,
          result: `📋 Hasil Pembacaan Slide #${ctx.slideNumber}:\nJudul: "${ctx.title}"\nIsi:\n${ctx.textContent}`,
        };
      }

      if (toolCall.name === 'create_slide') {
        const { title, bullets, notes, layout } = toolCall.arguments || {};
        if (!title) return { success: false, error: 'Judul slide wajib diisi.' };
        if (driver.createSlide) {
          const res = await driver.createSlide(title, Array.isArray(bullets) ? bullets : [], notes, layout);
          return { success: true, result: `Slide #${res.slideNumber} ("${title}") berhasil dibuat.` };
        }
        await driver.addSlide(layout || 'TitleAndContent');
        await driver.insertSlideContent(title, Array.isArray(bullets) ? bullets : []);
        if (notes) await driver.setSpeakerNotes(notes);
        return { success: true, result: `Slide ("${title}") berhasil dibuat.` };
      }

      if (toolCall.name === 'create_presentation_deck') {
        const { slides } = toolCall.arguments || {};
        if (!Array.isArray(slides) || slides.length === 0) {
          return { success: false, error: 'Daftar slides tidak boleh kosong.' };
        }
        if (driver.createPresentationDeck) {
          const res = await driver.createPresentationDeck(slides);
          return { success: true, result: `${res.createdCount} slide presentasi berhasil dibuat ke PowerPoint.` };
        }
        for (const s of slides) {
          await driver.addSlide(s.layout || 'TitleAndContent');
          await driver.insertSlideContent(s.title, Array.isArray(s.bullets) ? s.bullets : []);
          if (s.notes) await driver.setSpeakerNotes(s.notes);
        }
        return { success: true, result: `${slides.length} slide presentasi berhasil dibuat.` };
      }

      if (toolCall.name === 'generate_themed_deck') {
        const { topic, theme, slides } = toolCall.arguments || {};
        if (!topic || !Array.isArray(slides) || slides.length === 0) {
          return { success: false, error: 'topic dan slides wajib diisi.' };
        }
        if (driver.generateThemedDeck) {
          const res = await driver.generateThemedDeck({ topic, theme, slides });
          return { success: true, result: `Deck bertema "${theme || 'corporate_blue'}" untuk topik "${topic}" berhasil dibuat (${res.createdCount} slide).` };
        }
        for (const s of slides) {
          const bullets = s.content || (s.metrics ? s.metrics.map((m: any) => `${m.label}: ${m.value}`) : []);
          await driver.addSlide(s.layout || 'TitleAndContent');
          await driver.insertSlideContent(s.title, bullets);
          if (s.notes) await driver.setSpeakerNotes(s.notes);
        }
        return { success: true, result: `Deck bertema "${theme || 'corporate_blue'}" berhasil dibuat (${slides.length} slide).` };
      }

      if (toolCall.name === 'add_slide') {
        const res = await driver.addSlide(toolCall.arguments.layout);
        return { success: true, result: `Slide #${res.slideNumber} berhasil ditambahkan.` };
      }
      if (toolCall.name === 'insert_slide_content') {
        const { title, bullets } = toolCall.arguments;
        await driver.insertSlideContent(title, bullets);
        return { success: true, result: 'Konten slide berhasil diisi.' };
      }
      if (toolCall.name === 'set_speaker_notes') {
        await driver.setSpeakerNotes(toolCall.arguments.notes);
        return { success: true, result: 'Speaker notes berhasil disimpan.' };
      }
      return { success: false, error: 'Tool tidak ditemukan.' };
    } catch (e: any) {
      return { success: false, error: e.message };
    }
  }
}
