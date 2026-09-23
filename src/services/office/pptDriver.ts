import { DiagramResult, DocToDeckOptions, DocToDeckResult, IDocumentDriver, InsertFlowchartOptions, ThemedDeckOptions } from './types';
import { synthesizeDocToDeck } from './docToDeckTransformer';
import { synthesizeFlowchartFromText, renderFlowchartToPngBase64 } from '../../utils/diagramRenderer';

declare const PowerPoint: any;
declare const Office: any;

export class PPTDriver implements Partial<IDocumentDriver> {
  hostType = 'PowerPoint';

  async readSlideData(
    slideNumber?: number,
    allSlides = true
  ): Promise<{
    activeSlideIndex: number;
    totalSlides: number;
    slides: Array<{ slideIndex: number; title: string; textContent: string; notes?: string }>;
  }> {
    let totalSlides = 1;
    let activeSlideIndex = 1;
    const slidesResult: Array<{ slideIndex: number; title: string; textContent: string; notes?: string }> = [];

    // 1. Coba deteksi selected text via Office API jika ada
    let selectedDocumentText = '';
    try {
      if (typeof Office !== 'undefined' && Office.context?.document?.getSelectedDataAsync) {
        selectedDocumentText = await new Promise<string>((resolve) => {
          try {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (asyncResult: any) => {
              if (asyncResult?.status === Office.AsyncResultStatus.Succeeded && asyncResult.value) {
                resolve(String(asyncResult.value).trim());
              } else {
                resolve('');
              }
            });
          } catch {
            resolve('');
          }
        });
      }
    } catch (_docErr) {}

    // 2. Baca data dari PowerPoint.run
    try {
      if (typeof PowerPoint !== 'undefined' && typeof PowerPoint.run === 'function') {
        // Langkah A: Ambil jumlah slide dan slide aktif dengan aman
        let slideIds: string[] = [];
        await PowerPoint.run(async (context: any) => {
          const presentation = context.presentation;
          const slideCollection = presentation.slides;
          slideCollection.load(['items']);
          await context.sync();

          if (slideCollection.items && slideCollection.items.length > 0) {
            totalSlides = slideCollection.items.length;
            slideIds = slideCollection.items.map((s: any) => s.id);
          }

          // Coba baca slide terpilih (PowerPointApi 1.4+)
          try {
            if (typeof presentation.getSelectedSlides === 'function') {
              const selectedSlides = presentation.getSelectedSlides();
              selectedSlides.load(['items']);
              await context.sync();
              if (selectedSlides.items && selectedSlides.items.length > 0) {
                const activeId = selectedSlides.items[0].id;
                const idx = slideIds.findIndex((id) => id === activeId);
                if (idx >= 0) activeSlideIndex = idx + 1;
              }
            }
          } catch (_selErr) {
            // Abaikan jika getSelectedSlides tidak didukung di versi ini
          }
        });

        // Tentukan indeks slide yang akan dibaca
        const indicesToInspect: number[] = [];
        if (slideNumber && !allSlides) {
          const targetIdx = slideNumber - 1;
          if (targetIdx >= 0 && targetIdx < totalSlides) {
            indicesToInspect.push(targetIdx);
          } else {
            indicesToInspect.push(0);
          }
        } else {
          const maxSlides = Math.min(totalSlides, 30);
          for (let i = 0; i < maxSlides; i++) indicesToInspect.push(i);
        }

        // Langkah B: Baca masing-masing slide secara terisolasi agar shape non-teks tidak menggagalkan slide lain
        for (const idx of indicesToInspect) {
          const slideNum = idx + 1;
          let slideTitle = `Slide ${slideNum}`;
          let textParts: string[] = [];

          try {
            await PowerPoint.run(async (slideCtx: any) => {
              const slide = slideCtx.presentation.slides.getItemAt(idx);
              const shapes = slide.shapes;
              shapes.load('items/id, items/name');
              await slideCtx.sync();

              const shapeItems = shapes.items || [];
              if (shapeItems.length === 0) return;

              // Percobaan 1: Coba batch load textRange
              let batchSuccess = false;
              try {
                for (const shape of shapeItems) {
                  if (shape.textFrame?.textRange) {
                    shape.textFrame.textRange.load(['text']);
                  }
                }
                await slideCtx.sync();
                for (const shape of shapeItems) {
                  const t = shape.textFrame?.textRange?.text;
                  if (t && typeof t === 'string' && t.trim()) {
                    textParts.push(t.trim());
                  }
                }
                batchSuccess = true;
              } catch (_batchErr) {
                batchSuccess = false;
              }

              // Percobaan 2: Jika batch gagal (karena ada gambar, connector, garis, tabel, dll), baca shape-by-shape
              if (!batchSuccess) {
                textParts = [];
                for (const sh of shapeItems) {
                  try {
                    await PowerPoint.run(async (shCtx: any) => {
                      const shapeObj = shCtx.presentation.slides.getItemAt(idx).shapes.getItem(sh.id);
                      const tr = shapeObj.textFrame.textRange;
                      tr.load(['text']);
                      await shCtx.sync();
                      if (tr.text && typeof tr.text === 'string' && tr.text.trim()) {
                        textParts.push(tr.text.trim());
                      }
                    });
                  } catch (_shErr) {
                    // Abaikan shape yang tidak memiliki textFrame (gambar, garis, shape dekoratif)
                  }
                }
              }
            });
          } catch (slideErr) {
            console.warn(`Gagal membaca slide #${slideNum}:`, slideErr);
          }

          // Susun struktur judul & konten
          let slideBody = '';
          if (textParts.length > 0) {
            slideTitle = textParts[0];
            const remaining = textParts.slice(1);
            if (remaining.length > 0) {
              slideBody = remaining.map((b) => '• ' + b).join('\n');
            }
          }

          // Tambahkan selectedDocumentText jika di slide aktif teks belum terdeteksi
          if (slideNum === activeSlideIndex && selectedDocumentText && !textParts.some((p) => p.includes(selectedDocumentText))) {
            if (slideBody) {
              slideBody = `[Teks Terseleksi]: ${selectedDocumentText}\n\n${slideBody}`;
            } else {
              slideBody = `[Teks Terseleksi]: ${selectedDocumentText}`;
            }
          }

          slidesResult.push({
            slideIndex: slideNum,
            title: slideTitle,
            textContent: slideBody || (textParts.length > 0 ? textParts.join('\n') : '(Tidak ada teks terdeteksi di slide ini)'),
          });
        }

        return {
          activeSlideIndex,
          totalSlides,
          slides: slidesResult,
        };
      }
    } catch (e) {
      console.warn('Gagal membaca slide via PowerPoint API:', e);
    }

    // Fallback jika bukan lingkungan Office.js aktif atau API gagal total
    const fallbackText = selectedDocumentText
      ? `[Teks Terseleksi]: ${selectedDocumentText}`
      : 'Slide aktif siap diisi materi.';

    return {
      activeSlideIndex: 1,
      totalSlides: 1,
      slides: [
        {
          slideIndex: 1,
          title: 'Active Slide',
          textContent: fallbackText,
        },
      ],
    };
  }

  async getSlideContext(): Promise<{
    slideNumber: number;
    title: string;
    textContent: string;
    slides?: Array<{ slideIndex: number; title: string; textContent: string; notes?: string }>;
  }> {
    const data = await this.readSlideData(undefined, true);
    const activeSlide =
      data.slides.find((s) => s.slideIndex === data.activeSlideIndex) ||
      data.slides[0] || {
        slideIndex: 1,
        title: 'Active Slide',
        textContent: 'Slide aktif siap diisi materi.',
      };

    return {
      slideNumber: activeSlide.slideIndex,
      title: activeSlide.title,
      textContent: activeSlide.textContent,
      slides: data.slides,
    };
  }

  async addSlide(_layout?: string): Promise<{ slideNumber: number }> {
    try {
      if (typeof PowerPoint !== 'undefined' && typeof PowerPoint.run === 'function') {
        return await PowerPoint.run(async (context: any) => {
          const slides = context.presentation.slides;
          slides.add();
          const count = slides.getCount();
          await context.sync();
          return { slideNumber: count.value };
        });
      }
    } catch (e: any) {
      console.warn('Gagal menambahkan slide kosong:', e);
      throw new Error(`Gagal menambahkan slide PowerPoint: ${e.message || String(e)}`);
    }
    return { slideNumber: 1 };
  }

  async insertSlideContent(title: string, bullets: string[]): Promise<void> {
    if (typeof Office !== 'undefined' && Office.context?.document) {
      const content = `${title}\n\n` + bullets.map((b) => `• ${b}`).join('\n');
      Office.context.document.setSelectedDataAsync(content, { coercionType: Office.CoercionType.Text });
    }
  }

  async setSpeakerNotes(_notes: string): Promise<void> {
    // Notes handler
  }

  async createSlide(title: string, bullets: string[], _notes?: string, _layout?: string): Promise<{ slideNumber: number }> {
    try {
      if (typeof PowerPoint !== 'undefined' && typeof PowerPoint.run === 'function') {
        return await PowerPoint.run(async (context: any) => {
          const slideCollection = context.presentation.slides;
          slideCollection.add();
          const count = slideCollection.getCount();
          await context.sync();

          const slide = slideCollection.getItemAt(count.value - 1);
          slide.load(['id']);
          await context.sync();

          try {
            if (slide.shapes) {
              if (typeof slide.shapes.addTextBox === 'function') {
                const titleBox = slide.shapes.addTextBox(title, {
                  left: 60,
                  top: 50,
                  width: 600,
                  height: 70,
                });
                if (titleBox.textFrame?.textRange) {
                  titleBox.textFrame.textRange.text = title;
                  titleBox.textFrame.textRange.font.size = 28;
                  titleBox.textFrame.textRange.font.bold = true;
                  titleBox.textFrame.textRange.font.color = '#1E293B';
                }

                if (bullets && bullets.length > 0) {
                  const bulletText = bullets.map(b => `• ${b}`).join('\n\n');
                  const bodyBox = slide.shapes.addTextBox(bulletText, {
                    left: 60,
                    top: 140,
                    width: 600,
                    height: 270,
                  });
                  if (bodyBox.textFrame?.textRange) {
                    bodyBox.textFrame.textRange.text = bulletText;
                    bodyBox.textFrame.textRange.font.size = 18;
                    bodyBox.textFrame.textRange.font.color = '#334155';
                  }
                }
              } else if (typeof slide.shapes.addGeometricShape === 'function') {
                const rectType = PowerPoint.GeometricShapeType?.rectangle || 'Rectangle';
                const rect = slide.shapes.addGeometricShape(rectType, {
                  left: 60,
                  top: 50,
                  width: 600,
                  height: 70,
                });
                if (rect.fill?.setSolidColor) {
                  rect.fill.setSolidColor('#F8FAFC');
                }
                if (rect.textFrame?.textRange) {
                  rect.textFrame.textRange.text = title;
                  rect.textFrame.textRange.font.size = 26;
                  rect.textFrame.textRange.font.bold = true;
                }

                if (bullets && bullets.length > 0) {
                  const bodyRect = slide.shapes.addGeometricShape(rectType, {
                    left: 60,
                    top: 140,
                    width: 600,
                    height: 270,
                  });
                  if (bodyRect.fill?.setSolidColor) {
                    bodyRect.fill.setSolidColor('#FFFFFF');
                  }
                  if (bodyRect.textFrame?.textRange) {
                    bodyRect.textFrame.textRange.text = bullets.map(b => `• ${b}`).join('\n\n');
                    bodyRect.textFrame.textRange.font.size = 18;
                  }
                }
              }
            }
          } catch (shapeErr) {
            console.warn('Gagal menambahkan shape ke slide:', shapeErr);
          }

          await context.sync();
          return { slideNumber: count.value };
        });
      }
    } catch (err: any) {
      console.error('PowerPoint.run createSlide error:', err);
      throw new Error(`Gagal membuat slide PowerPoint: ${err.message || String(err)}`);
    }

    return { slideNumber: 1 };
  }

  async createPresentationDeck(slides: Array<{ title: string; bullets: string[]; notes?: string; layout?: string }>): Promise<{ createdCount: number }> {
    if (!slides || slides.length === 0) return { createdCount: 0 };

    try {
      if (typeof PowerPoint !== 'undefined' && typeof PowerPoint.run === 'function') {
        return await PowerPoint.run(async (context: any) => {
          const slideCollection = context.presentation.slides;

          for (const s of slides) {
            slideCollection.add();
            const count = slideCollection.getCount();
            await context.sync();

            const slide = slideCollection.getItemAt(count.value - 1);
            slide.load(['id']);
            await context.sync();

            try {
              if (slide.shapes) {
                if (typeof slide.shapes.addTextBox === 'function') {
                  const titleBox = slide.shapes.addTextBox(s.title, {
                    left: 60,
                    top: 50,
                    width: 600,
                    height: 70,
                  });
                  if (titleBox.textFrame?.textRange) {
                    titleBox.textFrame.textRange.text = s.title;
                    titleBox.textFrame.textRange.font.size = 28;
                    titleBox.textFrame.textRange.font.bold = true;
                    titleBox.textFrame.textRange.font.color = '#1E293B';
                  }

                  if (s.bullets && s.bullets.length > 0) {
                    const bulletText = s.bullets.map(b => `• ${b}`).join('\n\n');
                    const bodyBox = slide.shapes.addTextBox(bulletText, {
                      left: 60,
                      top: 140,
                      width: 600,
                      height: 270,
                    });
                    if (bodyBox.textFrame?.textRange) {
                      bodyBox.textFrame.textRange.text = bulletText;
                      bodyBox.textFrame.textRange.font.size = 18;
                      bodyBox.textFrame.textRange.font.color = '#334155';
                    }
                  }
                } else if (typeof slide.shapes.addGeometricShape === 'function') {
                  const rectType = PowerPoint.GeometricShapeType?.rectangle || 'Rectangle';
                  const rect = slide.shapes.addGeometricShape(rectType, {
                    left: 60,
                    top: 50,
                    width: 600,
                    height: 70,
                  });
                  if (rect.fill?.setSolidColor) {
                    rect.fill.setSolidColor('#F8FAFC');
                  }
                  if (rect.textFrame?.textRange) {
                    rect.textFrame.textRange.text = s.title;
                    rect.textFrame.textRange.font.size = 26;
                    rect.textFrame.textRange.font.bold = true;
                  }

                  if (s.bullets && s.bullets.length > 0) {
                    const bodyRect = slide.shapes.addGeometricShape(rectType, {
                      left: 60,
                      top: 140,
                      width: 600,
                      height: 270,
                    });
                    if (bodyRect.fill?.setSolidColor) {
                      bodyRect.fill.setSolidColor('#FFFFFF');
                    }
                    if (bodyRect.textFrame?.textRange) {
                      bodyRect.textFrame.textRange.text = s.bullets.map(b => `• ${b}`).join('\n\n');
                      bodyRect.textFrame.textRange.font.size = 18;
                    }
                  }
                }
              }
            } catch (shapeErr) {
              console.warn('Gagal menambahkan shape ke slide deck:', shapeErr);
            }

            await context.sync();
          }

          return { createdCount: slides.length };
        });
      }
    } catch (err: any) {
      console.error('PowerPoint.run createPresentationDeck error:', err);
      throw new Error(`Gagal membuat deck presentasi PowerPoint: ${err.message || String(err)}`);
    }

    return { createdCount: slides.length };
  }

  async generateThemedDeck(options: ThemedDeckOptions): Promise<{ success: boolean; createdCount: number }> {
    if (!options.slides || options.slides.length === 0) return { success: true, createdCount: 0 };
    const theme = options.theme || 'corporate_blue';

    try {
      if (typeof PowerPoint !== 'undefined' && typeof PowerPoint.run === 'function') {
        return await PowerPoint.run(async (context: any) => {
          const slideCollection = context.presentation.slides;

          const themePalette: Record<string, { primary: string; secondary: string; bg: string; text: string }> = {
            corporate_blue: { primary: '#1E3A8A', secondary: '#3B82F6', bg: '#F8FAFC', text: '#1E293B' },
            emerald_executive: { primary: '#065F46', secondary: '#10B981', bg: '#F0FDF4', text: '#0F172A' },
            modern_dark: { primary: '#38BDF8', secondary: '#818CF8', bg: '#0F172A', text: '#F8FAFC' },
            minimalist_clean: { primary: '#18181B', secondary: '#71717A', bg: '#FFFFFF', text: '#27272A' },
          };
          const colors = themePalette[theme] || themePalette.corporate_blue;

          for (const s of options.slides) {
            slideCollection.add();
            const count = slideCollection.getCount();
            await context.sync();

            const slide = slideCollection.getItemAt(count.value - 1);
            slide.load(['id']);
            await context.sync();

            try {
              if (slide.shapes) {
                // Background shape if GeometricShape is available
                if (typeof slide.shapes.addGeometricShape === 'function') {
                  const rectType = PowerPoint.GeometricShapeType?.rectangle || 'Rectangle';
                  const bgRect = slide.shapes.addGeometricShape(rectType, {
                    left: 0,
                    top: 0,
                    width: 720,
                    height: 405,
                  });
                  if (bgRect.fill?.setSolidColor) {
                    bgRect.fill.setSolidColor(colors.bg);
                  }
                }

                if (typeof slide.shapes.addTextBox === 'function') {
                  if (s.layout === 'title_cover') {
                    const titleBox = slide.shapes.addTextBox(s.title, {
                      left: 60,
                      top: 120,
                      width: 600,
                      height: 100,
                    });
                    if (titleBox.textFrame?.textRange) {
                      titleBox.textFrame.textRange.text = s.title;
                      titleBox.textFrame.textRange.font.size = 36;
                      titleBox.textFrame.textRange.font.bold = true;
                      titleBox.textFrame.textRange.font.color = colors.primary;
                    }
                    if (s.content && s.content.length > 0) {
                      const subtitleText = s.content.join('  •  ');
                      const subBox = slide.shapes.addTextBox(subtitleText, {
                        left: 60,
                        top: 230,
                        width: 600,
                        height: 60,
                      });
                      if (subBox.textFrame?.textRange) {
                        subBox.textFrame.textRange.text = subtitleText;
                        subBox.textFrame.textRange.font.size = 20;
                        subBox.textFrame.textRange.font.color = colors.secondary;
                      }
                    }
                  } else if (s.layout === 'metric_highlights' && s.metrics && s.metrics.length > 0) {
                    const titleBox = slide.shapes.addTextBox(s.title, {
                      left: 60,
                      top: 40,
                      width: 600,
                      height: 60,
                    });
                    if (titleBox.textFrame?.textRange) {
                      titleBox.textFrame.textRange.text = s.title;
                      titleBox.textFrame.textRange.font.size = 28;
                      titleBox.textFrame.textRange.font.bold = true;
                      titleBox.textFrame.textRange.font.color = colors.primary;
                    }

                    const cardWidth = Math.min(180, Math.floor(600 / s.metrics.length) - 10);
                    s.metrics.forEach((m, mIdx) => {
                      const mBox = slide.shapes.addTextBox(`${m.value}\n${m.label}`, {
                        left: 60 + mIdx * (cardWidth + 15),
                        top: 140,
                        width: cardWidth,
                        height: 120,
                      });
                      if (mBox.textFrame?.textRange) {
                        mBox.textFrame.textRange.text = `${m.value}\n${m.label}`;
                        mBox.textFrame.textRange.font.size = 22;
                        mBox.textFrame.textRange.font.bold = true;
                        mBox.textFrame.textRange.font.color = colors.secondary;
                      }
                    });
                  } else {
                    const titleBox = slide.shapes.addTextBox(s.title, {
                      left: 60,
                      top: 40,
                      width: 600,
                      height: 60,
                    });
                    if (titleBox.textFrame?.textRange) {
                      titleBox.textFrame.textRange.text = s.title;
                      titleBox.textFrame.textRange.font.size = 28;
                      titleBox.textFrame.textRange.font.bold = true;
                      titleBox.textFrame.textRange.font.color = colors.primary;
                    }

                    if (s.content && s.content.length > 0) {
                      const bodyBox = slide.shapes.addTextBox(s.content.map(b => `• ${b}`).join('\n\n'), {
                        left: 60,
                        top: 120,
                        width: 600,
                        height: 250,
                      });
                      if (bodyBox.textFrame?.textRange) {
                        bodyBox.textFrame.textRange.text = s.content.map(b => `• ${b}`).join('\n\n');
                        bodyBox.textFrame.textRange.font.size = 18;
                        bodyBox.textFrame.textRange.font.color = colors.text;
                      }
                    }
                  }
                }
              }
            } catch (shapeErr) {
              console.warn('Gagal menambahkan shape ber-tema ke slide:', shapeErr);
            }

            await context.sync();
          }

          return { success: true, createdCount: options.slides.length };
        });
      }
    } catch (err: any) {
      console.error('PowerPoint.run generateThemedDeck error:', err);
      throw new Error(`Gagal membuat themed deck PowerPoint: ${err.message || String(err)}`);
    }

    return { success: true, createdCount: options.slides.length };
  }

  async transformDocToDeck(options?: DocToDeckOptions): Promise<DocToDeckResult> {
    const result = synthesizeDocToDeck(options?.documentText, options);
    const theme = options?.theme || 'corporate_blue';

    try {
      if (typeof PowerPoint !== 'undefined' && typeof PowerPoint.run === 'function') {
        await PowerPoint.run(async (context: any) => {
          const slideCollection = context.presentation.slides;

          const themePalette: Record<string, { primary: string; secondary: string; bg: string; text: string }> = {
            corporate_blue: { primary: '#1E3A8A', secondary: '#3B82F6', bg: '#F8FAFC', text: '#1E293B' },
            emerald_executive: { primary: '#065F46', secondary: '#10B981', bg: '#F0FDF4', text: '#0F172A' },
            modern_dark: { primary: '#38BDF8', secondary: '#818CF8', bg: '#0F172A', text: '#F8FAFC' },
            minimalist_clean: { primary: '#18181B', secondary: '#71717A', bg: '#FFFFFF', text: '#27272A' },
          };
          const colors = themePalette[theme] || themePalette.corporate_blue;

          for (const s of result.slides) {
            slideCollection.add();
            const count = slideCollection.getCount();
            await context.sync();

            const slide = slideCollection.getItemAt(count.value - 1);
            slide.load(['id']);
            await context.sync();

            try {
              if (slide.shapes) {
                if (typeof slide.shapes.addGeometricShape === 'function') {
                  const rectType = PowerPoint.GeometricShapeType?.rectangle || 'Rectangle';
                  const bgRect = slide.shapes.addGeometricShape(rectType, {
                    left: 0,
                    top: 0,
                    width: 720,
                    height: 405,
                  });
                  if (bgRect.fill?.setSolidColor) {
                    bgRect.fill.setSolidColor(colors.bg);
                  }
                }

                if (typeof slide.shapes.addTextBox === 'function') {
                  if (s.category === 'cover') {
                    const titleBox = slide.shapes.addTextBox(s.title, {
                      left: 60,
                      top: 120,
                      width: 600,
                      height: 100,
                    });
                    if (titleBox.textFrame?.textRange) {
                      titleBox.textFrame.textRange.text = s.title;
                      titleBox.textFrame.textRange.font.size = 36;
                      titleBox.textFrame.textRange.font.bold = true;
                      titleBox.textFrame.textRange.font.color = colors.primary;
                    }
                    if (s.bullets && s.bullets.length > 0) {
                      const subtitleText = s.bullets.join('  •  ');
                      const subBox = slide.shapes.addTextBox(subtitleText, {
                        left: 60,
                        top: 230,
                        width: 600,
                        height: 60,
                      });
                      if (subBox.textFrame?.textRange) {
                        subBox.textFrame.textRange.text = subtitleText;
                        subBox.textFrame.textRange.font.size = 20;
                        subBox.textFrame.textRange.font.color = colors.secondary;
                      }
                    }
                  } else {
                    const titleBox = slide.shapes.addTextBox(s.title, {
                      left: 60,
                      top: 40,
                      width: 600,
                      height: 60,
                    });
                    if (titleBox.textFrame?.textRange) {
                      titleBox.textFrame.textRange.text = s.title;
                      titleBox.textFrame.textRange.font.size = 28;
                      titleBox.textFrame.textRange.font.bold = true;
                      titleBox.textFrame.textRange.font.color = colors.primary;
                    }

                    if (s.bullets && s.bullets.length > 0) {
                      const bulletText = s.bullets.map((b) => `• ${b}`).join('\n\n');
                      const bodyBox = slide.shapes.addTextBox(bulletText, {
                        left: 60,
                        top: 120,
                        width: 600,
                        height: 250,
                      });
                      if (bodyBox.textFrame?.textRange) {
                        bodyBox.textFrame.textRange.text = bulletText;
                        bodyBox.textFrame.textRange.font.size = 18;
                        bodyBox.textFrame.textRange.font.color = colors.text;
                      }
                    }
                  }
                }
              }
            } catch (shapeErr) {
              console.warn('Gagal menambahkan shape ber-tema ke slide:', shapeErr);
            }

            await context.sync();
          }
        });
      }
    } catch (err: any) {
      console.error('PowerPoint.run transformDocToDeck error:', err);
      throw new Error(`Gagal mentransformasi dokumen ke slide PowerPoint: ${err.message || String(err)}`);
    }

    return result;
  }

  async insertProcessFlowchart(options: InsertFlowchartOptions): Promise<DiagramResult> {
    const def = options.definition || synthesizeFlowchartFromText(options.textOrSteps, options);
    const base64 = renderFlowchartToPngBase64(def, options);
    const title = def.title || options.title || 'Alur Proses';

    try {
      if (typeof PowerPoint !== 'undefined' && typeof PowerPoint.run === 'function') {
        await PowerPoint.run(async (context: any) => {
          const slideCollection = context.presentation.slides;
          slideCollection.add();
          const count = slideCollection.getCount();
          await context.sync();

          const slide = slideCollection.getItemAt(count.value - 1);
          slide.load(['id']);
          await context.sync();

          try {
            if (slide.shapes) {
              if (typeof slide.shapes.addTextBox === 'function') {
                const titleBox = slide.shapes.addTextBox(title, {
                  left: 60,
                  top: 30,
                  width: 600,
                  height: 50,
                });
                if (titleBox.textFrame?.textRange) {
                  titleBox.textFrame.textRange.text = title;
                  titleBox.textFrame.textRange.font.size = 24;
                  titleBox.textFrame.textRange.font.bold = true;
                  titleBox.textFrame.textRange.font.color = '#1E293B';
                }
              }

              if (typeof slide.shapes.addImage === 'function') {
                slide.shapes.addImage(base64, {
                  left: 60,
                  top: 90,
                  width: 600,
                  height: 350,
                });
              } else if (typeof slide.shapes.addImageFromBase64 === 'function') {
                slide.shapes.addImageFromBase64(base64, {
                  left: 60,
                  top: 90,
                  width: 600,
                  height: 350,
                });
              }
            }
          } catch (shapeErr) {
            console.warn('Gagal menambahkan shape flowchart ke slide:', shapeErr);
          }

          await context.sync();
        });
      }
    } catch (err: any) {
      console.warn('PowerPoint.run insertProcessFlowchart fallback:', err);
    }

    return {
      title,
      nodeCount: def.nodes.length,
      edgeCount: def.edges.length,
      appliedTheme: def.theme || options.theme || 'corporate_navy',
      base64Png: base64,
      inserted: true,
    };
  }
}
