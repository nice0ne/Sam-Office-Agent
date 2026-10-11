import { describe, it, expect, beforeEach, vi } from 'vitest';
import { PPTDriver, PPT_CANVAS_WIDTH, PPT_CANVAS_HEIGHT, PPT_MARGIN_HORIZONTAL, PPT_CONTENT_WIDTH } from '../pptDriver';

describe('PPTDriver Layout & Slide Geometry', () => {
  let driver: PPTDriver;
  let addedShapes: any[];
  let mockSlides: any[];

  beforeEach(() => {
    driver = new PPTDriver();
    addedShapes = [];
    mockSlides = [];

    const mockSlide = {
      id: 'slide-1',
      load: vi.fn(),
      shapes: {
        addGeometricShape: vi.fn((type: string, options: any) => {
          const shape = {
            type,
            options,
            lineFormat: { visible: true },
            fill: {
              color: '',
              setSolidColor: vi.fn((c: string) => { shape.fill.color = c; }),
            },
            textFrame: {
              textRange: {
                text: '',
                font: { size: 14, bold: false, color: '' },
              },
            },
          };
          addedShapes.push({ kind: 'geometric', ...shape });
          return shape;
        }),
        addTextBox: vi.fn((text: string, options: any) => {
          const shape = {
            text,
            options,
            textFrame: {
              textRange: {
                text,
                font: { size: 14, bold: false, color: '' },
              },
            },
          };
          addedShapes.push({ kind: 'textbox', ...shape });
          return shape;
        }),
        addImage: vi.fn((data: string, options: any) => {
          const shape = { data, options };
          addedShapes.push({ kind: 'image', ...shape });
          return shape;
        }),
      },
    };

    mockSlides.push(mockSlide);

    // Mock global PowerPoint runtime
    (globalThis as any).PowerPoint = {
      GeometricShapeType: { rectangle: 'Rectangle' },
      run: async (callback: (context: any) => Promise<any>) => {
        const context = {
          presentation: {
            slides: {
              add: vi.fn(() => {
                const s = { ...mockSlide, id: `slide-${mockSlides.length + 1}` };
                mockSlides.push(s);
                return s;
              }),
              getCount: vi.fn(() => ({ value: mockSlides.length })),
              getItemAt: vi.fn((idx: number) => mockSlides[idx] || mockSlide),
            },
          },
          sync: vi.fn().mockResolvedValue(undefined),
        };
        return await callback(context);
      },
    };
  });

  it('exposes standard 16:9 widescreen canvas constants (960x540 pt)', () => {
    expect(PPT_CANVAS_WIDTH).toBe(960);
    expect(PPT_CANVAS_HEIGHT).toBe(540);
    expect(PPT_MARGIN_HORIZONTAL).toBe(80);
    expect(PPT_CONTENT_WIDTH).toBe(800);
  });

  it('creates full-bleed background shape matching 960x540 canvas without visible borders in generateThemedDeck', async () => {
    await driver.generateThemedDeck({
      topic: 'Strategi Pertumbuhan Bisnis',
      theme: 'mckinsey_consulting',
      slides: [
        {
          title: 'Strategi Pertumbuhan Bisnis & Capaian Kinerja',
          layout: 'title_cover',
          content: ['Review Manajemen', 'Kuartal 3'],
        },
      ],
    });

    const bgShape = addedShapes.find(s => s.kind === 'geometric' && s.options.left === 0 && s.options.top === 0);
    expect(bgShape).toBeDefined();
    expect(bgShape.options.width).toBe(960);
    expect(bgShape.options.height).toBe(540);
    expect(bgShape.lineFormat.visible).toBe(false);
    expect(bgShape.fill.color).toBe('#FFFFFF'); // mckinsey_consulting bg

    // Cover Title & Subtitle should be symmetrically centered/proportional with width 800
    const titleBox = addedShapes.find(s => s.kind === 'textbox' && s.text === 'Strategi Pertumbuhan Bisnis & Capaian Kinerja');
    expect(titleBox).toBeDefined();
    expect(titleBox.options.left).toBe(80);
    expect(titleBox.options.width).toBe(800);
    expect(titleBox.options.top).toBeGreaterThanOrEqual(140);
    expect(titleBox.textFrame.textRange.font.size).toBeGreaterThanOrEqual(36);

    const subBox = addedShapes.find(s => s.kind === 'textbox' && s.text.includes('Review Manajemen'));
    expect(subBox).toBeDefined();
    expect(subBox.options.left).toBe(80);
    expect(subBox.options.width).toBe(800);
    expect(subBox.options.top).toBeGreaterThan(titleBox.options.top);
  });

  it('creates proportional 800pt content layouts for bullet points and metric highlights', async () => {
    await driver.generateThemedDeck({
      topic: 'Metrik & Capaian',
      theme: 'corporate_blue',
      slides: [
        {
          title: 'Key Performance Indicators',
          layout: 'metric_highlights',
          metrics: [
            { label: 'Pertumbuhan ARR', value: '+142%' },
            { label: 'Retention Rate', value: '98.4%' },
            { label: 'NPS Score', value: '76' },
          ],
        },
        {
          title: 'Rencana Eksekusi',
          layout: 'bullet_points',
          content: ['Inisiatif 1', 'Inisiatif 2'],
        },
      ],
    });

    // Content slide titles should be placed at left 80, width 800
    const kpiTitle = addedShapes.find(s => s.kind === 'textbox' && s.text === 'Key Performance Indicators');
    expect(kpiTitle).toBeDefined();
    expect(kpiTitle.options.left).toBe(80);
    expect(kpiTitle.options.width).toBe(800);

    // Metric items should span across the 800pt canvas
    const metricBoxes = addedShapes.filter(s => s.kind === 'textbox' && s.text.includes('Pertumbuhan ARR'));
    expect(metricBoxes.length).toBeGreaterThan(0);
    expect(metricBoxes[0].options.left).toBeGreaterThanOrEqual(80);
  });

  it('renders transformDocToDeck with 960x540 canvas and borderless background', async () => {
    const res = await driver.transformDocToDeck({
      documentText: 'Ringkasan Laporan Tahunan.\nEkspansi pasar sukses besar.\nPendapatan melonjak.',
      theme: 'linear_dark',
    });

    expect(res.slides.length).toBeGreaterThan(0);
    const bgShape = addedShapes.find(s => s.kind === 'geometric' && s.options.width === 960 && s.options.height === 540);
    expect(bgShape).toBeDefined();
    expect(bgShape.lineFormat.visible).toBe(false);
    expect(bgShape.fill.color).toBe('#08090A'); // linear_dark bg
  });
});
