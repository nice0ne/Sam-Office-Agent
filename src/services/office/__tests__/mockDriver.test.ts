import { describe, it, expect, beforeEach } from 'vitest';
import { MockOfficeDriver } from '../mockDriver';
import { getOfficeDriver, resetOfficeDriver } from '../index';

describe('MockOfficeDriver', () => {
  let driver: MockOfficeDriver;

  beforeEach(() => {
    driver = new MockOfficeDriver();
  });

  it('handles Excel write and read operations in mock mode', async () => {
    await driver.writeCells('A1:B2', [['Item', 'Price'], ['Apple', 10]]);
    const range = await driver.readActiveRange();
    expect(range.values.length).toBeGreaterThan(0);
    expect(range.values[0][0]).toBe('Item');
  });

  it('returns default active range when no cells have been written', async () => {
    const freshDriver = new MockOfficeDriver();
    const range = await freshDriver.readActiveRange();
    expect(range.address).toBe('A1:C5');
    expect(range.values[0][0]).toBe('Produk');
    expect(range.formulas.length).toBe(5);
  });

  it('handles Excel formatting and chart creation without errors', async () => {
    await expect(driver.formatRange('A1:B2', { bold: true, color: '#FF0000' })).resolves.toBeUndefined();
    await expect(driver.createChart('columnClustered', 'A1:B2', 'Sales Chart')).resolves.toBeUndefined();
  });

  it('handles cleanData operation with trimming, deduplication, and empty filling', async () => {
    await driver.writeCells('A1:B4', [
      [' Nama ', ' Nilai '],
      [' Andi ', ''],
      [' Budi ', 80],
      [' Andi ', ''],
    ]);

    const result = await driver.cleanData({
      range: 'A1:B4',
      trimWhitespace: true,
      removeDuplicates: true,
      fillEmptyValues: 0,
    });

    expect(result.removedDuplicatesCount).toBe(1);
    expect(result.trimmedCellsCount).toBeGreaterThan(0);
    expect(result.filledCellsCount).toBe(2);
    expect(result.cleanedRows).toBe(3);

    const updated = await driver.readActiveRange();
    expect(updated.values[1][0]).toBe('Andi');
    expect(updated.values[1][1]).toBe(0);
  });

  it('handles applyConditionalFormatting on mock driver', async () => {
    const res = await driver.applyConditionalFormatting({
      range: 'B2:B10',
      type: 'data_bar',
      color: '#3B82F6',
    });
    expect(res.success).toBe(true);
    expect(res.rule).toBe('data_bar on B2:B10');
  });

  it('handles Word content insertion and selection replacement', async () => {
    await driver.insertContent('start', 'Halo Dunia', 'paragraph');
    const outline = await driver.getWordOutline();
    expect(outline).toContain('Halo Dunia');

    await driver.insertContent('end', 'Paragraph Kedua', 'paragraph');
    const updatedOutline = await driver.getWordOutline();
    expect(updatedOutline).toContain('Paragraph Kedua');

    await driver.replaceSelection('Teks Pengganti');
    expect(await driver.getWordOutline()).toBe('Teks Pengganti');

    await driver.insertTable(3, 3);
    expect(await driver.getWordOutline()).toContain('[Tabel 3x3]');
  });

  it('returns empty document notice when Word content is empty', async () => {
    const freshDriver = new MockOfficeDriver();
    const outline = await freshDriver.getWordOutline();
    expect(outline).toBe('Dokumen Word Kosong.');
  });

  it('handles PowerPoint slide creation and manipulation', async () => {
    const slide = await driver.addSlide('TitleAndContent');
    expect(slide.slideNumber).toBe(1);

    await driver.insertSlideContent('Slide Utama', ['Point 1', 'Point 2']);
    await driver.setSpeakerNotes('Catatan pembicara untuk presentasi');

    const context = await driver.getSlideContext();
    expect(context.slideNumber).toBe(1);
    expect(context.title).toBe('Slide Utama');
    expect(context.textContent).toContain('Point 1');
  });

  it('audits sheet data and detects formula errors and inconsistent overrides', async () => {
    const driver = new MockOfficeDriver();
    driver.mockData = [
      ['Header1', 'Header2', 'Total'],
      [10, 20, 200],
      [15, 0, '#DIV/0!'],
      [20, 5, 100],
      [25, 4, 'manual_text_instead_of_number'],
    ];

    const audit = await driver.auditSheetData!();
    expect(audit.totalCellsAudited).toBeGreaterThan(0);
    expect(audit.totalErrorsFound).toBeGreaterThanOrEqual(1);
    const divZero = audit.criticalIssues.find(i => i.currentValue === '#DIV/0!');
    expect(divZero).toBeDefined();
    expect(divZero?.address).toBe('C3');
    expect(divZero?.suggestion).toBeDefined();
  });

  it('generates executive data story with metrics and recommendations', async () => {
    const driver = new MockOfficeDriver();
    driver.mockData = [
      ['Wilayah', 'Penjualan Q1', 'Penjualan Q2'],
      ['Barat', 1000, 1200],
      ['Timur', 800, 1100],
      ['Pusat', 1500, 1400],
    ];

    const story = await driver.generateDataStory!({ focusMetric: 'Penjualan Q2' });
    expect(story.headline).toBeDefined();
    expect(story.keyFindings.length).toBeGreaterThan(0);
    expect(story.metrics.length).toBeGreaterThan(0);
    expect(story.recommendations.length).toBeGreaterThan(0);
  });
});

describe('getOfficeDriver', () => {
  beforeEach(() => {
    resetOfficeDriver();
    delete (globalThis as any).Office;
    delete (globalThis as any).window?.Office;
  });

  it('returns MockOfficeDriver by default in browser/node environment', () => {
    const driver = getOfficeDriver();
    expect(driver).toBeDefined();
    expect(driver.hostType).toBe('BrowserDev');
  });

  it('returns singleton instance on subsequent calls', () => {
    const driver1 = getOfficeDriver();
    const driver2 = getOfficeDriver();
    expect(driver1).toBe(driver2);
  });

  it('returns ExcelDriver when Office host is Excel', () => {
    (globalThis as any).Office = {
      context: { host: 'Excel' },
      HostType: { Excel: 'Excel', Word: 'Word', PowerPoint: 'PowerPoint' },
    };
    const driver = getOfficeDriver();
    expect(driver.hostType).toBe('Excel');
  });

  it('returns WordDriver when Office host is Word', () => {
    (globalThis as any).Office = {
      context: { host: 'Word' },
      HostType: { Excel: 'Excel', Word: 'Word', PowerPoint: 'PowerPoint' },
    };
    const driver = getOfficeDriver();
    expect(driver.hostType).toBe('Word');
  });

  it('returns PPTDriver when Office host is PowerPoint', () => {
    (globalThis as any).Office = {
      context: { host: 'PowerPoint' },
      HostType: { Excel: 'Excel', Word: 'Word', PowerPoint: 'PowerPoint' },
    };
    const driver = getOfficeDriver();
    expect(driver.hostType).toBe('PowerPoint');
  });
});
