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
