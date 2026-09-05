import { IDocumentDriver } from './types';

export class MockOfficeDriver implements IDocumentDriver {
  hostType = 'BrowserDev';
  private excelGrid: Record<string, any[][]> = {};
  private wordContent: string[] = [];
  private slides: Array<{ title: string; bullets: string[]; notes: string; layout?: string }> = [];

  async readActiveRange() {
    const writtenKeys = Object.keys(this.excelGrid);
    if (writtenKeys.length > 0) {
      const lastKey = writtenKeys[writtenKeys.length - 1];
      const values = this.excelGrid[lastKey] || [];
      return {
        address: lastKey,
        values,
        formulas: values.map((row: any[]) => (Array.isArray(row) ? row.map(() => '') : [''])),
      };
    }
    return {
      address: 'A1:C5',
      values: [
        ['Produk', 'Qty', 'Harga'],
        ['Laptop', 5, 15000000],
        ['Mouse', 20, 150000],
        ['Keyboard', 12, 500000],
        ['Monitor', 8, 2500000],
      ],
      formulas: [['', '', ''], ['', '', ''], ['', '', ''], ['', '', ''], ['', '', '']],
    };
  }

  async readActiveSheetData(range?: string) {
    const res = await this.readActiveRange();
    return {
      sheetName: 'Sheet1',
      address: range || res.address,
      rowCount: res.values.length,
      columnCount: res.values[0]?.length || 0,
      values: res.values,
      formulas: res.formulas,
    };
  }

  async writeCells(range: string, values?: any[][], formulas?: string[][]) {
    if (values) {
      this.excelGrid[range] = values;
    } else if (formulas) {
      this.excelGrid[range] = formulas;
    }
  }

  async formatRange(_range: string, _styles: Record<string, any>) {
    // Mock styling
  }

  async createChart(_type: string, _dataRange: string, _title?: string) {
    // Mock chart
  }

  async getWordOutline() {
    return this.wordContent.join('\n') || 'Dokumen Word Kosong.';
  }

  async insertContent(position: 'start' | 'end' | 'cursor', text: string, _type: string) {
    if (position === 'start') this.wordContent.unshift(text);
    else this.wordContent.push(text);
  }

  async replaceSelection(newText: string) {
    this.wordContent = [newText];
  }

  async insertTable(rows: number, cols: number, _data?: string[][]) {
    this.wordContent.push(`[Tabel ${rows}x${cols}]`);
  }

  async getSlideContext() {
    const active = this.slides[0] || { title: 'Slide 1', bullets: [], notes: '' };
    return {
      slideNumber: 1,
      title: active.title,
      textContent: active.bullets.join('\n'),
    };
  }

  async addSlide(layout: string) {
    this.slides.push({ title: `Slide ${this.slides.length + 1}`, bullets: [], notes: '', layout });
    return { slideNumber: this.slides.length };
  }

  async insertSlideContent(title: string, bullets: string[]) {
    if (this.slides.length === 0) await this.addSlide('TitleAndContent');
    const current = this.slides[this.slides.length - 1];
    current.title = title;
    current.bullets = bullets;
  }

  async setSpeakerNotes(notes: string) {
    if (this.slides.length === 0) await this.addSlide('TitleAndContent');
    this.slides[this.slides.length - 1].notes = notes;
  }
}
