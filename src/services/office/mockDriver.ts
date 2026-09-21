import { CleanDataOptions, CleanDataResult, ConditionalFormattingOptions, IDocumentDriver } from './types';

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

  async addWorksheet(sheetName: string) {
    return { sheetName };
  }

  async autoFitColumns(_range?: string) {
    // Mock auto-fit
  }

  async sortRange(
    range: string,
    columnIndex: number,
    ascending: boolean = true,
    hasHeaders: boolean = true,
    _enableAutoFilter?: boolean
  ) {
    const grid = this.excelGrid[range];
    if (grid && grid.length > 0) {
      const headers = hasHeaders ? grid.slice(0, 1) : [];
      const body = hasHeaders ? grid.slice(1) : [...grid];
      body.sort((a, b) => {
        const valA = a[columnIndex] ?? '';
        const valB = b[columnIndex] ?? '';
        if (valA < valB) return ascending ? -1 : 1;
        if (valA > valB) return ascending ? 1 : -1;
        return 0;
      });
      this.excelGrid[range] = [...headers, ...body];
    }
  }

  async cleanData(options?: CleanDataOptions): Promise<CleanDataResult> {
    const rangeAddress = options?.range;
    let targetKey = rangeAddress;
    let gridData: any[][];

    if (targetKey && this.excelGrid[targetKey]) {
      gridData = this.excelGrid[targetKey];
    } else {
      const writtenKeys = Object.keys(this.excelGrid);
      if (writtenKeys.length > 0) {
        targetKey = targetKey || writtenKeys[writtenKeys.length - 1];
        gridData = this.excelGrid[targetKey] || (await this.readActiveRange()).values;
      } else {
        targetKey = targetKey || 'A1:C5';
        gridData = (await this.readActiveRange()).values;
      }
    }

    let rows = gridData.map(row => (Array.isArray(row) ? [...row] : [row]));
    let trimmedCellsCount = 0;
    let filledCellsCount = 0;
    let removedDuplicatesCount = 0;

    // Trim whitespace
    if (options?.trimWhitespace) {
      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < rows[r].length; c++) {
          const val = rows[r][c];
          if (typeof val === 'string') {
            const trimmed = val.trim();
            if (trimmed !== val) {
              rows[r][c] = trimmed;
              trimmedCellsCount++;
            }
          }
        }
      }
    }

    // Fill empty values
    if (options?.fillEmptyValues !== undefined) {
      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < rows[r].length; c++) {
          const val = rows[r][c];
          if (val === null || val === undefined || (typeof val === 'string' && val.trim() === '') || val === '') {
            rows[r][c] = options.fillEmptyValues;
            filledCellsCount++;
          }
        }
      }
    }

    // Remove duplicates
    if (options?.removeDuplicates) {
      const seen = new Set<string>();
      const uniqueRows: any[][] = [];
      for (const row of rows) {
        const key = JSON.stringify(row);
        if (seen.has(key)) {
          removedDuplicatesCount++;
        } else {
          seen.add(key);
          uniqueRows.push(row);
        }
      }
      rows = uniqueRows;
    }

    this.excelGrid[targetKey] = rows;

    return {
      cleanedRows: rows.length,
      removedDuplicatesCount,
      trimmedCellsCount,
      filledCellsCount,
    };
  }

  async applyConditionalFormatting(options: ConditionalFormattingOptions): Promise<{ success: boolean; rule: string }> {
    return {
      success: true,
      rule: `${options.type} on ${options.range}`,
    };
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

  async createSlide(title: string, bullets: string[], notes?: string, layout?: string) {
    this.slides.push({
      title,
      bullets,
      notes: notes || '',
      layout: layout || 'TitleAndContent',
    });
    return { slideNumber: this.slides.length };
  }

  async createPresentationDeck(slides: Array<{ title: string; bullets: string[]; notes?: string; layout?: string }>) {
    for (const s of slides) {
      await this.createSlide(s.title, s.bullets, s.notes, s.layout);
    }
    return { createdCount: slides.length };
  }

  async readSlideData(slideNumber?: number, allSlides = true): Promise<{
    activeSlideIndex: number;
    totalSlides: number;
    slides: Array<{ slideIndex: number; title: string; textContent: string; notes?: string }>;
  }> {
    const total = this.slides.length || 1;
    const all = this.slides.length > 0 ? this.slides : [{ title: 'Slide 1', bullets: [], notes: '', layout: 'TitleAndContent' }];

    let target = all;
    if (slideNumber && !allSlides) {
      const found = all[slideNumber - 1];
      target = found ? [found] : all;
    }

    const slidesList = target.map((s, idx) => ({
      slideIndex: slideNumber && !allSlides ? slideNumber : idx + 1,
      title: s.title,
      textContent: s.bullets && s.bullets.length > 0 ? s.bullets.map((b) => '• ' + b).join('\n') : '(Tidak ada poin teks)',
      notes: s.notes,
    }));

    return {
      activeSlideIndex: 1,
      totalSlides: total,
      slides: slidesList,
    };
  }
}
