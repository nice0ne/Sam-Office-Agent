import {
  CleanDataOptions,
  CleanDataResult,
  ComplianceReviewOptions,
  ComplianceReviewResult,
  ConditionalFormattingOptions,
  CorporateStyleOptions,
  CorporateStyleResult,
  DataStoryMetric,
  DataStoryOptions,
  DataStoryResult,
  DiagramResult,
  DocToDeckOptions,
  DocToDeckResult,
  FormulaIssue,
  IDocumentDriver,
  InsertFlowchartOptions,
  PolishTextOptions,
  SheetAuditResult,
  StructuredDocOptions,
  ThemedDeckOptions,
} from './types';
import { analyzeContractCompliance } from './complianceReviewer';
import { synthesizeDocToDeck } from './docToDeckTransformer';
import { synthesizeFlowchartFromText, renderFlowchartToPngBase64 } from '../../utils/diagramRenderer';

export class MockOfficeDriver implements IDocumentDriver {
  hostType = 'BrowserDev';
  mockData?: any[][];
  mockWordBody?: string;
  private excelGrid: Record<string, any[][]> = {};
  private wordContent: string[] = [];
  slides: Array<{ title: string; bullets: string[]; notes: string; layout?: string }> = [];
  diagrams: DiagramResult[] = [];

  async readActiveRange() {
    if (this.mockData && this.mockData.length > 0) {
      const rowCount = this.mockData.length;
      const colCount = this.mockData[0]?.length || 1;
      const endCol = colIndexToLetter(colCount - 1);
      return {
        address: `A1:${endCol}${rowCount}`,
        values: this.mockData,
        formulas: this.mockData.map((row: any[]) => (Array.isArray(row) ? row.map(() => '') : [''])),
      };
    }
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

  async auditSheetData(options?: { range?: string }): Promise<SheetAuditResult> {
    let values: any[][];
    let formulas: string[][] = [];

    if (options?.range && this.excelGrid[options.range]) {
      values = this.excelGrid[options.range];
    } else if (this.mockData) {
      values = this.mockData;
    } else {
      const res = await this.readActiveRange();
      values = res.values;
      formulas = res.formulas;
    }

    return performSheetAudit('Sheet1', values, formulas, 0, 0);
  }

  async generateDataStory(options?: DataStoryOptions): Promise<DataStoryResult> {
    let values: any[][];

    if (options?.range && this.excelGrid[options.range]) {
      values = this.excelGrid[options.range];
    } else if (this.mockData) {
      values = this.mockData;
    } else {
      const res = await this.readActiveRange();
      values = res.values;
    }

    return buildDataStoryFromResult(values, options);
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

  async generateStructuredDoc(options: StructuredDocOptions): Promise<{ success: boolean; message: string }> {
    this.wordContent.push(`[${options.templateType}] ${options.title}`);
    for (const section of options.sections) {
      this.wordContent.push(`## ${section.heading}`);
      if (section.content) this.wordContent.push(section.content);
      if (section.bullets && section.bullets.length > 0) {
        this.wordContent.push(...section.bullets.map((b) => `• ${b}`));
      }
      if (section.table) {
        this.wordContent.push(`[Tabel: ${section.table.headers.join(', ')}]`);
      }
    }
    return {
      success: true,
      message: `Dokumen ${options.templateType} "${options.title}" berhasil dibuat dengan ${options.sections.length} bagian.`,
    };
  }

  async polishDocumentText(options: PolishTextOptions): Promise<{ success: boolean; polishedText: string }> {
    const textToSet = options.customText || 'Teks telah berhasil dipoles dan disempurnakan.';
    this.wordContent = [textToSet];
    return {
      success: true,
      polishedText: textToSet,
    };
  }

  async reviewComplianceClauses(options?: ComplianceReviewOptions): Promise<ComplianceReviewResult> {
    const text = this.mockWordBody || this.wordContent.join('\n');
    return analyzeContractCompliance(text, options);
  }

  async applyCorporateStyle(options?: CorporateStyleOptions): Promise<CorporateStyleResult> {
    const theme = options?.theme || 'corporate_navy';
    const fontFamily = options?.fontFamily || (theme === 'official_government' ? 'Times New Roman' : 'Calibri');
    const text = this.mockWordBody || this.wordContent.join('\n');
    const paragraphs = text.split(/\r?\n/).filter(p => p.trim().length > 0);
    const headingsCount = paragraphs.filter(p => /^(bab|pasal|section|judul|#)/i.test(p.trim())).length || 1;
    const styledParagraphsCount = paragraphs.length || 1;

    return {
      appliedTheme: theme,
      fontFamily,
      styledParagraphsCount,
      headingsCount,
      message: `Berhasil menerapkan gaya korporat ${theme} (${fontFamily}) pada ${styledParagraphsCount} paragraf.`,
    };
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

  async generateThemedDeck(options: ThemedDeckOptions): Promise<{ success: boolean; createdCount: number }> {
    for (const s of options.slides) {
      const bullets = s.content || (s.metrics ? s.metrics.map(m => `${m.label}: ${m.value}`) : []);
      this.slides.push({
        title: s.title,
        bullets,
        notes: s.notes || '',
        layout: s.layout || 'bullet_points',
      });
    }
    return {
      success: true,
      createdCount: options.slides.length,
    };
  }

  async transformDocToDeck(options?: DocToDeckOptions): Promise<DocToDeckResult> {
    const result = synthesizeDocToDeck(options?.documentText, options);
    for (const slide of result.slides) {
      const notes = [
        slide.speakerScript.hook,
        ...slide.speakerScript.keyTalkingPoints,
        slide.speakerScript.transition,
      ]
        .filter(Boolean)
        .join('\n\n');

      this.slides.push({
        title: slide.title,
        bullets: slide.bullets,
        notes,
        layout: slide.category === 'cover' ? 'title_cover' : 'bullet_points',
      });
    }
    return result;
  }

  async insertProcessFlowchart(options: InsertFlowchartOptions): Promise<DiagramResult> {
    const def = options.definition || synthesizeFlowchartFromText(options.textOrSteps, options);
    const base64 = renderFlowchartToPngBase64(def, options);
    const result: DiagramResult = {
      title: def.title || options.title || 'Alur Proses',
      nodeCount: def.nodes.length,
      edgeCount: def.edges.length,
      appliedTheme: def.theme || options.theme || 'corporate_navy',
      base64Png: base64,
      inserted: true,
    };
    this.diagrams.push(result);
    return result;
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

export function colIndexToLetter(colIndex: number): string {
  let temp = colIndex;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

export function getCellAddress(colIndex: number, rowIndex: number): string {
  return `${colIndexToLetter(colIndex)}${rowIndex + 1}`;
}

export const FORMULA_ERROR_REGEX = /^#(REF!|VALUE!|DIV\/0!|N\/A|NAME\?|NUM!|NULL!)/i;

export function isFormulaError(val: any): boolean {
  if (typeof val === 'string') {
    return FORMULA_ERROR_REGEX.test(val.trim());
  }
  return false;
}

export function getFormulaErrorSuggestion(errorStr: string, address: string): string {
  const upper = errorStr.toUpperCase();
  if (upper.includes('#DIV/0!')) {
    return `Sel ${address} menghasilkan error pembagian dengan nol (#DIV/0!). Gunakan IFERROR(...) atau pastikan sel pembagi tidak bernilai 0.`;
  }
  if (upper.includes('#REF!')) {
    return `Sel ${address} mereferensikan sel yang tidak valid atau telah dihapus (#REF!). Periksa dan perbarui referensi sel.`;
  }
  if (upper.includes('#VALUE!')) {
    return `Sel ${address} mengalami kesalahan tipe data (#VALUE!). Pastikan semua argumen bernilai sesuai (misal angka vs teks).`;
  }
  if (upper.includes('#N/A')) {
    return `Sel ${address} tidak menemukan nilai yang dicari (#N/A). Pastikan nilai pencarian tersedia atau bungkus dengan IFNA/IFERROR.`;
  }
  if (upper.includes('#NAME?')) {
    return `Sel ${address} memuat nama fungsi atau rentang yang salah ketik (#NAME?). Periksa kembali penulisan formula.`;
  }
  if (upper.includes('#NUM!')) {
    return `Sel ${address} mengalami kalkulasi di luar batas valid numerik (#NUM!).`;
  }
  if (upper.includes('#NULL!')) {
    return `Sel ${address} memiliki perpotongan rentang yang tidak valid (#NULL!).`;
  }
  return `Perbaiki formula pada sel ${address} untuk mengatasi error ${errorStr}.`;
}

export function performSheetAudit(
  sheetName: string,
  values: any[][],
  formulas: string[][] = [],
  startRow: number = 0,
  startCol: number = 0,
  valueTypes: any[][] = []
): SheetAuditResult {
  if (!values || values.length === 0 || values.every(row => !row || row.length === 0)) {
    return {
      sheetName,
      totalCellsAudited: 0,
      totalErrorsFound: 0,
      criticalIssues: [],
      warnings: [],
      summary: `Sheet "${sheetName}" kosong. Tidak ada sel yang diaudit.`,
    };
  }

  const rowCount = values.length;
  const colCount = Math.max(...values.map(r => (Array.isArray(r) ? r.length : 0)));
  const totalCellsAudited = rowCount * colCount;

  const criticalIssues: FormulaIssue[] = [];
  const warnings: FormulaIssue[] = [];

  // 1. Scan for formula errors
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const val = values[r]?.[c];
      const formula = formulas[r]?.[c] || '';
      const vType = valueTypes[r]?.[c];
      const cellAddress = getCellAddress(startCol + c, startRow + r);

      const isErr = isFormulaError(val) || vType === 'Error' || (typeof val === 'string' && val.trim().startsWith('#'));
      if (isErr) {
        criticalIssues.push({
          address: cellAddress,
          type: 'formula_error',
          severity: 'critical',
          currentValue: val,
          formula: formula || undefined,
          suggestion: getFormulaErrorSuggestion(String(val), cellAddress),
        });
      }
    }
  }

  // 2. Scan for hardcoded overrides in formula/numeric columns
  if (rowCount >= 2) {
    for (let c = 0; c < colCount; c++) {
      const headerName = String(values[0]?.[c] ?? '').trim();
      const colLetter = colIndexToLetter(startCol + c);

      let formulaCount = 0;
      let numericCount = 0;
      let errorCount = 0;

      for (let r = 1; r < rowCount; r++) {
        const f = formulas[r]?.[c];
        const v = values[r]?.[c];

        if (typeof f === 'string' && f.trim().startsWith('=')) {
          formulaCount++;
        }
        if (typeof v === 'number' && !isNaN(v)) {
          numericCount++;
        }
        if (isFormulaError(v) || (typeof v === 'string' && v.trim().startsWith('#'))) {
          errorCount++;
        }
      }

      const dataRowCount = rowCount - 1;
      const isPredominantlyFormula = dataRowCount > 1 && formulaCount / dataRowCount >= 0.5;
      const isPredominantlyNumeric = dataRowCount > 1 && (numericCount + errorCount + formulaCount) / dataRowCount >= 0.5;

      if (isPredominantlyFormula) {
        for (let r = 1; r < rowCount; r++) {
          const f = formulas[r]?.[c];
          const v = values[r]?.[c];
          const cellAddress = getCellAddress(startCol + c, startRow + r);
          const isErr = isFormulaError(v) || (typeof v === 'string' && v.trim().startsWith('#'));

          if (!isErr && (!f || !f.trim().startsWith('=')) && v !== '' && v !== null && v !== undefined) {
            warnings.push({
              address: cellAddress,
              type: 'hardcoded_override',
              severity: 'warning',
              currentValue: v,
              suggestion: `Kolom ${headerName || colLetter} mayoritas menggunakan formula, namun sel ${cellAddress} diisi secara manual (${v}). Pertimbangkan menyamakan rumus.`,
            });
          }
        }
      } else if (isPredominantlyNumeric) {
        for (let r = 1; r < rowCount; r++) {
          const v = values[r]?.[c];
          const cellAddress = getCellAddress(startCol + c, startRow + r);
          const isErr = isFormulaError(v) || (typeof v === 'string' && v.trim().startsWith('#'));

          if (!isErr && typeof v === 'string' && v.trim() !== '' && isNaN(Number(v))) {
            warnings.push({
              address: cellAddress,
              type: 'hardcoded_override',
              severity: 'warning',
              currentValue: v,
              suggestion: `Kolom ${headerName || colLetter} mayoritas bernilai numerik, namun sel ${cellAddress} berisi teks manual ("${v}").`,
            });
          }
        }
      }
    }
  }

  const totalErrorsFound = criticalIssues.length + warnings.length;
  let summary = '';
  if (totalErrorsFound === 0) {
    summary = `Audit selesai pada ${totalCellsAudited} sel di sheet "${sheetName}". Tidak ditemukan error formula ataupun inkonsistensi data.`;
  } else {
    summary = `Audit sheet "${sheetName}": Ditemukan ${criticalIssues.length} error formula kritis dan ${warnings.length} peringatan inkonsistensi dari total ${totalCellsAudited} sel yang diaudit.`;
  }

  return {
    sheetName,
    totalCellsAudited,
    totalErrorsFound,
    criticalIssues,
    warnings,
    summary,
  };
}

export function buildDataStoryFromResult(values: any[][], options?: DataStoryOptions): DataStoryResult {
  if (!values || values.length === 0 || values.every(row => !row || row.length === 0)) {
    return {
      headline: 'Data Tidak Tersedia untuk Analisis',
      keyFindings: ['Rentang data kosong atau tidak memuat informasi yang dapat dianalisis.'],
      metrics: [],
      recommendations: options?.includeRecommendations !== false
        ? ['Pastikan rentang yang dipilih memuat data numerik dan label kategori.']
        : [],
    };
  }

  const headers = (values[0] || []).map(h => String(h ?? '').trim());
  const dataRows = values.slice(1).filter(r => Array.isArray(r) && r.some(c => c !== null && c !== ''));

  if (dataRows.length === 0) {
    return {
      headline: 'Hanya Header Ditemukan',
      keyFindings: ['Tabel hanya memiliki baris header tanpa baris data untuk dianalisis.'],
      metrics: [],
      recommendations: options?.includeRecommendations !== false
        ? ['Tambahkan baris data transaksi atau metrik pada tabel.']
        : [],
    };
  }

  // Identify numeric columns
  const numericColIndices: number[] = [];
  for (let c = 0; c < headers.length; c++) {
    let numCount = 0;
    for (const row of dataRows) {
      const v = row[c];
      const parsed = typeof v === 'number' ? v : parseFloat(String(v).replace(/[^0-9.-]+/g, ''));
      if (!isNaN(parsed) && typeof v !== 'boolean') {
        numCount++;
      }
    }
    if (numCount >= Math.ceil(dataRows.length * 0.4) && numCount > 0) {
      numericColIndices.push(c);
    }
  }

  // Determine focus column
  let focusColIdx = -1;
  if (options?.focusMetric) {
    const focusTarget = options.focusMetric.toLowerCase().trim();
    focusColIdx = headers.findIndex(h => h.toLowerCase().includes(focusTarget));
  }
  if (focusColIdx === -1 && numericColIndices.length > 0) {
    focusColIdx = numericColIndices[numericColIndices.length - 1];
  }
  if (focusColIdx === -1) {
    focusColIdx = headers.length > 1 ? 1 : 0;
  }

  const focusHeader = headers[focusColIdx] || `Kolom ${focusColIdx + 1}`;

  // Collect entries for focus column
  const categoryColIdx = 0;
  const entries: Array<{ label: string; value: number }> = dataRows.map((r, i) => {
    const rawVal = r[focusColIdx];
    const num = typeof rawVal === 'number' ? rawVal : parseFloat(String(rawVal).replace(/[^0-9.-]+/g, '')) || 0;
    const cat = categoryColIdx !== focusColIdx && r[categoryColIdx] !== undefined && r[categoryColIdx] !== ''
      ? String(r[categoryColIdx])
      : `Item ${i + 1}`;
    return { label: cat, value: num };
  });

  const validValues = entries.map(e => e.value);
  const sum = validValues.reduce((a, b) => a + b, 0);
  const avg = validValues.length > 0 ? sum / validValues.length : 0;
  const min = validValues.length > 0 ? Math.min(...validValues) : 0;
  const max = validValues.length > 0 ? Math.max(...validValues) : 0;

  const maxEntry = entries.find(e => e.value === max);
  const minEntry = entries.find(e => e.value === min);

  // Check for preceding comparison numeric column
  const prevNumericColIdx = numericColIndices.filter(idx => idx < focusColIdx).pop();
  let deltaPercent: number | undefined;
  let trend: 'up' | 'down' | 'neutral' = 'neutral';

  if (prevNumericColIdx !== undefined) {
    const prevSum = dataRows.reduce((acc, r) => {
      const v = r[prevNumericColIdx];
      const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/[^0-9.-]+/g, '')) || 0;
      return acc + n;
    }, 0);
    if (prevSum !== 0) {
      deltaPercent = Math.round(((sum - prevSum) / Math.abs(prevSum)) * 1000) / 10;
      trend = deltaPercent > 0 ? 'up' : deltaPercent < 0 ? 'down' : 'neutral';
    }
  }

  const formatNum = (n: number) => {
    if (Number.isInteger(n)) return n.toLocaleString('id-ID');
    return n.toLocaleString('id-ID', { minimumFractionDigits: 1, maximumFractionDigits: 2 });
  };

  const metrics: DataStoryMetric[] = [
    {
      label: `Total ${focusHeader}`,
      value: formatNum(sum),
      changePercent: deltaPercent,
      trend,
    },
    {
      label: `Rata-rata ${focusHeader}`,
      value: formatNum(avg),
      trend: 'neutral',
    },
  ];

  if (maxEntry) {
    metrics.push({
      label: `Tertinggi (${maxEntry.label})`,
      value: formatNum(max),
      trend: 'up',
    });
  }

  if (minEntry && minEntry.label !== maxEntry?.label) {
    metrics.push({
      label: `Terendah (${minEntry.label})`,
      value: formatNum(min),
      trend: 'down',
    });
  }

  let headline = '';
  if (deltaPercent !== undefined && deltaPercent !== 0) {
    const dir = deltaPercent > 0 ? `Pertumbuhan +${deltaPercent}%` : `Penurunan ${deltaPercent}%`;
    headline = `${focusHeader} Mencapai ${formatNum(sum)} (${dir}) dengan Kontributor Utama ${maxEntry?.label || 'Teratas'}`;
  } else {
    headline = `Total ${focusHeader} Mencapai ${formatNum(sum)} Dipimpin oleh ${maxEntry?.label || 'Kontributor Utama'}`;
  }

  const keyFindings: string[] = [
    `Total capaian ${focusHeader} adalah ${formatNum(sum)} dengan nilai rata-rata ${formatNum(avg)} per entitas.`,
    `Kontributor terbesar diraih oleh ${maxEntry?.label || 'entitas utama'} sebesar ${formatNum(max)} (${Math.round((max / (sum || 1)) * 100)}% dari total).`,
  ];

  if (deltaPercent !== undefined && deltaPercent !== 0) {
    const growthDesc = deltaPercent > 0 ? 'pertumbuhan positif' : 'kontraksi';
    keyFindings.push(`Terjadi ${growthDesc} sebesar ${Math.abs(deltaPercent)}% dibandingkan kolom periode sebelumnya.`);
  }

  if (minEntry && minEntry.label !== maxEntry?.label) {
    keyFindings.push(`Capaian terendah berada pada ${minEntry.label} sebesar ${formatNum(min)}.`);
  }

  const risksOrAnomalies: string[] = [];
  if (minEntry && min < avg * 0.5 && entries.length > 2) {
    risksOrAnomalies.push(`Capaian ${minEntry.label} (${formatNum(min)}) berada lebih dari 50% di bawah rata-rata grup.`);
  }
  if (deltaPercent !== undefined && deltaPercent < -10) {
    risksOrAnomalies.push(`Penurunan tajam ${deltaPercent}% pada ${focusHeader} mengindikasikan deviasi target yang membutuhkan perhatian khusus.`);
  }

  const recommendations: string[] = [];
  if (options?.includeRecommendations !== false) {
    if (maxEntry) {
      recommendations.push(`Pertahankan strategi dan alokasi sumber daya pada ${maxEntry.label} sebagai penopang utama portofolio.`);
    }
    if (minEntry && minEntry.label !== maxEntry?.label) {
      recommendations.push(`Lakukan tinjauan operasional dan program akselerasi perbaikan untuk ${minEntry.label}.`);
    }
    if (deltaPercent !== undefined && deltaPercent > 0) {
      recommendations.push(`Manfaatkan tren pertumbuhan positif ${focusHeader} untuk memperluas skala target periode mendatang.`);
    } else {
      recommendations.push(`Analisis hambatan utama penurunan performa ${focusHeader} dan siapkan rencana pemulihan terukur.`);
    }
  }

  return {
    headline,
    keyFindings,
    metrics,
    risksOrAnomalies: risksOrAnomalies.length > 0 ? risksOrAnomalies : undefined,
    recommendations,
  };
}
