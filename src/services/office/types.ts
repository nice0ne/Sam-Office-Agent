export interface IDocumentDriver {
  hostType: string;
  // Excel operations
  readActiveRange(): Promise<{ address: string; values: any[][]; formulas: string[][] }>;
  readActiveSheetData?(range?: string): Promise<{
    sheetName: string;
    address: string;
    rowCount: number;
    columnCount: number;
    values: any[][];
    formulas: string[][];
  }>;
  writeCells(range: string, values?: any[][], formulas?: string[][]): Promise<void>;
  formatRange(range: string, styles: Record<string, any>): Promise<void>;
  createChart(type: string, dataRange: string, title?: string): Promise<void>;
  addWorksheet?(sheetName: string): Promise<{ sheetName: string }>;
  autoFitColumns?(range?: string): Promise<void>;
  sortRange?(
    range: string,
    columnIndex: number,
    ascending?: boolean,
    hasHeaders?: boolean,
    enableAutoFilter?: boolean
  ): Promise<void>;
  cleanData?(options?: CleanDataOptions): Promise<CleanDataResult>;
  applyConditionalFormatting?(options: ConditionalFormattingOptions): Promise<{ success: boolean; rule: string }>;
  // Word operations
  getWordOutline(): Promise<string>;
  getWordContext?(): Promise<{ bodyText: string; selectedText: string }>;
  insertContent(position: 'start' | 'end' | 'cursor', text: string, type: string): Promise<void>;
  replaceSelection(newText: string): Promise<void>;
  insertTable(rows: number, cols: number, data?: string[][]): Promise<void>;
  insertPageBreak?(breakType?: 'page' | 'section'): Promise<void>;
  findAndReplace?(findText: string, replaceText: string, matchCase?: boolean): Promise<{ count: number }>;
  insertPictureBase64?(base64Image: string): Promise<void>;
  generateStructuredDoc?(options: StructuredDocOptions): Promise<{ success: boolean; message: string }>;
  polishDocumentText?(options: PolishTextOptions): Promise<{ success: boolean; polishedText: string }>;
  // PowerPoint operations
  getSlideContext(): Promise<{
    slideNumber: number;
    title: string;
    textContent: string;
    slides?: Array<{ slideIndex: number; title: string; textContent: string; notes?: string }>;
  }>;
  readSlideData?(
    slideNumber?: number,
    allSlides?: boolean
  ): Promise<{
    activeSlideIndex: number;
    totalSlides: number;
    slides: Array<{ slideIndex: number; title: string; textContent: string; notes?: string }>;
  }>;
  addSlide(layout: string): Promise<{ slideNumber: number }>;
  insertSlideContent(title: string, bullets: string[]): Promise<void>;
  setSpeakerNotes(notes: string): Promise<void>;
  createSlide?(title: string, bullets: string[], notes?: string, layout?: string): Promise<{ slideNumber: number }>;
  createPresentationDeck?(slides: Array<{ title: string; bullets: string[]; notes?: string; layout?: string }>): Promise<{ createdCount: number }>;
  generateThemedDeck?(options: ThemedDeckOptions): Promise<{ success: boolean; createdCount: number }>;
}

export interface CleanDataOptions {
  range?: string;
  removeDuplicates?: boolean;
  trimWhitespace?: boolean;
  fillEmptyValues?: string | number;
}

export interface CleanDataResult {
  cleanedRows: number;
  removedDuplicatesCount: number;
  trimmedCellsCount: number;
  filledCellsCount: number;
}

export interface ConditionalFormattingOptions {
  range: string;
  type: 'color_scale' | 'data_bar' | 'highlight_threshold';
  color?: string;
  thresholdValue?: number;
}

export interface StructuredDocOptions {
  templateType: 'SOP' | 'MoM' | 'SPK' | 'PRD' | 'FormalMemo';
  title: string;
  sections: Array<{
    heading: string;
    content?: string;
    bullets?: string[];
    table?: { headers: string[]; rows: string[][] };
  }>;
  author?: string;
}

export interface PolishTextOptions {
  scope: 'selection' | 'document';
  tone: 'formal_indonesia' | 'executive_english' | 'concise';
  instruction?: string;
  customText?: string;
}

export interface ThemedDeckOptions {
  topic: string;
  theme?: 'corporate_blue' | 'emerald_executive' | 'modern_dark' | 'minimalist_clean';
  slides: Array<{
    title: string;
    layout?: 'title_cover' | 'split_comparison' | 'bullet_points' | 'metric_highlights';
    content?: string[];
    metrics?: Array<{ label: string; value: string }>;
    notes?: string;
  }>;
}

