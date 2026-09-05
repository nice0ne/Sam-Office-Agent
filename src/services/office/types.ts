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
  // Word operations
  getWordOutline(): Promise<string>;
  insertContent(position: 'start' | 'end' | 'cursor', text: string, type: string): Promise<void>;
  replaceSelection(newText: string): Promise<void>;
  insertTable(rows: number, cols: number, data?: string[][]): Promise<void>;
  // PowerPoint operations
  getSlideContext(): Promise<{ slideNumber: number; title: string; textContent: string }>;
  addSlide(layout: string): Promise<{ slideNumber: number }>;
  insertSlideContent(title: string, bullets: string[]): Promise<void>;
  setSpeakerNotes(notes: string): Promise<void>;
}
