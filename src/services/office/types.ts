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
  auditSheetData?(options?: { range?: string }): Promise<SheetAuditResult>;
  generateDataStory?(options?: DataStoryOptions): Promise<DataStoryResult>;
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
  reviewComplianceClauses?(options?: ComplianceReviewOptions): Promise<ComplianceReviewResult>;
  applyCorporateStyle?(options?: CorporateStyleOptions): Promise<CorporateStyleResult>;
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

export interface ComplianceClause {
  category: 'payment_terms' | 'sla_performance' | 'liability_indemnity' | 'termination_cancellation' | 'confidentiality_nda' | 'dispute_resolution' | 'force_majeure' | 'general';
  excerpt: string;
  status: 'compliant' | 'warning' | 'high_risk' | 'missing';
  analysis: string;
  recommendation?: string;
}

export interface ComplianceReviewOptions {
  scope?: 'selection' | 'document';
  contractType?: 'vendor_service' | 'employment' | 'nda' | 'procurement' | 'general';
  strictness?: 'standard' | 'strict';
}

export interface ComplianceReviewResult {
  contractType: string;
  overallRiskLevel: 'low' | 'medium' | 'high';
  clausesReviewedCount: number;
  identifiedClauses: ComplianceClause[];
  missingCriticalClauses: string[];
  executiveSummary: string;
  actionableRecommendations: string[];
}

export interface CorporateStyleOptions {
  theme?: 'corporate_navy' | 'executive_emerald' | 'modern_minimalist' | 'official_government';
  scope?: 'selection' | 'document';
  fontFamily?: string;
  applyHeadingHierarchy?: boolean;
}

export interface CorporateStyleResult {
  appliedTheme: string;
  fontFamily: string;
  styledParagraphsCount: number;
  headingsCount: number;
  message: string;
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

export type IOfficeDriver = IDocumentDriver;

export interface FormulaIssue {
  address: string;
  type: 'formula_error' | 'inconsistent_formula' | 'hardcoded_override' | 'suspicious_blank' | 'statistical_outlier';
  severity: 'critical' | 'warning' | 'info';
  currentValue?: any;
  formula?: string;
  expectedPattern?: string;
  suggestion: string;
}

export interface SheetAuditResult {
  sheetName: string;
  totalCellsAudited: number;
  totalErrorsFound: number;
  criticalIssues: FormulaIssue[];
  warnings: FormulaIssue[];
  summary: string;
}

export interface DataStoryOptions {
  range?: string;
  focusMetric?: string;
  includeRecommendations?: boolean;
}

export interface DataStoryMetric {
  label: string;
  value: string | number;
  changePercent?: number;
  trend?: 'up' | 'down' | 'neutral';
}

export interface DataStoryResult {
  headline: string;
  keyFindings: string[];
  metrics: DataStoryMetric[];
  risksOrAnomalies?: string[];
  recommendations: string[];
}

export interface PendingActionProposal {
  id: string;
  actionType: 'modify_cells' | 'clean_data' | 'format_cells';
  description: string;
  affectedCellsCount: number;
  targetRange?: string;
  payload: any;
}

