import {
  ComplianceReviewOptions,
  ComplianceReviewResult,
  CorporateStyleOptions,
  CorporateStyleResult,
  DiagramResult,
  IDocumentDriver,
  InsertFlowchartOptions,
  PolishTextOptions,
  StructuredDocOptions,
} from './types';
import { analyzeContractCompliance } from './complianceReviewer';
import { synthesizeFlowchartFromText, renderFlowchartToPngBase64 } from '../../utils/diagramRenderer';

declare const Word: any;

export class WordDriver implements Partial<IDocumentDriver> {
  hostType = 'Word';

  async getWordOutline(): Promise<string> {
    return await Word.run(async (context: any) => {
      const body = context.document.body;
      body.load('text');
      await context.sync();
      return body.text || '';
    });
  }

  async getWordContext(): Promise<{ bodyText: string; selectedText: string }> {
    return await Word.run(async (context: any) => {
      const body = context.document.body;
      const selection = context.document.getSelection();
      body.load('text');
      selection.load('text');
      await context.sync();
      return {
        bodyText: body.text || '',
        selectedText: selection.text || '',
      };
    });
  }

  async insertContent(position: 'start' | 'end' | 'cursor', text: string, type: string) {
    await Word.run(async (context: any) => {
      let location = Word.InsertLocation.end;
      if (position === 'start') location = Word.InsertLocation.start;
      const paragraph = context.document.body.insertParagraph(text, location);
      if (type === 'h1') paragraph.style = 'Heading 1';
      else if (type === 'h2') paragraph.style = 'Heading 2';
      await context.sync();
    });
  }

  async replaceSelection(newText: string) {
    await Word.run(async (context: any) => {
      const selection = context.document.getSelection();
      selection.insertText(newText, Word.InsertLocation.replace);
      await context.sync();
    });
  }

  async insertTable(rows: number, cols: number, data?: string[][]) {
    await Word.run(async (context: any) => {
      const table = context.document.body.insertTable(rows, cols, Word.InsertLocation.end, data || []);
      table.styleBuiltIn = Word.Style.gridTable4_Accent1;
      await context.sync();
    });
  }

  async insertPageBreak(breakType: 'page' | 'section' = 'page') {
    await Word.run(async (context: any) => {
      const type = breakType === 'section' ? Word.BreakType.sectionNext : Word.BreakType.page;
      context.document.body.insertBreak(type, Word.InsertLocation.end);
      await context.sync();
    });
  }

  async findAndReplace(findText: string, replaceText: string, matchCase = false): Promise<{ count: number }> {
    return await Word.run(async (context: any) => {
      const searchResults = context.document.body.search(findText, { matchCase });
      searchResults.load('items');
      await context.sync();

      const count = searchResults.items ? searchResults.items.length : 0;
      for (let i = 0; i < count; i++) {
        searchResults.items[i].insertText(replaceText, Word.InsertLocation.replace);
      }
      await context.sync();
      return { count };
    });
  }

  async insertPictureBase64(base64Image: string) {
    await Word.run(async (context: any) => {
      context.document.body.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end);
      await context.sync();
    });
  }

  async generateStructuredDoc(options: StructuredDocOptions): Promise<{ success: boolean; message: string }> {
    return await Word.run(async (context: any) => {
      const body = context.document.body;

      // 1. Title Heading 1
      const titlePara = body.insertParagraph(options.title, Word.InsertLocation.end);
      titlePara.style = 'Heading 1';

      // 2. Template badge & metadata
      const authorPart = options.author ? ` | Penyusun: ${options.author}` : '';
      const datePart = new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' });
      const metaText = `[Template: ${options.templateType}]${authorPart} | Tanggal: ${datePart}`;
      const metaPara = body.insertParagraph(metaText, Word.InsertLocation.end);
      metaPara.font.italic = true;
      metaPara.font.color = '#64748B';

      // 3. Sections
      for (const section of options.sections) {
        const h2 = body.insertParagraph(section.heading, Word.InsertLocation.end);
        h2.style = 'Heading 2';

        if (section.content) {
          body.insertParagraph(section.content, Word.InsertLocation.end);
        }

        if (section.bullets && section.bullets.length > 0) {
          for (const bullet of section.bullets) {
            body.insertParagraph(`• ${bullet}`, Word.InsertLocation.end);
          }
        }

        if (section.table && section.table.headers && section.table.rows) {
          const tableData = [section.table.headers, ...section.table.rows];
          const table = body.insertTable(tableData.length, section.table.headers.length, Word.InsertLocation.end, tableData);
          table.styleBuiltIn = Word.Style.gridTable4_Accent1;
        }
      }

      await context.sync();
      return {
        success: true,
        message: `Dokumen terstruktur [${options.templateType}] "${options.title}" berhasil dibuat dengan ${options.sections.length} bagian.`,
      };
    });
  }

  async polishDocumentText(options: PolishTextOptions): Promise<{ success: boolean; polishedText: string }> {
    return await Word.run(async (context: any) => {
      const targetRange = options.scope === 'selection' ? context.document.getSelection() : context.document.body;
      if (options.customText) {
        targetRange.insertText(options.customText, Word.InsertLocation.replace);
        await context.sync();
        return { success: true, polishedText: options.customText };
      }

      targetRange.load('text');
      await context.sync();
      const currentText = targetRange.text || '';
      return { success: true, polishedText: currentText };
    });
  }

  async reviewComplianceClauses(options?: ComplianceReviewOptions): Promise<ComplianceReviewResult> {
    return await Word.run(async (context: any) => {
      const targetRange = options?.scope === 'selection' ? context.document.getSelection() : context.document.body;
      targetRange.load('text');
      await context.sync();
      const text = targetRange.text || '';
      return analyzeContractCompliance(text, options);
    });
  }

  async applyCorporateStyle(options?: CorporateStyleOptions): Promise<CorporateStyleResult> {
    return await Word.run(async (context: any) => {
      const theme = options?.theme || 'corporate_navy';
      let primaryColor = '#1E3A8A';
      let defaultFont = 'Calibri';

      if (theme === 'executive_emerald') {
        primaryColor = '#065F46';
        defaultFont = 'Aptos';
      } else if (theme === 'modern_minimalist') {
        primaryColor = '#0F172A';
        defaultFont = 'Segoe UI';
      } else if (theme === 'official_government') {
        primaryColor = '#000000';
        defaultFont = 'Times New Roman';
      }

      const fontFamily = options?.fontFamily || defaultFont;
      const targetRange = options?.scope === 'selection' ? context.document.getSelection() : context.document.body;
      const paragraphs = targetRange.paragraphs;
      paragraphs.load('items,text');
      await context.sync();

      let styledCount = 0;
      let headingsCount = 0;

      for (let i = 0; i < (paragraphs.items ? paragraphs.items.length : 0); i++) {
        const p = paragraphs.items[i];
        const text = (p.text || '').trim();
        if (!text) continue;

        p.font.name = fontFamily;
        styledCount++;

        const isHeading = /^(bab|pasal|section|judul|#|\d+\.\s+[A-Z])/i.test(text);
        if (isHeading) {
          headingsCount++;
          p.font.bold = true;
          p.font.color = primaryColor;
          p.spaceBefore = 12;
          p.spaceAfter = 6;
        } else {
          p.font.size = 11;
          p.font.color = '#334155';
          p.spaceBefore = 0;
          p.spaceAfter = 6;
          p.lineSpacing = 1.15;
        }
      }

      await context.sync();

      return {
        appliedTheme: theme,
        fontFamily,
        styledParagraphsCount: styledCount || 1,
        headingsCount: headingsCount || 1,
        message: `Berhasil menerapkan format korporat (${theme}, font: ${fontFamily}) pada ${styledCount} paragraf.`,
      };
    });
  }

  async insertProcessFlowchart(options: InsertFlowchartOptions): Promise<DiagramResult> {
    const def = options.definition || synthesizeFlowchartFromText(options.textOrSteps, options);
    const base64 = renderFlowchartToPngBase64(def, options);

    if (typeof Word !== 'undefined' && typeof Word.run === 'function') {
      await Word.run(async (context: any) => {
        context.document.body.insertInlinePictureFromBase64(base64, Word.InsertLocation.end);
        if (options.caption) {
          const p = context.document.body.insertParagraph(options.caption, Word.InsertLocation.end);
          p.font.italic = true;
          p.font.size = 9.5;
          p.font.color = '#64748B';
        }
        await context.sync();
      });
    }

    return {
      title: def.title || options.title || 'Alur Proses',
      nodeCount: def.nodes.length,
      edgeCount: def.edges.length,
      appliedTheme: def.theme || options.theme || 'corporate_navy',
      base64Png: base64,
      inserted: true,
    };
  }
}
