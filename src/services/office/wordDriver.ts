import { IDocumentDriver } from './types';

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
}
