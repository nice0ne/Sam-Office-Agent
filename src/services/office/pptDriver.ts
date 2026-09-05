import { IDocumentDriver } from './types';

declare const Office: any;

export class PPTDriver implements Partial<IDocumentDriver> {
  hostType = 'PowerPoint';

  async getSlideContext() {
    return {
      slideNumber: 1,
      title: 'Active Slide',
      textContent: 'Slide content',
    };
  }

  async addSlide(_layout: string) {
    return { slideNumber: 1 };
  }

  async insertSlideContent(title: string, bullets: string[]) {
    // In PowerPoint Office.js Web Add-in, text insertion uses Office.context.document.setSelectedDataAsync
    if (typeof Office !== 'undefined' && Office.context?.document) {
      const content = `${title}\n\n` + bullets.map((b) => `• ${b}`).join('\n');
      Office.context.document.setSelectedDataAsync(content, { coercionType: Office.CoercionType.Text });
    }
  }

  async setSpeakerNotes(_notes: string) {
    // Notes handler
  }
}
