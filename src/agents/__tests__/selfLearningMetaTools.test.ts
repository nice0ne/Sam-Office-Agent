import { describe, it, expect, beforeEach } from 'vitest';
import {
  clearCustomTools,
  getCustomTools,
  saveCustomTool,
  deleteCustomTool,
  incrementToolUsage,
} from '../../services/storage/customToolStorage';
import {
  executeDynamicOfficeScript,
  sanitizeScriptCode,
} from '../../services/office/scriptExecutor';
import {
  executeMetaOrCustomTool,
  getLearnedAndMetaTools,
  EXECUTE_OFFICE_SCRIPT_TOOL,
  SAVE_CUSTOM_TOOL,
} from '../metaTools';
import { ExcelAgent } from '../excel/excelAgent';
import { WordAgent } from '../word/wordAgent';
import { PPTAgent } from '../powerpoint/pptAgent';
import { SamCoordinator } from '../coordinator/samCoordinator';

describe('Self-Learning & Dynamic Meta-Tooling Architecture', () => {
  beforeEach(() => {
    clearCustomTools();
  });

  describe('Custom Tool Storage & Persistence', () => {
    it('saves a new custom tool and normalizes snake_case name', () => {
      const tool = saveCustomTool({
        name: 'Calculate Pearson Correlation!',
        description: 'Menghitung korelasi Pearson 2 kolom',
        hostType: 'Excel',
        parameters: {
          type: 'object',
          properties: {
            rangeA: { type: 'string', description: 'Kolom 1' },
            rangeB: { type: 'string', description: 'Kolom 2' },
          },
        },
        code: 'return 0.85;',
      });

      expect(tool.name).toBe('calculate_pearson_correlation_');
      expect(tool.hostType).toBe('Excel');
      expect(tool.usageCount).toBe(0);

      const all = getCustomTools('Excel');
      expect(all.length).toBe(1);
      expect(all[0].name).toBe('calculate_pearson_correlation_');
    });

    it('filters custom tools by hostType and includes Universal tools', () => {
      saveCustomTool({
        name: 'excel_tool',
        description: 'Excel only',
        hostType: 'Excel',
        parameters: { type: 'object', properties: {} },
        code: 'return 1;',
      });

      saveCustomTool({
        name: 'word_tool',
        description: 'Word only',
        hostType: 'Word',
        parameters: { type: 'object', properties: {} },
        code: 'return 2;',
      });

      saveCustomTool({
        name: 'universal_tool',
        description: 'Universal',
        hostType: 'Universal',
        parameters: { type: 'object', properties: {} },
        code: 'return 3;',
      });

      const excelTools = getCustomTools('Excel');
      expect(excelTools.some(t => t.name === 'excel_tool')).toBe(true);
      expect(excelTools.some(t => t.name === 'universal_tool')).toBe(true);
      expect(excelTools.some(t => t.name === 'word_tool')).toBe(false);

      const wordTools = getCustomTools('Word');
      expect(wordTools.some(t => t.name === 'word_tool')).toBe(true);
      expect(wordTools.some(t => t.name === 'universal_tool')).toBe(true);
      expect(wordTools.some(t => t.name === 'excel_tool')).toBe(false);
    });

    it('increments tool usage count and deletes custom tool', () => {
      const tool = saveCustomTool({
        name: 'counter_tool',
        description: 'Counter',
        hostType: 'Excel',
        parameters: { type: 'object', properties: {} },
        code: 'return true;',
      });

      incrementToolUsage(tool.id);
      incrementToolUsage(tool.name);

      let tools = getCustomTools('Excel');
      expect(tools[0].usageCount).toBe(2);

      const deleted = deleteCustomTool(tool.name);
      expect(deleted).toBe(true);

      tools = getCustomTools('Excel');
      expect(tools.length).toBe(0);
    });
  });

  describe('Script Sanitizer & Dynamic Executor', () => {
    it('sanitizes markdown code blocks from LLM output', () => {
      const rawWithJs = '```javascript\nconst a = 10;\nreturn a * 2;\n```';
      expect(sanitizeScriptCode(rawWithJs)).toBe('const a = 10;\nreturn a * 2;');

      const rawWithTs = '```typescript\nreturn "hello";\n```';
      expect(sanitizeScriptCode(rawWithTs)).toBe('return "hello";');

      const plain = 'return 42;';
      expect(sanitizeScriptCode(plain)).toBe('return 42;');
    });

    it('executes dynamic script in mock environment and passes parameters', async () => {
      const code = `
        const { multiplier } = params;
        return 10 * (multiplier || 1);
      `;
      const res = await executeDynamicOfficeScript('Excel', code, { multiplier: 5 });
      expect(res.success).toBe(true);
      expect(res.result).toBe(50);
      expect(res.executionTimeMs).toBeGreaterThanOrEqual(0);
    });

    it('automatically invokes function main if the script wrapped execution in main', async () => {
      const codeWrapped = `
        async function main(context) {
          return { copied: true, count: 42 };
        }
      `;
      const res = await executeDynamicOfficeScript('Excel', codeWrapped);
      expect(res.success).toBe(true);
      expect(res.result).toEqual({ copied: true, count: 42 });
    });

    it('strips TypeScript type annotations such as context: Excel.RequestContext without error', async () => {
      const tsCode = `
        async function main(context: Excel.RequestContext): Promise<any> {
          const count: number = 99;
          return { normalized: true, count };
        }
      `;
      const res = await executeDynamicOfficeScript('Excel', tsCode);
      expect(res.success).toBe(true);
      expect(res.result).toEqual({ normalized: true, count: 99 });
    });

    it('catches syntax and runtime errors cleanly for self-debugging', async () => {
      const invalidCode = `
        throw new Error("Formula context column mismatch");
      `;
      const res = await executeDynamicOfficeScript('Excel', invalidCode);
      expect(res.success).toBe(false);
      expect(res.error).toContain('Formula context column mismatch');
    });
  });

  describe('Meta-Tools & Dynamic Dispatcher', () => {
    it('exposes execute_office_script and save_custom_tool in getLearnedAndMetaTools', () => {
      const tools = getLearnedAndMetaTools('Excel');
      const toolNames = tools.map(t => t.name);
      expect(toolNames).toContain(EXECUTE_OFFICE_SCRIPT_TOOL.name);
      expect(toolNames).toContain(SAVE_CUSTOM_TOOL.name);
    });

    it('handles execute_office_script tool call directly', async () => {
      const toolCall = {
        id: 'call-dyn-1',
        name: 'execute_office_script',
        arguments: {
          script: 'return params.a + params.b;',
          explanation: 'Menjumlahkan parameter',
          params: { a: 15, b: 25 },
        },
        status: 'pending' as const,
      };

      const res = await executeMetaOrCustomTool('Excel', toolCall, { host: 'Excel' });
      expect(res.handled).toBe(true);
      expect(res.result?.success).toBe(true);
      expect(res.result?.result).toBe(40);
    });

    it('handles save_custom_tool and persists learned tool for subsequent turns', async () => {
      const saveCall = {
        id: 'call-save-1',
        name: 'save_custom_tool',
        arguments: {
          name: 'highlight_top_performers',
          description: 'Memberi warna emas pada 3 nilai teratas',
          parameters: {
            type: 'object',
            properties: {
              range: { type: 'string', description: 'Range data' },
            },
          },
          code: 'return `Highlighting top performers in ${params.range}`;',
        },
        status: 'pending' as const,
      };

      const saveRes = await executeMetaOrCustomTool('Excel', saveCall, { host: 'Excel' });
      expect(saveRes.handled).toBe(true);
      expect(saveRes.result?.success).toBe(true);
      expect(saveRes.result?.result).toContain('Keahlian baru "highlight_top_performers" berhasil dipelajari');

      // Check that it is now listed in getLearnedAndMetaTools
      const updatedTools = getLearnedAndMetaTools('Excel');
      expect(updatedTools.some(t => t.name === 'highlight_top_performers')).toBe(true);

      // Now execute the newly learned tool!
      const learnedCall = {
        id: 'call-exec-learned-1',
        name: 'highlight_top_performers',
        arguments: { range: 'C2:C50' },
        status: 'pending' as const,
      };

      const execRes = await executeMetaOrCustomTool('Excel', learnedCall, { host: 'Excel' });
      expect(execRes.handled).toBe(true);
      expect(execRes.result?.success).toBe(true);
      expect(execRes.result?.result).toBe('Highlighting top performers in C2:C50');
    });

    it('returns handled: false for standard built-in tools', async () => {
      const standardCall = {
        id: 'call-std-1',
        name: 'write_cells',
        arguments: { range: 'A1', values: [[1]] },
        status: 'pending' as const,
      };

      const res = await executeMetaOrCustomTool('Excel', standardCall, { host: 'Excel' });
      expect(res.handled).toBe(false);
    });
  });

  describe('Specialists Integration & Coordinator Prompt', () => {
    it('ExcelAgent exposes meta-tools and executes execute_office_script', async () => {
      const agent = new ExcelAgent();
      const toolNames = agent.getTools().map(t => t.name);
      expect(toolNames).toContain('execute_office_script');
      expect(toolNames).toContain('save_custom_tool');

      const res = await agent.executeTool(
        {
          id: 'call-dyn-excel',
          name: 'execute_office_script',
          arguments: {
            script: 'return "Excel script custom executed!";',
            explanation: 'Test',
          },
          status: 'pending',
        },
        { host: 'Excel' }
      );

      expect(res.success).toBe(true);
      expect(res.result).toBe('Excel script custom executed!');
    });

    it('WordAgent and PPTAgent expose meta-tools', () => {
      const wordAgent = new WordAgent();
      const wordTools = wordAgent.getTools().map(t => t.name);
      expect(wordTools).toContain('execute_office_script');
      expect(wordTools).toContain('save_custom_tool');

      const pptAgent = new PPTAgent();
      const pptTools = pptAgent.getTools().map(t => t.name);
      expect(pptTools).toContain('execute_office_script');
      expect(pptTools).toContain('save_custom_tool');
    });

    it('SamCoordinator prompt includes rule 4 for Self-Learning & Meta-Tooling', () => {
      const coordinator = new SamCoordinator();
      const prompt = coordinator.buildSystemPrompt('Excel', { host: 'Excel' });
      expect(prompt).toContain('KEMAMPUAN SELF-LEARNING & SCRIPT DINAMIS (META-TOOLING)');
      expect(prompt).toContain('execute_office_script');
      expect(prompt).toContain('save_custom_tool');
    });
  });
});
