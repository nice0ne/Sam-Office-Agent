import { HostType } from '../../types';

declare const Excel: any;
declare const Word: any;
declare const PowerPoint: any;
declare const Office: any;

export interface ScriptExecutionResult {
  success: boolean;
  result?: any;
  error?: string;
  executionTimeMs?: number;
}

/**
 * Strip common TypeScript syntax (type annotations, interfaces, assertions)
 * so that scripts run cleanly in raw JavaScript runtime engines.
 */
export function stripTypeScriptTypes(rawCode: string): string {
  let code = rawCode;

  // 1. Remove interfaces and type aliases
  code = code.replace(/\binterface\s+[a-zA-Z0-9_$]+(?:\s*<[^>]+>)?\s*\{[\s\S]*?\}/g, '');
  code = code.replace(/\btype\s+[a-zA-Z0-9_$]+(?:\s*<[^>]+>)?\s*=[^;]+;/g, '');

  // 2. Remove return type annotations on functions, e.g. function main(...): Promise<void> { or ): void {
  code = code.replace(/\)\s*:\s*(?:Promise\s*<[^>]+>|[a-zA-Z0-9_$.<>]+(?:\[\])?)\s*\{/g, ') {');

  // 3. Remove parameter type annotations in function signatures, e.g. main(context: Excel.RequestContext, params: any)
  code = code.replace(/(function\s*(?:[a-zA-Z0-9_$]+)?\s*\([^)]*\)|\((?:[a-zA-Z0-9_$,\s:]+)\)\s*=>)/g, match => {
    return match.replace(/([a-zA-Z0-9_$]+)\s*:\s*[a-zA-Z0-9_$.<>]+(?:\[\])?/g, '$1');
  });

  // 4. Remove variable type annotations, e.g. const range: Excel.Range = ...
  code = code.replace(/\b(const|let|var)\s+([a-zA-Z0-9_$]+)\s*:\s*(?:[a-zA-Z0-9_$.<>]+(?:\[\])?)\s*=/g, '$1 $2 =');

  // 5. Remove type assertions, e.g. val as string or range as Excel.Range
  code = code.replace(/\s+as\s+[a-zA-Z0-9_$.<>]+(?:\[\])?/g, '');

  return code;
}

/**
 * Clean up script code from markdown code fences, TypeScript annotations, and auto-invoke main if wrapped.
 */
export function sanitizeScriptCode(rawCode: string): string {
  let code = rawCode.trim();
  // Remove markdown code blocks like ```javascript or ```typescript or ```js
  if (code.startsWith('```')) {
    code = code.replace(/^```[a-zA-Z0-9_-]*\n?/, '').replace(/\n?```$/, '');
  }
  code = code.trim();

  // Strip TypeScript annotations so raw JS runtime does not throw "Unexpected token ':'"
  code = stripTypeScriptTypes(code);

  // If the script defines function main(...) or async function main(...) without calling it:
  if (/(\basync\s+)?function\s+main\s*\(/.test(code)) {
    // Check if main(...) call already exists after the declaration
    const afterMainDecl = code.slice(code.indexOf('function main') + 13);
    const hasCall = /(?:return\s+)?(?:await\s+)?main\s*\(/.test(afterMainDecl);
    if (!hasCall) {
      code += '\nreturn await main(context, params, hostApi);';
    }
  }

  return code;
}

/**
 * Execute dynamic Office.js script safely within the target host context.
 */
export async function executeDynamicOfficeScript(
  host: HostType,
  scriptCode: string,
  params: Record<string, any> = {}
): Promise<ScriptExecutionResult> {
  const startTime = Date.now();
  const cleanCode = sanitizeScriptCode(scriptCode);

  if (!cleanCode) {
    return { success: false, error: 'Kode script kosong.' };
  }

  try {
    const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
    const runner = new AsyncFunction('context', 'params', 'hostApi', cleanCode);

    if (host === 'Excel' && typeof Excel !== 'undefined' && Excel.run) {
      const execResult = await Excel.run(async (context: any) => {
        const out = await runner(context, params, Excel);
        await context.sync();
        return out;
      });
      return {
        success: true,
        result: execResult !== undefined ? execResult : 'Script Excel berhasil dijalankan.',
        executionTimeMs: Date.now() - startTime,
      };
    }

    if (host === 'Word' && typeof Word !== 'undefined' && Word.run) {
      const execResult = await Word.run(async (context: any) => {
        const out = await runner(context, params, Word);
        await context.sync();
        return out;
      });
      return {
        success: true,
        result: execResult !== undefined ? execResult : 'Script Word berhasil dijalankan.',
        executionTimeMs: Date.now() - startTime,
      };
    }

    if (host === 'PowerPoint' && typeof PowerPoint !== 'undefined' && PowerPoint.run) {
      const execResult = await PowerPoint.run(async (context: any) => {
        const out = await runner(context, params, PowerPoint);
        await context.sync();
        return out;
      });
      return {
        success: true,
        result: execResult !== undefined ? execResult : 'Script PowerPoint berhasil dijalankan.',
        executionTimeMs: Date.now() - startTime,
      };
    }

    // Fallback for Mock / BrowserDev environment
    const mockContext = {
      workbook: {
        worksheets: {
          getActiveWorksheet: () => ({
            name: 'MockSheet',
            getRange: () => ({ values: [], format: { autofitColumns: () => {} } }),
          }),
        },
      },
      sync: async () => {},
    };

    const fallbackResult = await runner(mockContext, params, {});
    return {
      success: true,
      result: fallbackResult !== undefined ? fallbackResult : 'Script berhasil dijalankan di lingkungan pengujian.',
      executionTimeMs: Date.now() - startTime,
    };
  } catch (err: any) {
    console.error('Error executing dynamic office script:', err);
    return {
      success: false,
      error: `Kesalahan saat menjalankan script Office: ${err.message || String(err)}`,
      executionTimeMs: Date.now() - startTime,
    };
  }
}
