import { HostType, ToolCall, ToolDefinition } from '../types';
import { executeDynamicOfficeScript } from '../services/office/scriptExecutor';
import {
  getCustomTools,
  incrementToolUsage,
  saveCustomTool,
} from '../services/storage/customToolStorage';
import { AgentContext } from './types';

export const EXECUTE_OFFICE_SCRIPT_TOOL: ToolDefinition = {
  name: 'execute_office_script',
  description:
    'Mengeksekusi kode script Office.js dinamis langsung ke aplikasi Office yang sedang dibuka saat pengguna meminta manipulasi tingkat lanjut atau aksi khusus yang belum didukung oleh tool bawaan standar. PENTING: Script harus benar-benar melakukan perubahan dokumen (misal membuat sheet baru, menyalin data ke range, atau memformat), bukan hanya membaca data!',
  parameters: {
    type: 'object',
    properties: {
      script: {
        type: 'string',
        description:
          'Kode fungsi async JavaScript Office.js. Gunakan objek `context` (misal context.workbook / context.document / context.presentation) dan panggil await context.sync(). Parameter dapat diakses via objek `params`.',
      },
      explanation: {
        type: 'string',
        description: 'Penjelasan singkat apa yang dikerjakan oleh script ini bagi pengguna.',
      },
      params: {
        type: 'object',
        description: 'Argumen/parameter opsional yang diteruskan ke script.',
      },
    },
    required: ['script', 'explanation'],
  },
};

export const SAVE_CUSTOM_TOOL: ToolDefinition = {
  name: 'save_custom_tool',
  description:
    'Menyimpan kemampuan baru (script Office.js) menjadi tool permanen ke registry agen (Self-Learning), sehingga tool ini otomatis dapat digunakan kembali di masa depan tanpa menulis ulang kode.',
  parameters: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description:
          'Nama tool baru dalam format snake_case (misal: "highlight_values_below_threshold" atau "insert_diagonal_watermark").',
      },
      description: {
        type: 'string',
        description: 'Deskripsi lengkap apa yang dilakukan tool ini dan kapan harus dipanggil.',
      },
      parameters: {
        type: 'object',
        description: 'Skema parameter tool (JSON Schema object dengan properties).',
      },
      code: {
        type: 'string',
        description:
          'Kode script Office.js yang dapat digunakan kembali, menerima parameter dari objek `params`.',
      },
    },
    required: ['name', 'description', 'code'],
  },
};

export function getLearnedAndMetaTools(host: HostType): ToolDefinition[] {
  const custom = getCustomTools(host);
  const learnedDefinitions: ToolDefinition[] = custom.map(t => ({
    name: t.name,
    description: `[Tool Kustom Dipelajari] ${t.description}`,
    parameters: t.parameters || { type: 'object', properties: {} },
  }));

  return [EXECUTE_OFFICE_SCRIPT_TOOL, SAVE_CUSTOM_TOOL, ...learnedDefinitions];
}

export async function executeMetaOrCustomTool(
  host: HostType,
  toolCall: ToolCall,
  _context: AgentContext
): Promise<{ handled: boolean; result?: { success: boolean; result?: any; error?: string } }> {
  if (toolCall.name === 'execute_office_script') {
    const { script, params } = toolCall.arguments || {};
    if (!script) {
      return {
        handled: true,
        result: { success: false, error: 'Parameter script tidak boleh kosong.' },
      };
    }
    const execRes = await executeDynamicOfficeScript(host, script, params || {});
    return {
      handled: true,
      result: {
        success: execRes.success,
        result: execRes.result,
        error: execRes.error,
      },
    };
  }

  if (toolCall.name === 'save_custom_tool') {
    const { name, description, parameters, code } = toolCall.arguments || {};
    if (!name || !description || !code) {
      return {
        handled: true,
        result: {
          success: false,
          error: 'Parameter name, description, dan code wajib diisi untuk menyimpan tool baru.',
        },
      };
    }

    const saved = saveCustomTool({
      name,
      description,
      hostType: host,
      parameters: parameters || { type: 'object', properties: {} },
      code,
    });

    return {
      handled: true,
      result: {
        success: true,
        result: `Keahlian baru "${saved.name}" berhasil dipelajari dan disimpan ke registry! Tool ini kini menjadi tool permanen yang siap digunakan kembali.`,
      },
    };
  }

  // Check if toolCall corresponds to a learned custom tool
  const learnedTools = getCustomTools(host);
  const foundTool = learnedTools.find(
    t => t.name.toLowerCase() === toolCall.name.toLowerCase()
  );

  if (foundTool) {
    incrementToolUsage(foundTool.id);
    const execRes = await executeDynamicOfficeScript(
      host,
      foundTool.code,
      toolCall.arguments || {}
    );
    return {
      handled: true,
      result: {
        success: execRes.success,
        result: execRes.result,
        error: execRes.error,
      },
    };
  }

  return { handled: false };
}
