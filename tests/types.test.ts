import { describe, it, expect } from 'vitest';
import {
  HostType,
  isValidHost,
  ProviderConfig,
  ChatMessage,
  ToolDefinition,
  ToolCall,
  ActionQueueItem,
  AgentAction,
  ExecutionMode,
} from '../src/types';
import { AgentContext, AgentResponse, IAgent, ExecutionContext } from '../src/agents/types';
import {
  FormulaIssue,
  SheetAuditResult,
  DataStoryOptions,
  DataStoryMetric,
  DataStoryResult,
  PendingActionProposal,
  ComplianceClause,
  ComplianceReviewResult,
  CorporateStyleResult,
  DocToDeckSlide,
  DocToDeckOptions,
  DocToDeckResult,
  FlowchartNode,
  FlowchartEdge,
  FlowchartDefinition,
  DiagramResult,
} from '../src/services/office/types';

describe('HostType & Guards', () => {
  it('validates known Office hosts correctly', () => {
    const excelHost: HostType = 'Excel';
    expect(isValidHost(excelHost)).toBe(true);
    expect(isValidHost('Word')).toBe(true);
    expect(isValidHost('PowerPoint')).toBe(true);
    expect(isValidHost('BrowserDev')).toBe(true);
    expect(isValidHost('UnknownHost')).toBe(false);
  });
});

describe('Domain Models & Contracts', () => {
  it('instantiates valid ProviderConfig objects', () => {
    const config: ProviderConfig = {
      id: 'gemini',
      name: 'Google Gemini',
      apiKey: 'test-api-key',
      selectedModel: 'gemini-2.0-flash',
      enabled: true,
    };
    expect(config.id).toBe('gemini');
    expect(config.enabled).toBe(true);
  });

  it('instantiates valid ChatMessage with tool calls', () => {
    const toolCall: ToolCall = {
      id: 'call-1',
      name: 'writeRange',
      arguments: { range: 'A1', values: [['Hello']] },
      status: 'pending',
    };

    const msg: ChatMessage = {
      id: 'msg-1',
      role: 'assistant',
      content: 'Writing data to range',
      timestamp: Date.now(),
      toolCalls: [toolCall],
      status: 'done',
    };

    expect(msg.role).toBe('assistant');
    expect(msg.toolCalls?.length).toBe(1);
    expect(msg.toolCalls?.[0].status).toBe('pending');
  });

  it('supports ToolDefinition schema contract', () => {
    const tool: ToolDefinition = {
      name: 'formatRange',
      description: 'Format specified cells in active worksheet',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'Target range like A1:B10' },
          bold: { type: 'boolean', description: 'Bold text formatting' },
        },
        required: ['range'],
      },
    };
    expect(tool.name).toBe('formatRange');
    expect(tool.parameters.required).toContain('range');
  });

  it('supports ActionQueueItem and AgentAction alias', () => {
    const item: ActionQueueItem = {
      id: 'action-1',
      stepNumber: 1,
      title: 'Set Header',
      description: 'Format row 1 with bold and fill color',
      toolCall: {
        id: 'tc-1',
        name: 'formatHeader',
        arguments: { row: 1 },
        status: 'pending',
      },
      status: 'pending',
    };

    const agentAction: AgentAction = item;
    expect(agentAction.stepNumber).toBe(1);
    expect(agentAction.status).toBe('pending');

    const mode: ExecutionMode = 'copilot';
    expect(mode).toBe('copilot');
  });

  it('supports AgentContext and IAgent implementation', async () => {
    const context: AgentContext = {
      host: 'Excel',
      activeCellOrRange: 'B2',
      documentSummary: 'Sales report 2026',
    };

    const execContext: ExecutionContext = context;
    expect(execContext.host).toBe('Excel');

    class MockAgent implements IAgent {
      id = 'mock-excel-agent';
      name = 'Mock Excel Agent';
      hostType: HostType = 'Excel';

      getSystemPrompt(ctx: AgentContext): string {
        return `Specialist for ${ctx.host}`;
      }

      getTools(): ToolDefinition[] {
        return [];
      }

      async executeTool(toolCall: ToolCall, _ctx: AgentContext) {
        return { success: true, result: toolCall.name };
      }
    }

    const agent = new MockAgent();
    expect(agent.getSystemPrompt(context)).toContain('Excel');
    const result = await agent.executeTool({ id: '1', name: 'ping', arguments: {}, status: 'pending' }, context);
    expect(result.success).toBe(true);
    expect(result.result).toBe('ping');

    const resp: AgentResponse = {
      message: 'Done',
      toolCalls: [],
    };
    expect(resp.message).toBe('Done');
  });
});

describe('Excel Deep Intelligence Types', () => {
  it('validates FormulaIssue and SheetAuditResult structure', () => {
    const issue: FormulaIssue = {
      address: 'C5',
      type: 'formula_error',
      severity: 'critical',
      formula: '=A5/B5',
      currentValue: '#DIV/0!',
      suggestion: 'B5 bernilai 0, gunakan IFERROR(A5/B5, 0)',
    };
    const audit: SheetAuditResult = {
      sheetName: 'Sheet1',
      totalCellsAudited: 50,
      totalErrorsFound: 1,
      criticalIssues: [issue],
      warnings: [],
      summary: 'Ditemukan 1 error kritis pada sheet Sheet1.',
    };
    expect(audit.totalErrorsFound).toBe(1);
    expect(audit.criticalIssues[0].type).toBe('formula_error');
    expect(audit.criticalIssues[0].address).toBe('C5');
  });

  it('validates DataStoryOptions, DataStoryMetric, and DataStoryResult structure', () => {
    const options: DataStoryOptions = {
      range: 'A1:C10',
      focusMetric: 'Total Revenue',
      includeRecommendations: true,
    };
    const metric: DataStoryMetric = {
      label: 'Total Revenue',
      value: 'Rp 1.250.000.000',
      changePercent: 24,
      trend: 'up',
    };
    const story: DataStoryResult = {
      headline: 'Pendapatan Q3 Meningkat 24%',
      keyFindings: ['Penjualan produk X naik tajam'],
      metrics: [metric],
      risksOrAnomalies: ['Keterlambatan pasokan di minggu 2'],
      recommendations: ['Tingkatkan alokasi stok untuk produk X'],
    };
    expect(options.focusMetric).toBe('Total Revenue');
    expect(metric.trend).toBe('up');
    expect(story.headline).toContain('Pendapatan');
    expect(story.metrics[0].value).toBe('Rp 1.250.000.000');
  });

  it('validates PendingActionProposal structure', () => {
    const action: PendingActionProposal = {
      id: 'act-1',
      actionType: 'modify_cells',
      description: 'Perbarui 20 formula di kolom Total',
      affectedCellsCount: 20,
      targetRange: 'D2:D21',
      payload: { formula: '=B2*C2' },
    };
    expect(action.affectedCellsCount).toBe(20);
    expect(action.actionType).toBe('modify_cells');
  });
});

describe('Word Compliance & Corporate Style Types', () => {
  it('validates ComplianceClause and ComplianceReviewResult structure', () => {
    const clause: ComplianceClause = {
      category: 'liability_indemnity',
      excerpt: 'Pihak kedua menanggung seluruh ganti rugi tanpa batas',
      status: 'high_risk',
      analysis: 'Klausul tanggung jawab tidak terbatas (unlimited liability)',
      recommendation: 'Batasi maksimal nilai ganti rugi senilai nilai kontrak',
    };
    const review: ComplianceReviewResult = {
      contractType: 'vendor_service',
      overallRiskLevel: 'high',
      clausesReviewedCount: 1,
      identifiedClauses: [clause],
      missingCriticalClauses: ['force_majeure'],
      executiveSummary: 'Ditemukan klausul risiko tinggi ganti rugi tanpa batas.',
      actionableRecommendations: ['Tambahkan pembatasan liability cap.'],
    };
    expect(review.overallRiskLevel).toBe('high');
    expect(review.identifiedClauses[0].status).toBe('high_risk');
  });

  it('validates CorporateStyleResult structure', () => {
    const res: CorporateStyleResult = {
      appliedTheme: 'corporate_navy',
      fontFamily: 'Calibri',
      styledParagraphsCount: 15,
      headingsCount: 3,
      message: 'Berhasil menerapkan tema Corporate Navy ke 15 paragraf',
    };
    expect(res.appliedTheme).toBe('corporate_navy');
    expect(res.headingsCount).toBe(3);
  });
});

describe('DocToDeck Types', () => {
  it('validates DocToDeckSlide and DocToDeckResult contracts', () => {
    const slide: DocToDeckSlide = {
      title: 'Tinjauan Strategis',
      category: 'context',
      bullets: ['Tantangan integrasi sistem', 'Peningkatan kebutuhan efisiensi'],
      speakerScript: {
        hook: 'Bapak/Ibu sekalian, mari kita mulai dengan memahami latar belakang urgensi proyek ini.',
        keyTalkingPoints: ['Sistem saat ini menghadapi tantangan integrasi.', 'Efisiensi operasional perlu ditingkatkan segera.'],
        transition: 'Selanjutnya, mari kita telusuri pilar solusi yang diusulkan.',
      },
    };
    const options: DocToDeckOptions = {
      documentText: 'Teks dokumen contoh',
      theme: 'corporate_blue',
      targetSlideCount: 5,
    };
    const result: DocToDeckResult = {
      deckTitle: 'Presentasi Strategis',
      appliedTheme: options.theme || 'corporate_blue',
      totalSlidesCreated: 1,
      slides: [slide],
      summaryMessage: 'Berhasil membuat 1 slide.',
    };
    expect(result.deckTitle).toBe('Presentasi Strategis');
    expect(result.slides[0].category).toBe('context');
    expect(result.slides[0].speakerScript.hook).toBeDefined();
    expect(result.appliedTheme).toBe('corporate_blue');
  });
});

describe('Diagram & Flowchart Type Contracts', () => {
  it('validates FlowchartDefinition and DiagramResult structure', () => {
    const node: FlowchartNode = {
      id: 'step-1',
      label: 'Pengajuan Cuti',
      type: 'start',
      subText: 'Karyawan mengisi form',
    };

    const edge: FlowchartEdge = {
      from: 'step-1',
      to: 'step-2',
      label: 'Kirim',
      style: 'solid',
    };

    const def: FlowchartDefinition = {
      title: 'SOP Cuti Karyawan',
      direction: 'TD',
      theme: 'corporate_navy',
      nodes: [node],
      edges: [edge],
    };

    const result: DiagramResult = {
      title: def.title!,
      nodeCount: def.nodes.length,
      edgeCount: def.edges.length,
      appliedTheme: def.theme!,
      base64Png: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
      inserted: true,
    };

    expect(result.nodeCount).toBe(1);
    expect(result.edgeCount).toBe(1);
    expect(result.appliedTheme).toBe('corporate_navy');
    expect(result.inserted).toBe(true);
  });
});


