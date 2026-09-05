import React from 'react';
import { renderToString } from 'react-dom/server';
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { App } from '../src/App';
import { renderApp, detectHostAndRender } from '../src/main';
import {
  getExecutionMode,
  setExecutionMode,
} from '../src/services/storage/settingsStorage';
import * as llmFactory from '../src/services/llm/factory';
import * as officeModule from '../src/services/office';
import { ToolCall, HostType } from '../src/types';
import { ILLMProvider, StreamEvent } from '../src/services/llm/types';

// Ambient definitions for Node utilities in DOM TypeScript environment
declare function require(module: string): any;
declare const __dirname: string;

const fs = require('fs');
const path = require('path');

// Helper to simulate component execution and hooks in node environment
function createHookHarness<P>(Component: React.FC<P>, initialProps: P) {
  const states: any[] = [];
  let stateIdx = 0;
  const effectSlots: { effect: () => void | (() => void); deps?: any[] }[] = [];
  let effectIdx = 0;
  let props = initialProps;

  const dispatcher = {
    useState: (initial: any) => {
      const i = stateIdx++;
      if (!(i in states)) {
        states[i] = typeof initial === 'function' ? initial() : initial;
      }
      const setState = (next: any) => {
        const prevVal = states[i];
        states[i] = typeof next === 'function' ? next(prevVal) : next;
      };
      return [states[i], setState];
    },
    useEffect: (effect: () => void | (() => void), deps?: any[]) => {
      const slot = effectSlots[effectIdx];
      let shouldRun = false;

      if (!slot) {
        shouldRun = true;
      } else if (!deps) {
        shouldRun = true;
      } else if (
        deps.length !== slot.deps?.length ||
        deps.some((d, idx) => !Object.is(d, slot.deps?.[idx]))
      ) {
        shouldRun = true;
      }

      effectSlots[effectIdx] = { effect, deps };
      effectIdx++;

      if (shouldRun) {
        effect();
      }
    },
    useRef: (val: any) => {
      const i = stateIdx++;
      if (!(i in states)) {
        states[i] = { current: val };
      }
      return states[i];
    },
    useCallback: (fn: any) => fn,
    useMemo: (fn: any) => fn(),
  };

  const render = (newProps?: P): React.ReactElement => {
    if (newProps !== undefined) {
      props = newProps;
    }
    stateIdx = 0;
    effectIdx = 0;
    const prev = (React as any).__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED.ReactCurrentDispatcher.current;
    try {
      (React as any).__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED.ReactCurrentDispatcher.current = dispatcher;
      return Component(props) as React.ReactElement;
    } finally {
      (React as any).__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED.ReactCurrentDispatcher.current = prev;
    }
  };

  return {
    render,
    getStates: () => states,
    setProps: (p: P) => {
      props = p;
    },
  };
}

beforeEach(() => {
  localStorage.clear();
  setExecutionMode('copilot');
  officeModule.resetOfficeDriver();
  delete (globalThis as any).window;
  delete (globalThis as any).Office;
  delete (globalThis as any).document;
});

afterEach(() => {
  vi.restoreAllMocks();
  officeModule.resetOfficeDriver();
  delete (globalThis as any).window;
  delete (globalThis as any).Office;
  delete (globalThis as any).document;
});

describe('Universal Office Add-in Manifest (manifest.xml)', () => {
  const manifestPath = path.resolve(__dirname, '../manifest.xml');

  it('manifest.xml file exists on disk', () => {
    expect(fs.existsSync(manifestPath)).toBe(true);
  });

  it('contains valid OfficeApp XML namespaces and TaskPaneApp type', () => {
    const xml = fs.readFileSync(manifestPath, 'utf8');
    expect(xml).toContain('xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"');
    expect(xml).toContain('xsi:type="TaskPaneApp"');
  });

  it('defines correct application ID, version, and metadata', () => {
    const xml = fs.readFileSync(manifestPath, 'utf8');
    expect(xml).toContain('<Id>d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23</Id>');
    expect(xml).toContain('<Version>1.0.0.0</Version>');
    expect(xml).toContain('<ProviderName>Sam Office</ProviderName>');
    expect(xml).toContain('<DefaultLocale>en-US</DefaultLocale>');
    expect(xml).toContain('<DisplayName DefaultValue="Sam Office Agent" />');
    expect(xml).toContain('Universal multi-agent AI assistant for Microsoft Excel, Word, and PowerPoint.');
  });

  it('supports all 3 universal Office hosts: Workbook, Document, Presentation', () => {
    const xml = fs.readFileSync(manifestPath, 'utf8');
    expect(xml).toContain('<Host Name="Workbook" />');
    expect(xml).toContain('<Host Name="Document" />');
    expect(xml).toContain('<Host Name="Presentation" />');
  });

  it('specifies localhost:5173 source location and ReadWriteDocument permissions', () => {
    const xml = fs.readFileSync(manifestPath, 'utf8');
    expect(xml).toContain('<SourceLocation DefaultValue="https://localhost:5173/" />');
    expect(xml).toContain('<Permissions>ReadWriteDocument</Permissions>');
  });
});

describe('App Component Mounting & SSR Rendering', () => {
  it('renders default Excel view with welcome message and Header', () => {
    const html = renderToString(<App />);
    expect(html).toContain('Sam Office');
    expect(html).toContain('Excel');
    expect(html).toContain('Halo! Saya Sam, asisten AI untuk Microsoft Office (Excel)');
    expect(html).toContain('Tanya Sam atau perintahkan sesuatu...');
  });

  it('renders Word host view when initialHost is Word', () => {
    const html = renderToString(<App initialHost="Word" />);
    expect(html).toContain('Sam Office');
    expect(html).toContain('Word');
    expect(html).toContain('Halo! Saya Sam, asisten AI untuk Microsoft Office (Word)');
  });

  it('renders PowerPoint host view when initialHost is PowerPoint', () => {
    const html = renderToString(<App initialHost="PowerPoint" />);
    expect(html).toContain('Sam Office');
    expect(html).toContain('PowerPoint');
    expect(html).toContain('Halo! Saya Sam, asisten AI untuk Microsoft Office (PowerPoint)');
  });

  it('detects host dynamically from Office context if available', () => {
    (globalThis as any).window = {
      Office: {
        context: {
          host: 'Word',
        },
      },
    };

    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    harness.render(); // Mounts and triggers useEffect which sets host to 'Word'
    const vdom = harness.render(); // Re-render with new state
    const header = vdom.props.children[0];
    expect(header.props.host).toBe('Word');
  });
});

describe('App Controls & Settings Interaction', () => {
  it('toggles execution mode between copilot and autopilot', () => {
    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    let vdom = harness.render();

    let header = vdom.props.children[0];
    expect(header.props.executionMode).toBe('copilot');
    expect(getExecutionMode()).toBe('copilot');

    // Click toggle button
    header.props.onToggleMode();
    vdom = harness.render();

    header = vdom.props.children[0];
    expect(header.props.executionMode).toBe('autopilot');
    expect(getExecutionMode()).toBe('autopilot');

    // Toggle again back to copilot
    header.props.onToggleMode();
    vdom = harness.render();

    header = vdom.props.children[0];
    expect(header.props.executionMode).toBe('copilot');
    expect(getExecutionMode()).toBe('copilot');
  });

  it('opens and closes settings modal', () => {
    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    let vdom = harness.render();

    let settingsModal = vdom.props.children[3];
    expect(settingsModal.props.isOpen).toBe(false);

    // Open settings via Header
    const header = vdom.props.children[0];
    header.props.onOpenSettings();

    vdom = harness.render();
    settingsModal = vdom.props.children[3];
    expect(settingsModal.props.isOpen).toBe(true);

    // Close settings via modal onClose
    settingsModal.props.onClose();

    vdom = harness.render();
    settingsModal = vdom.props.children[3];
    expect(settingsModal.props.isOpen).toBe(false);
  });

  it('re-renders messages on settings onSaved callback', () => {
    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    let vdom = harness.render();

    const settingsModal = vdom.props.children[3];
    expect(() => settingsModal.props.onSaved()).not.toThrow();

    vdom = harness.render();
    expect(vdom).toBeDefined();
  });
});

describe('App Chat & LLM Streaming Integration', () => {
  it('ignores empty message sends', async () => {
    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    let vdom = harness.render();

    const inputBar = vdom.props.children[2];
    await inputBar.props.onSendMessage('   ');

    vdom = harness.render();
    const chatContainer = vdom.props.children[1];
    expect(chatContainer.props.messages.length).toBe(1); // Only initial welcome message
  });

  it('streams content_delta chunks into assistant response message', async () => {
    async function* mockStream(): AsyncIterable<StreamEvent> {
      yield { type: 'content_delta', delta: 'Halo, saya telah ' };
      yield { type: 'content_delta', delta: 'menganalisis spreadsheet Anda.' };
      yield { type: 'done' };
    }

    const mockProvider: ILLMProvider = {
      id: 'gemini',
      name: 'Google Gemini',
      sendMessage: vi.fn().mockReturnValue(mockStream()),
      testConnection: vi.fn().mockResolvedValue({ success: true, message: 'OK' }),
    };

    vi.spyOn(llmFactory, 'getLLMProvider').mockReturnValue(mockProvider);

    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    let vdom = harness.render();

    const inputBar = vdom.props.children[2];
    await inputBar.props.onSendMessage('Analisis data kolom A');

    vdom = harness.render();
    const chatContainer = vdom.props.children[1];
    const messages = chatContainer.props.messages;

    expect(messages.length).toBe(3);
    expect(messages[1].role).toBe('user');
    expect(messages[1].content).toBe('Analisis data kolom A');

    expect(messages[2].role).toBe('assistant');
    expect(messages[2].content).toBe('Halo, saya telah menganalisis spreadsheet Anda.');

    expect(mockProvider.sendMessage).toHaveBeenCalled();
  });

  it('handles error events streamed from provider', async () => {
    async function* mockStream(): AsyncIterable<StreamEvent> {
      yield { type: 'content_delta', delta: 'Sedang memproses... ' };
      yield { type: 'error', error: 'Rate limit provider terlampaui' };
    }

    const mockProvider: ILLMProvider = {
      id: 'gemini',
      name: 'Google Gemini',
      sendMessage: vi.fn().mockReturnValue(mockStream()),
      testConnection: vi.fn().mockResolvedValue({ success: true, message: 'OK' }),
    };

    vi.spyOn(llmFactory, 'getLLMProvider').mockReturnValue(mockProvider);

    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    let vdom = harness.render();

    const inputBar = vdom.props.children[2];
    await inputBar.props.onSendMessage('Buat formula');

    vdom = harness.render();
    const chatContainer = vdom.props.children[1];
    const messages = chatContainer.props.messages;

    expect(messages[2].content).toContain('[Error: Rate limit provider terlampaui]');
  });

  it('handles exception thrown by provider sendMessage stream', async () => {
    const mockProvider: ILLMProvider = {
      id: 'gemini',
      name: 'Google Gemini',
      sendMessage: vi.fn().mockImplementation(() => {
        throw new Error('Jaringan terputus ke LLM endpoint');
      }),
      testConnection: vi.fn().mockResolvedValue({ success: true, message: 'OK' }),
    };

    vi.spyOn(llmFactory, 'getLLMProvider').mockReturnValue(mockProvider);

    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    let vdom = harness.render();

    const inputBar = vdom.props.children[2];
    await inputBar.props.onSendMessage('Hitung total');

    vdom = harness.render();
    const chatContainer = vdom.props.children[1];
    const messages = chatContainer.props.messages;

    expect(messages[2].content).toContain('[Error: Jaringan terputus ke LLM endpoint]');
  });

  it('locks concurrency by updating isBusy state during sendMessage', async () => {
    let resolveStream: () => void;
    const streamGate = new Promise<void>(resolve => {
      resolveStream = resolve;
    });

    async function* controlledStream(): AsyncIterable<StreamEvent> {
      yield { type: 'content_delta', delta: 'Wait...' };
      await streamGate;
      yield { type: 'done' };
    }

    const mockProvider: ILLMProvider = {
      id: 'gemini',
      name: 'Google Gemini',
      sendMessage: vi.fn().mockReturnValue(controlledStream()),
      testConnection: vi.fn().mockResolvedValue({ success: true, message: 'OK' }),
    };

    vi.spyOn(llmFactory, 'getLLMProvider').mockReturnValue(mockProvider);

    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    let vdom = harness.render();

    let inputBar = vdom.props.children[2];
    expect(inputBar.props.disabled).toBe(false);

    const sendPromise = inputBar.props.onSendMessage('Test busy lock');

    vdom = harness.render();
    inputBar = vdom.props.children[2];
    expect(inputBar.props.disabled).toBe(true);

    resolveStream!();
    await sendPromise;

    vdom = harness.render();
    inputBar = vdom.props.children[2];
    expect(inputBar.props.disabled).toBe(false);
  });
});

describe('Copilot Mode Tool Call Workflow (Preview & Manual Apply)', () => {
  it('receives tool_call with status pending without auto-executing in copilot mode', async () => {
    const testToolCall: ToolCall = {
      id: 'tc-copilot-1',
      name: 'write_cells',
      arguments: { range: 'C10', formula: '=SUM(C1:C9)' },
      status: 'pending',
    };

    async function* mockStream(): AsyncIterable<StreamEvent> {
      yield { type: 'content_delta', delta: 'Berikut formula yang disiapkan:' };
      yield { type: 'tool_call', toolCall: testToolCall };
      yield { type: 'done' };
    }

    const mockProvider: ILLMProvider = {
      id: 'gemini',
      name: 'Google Gemini',
      sendMessage: vi.fn().mockReturnValue(mockStream()),
      testConnection: vi.fn().mockResolvedValue({ success: true, message: 'OK' }),
    };
    vi.spyOn(llmFactory, 'getLLMProvider').mockReturnValue(mockProvider);

    const mockDriver = new officeModule.MockOfficeDriver();
    const writeCellsSpy = vi.spyOn(mockDriver, 'writeCells');
    officeModule.setOfficeDriver(mockDriver);

    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    let vdom = harness.render();

    const inputBar = vdom.props.children[2];
    await inputBar.props.onSendMessage('Tolong buat formula jumlah');

    vdom = harness.render();
    const chatContainer = vdom.props.children[1];
    const assistantMsg = chatContainer.props.messages[2];

    expect(assistantMsg.toolCalls?.length).toBe(1);
    expect(assistantMsg.toolCalls?.[0].status).toBe('pending');
    expect(writeCellsSpy).not.toHaveBeenCalled(); // NOT auto-executed in copilot mode

    // Now user clicks "Terapkan ke Dokumen" (calling onApplyToolCall)
    await chatContainer.props.onApplyToolCall(assistantMsg.toolCalls![0]);

    vdom = harness.render();
    expect(assistantMsg.toolCalls?.[0].status).toBe('applied');
    expect(writeCellsSpy).toHaveBeenCalledWith('C10', undefined, [['=SUM(C1:C9)']]);
  });

  it('updates toolCall status to failed when formula validation fails during manual apply', async () => {
    const invalidToolCall: ToolCall = {
      id: 'tc-copilot-err',
      name: 'write_cells',
      arguments: { range: 'A1', formula: '=SUM(invalid formula syntax ++' },
      status: 'pending',
    };

    async function* mockStream(): AsyncIterable<StreamEvent> {
      yield { type: 'tool_call', toolCall: invalidToolCall };
    }

    const mockProvider: ILLMProvider = {
      id: 'gemini',
      name: 'Google Gemini',
      sendMessage: vi.fn().mockReturnValue(mockStream()),
      testConnection: vi.fn().mockResolvedValue({ success: true, message: 'OK' }),
    };
    vi.spyOn(llmFactory, 'getLLMProvider').mockReturnValue(mockProvider);

    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    let vdom = harness.render();

    const inputBar = vdom.props.children[2];
    await inputBar.props.onSendMessage('Tulis formula salah');

    vdom = harness.render();
    const chatContainer = vdom.props.children[1];
    const toolCall = chatContainer.props.messages[2].toolCalls![0];

    // Manual apply
    await chatContainer.props.onApplyToolCall(toolCall);

    vdom = harness.render();
    expect(toolCall.status).toBe('failed');
    expect(toolCall.error).toBeDefined();
  });
});

describe('Autopilot Mode Tool Call Workflow (Direct Auto Apply)', () => {
  beforeEach(() => {
    setExecutionMode('autopilot');
  });

  it('automatically applies tool calls immediately upon receipt in autopilot mode', async () => {
    const autoToolCall: ToolCall = {
      id: 'tc-auto-1',
      name: 'write_cells',
      arguments: { range: 'B5', formula: '=AVERAGE(B1:B4)' },
      status: 'pending',
    };

    async function* mockStream(): AsyncIterable<StreamEvent> {
      yield { type: 'content_delta', delta: 'Menghitung rata-rata...' };
      yield { type: 'tool_call', toolCall: autoToolCall };
      yield { type: 'done' };
    }

    const mockProvider: ILLMProvider = {
      id: 'gemini',
      name: 'Google Gemini',
      sendMessage: vi.fn().mockReturnValue(mockStream()),
      testConnection: vi.fn().mockResolvedValue({ success: true, message: 'OK' }),
    };
    vi.spyOn(llmFactory, 'getLLMProvider').mockReturnValue(mockProvider);

    const mockDriver = new officeModule.MockOfficeDriver();
    const writeCellsSpy = vi.spyOn(mockDriver, 'writeCells');
    officeModule.setOfficeDriver(mockDriver);

    const harness = createHookHarness(App, { initialHost: 'Excel' as HostType });
    let vdom = harness.render();

    const inputBar = vdom.props.children[2];
    await inputBar.props.onSendMessage('Hitung rata-rata');

    vdom = harness.render();
    const chatContainer = vdom.props.children[1];
    const assistantMsg = chatContainer.props.messages[2];

    expect(assistantMsg.toolCalls?.length).toBe(1);
    // In autopilot mode, status should be automatically applied without calling onApplyToolCall manually!
    expect(assistantMsg.toolCalls?.[0].status).toBe('applied');
    expect(writeCellsSpy).toHaveBeenCalledWith('B5', undefined, [['=AVERAGE(B1:B4)']]);
  });
});

describe('Entrypoint Lifecycle & Office.onReady (src/main.tsx)', () => {
  it('falls back to BrowserDev host when Office is not defined', () => {
    let renderedHost: HostType | null = null;
    detectHostAndRender(host => {
      renderedHost = host;
    });
    expect(renderedHost).toBe('BrowserDev');
  });

  it('detects Excel host when Office.onReady returns Excel host', () => {
    let renderedHost: HostType | null = null;
    (globalThis as any).window = {
      Office: {
        HostType: {
          Excel: 'Excel',
          Word: 'Word',
          PowerPoint: 'PowerPoint',
        },
        onReady: vi.fn((callback: (info: any) => void) => {
          callback({ host: 'Excel' });
        }),
      },
    };

    detectHostAndRender(host => {
      renderedHost = host;
    });

    expect(renderedHost).toBe('Excel');
  });

  it('detects Word host when Office.onReady returns Word host', () => {
    let renderedHost: HostType | null = null;
    (globalThis as any).window = {
      Office: {
        HostType: {
          Excel: 'Excel',
          Word: 'Word',
          PowerPoint: 'PowerPoint',
        },
        onReady: vi.fn((callback: (info: any) => void) => {
          callback({ host: 'Word' });
        }),
      },
    };

    detectHostAndRender(host => {
      renderedHost = host;
    });

    expect(renderedHost).toBe('Word');
  });

  it('detects PowerPoint host when Office.onReady returns PowerPoint host', () => {
    let renderedHost: HostType | null = null;
    (globalThis as any).window = {
      Office: {
        HostType: {
          Excel: 'Excel',
          Word: 'Word',
          PowerPoint: 'PowerPoint',
        },
        onReady: vi.fn((callback: (info: any) => void) => {
          callback({ host: 'PowerPoint' });
        }),
      },
    };

    detectHostAndRender(host => {
      renderedHost = host;
    });

    expect(renderedHost).toBe('PowerPoint');
  });

  it('renderApp renders into targetElement when provided', () => {
    let renderedElem: React.ReactElement | null = null;
    const mockRoot = {
      render: vi.fn((elem: React.ReactElement) => {
        renderedElem = elem;
      }),
    };

    renderApp('Excel', mockRoot);
    expect(mockRoot.render).toHaveBeenCalled();
    expect(renderedElem).toBeDefined();
  });

  it('renderApp gracefully does nothing if document or root element is missing', () => {
    expect(() => renderApp('BrowserDev')).not.toThrow();

    (globalThis as any).document = {
      getElementById: vi.fn().mockReturnValue(null),
    };
    expect(() => renderApp('BrowserDev')).not.toThrow();
  });
});
