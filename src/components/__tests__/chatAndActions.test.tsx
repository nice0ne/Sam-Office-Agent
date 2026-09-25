import React from 'react';
import { renderToString } from 'react-dom/server';
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ActionCard } from '../Actions/ActionCard';
import { MessageBubble } from '../Chat/MessageBubble';
import { InputBar } from '../Chat/InputBar';
import { ChatContainer } from '../Chat/ChatContainer';
import { QuickActionPresets } from '../Chat/QuickActionPresets';
import { ActionConfirmationCard } from '../Chat/ActionConfirmationCard';
import { CrossAppSnapshotCard } from '../Chat/CrossAppSnapshotCard';
import {
  saveCrossAppSnapshot,
  clearCrossAppSnapshots,
  isCrossAppSnapshotDismissed,
} from '../../services/storage/crossAppBridge';
import { ToolCall, ChatMessage, PendingActionProposal } from '../../types';

// Helper to strip React SSR comments from html output
function stripComments(html: string): string {
  return html.replace(/<!--.*?-->/g, '');
}

interface CustomMatchers<R = unknown> {
  toBeInTheDocument(): R;
}

declare module 'vitest' {
  interface Assertion<T = any> extends CustomMatchers<T> {}
  interface AsymmetricMatchersContaining extends CustomMatchers {}
}

let lastRenderedHtml = '';
function render(ui: React.ReactElement) {
  lastRenderedHtml = stripComments(renderToString(ui));
  return {
    html: lastRenderedHtml,
    rerender: (newUi: React.ReactElement) => {
      lastRenderedHtml = stripComments(renderToString(newUi));
      return { html: lastRenderedHtml };
    },
  };
}

const screen = {
  getByText: (pattern: RegExp | string) => {
    const decoded = lastRenderedHtml.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
    const regex = typeof pattern === 'string' ? new RegExp(pattern) : pattern;
    const match = regex.test(decoded) || regex.test(lastRenderedHtml);
    return match ? { inDocument: true } : null;
  },
  getByTitle: (pattern: RegExp | string) => {
    const decoded = lastRenderedHtml.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
    const regex = typeof pattern === 'string' ? new RegExp(pattern) : pattern;
    const titleRegex = /title="([^"]*)"/g;
    let matchFound = false;
    let match;
    while ((match = titleRegex.exec(decoded)) !== null) {
      if (regex.test(match[1])) {
        matchFound = true;
        break;
      }
    }
    return matchFound ? { inDocument: true } : null;
  },
};

expect.extend({
  toBeInTheDocument(received: any) {
    const pass = Boolean(received && received.inDocument === true);
    return {
      pass,
      message: () => `expected element ${pass ? 'not ' : ''}to be in document`,
    };
  },
});

// Helper to recursively find an element in a ReactElement tree (handles nested arrays from .map)
function findElement(node: any, predicate: (elem: any) => boolean): any {
  if (!node) return null;
  if (Array.isArray(node)) {
    for (const child of node) {
      const found = findElement(child, predicate);
      if (found) return found;
    }
    return null;
  }
  if (typeof node !== 'object') return null;
  if (predicate(node)) return node;
  if (node.props && node.props.children) {
    return findElement(node.props.children, predicate);
  }
  return null;
}

// Helper to recursively find all elements matching a predicate (handles nested arrays from .map)
function findAllElements(node: any, predicate: (elem: any) => boolean): any[] {
  const results: any[] = [];
  function search(item: any) {
    if (!item) return;
    if (Array.isArray(item)) {
      item.forEach(search);
      return;
    }
    if (typeof item !== 'object') return;
    if (predicate(item)) results.push(item);
    if (item.props && item.props.children) {
      search(item.props.children);
    }
  }
  search(node);
  return results;
}

// Helper to simulate component render with React hooks in node environment
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
        states[i] = typeof next === 'function' ? next(states[i]) : next;
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

describe('ActionCard Component', () => {
  const mockToolCall: ToolCall = {
    id: 'tc-1',
    name: 'write_cells',
    arguments: { range: 'A1:B10', formula: '=SUM(C1:C10)' },
    status: 'pending',
  };

  it('renders tool name with underscores replaced and styled capitalize', () => {
    const html = renderToString(<ActionCard toolCall={mockToolCall} onApply={() => {}} />);
    expect(html).toContain('write cells');
    expect(html).toContain('capitalize');
  });

  it('renders write_cells specific arguments (range and formula)', () => {
    const html = renderToString(<ActionCard toolCall={mockToolCall} onApply={() => {}} />);
    expect(html).toContain('Range:');
    expect(html).toContain('A1:B10');
    expect(html).toContain('Formula:');
    expect(html).toContain('=SUM(C1:C10)');
  });

  it('renders write_cells without formula if not provided', () => {
    const noFormula: ToolCall = {
      id: 'tc-2',
      name: 'write_cells',
      arguments: { range: 'D5:E5' },
      status: 'pending',
    };
    const html = renderToString(<ActionCard toolCall={noFormula} onApply={() => {}} />);
    expect(html).toContain('Range:');
    expect(html).toContain('D5:E5');
    expect(html).not.toContain('Formula:');
  });

  it('renders format_range specific arguments (range and numberFormat)', () => {
    const formatTool: ToolCall = {
      id: 'tc-3',
      name: 'format_range',
      arguments: { range: 'C2:C20', numberFormat: 'Rp#,##0' },
      status: 'pending',
    };
    const html = renderToString(<ActionCard toolCall={formatTool} onApply={() => {}} />);
    expect(html).toContain('format range');
    expect(html).toContain('Range:');
    expect(html).toContain('C2:C20');
    expect(html).toContain('Format:');
    expect(html).toContain('Rp#,##0');
  });

  it('renders create_chart specific arguments (chartType and dataRange)', () => {
    const chartTool: ToolCall = {
      id: 'tc-4',
      name: 'create_chart',
      arguments: { chartType: 'ColumnClustered', dataRange: 'A1:D10' },
      status: 'pending',
    };
    const html = stripComments(
      renderToString(<ActionCard toolCall={chartTool} onApply={() => {}} />)
    );
    expect(html).toContain('create chart');
    expect(html).toContain('Chart:');
    expect(html).toContain('ColumnClustered (A1:D10)');
  });

  it('renders fallback JSON pre block for custom or non-standard tools', () => {
    const customTool: ToolCall = {
      id: 'tc-5',
      name: 'insert_paragraph',
      arguments: { text: 'Halo dunia dokumen', position: 'end' },
      status: 'pending',
    };
    const html = renderToString(<ActionCard toolCall={customTool} onApply={() => {}} />);
    expect(html).toContain('insert paragraph');
    expect(html).toContain('&quot;text&quot;: &quot;Halo dunia dokumen&quot;');
    expect(html).toContain('&quot;position&quot;: &quot;end&quot;');
  });

  it('handles empty arguments safely without crashing', () => {
    const emptyArgsTool: ToolCall = {
      id: 'tc-empty',
      name: 'unknown_tool',
      arguments: {},
      status: 'pending',
    };
    const html = renderToString(<ActionCard toolCall={emptyArgsTool} onApply={() => {}} />);
    expect(html).toContain('unknown tool');
  });

  it('renders pending status with apply button and no badges', () => {
    const html = renderToString(<ActionCard toolCall={mockToolCall} onApply={() => {}} />);
    expect(html).toContain('Terapkan ke Dokumen');
    expect(html).not.toContain('Diterapkan');
    expect(html).not.toContain('Gagal');
  });

  it('renders applied status with CheckCircle2 badge and hides apply button', () => {
    const appliedTool: ToolCall = { ...mockToolCall, status: 'applied' };
    const html = renderToString(<ActionCard toolCall={appliedTool} onApply={() => {}} />);
    expect(html).toContain('Diterapkan');
    expect(html).not.toContain('Terapkan ke Dokumen');
  });

  it('renders failed status with AlertCircle badge, error message, and allows apply retry', () => {
    const failedTool: ToolCall = {
      ...mockToolCall,
      status: 'failed',
      error: 'Invalid formula syntax',
    };
    const html = renderToString(<ActionCard toolCall={failedTool} onApply={() => {}} />);
    expect(html).toContain('Gagal');
    expect(html).toContain('Invalid formula syntax');
    expect(html).toContain('Terapkan ke Dokumen');
  });

  it('triggers onApply callback with toolCall when apply button is clicked', () => {
    const onApply = vi.fn();
    const vdom = ActionCard({
      toolCall: mockToolCall,
      onApply,
    }) as React.ReactElement;

    const button = findElement(vdom, el => el.type === 'button');
    expect(button).toBeDefined();
    button.props.onClick();

    expect(onApply).toHaveBeenCalledTimes(1);
    expect(onApply).toHaveBeenCalledWith(mockToolCall);
  });

  it('disables apply button and shows Menerapkan... when isExecuting is true', () => {
    const html = renderToString(
      <ActionCard toolCall={mockToolCall} onApply={() => {}} isExecuting={true} />
    );
    expect(html).toContain('Menerapkan...');
    expect(html).toContain('disabled=""');
  });
});

describe('MessageBubble Component', () => {
  const userMessage: ChatMessage = {
    id: 'msg-1',
    role: 'user',
    content: 'Tolong hitung total penjualan kolom B.',
    timestamp: Date.now(),
  };

  const assistantMessage: ChatMessage = {
    id: 'msg-2',
    role: 'assistant',
    content: 'Tentu, saya telah menyiapkan formula =SUM(B2:B20) untuk Anda.',
    timestamp: Date.now(),
    toolCalls: [
      {
        id: 'tc-sub-1',
        name: 'write_cells',
        arguments: { range: 'B21', formula: '=SUM(B2:B20)' },
        status: 'pending',
      },
    ],
  };

  it('renders user message justified to end with blue styling and user avatar', () => {
    const html = renderToString(
      <MessageBubble message={userMessage} onApplyToolCall={() => {}} />
    );
    expect(html).toContain('justify-end');
    expect(html).toContain('bg-blue-600 text-white rounded-tr-none');
    expect(html).toContain('Tolong hitung total penjualan kolom B.');
  });

  it('renders assistant message justified to start with bot avatar and card styling', () => {
    const html = renderToString(
      <MessageBubble message={assistantMessage} onApplyToolCall={() => {}} />
    );
    expect(html).toContain('justify-start');
    expect(html).toContain('rounded-tl-none shadow-sm');
    expect(html).toContain('Tentu, saya telah menyiapkan formula');
  });

  it('renders rich markdown content including bold, headings, and tables', () => {
    const mdMsg: ChatMessage = {
      id: 'msg-md',
      role: 'assistant',
      content: '# Ringkasan Proyek\nBerikut data **rekapitulasi**:\n\n| Item | Nilai |\n|---|---|\n| Server | 10 Juta |',
      timestamp: Date.now(),
    };
    const html = renderToString(<MessageBubble message={mdMsg} onApplyToolCall={() => {}} />);
    expect(html).toContain('<h1');
    expect(html).toContain('Ringkasan Proyek');
    expect(html).toContain('<strong');
    expect(html).toContain('rekapitulasi</strong>');
    expect(html).toContain('<table');
    expect(html).toContain('Server');
    expect(html).toContain('10 Juta');
  });

  it('renders embedded action cards when message has toolCalls', () => {
    const html = renderToString(
      <MessageBubble message={assistantMessage} onApplyToolCall={() => {}} />
    );
    expect(html).toContain('write cells');
    expect(html).toContain('B21');
    expect(html).toContain('=SUM(B2:B20)');
    expect(html).toContain('Terapkan ke Dokumen');
  });

  it('renders message error message if present', () => {
    const errorMsg: ChatMessage = {
      id: 'msg-err',
      role: 'assistant',
      content: 'Maaf, terjadi kendala.',
      error: 'Quota API terlampaui',
      timestamp: Date.now(),
    };
    const html = renderToString(<MessageBubble message={errorMsg} onApplyToolCall={() => {}} />);
    expect(html).toContain('Quota API terlampaui');
  });

  it('renders thinking bouncing dots animation when assistant message is empty', () => {
    const thinkingMsg: ChatMessage = {
      id: 'msg-think',
      role: 'assistant',
      content: '',
      timestamp: Date.now(),
      toolCalls: [],
    };
    const html = renderToString(<MessageBubble message={thinkingMsg} onApplyToolCall={() => {}} />);
    expect(html).toContain('Sam sedang berpikir...');
    expect(html).toContain('animate-dot-1');
    expect(html).toContain('animate-dot-2');
    expect(html).toContain('animate-dot-3');
  });

  it('passes onApplyToolCall callback to embedded ActionCard', () => {
    const onApplyToolCall = vi.fn();
    const harness = createHookHarness(MessageBubble, {
      message: assistantMessage,
      onApplyToolCall,
    });
    const vdom = harness.render();

    // ActionCard element inside MessageBubble tree
    const actionCard = findElement(vdom, el => el.type === ActionCard);
    expect(actionCard).toBeDefined();
    expect(actionCard.props.toolCall).toBe(assistantMessage.toolCalls![0]);
    expect(actionCard.props.onApply).toBe(onApplyToolCall);

    // Invoke onApply via ActionCard props
    actionCard.props.onApply(actionCard.props.toolCall);
    expect(onApplyToolCall).toHaveBeenCalledTimes(1);
    expect(onApplyToolCall).toHaveBeenCalledWith(assistantMessage.toolCalls![0]);

    // Render ActionCard directly to simulate click on internal button
    const actionCardVdom = ActionCard(actionCard.props) as React.ReactElement;
    const button = findElement(actionCardVdom, el => el.type === 'button');
    expect(button).toBeDefined();
    button.props.onClick();
    expect(onApplyToolCall).toHaveBeenCalledTimes(2);
  });

  it('renders copy button and copies message content to clipboard on click', async () => {
    const writeText = vi.fn().mockResolvedValue(undefined);
    Object.assign(navigator, {
      clipboard: {
        writeText,
      },
    });

    const harness = createHookHarness(MessageBubble, {
      message: userMessage,
      onApplyToolCall: () => {},
    });
    const vdom = harness.render();
    const copyButton = findElement(vdom, el => el.type === 'button' && (el.props['aria-label'] === 'Salin pesan' || el.props.title === 'Salin pesan'));
    expect(copyButton).toBeDefined();

    await copyButton.props.onClick();
    expect(writeText).toHaveBeenCalledWith(userMessage.content);
  });

  it('renders download button on assistant messages and triggers file download', () => {
    const mockLink = {
      href: '',
      download: '',
      click: vi.fn(),
    };
    const mockDoc = {
      createElement: vi.fn().mockReturnValue(mockLink),
      body: {
        appendChild: vi.fn(),
        removeChild: vi.fn(),
      },
    };
    const origDoc = (globalThis as any).document;
    const origURL = (globalThis as any).URL;
    (globalThis as any).document = mockDoc;
    (globalThis as any).URL = {
      createObjectURL: vi.fn().mockReturnValue('blob:mock-url'),
      revokeObjectURL: vi.fn(),
    };

    try {
      const harness = createHookHarness(MessageBubble, {
        message: assistantMessage,
        onApplyToolCall: () => {},
      });
      const vdom = harness.render();
      const downloadButton = findElement(
        vdom,
        el => el.type === 'button' && (el.props['aria-label'] === 'Unduh file' || el.props.title?.includes('Unduh'))
      );
      expect(downloadButton).toBeDefined();

      downloadButton.props.onClick();
      expect((globalThis as any).URL.createObjectURL).toHaveBeenCalled();
      expect(mockDoc.body.appendChild).toHaveBeenCalled();
      expect(mockLink.click).toHaveBeenCalled();
      expect(mockDoc.body.removeChild).toHaveBeenCalled();
      expect((globalThis as any).URL.revokeObjectURL).toHaveBeenCalledWith('blob:mock-url');
    } finally {
      (globalThis as any).document = origDoc;
      (globalThis as any).URL = origURL;
    }
  });
});

describe('InputBar Component', () => {
  it('renders input with default placeholder and send button', () => {
    const html = renderToString(<InputBar onSendMessage={() => {}} />);
    expect(html).toContain('Tanya Sam atau perintahkan sesuatu...');
    expect(html).toContain('disabled=""'); // empty input starts disabled
  });

  it('renders custom placeholder when specified', () => {
    const html = renderToString(
      <InputBar onSendMessage={() => {}} placeholder="Tulis instruksi kustom..." />
    );
    expect(html).toContain('Tulis instruksi kustom...');
  });

  it('renders Stop button when isBusy and onStop are provided, and clicking it triggers onStop', () => {
    const onStop = vi.fn();
    const html = renderToString(
      <InputBar onSendMessage={() => {}} onStop={onStop} isBusy={true} />
    );
    expect(html).toContain('Hentikan respons');
    expect(html).toContain('bg-red-600');

    const harness = createHookHarness(InputBar, { onSendMessage: () => {}, onStop, isBusy: true });
    const vdom = harness.render();
    const stopButton = findElement(vdom, el => el.type === 'button' && el.props['aria-label'] === 'Hentikan respons');
    expect(stopButton).toBeDefined();
    stopButton.props.onClick();
    expect(onStop).toHaveBeenCalledTimes(1);
  });

  it('enables send button when text is entered and sends trimmed text on click', () => {
    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, { onSendMessage });

    // Initial render
    let vdom = harness.render();
    const inputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');
    const buttonEl = findElement(vdom, el => el.type === 'button' && el.props['aria-label'] === 'Kirim Pesan');
    expect(buttonEl.props.disabled).toBe(true);

    // Simulate typing
    inputEl.props.onChange({ target: { value: '  Buatkan chart dari kolom A ke D  ' } });

    // Re-render after state update
    vdom = harness.render();
    const updatedInputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');
    const updatedButtonEl = findElement(vdom, el => el.type === 'button' && el.props['aria-label'] === 'Kirim Pesan');
    expect(updatedInputEl.props.value).toBe('  Buatkan chart dari kolom A ke D  ');
    expect(updatedButtonEl.props.disabled).toBe(false);

    // Click send button
    updatedButtonEl.props.onClick();

    // Verify callback
    expect(onSendMessage).toHaveBeenCalledTimes(1);
    expect(onSendMessage).toHaveBeenCalledWith('Buatkan chart dari kolom A ke D');

    // Re-render to confirm input is reset to empty string
    vdom = harness.render();
    const resetInputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');
    const resetButtonEl = findElement(vdom, el => el.type === 'button' && el.props['aria-label'] === 'Kirim Pesan');
    expect(resetInputEl.props.value).toBe('');
    expect(resetButtonEl.props.disabled).toBe(true);
  });

  it('sends message on Enter key without shiftKey', () => {
    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, { onSendMessage });

    let vdom = harness.render();
    const inputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');

    // Type text
    inputEl.props.onChange({ target: { value: 'Hitung rata-rata C1:C10' } });
    vdom = harness.render();
    const updatedInputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');

    // Press Enter
    const preventDefault = vi.fn();
    updatedInputEl.props.onKeyDown({ key: 'Enter', shiftKey: false, preventDefault });

    expect(preventDefault).toHaveBeenCalledTimes(1);
    expect(onSendMessage).toHaveBeenCalledTimes(1);
    expect(onSendMessage).toHaveBeenCalledWith('Hitung rata-rata C1:C10');
  });

  it('does not send message on Shift+Enter key', () => {
    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, { onSendMessage });

    let vdom = harness.render();
    const inputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');

    inputEl.props.onChange({ target: { value: 'Baris pertama' } });
    vdom = harness.render();
    const updatedInputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');

    const preventDefault = vi.fn();
    updatedInputEl.props.onKeyDown({ key: 'Enter', shiftKey: true, preventDefault });

    expect(preventDefault).not.toHaveBeenCalled();
    expect(onSendMessage).not.toHaveBeenCalled();
  });

  it('does not send message when input contains only whitespace', () => {
    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, { onSendMessage });

    let vdom = harness.render();
    const inputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');

    inputEl.props.onChange({ target: { value: '     ' } });
    vdom = harness.render();
    const updatedButtonEl = findElement(vdom, el => el.type === 'button' && el.props['aria-label'] === 'Kirim Pesan');
    const updatedInputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');

    expect(updatedButtonEl.props.disabled).toBe(true);
    updatedButtonEl.props.onClick();
    expect(onSendMessage).not.toHaveBeenCalled();

    updatedInputEl.props.onKeyDown({ key: 'Enter', shiftKey: false, preventDefault: vi.fn() });
    expect(onSendMessage).not.toHaveBeenCalled();
  });

  it('disables input and button and ignores events when disabled prop is true', () => {
    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, { onSendMessage, disabled: true });

    let vdom = harness.render();
    const inputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');
    const buttonEl = findElement(vdom, el => el.type === 'button');

    expect(inputEl.props.disabled).toBe(true);
    expect(buttonEl.props.disabled).toBe(true);

    // Attempt typing and sending while disabled
    inputEl.props.onChange({ target: { value: 'Test message' } });
    vdom = harness.render();
    const updatedButtonEl = findElement(vdom, el => el.type === 'button');
    const updatedInputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');

    expect(updatedButtonEl.props.disabled).toBe(true);
    updatedButtonEl.props.onClick();
    expect(onSendMessage).not.toHaveBeenCalled();

    updatedInputEl.props.onKeyDown({ key: 'Enter', shiftKey: false, preventDefault: vi.fn() });
    expect(onSendMessage).not.toHaveBeenCalled();
  });

  it('renders microphone button and language toggle', () => {
    (globalThis as any).webkitSpeechRecognition = class MockSpeech {
      start = vi.fn();
      stop = vi.fn();
    };
    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, { onSendMessage });
    const vdom = harness.render();

    const micButton = findElement(vdom, el => el.props && el.props['aria-label'] === 'Perintah Suara');
    expect(micButton).toBeDefined();

    const langButton = findElement(vdom, el => el.props && el.props['aria-label'] === 'Pilihan Bahasa Suara');
    expect(langButton).toBeDefined();
    expect(langButton.props.children).toContain('ID');
  });

  it('toggles voice recognition and language in InputBar', () => {
    let mockInstance: any = null;
    class MockSpeech {
      continuous = false;
      interimResults = false;
      lang = 'id-ID';
      onstart: any = null;
      onend: any = null;
      onresult: any = null;
      constructor() {
        mockInstance = this;
      }
      start = vi.fn().mockImplementation(() => {
        if (this.onstart) this.onstart();
      });
      stop = vi.fn().mockImplementation(() => {
        if (this.onend) this.onend();
      });
    }
    (globalThis as any).webkitSpeechRecognition = MockSpeech;

    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, { onSendMessage });
    let vdom = harness.render();

    // Click language toggle
    const langButton = findElement(vdom, el => el.props && el.props['aria-label'] === 'Pilihan Bahasa Suara');
    langButton.props.onClick();
    vdom = harness.render();
    const updatedLangButton = findElement(vdom, el => el.props && el.props['aria-label'] === 'Pilihan Bahasa Suara');
    expect(updatedLangButton.props.children).toContain('EN');

    // Click mic button to start
    const micButton = findElement(vdom, el => el.props && el.props['aria-label'] === 'Perintah Suara');
    micButton.props.onClick();
    vdom = harness.render();

    expect(mockInstance.start).toHaveBeenCalled();

    // Send speech recognition result
    mockInstance.onresult({
      resultIndex: 0,
      results: [
        {
          0: { transcript: 'buatkan grafik baris' },
          isFinal: true,
          length: 1,
        },
      ],
    });
    vdom = harness.render();

    const inputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');
    expect(inputEl.props.value).toContain('buatkan grafik baris');
  });

  it('clears input text and resets transcript when clicking microphone to start new recording', () => {
    let mockInstance: any = null;
    class MockSpeech {
      continuous = false;
      interimResults = false;
      lang = 'id-ID';
      onstart: any = null;
      onend: any = null;
      onresult: any = null;
      constructor() {
        mockInstance = this;
      }
      start = vi.fn().mockImplementation(() => {
        if (this.onstart) this.onstart();
      });
      stop = vi.fn().mockImplementation(() => {
        if (this.onend) this.onend();
      });
    }
    (globalThis as any).webkitSpeechRecognition = MockSpeech;

    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, { onSendMessage });
    let vdom = harness.render();

    // Type some text first or have previous input
    const inputEl = findElement(vdom, el => el.type === 'textarea' || el.type === 'input');
    inputEl.props.onChange({ target: { value: 'teks lama' } });
    vdom = harness.render();
    expect(findElement(vdom, el => el.type === 'textarea' || el.type === 'input').props.value).toBe('teks lama');

    // Click mic button to start a new voice session
    const micButton = findElement(vdom, el => el.props && el.props['aria-label'] === 'Perintah Suara');
    micButton.props.onClick();
    vdom = harness.render();

    // Input must be cleared upon starting new voice recognition
    expect(findElement(vdom, el => el.type === 'textarea' || el.type === 'input').props.value).toBe('');
    expect(mockInstance.start).toHaveBeenCalled();
  });

  it('renders file attachment button in InputBar and RAG quick action preset', () => {
    render(<InputBar onSendMessage={vi.fn()} host="Word" />);
    expect(screen.getByTitle(/Lampirkan berkas referensi/i)).toBeInTheDocument();

    render(<QuickActionPresets host="Word" onSelectPreset={vi.fn()} />);
    expect(screen.getByText(/Analisis Berkas Referensi/i)).toBeInTheDocument();
  });
});

describe('ChatContainer Component', () => {
  const sampleMessages: ChatMessage[] = [
    {
      id: 'm1',
      role: 'user',
      content: 'Halo Sam!',
      timestamp: 1000,
    },
    {
      id: 'm2',
      role: 'assistant',
      content: 'Halo! Ada yang bisa saya bantu di spreadsheet Anda?',
      timestamp: 1001,
      toolCalls: [
        {
          id: 'tc-chat-1',
          name: 'format_range',
          arguments: { range: 'A1:D1', bold: true },
          status: 'pending',
        },
      ],
    },
    {
      id: 'm3',
      role: 'user',
      content: 'Tolong tebalkan judul tabel.',
      timestamp: 1002,
    },
  ];

  it('renders empty message container without error', () => {
    const html = renderToString(<ChatContainer messages={[]} onApplyToolCall={() => {}} />);
    expect(html).toContain('flex-1 overflow-y-auto');
  });

  it('renders message bubbles for all provided messages in order', () => {
    const html = renderToString(
      <ChatContainer messages={sampleMessages} onApplyToolCall={() => {}} />
    );
    expect(html).toContain('Halo Sam!');
    expect(html).toContain('Halo! Ada yang bisa saya bantu di spreadsheet Anda?');
    expect(html).toContain('Tolong tebalkan judul tabel.');
    expect(html).toContain('format range');
  });

  it('propagates onApplyToolCall from messages to container callback', () => {
    const onApplyToolCall = vi.fn();
    const harness = createHookHarness(ChatContainer, {
      messages: sampleMessages,
      onApplyToolCall,
    });
    const vdom = harness.render();

    // Find the MessageBubble for the message with toolCalls
    const messageBubbles = findAllElements(vdom, el => el.type === MessageBubble);
    expect(messageBubbles.length).toBe(3);

    const assistantBubble = messageBubbles.find(
      b => b.props.message.id === 'm2'
    );
    expect(assistantBubble).toBeDefined();
    expect(assistantBubble.props.onApplyToolCall).toBe(onApplyToolCall);

    // Render the assistant bubble and simulate tool apply
    const bubbleHarness = createHookHarness(MessageBubble, assistantBubble.props);
    const bubbleVdom = bubbleHarness.render();
    const actionCard = findElement(bubbleVdom, el => el.type === ActionCard);
    expect(actionCard).toBeDefined();

    actionCard.props.onApply(actionCard.props.toolCall);
    expect(onApplyToolCall).toHaveBeenCalledTimes(1);
    expect(onApplyToolCall).toHaveBeenCalledWith(sampleMessages[1].toolCalls![0]);
  });

  it('invokes scrollIntoView on bottomRef on render with messages', () => {
    const scrollIntoView = vi.fn();
    const harness = createHookHarness(ChatContainer, {
      messages: sampleMessages,
      onApplyToolCall: () => {},
    });

    // Initial render
    const vdom = harness.render();
    expect(vdom).toBeDefined();

    // The states array stores the ref object at index 0
    const states = harness.getStates();
    expect(states.length).toBeGreaterThan(0);
    const refSlot = states[0];
    expect(refSlot).toHaveProperty('current');

    // Attach mock DOM node with scrollIntoView
    refSlot.current = { scrollIntoView };

    // Re-render to trigger useEffect hook
    harness.render({
      messages: [
        ...sampleMessages,
        { id: 'm4', role: 'user', content: 'Pesan baru', timestamp: 1003 },
      ],
      onApplyToolCall: () => {},
    });

    expect(scrollIntoView).toHaveBeenCalledWith({ behavior: 'smooth' });
  });

  it('propagates onConfirmAction and onCancelAction to MessageBubble', () => {
    const onConfirmAction = vi.fn();
    const onCancelAction = vi.fn();
    const sampleMsgWithProp: ChatMessage = {
      id: 'm-prop',
      role: 'assistant',
      content: 'Ada proposal',
      timestamp: 1000,
      pendingAction: {
        id: 'p-1',
        actionType: 'modify_cells',
        description: 'Ubah formula',
        affectedCellsCount: 12,
        payload: {},
      },
    };

    const harness = createHookHarness(ChatContainer, {
      messages: [sampleMsgWithProp],
      onApplyToolCall: () => {},
      onConfirmAction,
      onCancelAction,
    });
    const vdom = harness.render();

    const bubble = findElement(vdom, el => el.type === MessageBubble);
    expect(bubble).toBeDefined();
    expect(bubble.props.onConfirmAction).toBe(onConfirmAction);
    expect(bubble.props.onCancelAction).toBe(onCancelAction);
  });
});

describe('QuickActionPresets Component', () => {
  it('renders QuickActionPresets for Excel and triggers preset prompt on click', () => {
    const onSelectPreset = vi.fn();
    const harness = createHookHarness(QuickActionPresets, {
      host: 'Excel',
      onSelectPreset,
      disabled: false,
    });
    const vdom = harness.render();

    const buttons = findAllElements(vdom, el => el.type === 'button');
    expect(buttons.length).toBe(9);

    const cleanButton = buttons.find(b => typeof b.props.children === 'string' && b.props.children.includes('Bersihkan Data'));
    expect(cleanButton).toBeDefined();

    cleanButton.props.onClick();
    expect(onSelectPreset).toHaveBeenCalledWith(
      'Bersihkan data di tabel aktif: hapus duplikat, rapikan spasi berlebih, dan isi sel yang kosong.'
    );

    const html = renderToString(<QuickActionPresets host="Excel" onSelectPreset={onSelectPreset} />);
    expect(html).toContain('Bersihkan Data');
    expect(html).toContain('Rekap &amp; Chart');
    expect(html).toContain('Format &amp; Color Scale');
    expect(html).toContain('Bagikan ke Hub');
  });

  it('renders new Excel presets for Audit Formula and Analisis Tren', () => {
    const onSelect = vi.fn();
    const harness = createHookHarness(QuickActionPresets, {
      host: 'Excel',
      onSelectPreset: onSelect,
      disabled: false,
    });
    const vdom = harness.render();

    const buttons = findAllElements(vdom, el => el.type === 'button');
    const auditBtn = buttons.find(b => typeof b.props.children === 'string' && b.props.children.includes('Audit Formula & Error'));
    const trendBtn = buttons.find(b => typeof b.props.children === 'string' && b.props.children.includes('Buat Analisis Tren'));
    const hubBtn = buttons.find(b => typeof b.props.children === 'string' && b.props.children.includes('Bagikan ke Hub'));

    expect(auditBtn).toBeDefined();
    expect(trendBtn).toBeDefined();
    expect(hubBtn).toBeDefined();

    auditBtn.props.onClick();
    expect(onSelect).toHaveBeenCalledWith(
      'Audit seluruh formula pada sheet ini: cari error #REF!, #VALUE!, #DIV/0!, atau sel dengan rumus tidak konsisten dan berikan rekomendasinya.'
    );

    trendBtn.props.onClick();
    expect(onSelect).toHaveBeenCalledWith(
      'Analisis data pada tabel aktif ini: buatkan ringkasan eksekutif, tren pertumbuhan, dan insight temuan utama.'
    );

    hubBtn.props.onClick();
    expect(onSelect).toHaveBeenCalledWith(
      'Kirim ringkasan data dan tabel aktif ini ke Universal Hub agar bisa langsung dipakai di Word atau PowerPoint.'
    );

    const html = renderToString(<QuickActionPresets host="Excel" onSelectPreset={onSelect} />);
    expect(html).toContain('Audit Formula &amp; Error');
    expect(html).toContain('Buat Analisis Tren');
    expect(html).toContain('Bagikan ke Hub');
  });

  it('renders QuickActionPresets for Word with correct labels and prompts', () => {
    const onSelectPreset = vi.fn();
    const harness = createHookHarness(QuickActionPresets, {
      host: 'Word',
      onSelectPreset,
      disabled: false,
    });
    const vdom = harness.render();

    const buttons = findAllElements(vdom, el => el.type === 'button');
    expect(buttons.length).toBe(10);

    const momButton = buttons.find(b => typeof b.props.children === 'string' && b.props.children.includes('Notulen Rapat (MoM)'));
    expect(momButton).toBeDefined();
    momButton.props.onClick();
    expect(onSelectPreset).toHaveBeenCalledWith(
      'Buatkan dokumen terstruktur Notulen Rapat (Minutes of Meeting) lengkap dengan agenda, pembahasan, dan action items.'
    );

    const html = renderToString(<QuickActionPresets host="Word" onSelectPreset={onSelectPreset} />);
    expect(html).toContain('Notulen Rapat (MoM)');
    expect(html).toContain('Buat SOP');
    expect(html).toContain('Poles Bahasa &amp; EYD');
    expect(html).toContain('Impor Data dari Excel Hub');
  });

  it('renders new Word presets for Review Kontrak, Format Brand Korporat, and Hub Import', () => {
    const onSelect = vi.fn();
    const harness = createHookHarness(QuickActionPresets, {
      host: 'Word',
      onSelectPreset: onSelect,
      disabled: false,
    });
    const vdom = harness.render();

    const buttons = findAllElements(vdom, el => el.type === 'button');
    const contractBtn = buttons.find(b => typeof b.props.children === 'string' && b.props.children.includes('Review Kontrak & Risiko'));
    const corporateBtn = buttons.find(b => typeof b.props.children === 'string' && b.props.children.includes('Format Brand Korporat'));
    const importHubBtn = buttons.find(b => typeof b.props.children === 'string' && b.props.children.includes('Impor Data dari Excel Hub'));

    expect(contractBtn).toBeDefined();
    expect(corporateBtn).toBeDefined();
    expect(importHubBtn).toBeDefined();

    contractBtn.props.onClick();
    expect(onSelect).toHaveBeenCalledWith(
      'Audit dan review kepatuhan klausul pada dokumen/kontrak ini: periksa SLA, denda, termin pembayaran, klausul risiko tinggi, dan berikan rekomendasi perbaikan.'
    );

    corporateBtn.props.onClick();
    expect(onSelect).toHaveBeenCalledWith(
      'Terapkan standarisasi gaya dan format korporat profesional (Corporate Navy) pada seluruh dokumen ini: tata hierarki heading, font, dan spasi yang rapi.'
    );

    importHubBtn.props.onClick();
    expect(onSelect).toHaveBeenCalledWith(
      'Ambil snapshot data terbaru dari Excel di Universal Hub, lalu buatkan tabel dan draf naskah laporannya di dokumen ini.'
    );

    const html = renderToString(<QuickActionPresets host="Word" onSelectPreset={onSelect} />);
    expect(html).toContain('Review Kontrak &amp; Risiko');
    expect(html).toContain('Format Brand Korporat');
    expect(html).toContain('Impor Data dari Excel Hub');
  });

  it('renders QuickActionPresets for PowerPoint with correct labels and prompts', () => {
    const onSelectPreset = vi.fn();
    const harness = createHookHarness(QuickActionPresets, {
      host: 'PowerPoint',
      onSelectPreset,
      disabled: false,
    });
    const vdom = harness.render();

    const buttons = findAllElements(vdom, el => el.type === 'button');
    expect(buttons.length).toBe(9);

    const deckButton = buttons.find(b => typeof b.props.children === 'string' && b.props.children.includes('Buat Deck 3 Slide'));
    expect(deckButton).toBeDefined();
    deckButton.props.onClick();
    expect(onSelectPreset).toHaveBeenCalledWith(
      'Buatkan 3 slide presentasi terstruktur: Slide 1 Judul Cover, Slide 2 Poin Materi Utama, Slide 3 Key Metrics Capaian.'
    );

    const pptHubBtn = buttons.find(b => typeof b.props.children === 'string' && b.props.children.includes('Buat Slide dari Hub'));
    expect(pptHubBtn).toBeDefined();
    pptHubBtn.props.onClick();
    expect(onSelectPreset).toHaveBeenCalledWith(
      'Ambil data atau ringkasan terbaru dari Universal Hub dan buatkan 5 slide presentasi eksekutif lengkap dengan speaker notes.'
    );

    const html = renderToString(<QuickActionPresets host="PowerPoint" onSelectPreset={onSelectPreset} />);
    expect(html).toContain('Buat Deck 3 Slide');
    expect(html).toContain('Ringkas Seluruh Slide');
    expect(html).toContain('Catatan Pemateri');
    expect(html).toContain('Dokumen ke Slide');
    expect(html).toContain('Buat Slide dari Hub');
  });

  it('renders ppt_doc_to_deck preset in PowerPoint host mode', () => {
    const onSelectPreset = vi.fn();
    const harness = createHookHarness(QuickActionPresets, {
      host: 'PowerPoint',
      onSelectPreset,
      disabled: false,
    });
    const vdom = harness.render();

    const buttons = findAllElements(vdom, el => el.type === 'button');
    const docToDeckBtn = buttons.find(b => typeof b.props.children === 'string' && b.props.children.includes('Dokumen ke Slide'));
    expect(docToDeckBtn).toBeDefined();
    docToDeckBtn.props.onClick();
    expect(onSelectPreset).toHaveBeenCalledWith(
      'Ubah teks/dokumen laporan ini menjadi 5 slide presentasi eksekutif terstruktur lengkap dengan naskah pembicara (speaker notes).'
    );

    render(<QuickActionPresets host="PowerPoint" onSelectPreset={vi.fn()} />);
    expect(screen.getByText(/Dokumen ke Slide/i)).toBeInTheDocument();
  });

  it('disables all preset buttons when disabled prop is true', () => {
    const onSelectPreset = vi.fn();
    const harness = createHookHarness(QuickActionPresets, {
      host: 'Excel',
      onSelectPreset,
      disabled: true,
    });
    const vdom = harness.render();

    const buttons = findAllElements(vdom, el => el.type === 'button');
    expect(buttons.every(b => b.props.disabled === true)).toBe(true);

    const html = renderToString(<QuickActionPresets host="Excel" onSelectPreset={onSelectPreset} disabled={true} />);
    expect(html).toContain('disabled=""');
  });

  it('renders web research presets for Excel, Word, and PowerPoint hosts', () => {
    const { rerender } = render(<QuickActionPresets host="Excel" onSelectPreset={vi.fn()} />);
    expect(screen.getByText(/Riset Web & Buat Tabel/i)).toBeInTheDocument();

    rerender(<QuickActionPresets host="Word" onSelectPreset={vi.fn()} />);
    expect(screen.getByText(/Riset Web & Tulis Laporan/i)).toBeInTheDocument();

    rerender(<QuickActionPresets host="PowerPoint" onSelectPreset={vi.fn()} />);
    expect(screen.getByText(/Riset Data & Buat Slide/i)).toBeInTheDocument();
  });

  it('renders flowchart preset in Word and PowerPoint hosts', () => {
    const { rerender } = render(
      <QuickActionPresets host="Word" onSelectPreset={vi.fn()} />
    );
    expect(screen.getByText(/Buat Diagram Alur/i)).toBeInTheDocument();

    rerender(
      <QuickActionPresets host="PowerPoint" onSelectPreset={vi.fn()} />
    );
    expect(screen.getByText(/Slide Flowchart Alur/i)).toBeInTheDocument();
  });

  it('renders universal quick action preset "Pelajari Preferensi Ini" across hosts', () => {
    const hosts: ('Excel' | 'Word' | 'PowerPoint' | 'BrowserDev')[] = ['Excel', 'Word', 'PowerPoint', 'BrowserDev'];
    for (const host of hosts) {
      const onSelect = vi.fn();
      const html = renderToString(<QuickActionPresets host={host} onSelectPreset={onSelect} />);
      expect(html).toContain('Pelajari Preferensi Ini');
    }
  });
});

describe('InputBar & QuickActionPresets Integration', () => {
  it('renders QuickActionPresets in InputBar and clicking preset triggers onSendMessage', () => {
    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, {
      onSendMessage,
      host: 'Excel',
    });
    const vdom = harness.render();

    const presetsComponent = findElement(vdom, el => el.type === QuickActionPresets);
    expect(presetsComponent).toBeDefined();
    expect(presetsComponent.props.host).toBe('Excel');

    // Simulate preset click
    presetsComponent.props.onSelectPreset('Bersihkan data di tabel aktif');
    expect(onSendMessage).toHaveBeenCalledWith('Bersihkan data di tabel aktif');
  });
});

describe('MessageBubble ReAct Step and Tool Badges', () => {
  it('renders ReAct step badge when stepProgress is defined on message', () => {
    const msg: ChatMessage = {
      id: 'step_msg_1',
      role: 'assistant',
      content: 'Menganalisis langkah...',
      timestamp: 100,
      stepProgress: { step: 2, maxSteps: 6, description: 'Mengeksekusi tool clean_data...' },
    };

    const html = renderToString(
      <MessageBubble message={msg} onApplyToolCall={() => {}} />
    );
    expect(html).toContain('Iterasi ReAct [2/6]');
    expect(html).toContain('Mengeksekusi tool clean_data...');
  });

  it('renders ReAct tool result badge when message role is tool', () => {
    const toolMsg: ChatMessage = {
      id: 'tool_msg_1',
      role: 'tool',
      toolName: 'clean_data',
      content: 'Berhasil membersihkan 15 baris data duplikat.',
      timestamp: 101,
      isError: false,
    };

    const html = renderToString(
      <MessageBubble message={toolMsg} onApplyToolCall={() => {}} />
    );
    expect(html).toContain('Tool Result');
    expect(html).toContain('clean_data');
    expect(html).toContain('Berhasil membersihkan 15 baris data duplikat.');
  });

  it('renders ReAct tool error badge when message role is tool with isError true', () => {
    const toolErrorMsg: ChatMessage = {
      id: 'tool_msg_err',
      role: 'tool',
      toolName: 'apply_conditional_formatting',
      content: 'Error: Range tidak valid',
      timestamp: 102,
      isError: true,
    };

    const html = renderToString(
      <MessageBubble message={toolErrorMsg} onApplyToolCall={() => {}} />
    );
    expect(html).toContain('Tool Error');
    expect(html).toContain('apply_conditional_formatting');
    expect(html).toContain('Error: Range tidak valid');
  });
});

describe('ActionConfirmationCard Component', () => {
  const mockProposal: PendingActionProposal = {
    id: 'prop-123',
    actionType: 'modify_cells',
    description: 'Menulis formula total penjualan dan diskon ke kolom E dan F',
    affectedCellsCount: 42,
    targetRange: 'E2:F22',
    payload: { range: 'E2:F22' },
  };

  it('renders proposal description, affected cells count badge, and target range', () => {
    const html = stripComments(
      renderToString(
        <ActionConfirmationCard
          proposal={mockProposal}
          onConfirm={() => {}}
          onCancel={() => {}}
        />
      )
    );
    expect(html).toContain('Menulis formula total penjualan dan diskon ke kolom E dan F');
    expect(html).toContain('42 sel');
    expect(html).toContain('E2:F22');
    expect(html).toContain('Terapkan Perubahan');
    expect(html).toContain('Batalkan');
  });

  it('triggers onConfirm when apply button is clicked', () => {
    const onConfirm = vi.fn();
    const onCancel = vi.fn();
    const vdom = ActionConfirmationCard({
      proposal: mockProposal,
      onConfirm,
      onCancel,
    }) as React.ReactElement;

    const buttons = findAllElements(vdom, el => el.type === 'button');
    const confirmBtn = buttons.find(b => {
      const text = JSON.stringify(b.props.children);
      return text.includes('Terapkan') || text.includes('Terapkan Perubahan');
    });
    expect(confirmBtn).toBeDefined();
    confirmBtn.props.onClick();

    expect(onConfirm).toHaveBeenCalledTimes(1);
    expect(onConfirm).toHaveBeenCalledWith('prop-123');
    expect(onCancel).not.toHaveBeenCalled();
  });

  it('triggers onCancel when cancel button is clicked', () => {
    const onConfirm = vi.fn();
    const onCancel = vi.fn();
    const vdom = ActionConfirmationCard({
      proposal: mockProposal,
      onConfirm,
      onCancel,
    }) as React.ReactElement;

    const buttons = findAllElements(vdom, el => el.type === 'button');
    const cancelBtn = buttons.find(b => {
      const text = JSON.stringify(b.props.children);
      return text.includes('Batalkan');
    });
    expect(cancelBtn).toBeDefined();
    cancelBtn.props.onClick();

    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(onCancel).toHaveBeenCalledWith('prop-123');
    expect(onConfirm).not.toHaveBeenCalled();
  });

  it('disables buttons when disabled prop is true', () => {
    const html = renderToString(
      <ActionConfirmationCard
        proposal={mockProposal}
        onConfirm={() => {}}
        onCancel={() => {}}
        disabled={true}
      />
    );
    expect(html).toContain('disabled');
  });
});

describe('MessageBubble with pendingAction proposal', () => {
  const proposal: PendingActionProposal = {
    id: 'prop-abc',
    actionType: 'clean_data',
    description: 'Membersihkan spasi ganda dan baris duplikat',
    affectedCellsCount: 15,
    payload: {},
  };

  const messageWithProposal: ChatMessage = {
    id: 'msg-prop-1',
    role: 'assistant',
    content: 'Saya telah menyiapkan aksi pembersihan data. Silakan konfirmasi untuk menerapkan.',
    timestamp: 1000,
    pendingAction: proposal,
  };

  it('renders ActionConfirmationCard when message has pendingAction', () => {
    const onConfirm = vi.fn();
    const onCancel = vi.fn();
    const html = stripComments(
      renderToString(
        <MessageBubble
          message={messageWithProposal}
          onApplyToolCall={() => {}}
          onConfirmAction={onConfirm}
          onCancelAction={onCancel}
        />
      )
    );
    expect(html).toContain('Membersihkan spasi ganda dan baris duplikat');
    expect(html).toContain('15 sel');
    expect(html).toContain('Terapkan Perubahan');
    expect(html).toContain('Batalkan');
  });

  it('passes onConfirm and onCancel callbacks through MessageBubble', () => {
    const onConfirm = vi.fn();
    const onCancel = vi.fn();
    const harness = createHookHarness(MessageBubble, {
      message: messageWithProposal,
      onApplyToolCall: () => {},
      onConfirmAction: onConfirm,
      onCancelAction: onCancel,
    });
    const vdom = harness.render();

    const actionCard = findElement(vdom, el => el.type === ActionConfirmationCard);
    expect(actionCard).toBeDefined();
    expect(actionCard.props.proposal).toBe(proposal);
    expect(actionCard.props.onConfirm).toBe(onConfirm);
    expect(actionCard.props.onCancel).toBe(onCancel);

    actionCard.props.onConfirm('prop-abc');
    expect(onConfirm).toHaveBeenCalledWith('prop-abc');

    actionCard.props.onCancel('prop-abc');
    expect(onCancel).toHaveBeenCalledWith('prop-abc');
  });
});

describe('CrossAppSnapshotCard Component', () => {
  beforeEach(() => {
    clearCrossAppSnapshots();
  });

  it('renders CrossAppSnapshotCard when fresh snapshot from another host exists', () => {
    saveCrossAppSnapshot({
      sourceHost: 'Excel',
      title: 'Rekap Penjualan Q3',
      artifactType: 'table_data',
      summaryText: 'Tabel 15 baris',
    });

    render(
      <ChatContainer
        host="Word"
        messages={[]}
        onSendMessage={vi.fn()}
        onApplyToolCall={vi.fn()}
      />
    );

    expect(screen.getByText(/Rekap Penjualan Q3/i)).toBeInTheDocument();
    expect(screen.getByText(/Ditemukan data dari Excel/i)).toBeInTheDocument();
  });

  it('triggers onSendMessage with appropriate prompt when action button is clicked', () => {
    saveCrossAppSnapshot({
      sourceHost: 'Excel',
      title: 'Rekap Penjualan Q3',
      artifactType: 'table_data',
      summaryText: 'Tabel 15 baris',
    });

    const onSendMessage = vi.fn();
    const harness = createHookHarness(CrossAppSnapshotCard, {
      host: 'Word',
      onSendMessage,
    });
    const vdom = harness.render();

    const buttons = findAllElements(vdom, el => el.type === 'button');
    const actionBtn = buttons.find(b => {
      const txt = typeof b.props.children === 'string' ? b.props.children : JSON.stringify(b.props.children);
      return txt.includes('Buat Laporan & Sisipkan Tabel');
    });
    expect(actionBtn).toBeDefined();

    actionBtn.props.onClick();
    expect(onSendMessage).toHaveBeenCalledTimes(1);
    expect(onSendMessage).toHaveBeenCalledWith(
      'Ambil snapshot data terbaru dari Excel di Universal Hub, lalu buatkan tabel dan draf naskah laporannya di dokumen ini.'
    );
  });

  it('hides card and records dismissal when dismiss button is clicked', () => {
    const snap = saveCrossAppSnapshot({
      sourceHost: 'Excel',
      title: 'Rekap Penjualan Q3',
      artifactType: 'table_data',
      summaryText: 'Tabel 15 baris',
    });

    const harness = createHookHarness(CrossAppSnapshotCard, {
      host: 'Word',
      onSendMessage: vi.fn(),
    });
    let vdom = harness.render();
    expect(vdom).not.toBeNull();

    const buttons = findAllElements(vdom, el => el.type === 'button');
    const dismissBtn = buttons.find(
      b => b.props['aria-label'] === 'Abaikan snapshot' || b.props.children === '✕'
    );
    expect(dismissBtn).toBeDefined();

    dismissBtn.props.onClick();

    vdom = harness.render();
    expect(vdom).toBeNull();
    expect(isCrossAppSnapshotDismissed(snap.id)).toBe(true);
  });
});

