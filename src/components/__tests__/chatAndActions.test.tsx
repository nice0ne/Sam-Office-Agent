import React from 'react';
import { renderToString } from 'react-dom/server';
import { describe, it, expect, vi } from 'vitest';
import { ActionCard } from '../Actions/ActionCard';
import { MessageBubble } from '../Chat/MessageBubble';
import { InputBar } from '../Chat/InputBar';
import { ChatContainer } from '../Chat/ChatContainer';
import { ToolCall, ChatMessage } from '../../types';

// Helper to strip React SSR comments from html output
function stripComments(html: string): string {
  return html.replace(/<!--.*?-->/g, '');
}

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

  it('enables send button when text is entered and sends trimmed text on click', () => {
    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, { onSendMessage });

    // Initial render
    let vdom = harness.render();
    const inputEl = findElement(vdom, el => el.type === 'input');
    const buttonEl = findElement(vdom, el => el.type === 'button' && el.props['aria-label'] === 'Kirim Pesan');
    expect(buttonEl.props.disabled).toBe(true);

    // Simulate typing
    inputEl.props.onChange({ target: { value: '  Buatkan chart dari kolom A ke D  ' } });

    // Re-render after state update
    vdom = harness.render();
    const updatedInputEl = findElement(vdom, el => el.type === 'input');
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
    const resetInputEl = findElement(vdom, el => el.type === 'input');
    const resetButtonEl = findElement(vdom, el => el.type === 'button' && el.props['aria-label'] === 'Kirim Pesan');
    expect(resetInputEl.props.value).toBe('');
    expect(resetButtonEl.props.disabled).toBe(true);
  });

  it('sends message on Enter key without shiftKey', () => {
    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, { onSendMessage });

    let vdom = harness.render();
    const inputEl = findElement(vdom, el => el.type === 'input');

    // Type text
    inputEl.props.onChange({ target: { value: 'Hitung rata-rata C1:C10' } });
    vdom = harness.render();
    const updatedInputEl = findElement(vdom, el => el.type === 'input');

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
    const inputEl = findElement(vdom, el => el.type === 'input');

    inputEl.props.onChange({ target: { value: 'Baris pertama' } });
    vdom = harness.render();
    const updatedInputEl = findElement(vdom, el => el.type === 'input');

    const preventDefault = vi.fn();
    updatedInputEl.props.onKeyDown({ key: 'Enter', shiftKey: true, preventDefault });

    expect(preventDefault).not.toHaveBeenCalled();
    expect(onSendMessage).not.toHaveBeenCalled();
  });

  it('does not send message when input contains only whitespace', () => {
    const onSendMessage = vi.fn();
    const harness = createHookHarness(InputBar, { onSendMessage });

    let vdom = harness.render();
    const inputEl = findElement(vdom, el => el.type === 'input');

    inputEl.props.onChange({ target: { value: '     ' } });
    vdom = harness.render();
    const updatedButtonEl = findElement(vdom, el => el.type === 'button' && el.props['aria-label'] === 'Kirim Pesan');
    const updatedInputEl = findElement(vdom, el => el.type === 'input');

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
    const inputEl = findElement(vdom, el => el.type === 'input');
    const buttonEl = findElement(vdom, el => el.type === 'button');

    expect(inputEl.props.disabled).toBe(true);
    expect(buttonEl.props.disabled).toBe(true);

    // Attempt typing and sending while disabled
    inputEl.props.onChange({ target: { value: 'Test message' } });
    vdom = harness.render();
    const updatedButtonEl = findElement(vdom, el => el.type === 'button');
    const updatedInputEl = findElement(vdom, el => el.type === 'input');

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

    const inputEl = findElement(vdom, el => el.type === 'input');
    expect(inputEl.props.value).toContain('buatkan grafik baris');
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
});
