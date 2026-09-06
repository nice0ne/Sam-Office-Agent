import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import React from 'react';
import { useSpeechRecognition } from '../useSpeechRecognition';

class MockSpeechRecognition {
  continuous = false;
  interimResults = false;
  lang = 'id-ID';
  onstart: (() => void) | null = null;
  onend: (() => void) | null = null;
  onerror: ((event: any) => void) | null = null;
  onresult: ((event: any) => void) | null = null;
  start = vi.fn().mockImplementation(() => {
    if (this.onstart) this.onstart();
  });
  stop = vi.fn().mockImplementation(() => {
    if (this.onend) this.onend();
  });
  abort = vi.fn().mockImplementation(() => {
    if (this.onend) this.onend();
  });
}

// Hook test runner helper
function renderSpeechHook(options?: Parameters<typeof useSpeechRecognition>[0]) {
  const states: any[] = [];
  let stateIdx = 0;
  const effectSlots: { effect: () => void | (() => void); deps?: any[] }[] = [];
  let effectIdx = 0;
  const refSlots: any[] = [];
  let refIdx = 0;

  let currentResult: any;

  const dispatcher = {
    useState: (initial: any) => {
      const i = stateIdx++;
      if (!(i in states)) {
        states[i] = typeof initial === 'function' ? initial() : initial;
      }
      const setState = (next: any) => {
        const prev = states[i];
        states[i] = typeof next === 'function' ? next(prev) : next;
      };
      return [states[i], setState];
    },
    useRef: (initial: any) => {
      const i = refIdx++;
      if (!(i in refSlots)) {
        refSlots[i] = { current: initial };
      }
      return refSlots[i];
    },
    useEffect: (effect: () => void | (() => void), deps?: any[]) => {
      const slot = effectSlots[effectIdx];
      let shouldRun = false;
      if (!slot) shouldRun = true;
      else if (!deps) shouldRun = true;
      else if (deps.length !== slot.deps?.length || deps.some((d, idx) => !Object.is(d, slot.deps?.[idx]))) {
        shouldRun = true;
      }
      effectSlots[effectIdx] = { effect, deps };
      effectIdx++;
      if (shouldRun) {
        effect();
      }
    },
    useCallback: (fn: any, _deps?: any[]) => fn,
  };

  const run = (opts = options) => {
    stateIdx = 0;
    effectIdx = 0;
    refIdx = 0;
    const prevDispatcher = (React as any).__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED?.ReactCurrentDispatcher?.current;
    try {
      if ((React as any).__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED) {
        (React as any).__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED.ReactCurrentDispatcher.current = dispatcher;
      }
      currentResult = useSpeechRecognition(opts);
      return currentResult;
    } finally {
      if ((React as any).__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED && prevDispatcher) {
        (React as any).__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED.ReactCurrentDispatcher.current = prevDispatcher;
      }
    }
  };

  run();

  return {
    get result() {
      return currentResult;
    },
    rerender: (opts?: any) => run(opts),
  };
}

describe('useSpeechRecognition', () => {
  const originalWebkit = (globalThis as any).webkitSpeechRecognition;
  const originalSpeech = (globalThis as any).SpeechRecognition;

  beforeEach(() => {
    delete (globalThis as any).webkitSpeechRecognition;
    delete (globalThis as any).SpeechRecognition;
  });

  afterEach(() => {
    (globalThis as any).webkitSpeechRecognition = originalWebkit;
    (globalThis as any).SpeechRecognition = originalSpeech;
  });

  it('reports isSupported as false when Web Speech API is missing', () => {
    const hook = renderSpeechHook();
    expect(hook.result.isSupported).toBe(false);
    expect(hook.result.isListening).toBe(false);
  });

  it('initializes correctly when webkitSpeechRecognition is available', () => {
    (globalThis as any).webkitSpeechRecognition = MockSpeechRecognition;
    const hook = renderSpeechHook();
    expect(hook.result.isSupported).toBe(true);
    expect(hook.result.isListening).toBe(false);
    expect(hook.result.language).toBe('id-ID');
  });

  it('allows starting and stopping recognition', () => {
    (globalThis as any).webkitSpeechRecognition = MockSpeechRecognition;
    const hook = renderSpeechHook();

    hook.result.startListening();
    hook.rerender();
    expect(hook.result.isListening).toBe(true);

    hook.result.stopListening();
    hook.rerender();
    expect(hook.result.isListening).toBe(false);
  });

  it('toggles language between id-ID and en-US', () => {
    (globalThis as any).webkitSpeechRecognition = MockSpeechRecognition;
    const hook = renderSpeechHook();

    expect(hook.result.language).toBe('id-ID');
    hook.result.setLanguage('en-US');
    hook.rerender();
    expect(hook.result.language).toBe('en-US');
  });

  it('handles incoming speech recognition results', () => {
    let mockInstance: MockSpeechRecognition | null = null;
    class SpyingMockSpeechRecognition extends MockSpeechRecognition {
      constructor() {
        super();
        mockInstance = this;
      }
    }
    (globalThis as any).webkitSpeechRecognition = SpyingMockSpeechRecognition;

    const onTranscriptChange = vi.fn();
    const hook = renderSpeechHook({ onTranscriptChange });

    hook.result.startListening();
    hook.rerender();

    expect(mockInstance).not.toBeNull();

    // Simulate interim result
    mockInstance!.onresult!({
      resultIndex: 0,
      results: [
        {
          0: { transcript: 'buat tabel' },
          isFinal: false,
          length: 1,
        },
      ],
    });
    hook.rerender();

    expect(hook.result.interimTranscript).toBe('buat tabel');
    expect(onTranscriptChange).toHaveBeenCalledWith('buat tabel', false);

    // Simulate final result
    mockInstance!.onresult!({
      resultIndex: 0,
      results: [
        {
          0: { transcript: 'buat tabel penjualan' },
          isFinal: true,
          length: 1,
        },
      ],
    });
    hook.rerender();

    expect(hook.result.transcript).toBe('buat tabel penjualan');
    expect(onTranscriptChange).toHaveBeenCalledWith('buat tabel penjualan', true);
  });

  it('handles errors gracefully like not-allowed', () => {
    let mockInstance: MockSpeechRecognition | null = null;
    class SpyingMockSpeechRecognition extends MockSpeechRecognition {
      constructor() {
        super();
        mockInstance = this;
      }
    }
    (globalThis as any).webkitSpeechRecognition = SpyingMockSpeechRecognition;

    const hook = renderSpeechHook();

    hook.result.startListening();
    hook.rerender();

    mockInstance!.onerror!({ error: 'not-allowed' });
    hook.rerender();

    expect(hook.result.error).toBe('not-allowed');
    expect(hook.result.isListening).toBe(false);
  });
});
