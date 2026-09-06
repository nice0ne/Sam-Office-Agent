import { useState, useEffect, useRef, useCallback } from 'react';

export type SpeechLanguage = 'id-ID' | 'en-US';

export interface UseSpeechRecognitionOptions {
  initialLanguage?: SpeechLanguage;
  onTranscriptChange?: (text: string, isFinal: boolean) => void;
  continuous?: boolean;
  interimResults?: boolean;
}

export interface UseSpeechRecognitionReturn {
  isSupported: boolean;
  isListening: boolean;
  transcript: string;
  interimTranscript: string;
  error: string | null;
  language: SpeechLanguage;
  startListening: () => void;
  stopListening: () => void;
  toggleListening: () => void;
  setLanguage: (lang: SpeechLanguage) => void;
  resetTranscript: () => void;
}

// Get SpeechRecognition constructor across browsers / Chromium WebView2
function getSpeechRecognitionClass(): any {
  if (typeof window !== 'undefined') {
    return (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition || null;
  }
  if (typeof globalThis !== 'undefined') {
    return (globalThis as any).SpeechRecognition || (globalThis as any).webkitSpeechRecognition || null;
  }
  return null;
}

export function useSpeechRecognition({
  initialLanguage = 'id-ID',
  onTranscriptChange,
  continuous = true,
  interimResults = true,
}: UseSpeechRecognitionOptions = {}): UseSpeechRecognitionReturn {
  const [language, setLanguageState] = useState<SpeechLanguage>(initialLanguage);
  const [isListening, setIsListening] = useState(false);
  const [transcript, setTranscript] = useState('');
  const [interimTranscript, setInterimTranscript] = useState('');
  const [error, setError] = useState<string | null>(null);

  const recognitionRef = useRef<any>(null);
  const isListeningRef = useRef(false);
  const onTranscriptChangeRef = useRef(onTranscriptChange);

  useEffect(() => {
    onTranscriptChangeRef.current = onTranscriptChange;
  }, [onTranscriptChange]);

  const SpeechRecognitionClass = getSpeechRecognitionClass();
  const isSupported = Boolean(SpeechRecognitionClass);

  const stopListening = useCallback(() => {
    isListeningRef.current = false;
    setIsListening(false);
    if (recognitionRef.current) {
      try {
        recognitionRef.current.stop();
      } catch (e) {
        // Ignore if already stopped
      }
    }
  }, []);

  const startListening = useCallback(() => {
    if (!isSupported) {
      setError('not-supported');
      return;
    }

    setError(null);

    try {
      if (recognitionRef.current) {
        try {
          recognitionRef.current.abort();
        } catch (e) {
          // Ignore
        }
      }

      const recognition = new SpeechRecognitionClass();
      recognition.continuous = continuous;
      recognition.interimResults = interimResults;
      recognition.lang = language;

      recognition.onstart = () => {
        isListeningRef.current = true;
        setIsListening(true);
      };

      recognition.onerror = (event: any) => {
        isListeningRef.current = false;
        setIsListening(false);
        setError(event.error || 'error');
      };

      recognition.onend = () => {
        isListeningRef.current = false;
        setIsListening(false);
      };

      recognition.onresult = (event: any) => {
        let currentInterim = '';
        let currentFinal = '';

        for (let i = event.resultIndex; i < event.results.length; ++i) {
          const item = event.results[i];
          if (item.isFinal) {
            currentFinal += item[0]?.transcript || '';
          } else {
            currentInterim += item[0]?.transcript || '';
          }
        }

        if (currentFinal) {
          setTranscript(prev => {
            const next = prev ? `${prev} ${currentFinal.trim()}` : currentFinal.trim();
            onTranscriptChangeRef.current?.(next, true);
            return next;
          });
          setInterimTranscript('');
        } else if (currentInterim) {
          setInterimTranscript(currentInterim);
          onTranscriptChangeRef.current?.(currentInterim, false);
        }
      };

      recognitionRef.current = recognition;
      recognition.start();
      isListeningRef.current = true;
      setIsListening(true);
    } catch (err: any) {
      isListeningRef.current = false;
      setIsListening(false);
      setError(err?.message || 'start-failed');
    }
  }, [isSupported, language, continuous, interimResults, SpeechRecognitionClass]);

  const toggleListening = useCallback(() => {
    if (isListeningRef.current) {
      stopListening();
    } else {
      startListening();
    }
  }, [stopListening, startListening]);

  const setLanguage = useCallback(
    (newLang: SpeechLanguage) => {
      setLanguageState(newLang);
      if (isListeningRef.current) {
        stopListening();
      }
    },
    [stopListening]
  );

  const resetTranscript = useCallback(() => {
    setTranscript('');
    setInterimTranscript('');
  }, []);

  useEffect(() => {
    return () => {
      if (recognitionRef.current) {
        try {
          recognitionRef.current.abort();
        } catch (e) {
          // Cleanup
        }
      }
    };
  }, []);

  return {
    isSupported,
    isListening,
    transcript,
    interimTranscript,
    error,
    language,
    startListening,
    stopListening,
    toggleListening,
    setLanguage,
    resetTranscript,
  };
}
