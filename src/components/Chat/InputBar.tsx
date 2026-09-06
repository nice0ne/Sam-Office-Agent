import React, { useState, KeyboardEvent } from 'react';
import { Send, Sparkles, Loader2, Mic, MicOff } from 'lucide-react';
import { useSpeechRecognition, SpeechLanguage } from '../../hooks/useSpeechRecognition';

export interface InputBarProps {
  onSendMessage: (text: string) => void;
  disabled?: boolean;
  placeholder?: string;
}

export const InputBar: React.FC<InputBarProps> = ({
  onSendMessage,
  disabled,
  placeholder = 'Tanya Sam atau perintahkan sesuatu...',
}) => {
  const [input, setInput] = useState('');

  const {
    isSupported,
    isListening,
    language,
    setLanguage,
    toggleListening,
    stopListening,
    error: speechError,
  } = useSpeechRecognition({
    initialLanguage: 'id-ID',
    onTranscriptChange: (text) => {
      setInput(text);
    },
  });

  const handleSend = () => {
    if (!input.trim() || disabled) return;
    if (isListening) {
      stopListening();
    }
    onSendMessage(input.trim());
    setInput('');
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleToggleLanguage = () => {
    const nextLang: SpeechLanguage = language === 'id-ID' ? 'en-US' : 'id-ID';
    setLanguage(nextLang);
  };

  const currentPlaceholder = disabled
    ? 'Sam sedang memproses...'
    : isListening
    ? 'Mendengarkan suara Anda...'
    : placeholder;

  return (
    <div className="p-3 bg-white dark:bg-gray-800 border-t border-gray-200 dark:border-gray-700 transition-colors">
      {speechError && (
        <div className="mb-2 px-2.5 py-1 text-xs rounded-lg bg-red-50 dark:bg-red-900/30 text-red-600 dark:text-red-400 border border-red-200 dark:border-red-800 flex items-center justify-between">
          <span>
            {speechError === 'not-allowed'
              ? 'Izin mikrofon ditolak. Aktifkan izin mikrofon di browser Anda.'
              : speechError === 'not-supported'
              ? 'Web Speech API tidak didukung di browser ini.'
              : `Peringatan suara: ${speechError}`}
          </span>
        </div>
      )}
      <div className="flex items-center gap-2 bg-gray-50 dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-xl px-3 py-1.5 focus-within:ring-2 focus-within:ring-blue-500 shadow-sm transition-all">
        <Sparkles className="w-4 h-4 text-blue-500 shrink-0" />
        <input
          type="text"
          value={input}
          onChange={e => setInput(e.target.value)}
          onKeyDown={handleKeyDown}
          disabled={Boolean(disabled)}
          placeholder={currentPlaceholder}
          className="w-full bg-transparent text-xs text-gray-900 dark:text-gray-100 placeholder-gray-400 dark:placeholder-gray-500 focus:outline-none disabled:opacity-60"
        />

        {/* Language selector toggle */}
        <button
          type="button"
          onClick={handleToggleLanguage}
          disabled={Boolean(disabled)}
          aria-label="Pilihan Bahasa Suara"
          title={`Ganti bahasa perintah suara (${language === 'id-ID' ? 'Bahasa Indonesia' : 'English'})`}
          className="px-1.5 py-0.5 text-[10px] font-semibold rounded bg-gray-200 dark:bg-gray-700 text-gray-700 dark:text-gray-300 hover:bg-gray-300 dark:hover:bg-gray-600 disabled:opacity-40 transition shrink-0"
        >
          {language === 'id-ID' ? 'ID' : 'EN'}
        </button>

        {/* Voice Command Microphone button */}
        <button
          type="button"
          onClick={toggleListening}
          disabled={Boolean(disabled)}
          aria-label="Perintah Suara"
          title={
            !isSupported
              ? 'Browser tidak mendukung perintah suara'
              : isListening
              ? 'Hentikan merekam suara'
              : 'Mulai perintah suara'
          }
          className={`p-1.5 rounded-lg transition shrink-0 ${
            isListening
              ? 'bg-red-500 text-white animate-pulse shadow-sm shadow-red-500/50'
              : 'text-gray-500 dark:text-gray-400 hover:bg-gray-200 dark:hover:bg-gray-700 disabled:opacity-40'
          }`}
        >
          {isListening ? (
            <MicOff className="w-4 h-4" />
          ) : (
            <Mic className="w-4 h-4" />
          )}
        </button>

        {/* Send message button */}
        <button
          type="button"
          onClick={handleSend}
          disabled={Boolean(!input.trim() || disabled)}
          aria-label="Kirim Pesan"
          className="p-1.5 rounded-lg text-blue-600 dark:text-blue-400 hover:bg-blue-50 dark:hover:bg-gray-800 disabled:opacity-40 transition shrink-0"
        >
          {disabled ? (
            <Loader2 className="w-4 h-4 animate-spin text-blue-500" />
          ) : (
            <Send className="w-4 h-4" />
          )}
        </button>
      </div>
    </div>
  );
};
