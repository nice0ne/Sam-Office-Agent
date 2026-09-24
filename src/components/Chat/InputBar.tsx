import React, { useState, KeyboardEvent, useRef, useEffect } from 'react';
import { Send, Sparkles, Loader2, Mic, MicOff, Paperclip, FileText, X } from 'lucide-react';
import { useSpeechRecognition, SpeechLanguage } from '../../hooks/useSpeechRecognition';
import { QuickActionPresets } from './QuickActionPresets';
import { addDocument, removeDocument, listDocuments } from '../../services/rag/ragEngine';
import { HostType } from '../../types';

export interface InputBarProps {
  onSendMessage: (text: string) => void;
  disabled?: boolean;
  placeholder?: string;
  host?: HostType;
}

export const InputBar: React.FC<InputBarProps> = ({
  onSendMessage,
  disabled,
  placeholder = 'Tanya Sam atau perintahkan sesuatu...',
  host = 'Excel',
}) => {
  const [input, setInput] = useState('');
  const [docs, setDocs] = useState(listDocuments());
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const {
    isSupported,
    isListening,
    language,
    setLanguage,
    toggleListening,
    stopListening,
    resetTranscript,
    error: speechError,
  } = useSpeechRecognition({
    initialLanguage: 'id-ID',
    onTranscriptChange: (text) => {
      setInput(text);
    },
  });

  // Auto-adjust height up to 128px
  useEffect(() => {
    if (textareaRef.current && typeof textareaRef.current.scrollHeight === 'number') {
      textareaRef.current.style.height = 'auto';
      const newHeight = Math.min(Math.max(textareaRef.current.scrollHeight, 24), 128);
      textareaRef.current.style.height = `${newHeight}px`;
    }
  }, [input]);

  const handleToggleVoice = () => {
    if (disabled) return;
    if (!isListening) {
      setInput('');
      resetTranscript();
    }
    toggleListening();
  };

  const handleSend = () => {
    if (!input.trim() || disabled) return;
    if (isListening) {
      stopListening();
    }
    onSendMessage(input.trim());
    setInput('');
    resetTranscript();
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
    }
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleToggleLanguage = () => {
    const nextLang: SpeechLanguage = language === 'id-ID' ? 'en-US' : 'id-ID';
    setLanguage(nextLang);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const content = event.target?.result as string;
      if (content) {
        addDocument(file.name, content);
        setDocs(listDocuments());
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  const currentPlaceholder = disabled
    ? 'Sam sedang memproses...'
    : isListening
    ? 'Mendengarkan suara Anda...'
    : placeholder;

  return (
    <div className="px-3 pt-1.5 pb-2 bg-white/95 dark:bg-gray-800/95 backdrop-blur-xs border-t border-gray-200 dark:border-gray-700 transition-colors shrink-0">
      <QuickActionPresets
        host={host}
        onSelectPreset={prompt => onSendMessage(prompt)}
        disabled={disabled}
      />
      {docs.length > 0 && (
        <div className="flex flex-wrap gap-1.5 px-3 py-1.5 bg-slate-50 dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700 mb-1.5 rounded-md">
          {docs.map(doc => (
            <span key={doc.id} className="inline-flex items-center gap-1 px-2 py-0.5 text-xs bg-blue-100 text-blue-800 dark:bg-blue-900 dark:text-blue-200 rounded-md">
              <FileText className="w-3 h-3" />
              <span className="font-medium truncate max-w-[140px]">{doc.name}</span>
              <span className="text-[10px] opacity-75">({doc.totalChunks} chunks)</span>
              <button
                type="button"
                onClick={() => { removeDocument(doc.id); setDocs(listDocuments()); }}
                className="hover:text-red-500 ml-0.5"
                title={`Hapus ${doc.name}`}
              >
                <X className="w-3 h-3" />
              </button>
            </span>
          ))}
        </div>
      )}
      {speechError && (
        <div className="mb-1.5 px-2 py-0.5 text-[11px] rounded bg-red-50 dark:bg-red-900/30 text-red-600 dark:text-red-400 border border-red-200 dark:border-red-800 flex items-center justify-between">
          <span>
            {speechError === 'not-allowed'
              ? 'Izin mikrofon ditolak. Aktifkan izin mikrofon di browser Anda.'
              : speechError === 'not-supported'
              ? 'Web Speech API tidak didukung di browser ini.'
              : `Peringatan suara: ${speechError}`}
          </span>
        </div>
      )}
      <div className="flex items-end gap-1.5 bg-gray-50 dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-lg px-2.5 py-1 focus-within:ring-1.5 focus-within:ring-blue-500 shadow-2xs transition-all">
        <Sparkles className="w-3.5 h-3.5 text-blue-500 shrink-0 mb-1" />
        <textarea
          ref={textareaRef}
          rows={1}
          value={input}
          onChange={e => setInput(e.target.value)}
          onKeyDown={handleKeyDown}
          disabled={Boolean(disabled)}
          placeholder={currentPlaceholder}
          className="w-full bg-transparent text-xs text-gray-900 dark:text-gray-100 placeholder-gray-400 dark:placeholder-gray-500 focus:outline-none disabled:opacity-60 resize-none min-h-[22px] max-h-28 py-0.5 leading-normal"
        />

        {/* Language selector toggle */}
        <button
          type="button"
          onClick={handleToggleLanguage}
          disabled={Boolean(disabled)}
          aria-label="Pilihan Bahasa Suara"
          title={`Ganti bahasa perintah suara (${language === 'id-ID' ? 'Bahasa Indonesia' : 'English'})`}
          className="px-1 py-0.5 mb-0.5 text-[9px] font-semibold rounded bg-gray-200 dark:bg-gray-700 text-gray-700 dark:text-gray-300 hover:bg-gray-300 dark:hover:bg-gray-600 disabled:opacity-40 transition shrink-0"
        >
          {language === 'id-ID' ? 'ID' : 'EN'}
        </button>

        {/* Voice Command Microphone button */}
        <button
          type="button"
          onClick={handleToggleVoice}
          disabled={Boolean(disabled)}
          aria-label="Perintah Suara"
          title={
            !isSupported
              ? 'Browser tidak mendukung perintah suara'
              : isListening
              ? 'Hentikan merekam suara'
              : 'Mulai perintah suara'
          }
          className={`p-1 mb-0.5 rounded-md transition shrink-0 ${
            isListening
              ? 'bg-red-500 text-white animate-pulse shadow-sm shadow-red-500/50'
              : 'text-gray-500 dark:text-gray-400 hover:bg-gray-200 dark:hover:bg-gray-700 disabled:opacity-40'
          }`}
        >
          {isListening ? (
            <MicOff className="w-3.5 h-3.5" />
          ) : (
            <Mic className="w-3.5 h-3.5" />
          )}
        </button>

        {/* Attachment button (Mini-RAG) */}
        <button
          type="button"
          onClick={() => fileInputRef.current?.click()}
          disabled={Boolean(disabled)}
          title="Lampirkan berkas referensi (Mini-RAG)"
          aria-label="Lampirkan berkas referensi"
          className="p-1 mb-0.5 rounded-md text-gray-500 dark:text-gray-400 hover:bg-gray-200 dark:hover:bg-gray-700 disabled:opacity-40 transition shrink-0"
        >
          <Paperclip className="w-4 h-4" />
        </button>
        <input
          ref={fileInputRef}
          type="file"
          accept=".txt,.md,.csv,.json,.html,.log"
          className="hidden"
          onChange={handleFileChange}
        />

        {/* Send message button */}
        <button
          type="button"
          onClick={handleSend}
          disabled={Boolean(!input.trim() || disabled)}
          aria-label="Kirim Pesan"
          className="p-1 mb-0.5 rounded-md text-blue-600 dark:text-blue-400 hover:bg-blue-50 dark:hover:bg-gray-800 disabled:opacity-40 transition shrink-0"
        >
          {disabled ? (
            <Loader2 className="w-3.5 h-3.5 animate-spin text-blue-500" />
          ) : (
            <Send className="w-3.5 h-3.5" />
          )}
        </button>
      </div>
    </div>
  );
};
