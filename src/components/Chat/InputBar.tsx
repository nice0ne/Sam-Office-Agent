import React, { useState, KeyboardEvent } from 'react';
import { Send, Sparkles, Loader2 } from 'lucide-react';

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

  const handleSend = () => {
    if (!input.trim() || disabled) return;
    onSendMessage(input.trim());
    setInput('');
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const currentPlaceholder = disabled ? 'Sam sedang memproses...' : placeholder;

  return (
    <div className="p-3 bg-white dark:bg-gray-800 border-t border-gray-200 dark:border-gray-700 transition-colors">
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
        <button
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
