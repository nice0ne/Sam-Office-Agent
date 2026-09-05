import React, { useState } from 'react';
import { ChatMessage, ToolCall } from '../../types';
import { ActionCard } from '../Actions/ActionCard';
import { Bot, User, Copy, Check } from 'lucide-react';

export interface MessageBubbleProps {
  message: ChatMessage;
  onApplyToolCall: (toolCall: ToolCall) => void;
  isExecuting?: boolean;
}

export const MessageBubble: React.FC<MessageBubbleProps> = ({
  message,
  onApplyToolCall,
  isExecuting,
}) => {
  const isUser = message.role === 'user';
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(message.content);
      } else {
        const textArea = document.createElement('textarea');
        textArea.value = message.content;
        textArea.style.position = 'fixed';
        textArea.style.opacity = '0';
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
      }
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.warn('Gagal menyalin pesan:', err);
    }
  };

  return (
    <div className={`flex gap-2 mb-3 group ${isUser ? 'justify-end' : 'justify-start'}`}>
      {!isUser && (
        <div className="w-6 h-6 rounded-full bg-blue-600 text-white flex items-center justify-center shrink-0 mt-1">
          <Bot className="w-3.5 h-3.5" />
        </div>
      )}

      <div
        className={`max-w-[85%] rounded-xl px-3.5 py-2.5 text-xs relative select-text ${
          isUser
            ? 'bg-blue-600 text-white rounded-tr-none'
            : 'bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 text-gray-800 dark:text-gray-200 rounded-tl-none shadow-sm'
        }`}
      >
        <div className="flex items-start justify-between gap-2">
          <div className="whitespace-pre-wrap leading-relaxed select-text cursor-text flex-1">
            {message.content}
          </div>

          {message.content && (
            <button
              type="button"
              onClick={handleCopy}
              title={copied ? "Tersalin!" : "Salin pesan"}
              aria-label={copied ? "Tersalin" : "Salin pesan"}
              className={`shrink-0 p-1 rounded transition-opacity opacity-70 hover:opacity-100 ${
                isUser
                  ? 'hover:bg-blue-700 text-blue-100'
                  : 'hover:bg-gray-100 dark:hover:bg-gray-700 text-gray-400 hover:text-gray-600 dark:hover:text-gray-200'
              }`}
            >
              {copied ? (
                <Check className="w-3.5 h-3.5 text-green-400" />
              ) : (
                <Copy className="w-3.5 h-3.5" />
              )}
            </button>
          )}
        </div>

        {message.error && (
          <div className="mt-1 text-red-500 font-medium text-[11px] select-text">{message.error}</div>
        )}

        {message.toolCalls && message.toolCalls.map(tc => (
          <ActionCard
            key={tc.id}
            toolCall={tc}
            onApply={onApplyToolCall}
            isExecuting={isExecuting}
          />
        ))}
      </div>

      {isUser && (
        <div className="w-6 h-6 rounded-full bg-gray-300 dark:bg-gray-700 text-gray-700 dark:text-gray-300 flex items-center justify-center shrink-0 mt-1">
          <User className="w-3.5 h-3.5" />
        </div>
      )}
    </div>
  );
};
