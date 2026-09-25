import React, { useState } from 'react';
import { ChatMessage, ToolCall } from '../../types';
import { ActionCard } from '../Actions/ActionCard';
import { ActionConfirmationCard } from './ActionConfirmationCard';
import { MarkdownRenderer } from './MarkdownRenderer';
import { Bot, User, Copy, Check, Download, Wrench } from 'lucide-react';

export interface MessageBubbleProps {
  message: ChatMessage;
  onApplyToolCall: (toolCall: ToolCall) => void;
  isExecuting?: boolean;
  onConfirmAction?: (proposalId: string) => void;
  onCancelAction?: (proposalId: string) => void;
}

export const MessageBubble: React.FC<MessageBubbleProps> = ({
  message,
  onApplyToolCall,
  isExecuting,
  onConfirmAction,
  onCancelAction,
}) => {
  const isUser = message.role === 'user';
  const isTool = message.role === 'tool';
  const displayText = message.displayContent || message.content;
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(displayText);
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

  const handleDownload = () => {
    try {
      if (typeof document === 'undefined') return;
      const lower = message.content.toLowerCase();
      let filename = 'Dokumen.md';
      if (lower.includes('prd') || lower.includes('product requirement')) {
        filename = 'PRD-Document.md';
      } else if (lower.includes('summary') || lower.includes('ringkasan')) {
        filename = 'Summary-Dokumen.md';
      } else if (lower.includes('# ')) {
        filename = 'Dokumen-Laporan.md';
      } else {
        filename = 'Catatan-Sam.md';
      }

      const blob = new Blob([message.content], { type: 'text/markdown;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.warn('Gagal mengunduh file:', err);
    }
  };

  const isThinking = !isUser && !isTool && !displayText && (!message.toolCalls || message.toolCalls.length === 0);

  return (
    <div className={`flex gap-2 mb-3 group ${isUser ? 'justify-end' : 'justify-start'}`}>
      {!isUser && (
        <div
          className={`w-6 h-6 rounded-full ${
            isTool
              ? 'bg-emerald-600 text-white'
              : 'bg-blue-600 text-white'
          } flex items-center justify-center shrink-0 mt-1 shadow-sm transition-all ${
            isThinking ? 'ring-2 ring-blue-400/60 ring-offset-1 dark:ring-offset-gray-900 animate-pulse' : ''
          }`}
        >
          {isTool ? <Wrench className="w-3.5 h-3.5" /> : <Bot className="w-3.5 h-3.5" />}
        </div>
      )}

      <div
        className={`max-w-[85%] rounded-xl px-3.5 py-2.5 text-xs relative select-text transition-all ${
          isUser
            ? 'bg-blue-600 text-white rounded-tr-none shadow-sm'
            : isTool
            ? 'bg-gray-50 dark:bg-gray-850 border border-gray-200 dark:border-gray-700/80 text-gray-800 dark:text-gray-200 rounded-tl-none shadow-sm font-mono text-[11px]'
            : 'bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 text-gray-800 dark:text-gray-200 rounded-tl-none shadow-sm'
        }`}
      >
        {/* ReAct Step Progress Badge */}
        {message.stepProgress && (
          <div className="inline-flex items-center gap-1.5 px-2 py-0.5 mb-2 rounded-full text-[10px] font-medium bg-blue-50 dark:bg-blue-900/30 text-blue-700 dark:text-blue-300 border border-blue-200 dark:border-blue-800 select-none">
            <span className="w-1.5 h-1.5 rounded-full bg-blue-500 animate-pulse" />
            <span>{`Iterasi ReAct [${message.stepProgress.step}/${message.stepProgress.maxSteps}]`}</span>
            {message.stepProgress.description && (
              <span className="text-gray-500 dark:text-gray-400">· {message.stepProgress.description}</span>
            )}
          </div>
        )}

        {/* ReAct Tool Result / Error Header */}
        {isTool && (
          <div className="flex items-center gap-1.5 mb-1.5 text-[11px] font-medium text-gray-500 dark:text-gray-400">
            <span
              className={`inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-semibold ${
                message.isError
                  ? 'bg-red-100 text-red-700 dark:bg-red-900/40 dark:text-red-300'
                  : 'bg-emerald-100 text-emerald-700 dark:bg-emerald-900/40 dark:text-emerald-300'
              }`}
            >
              {message.isError ? 'Tool Error' : 'Tool Result'}
            </span>
            {message.toolName && (
              <span className="font-mono text-[10px] text-gray-600 dark:text-gray-300">
                {message.toolName}
              </span>
            )}
          </div>
        )}

        {isThinking ? (
          <div className="flex items-center gap-2.5 py-1 px-1">
            <div className="flex items-center space-x-1.5">
              <span className="w-2 h-2 rounded-full bg-blue-500 dark:bg-blue-400 animate-dot-1" />
              <span className="w-2 h-2 rounded-full bg-blue-500 dark:bg-blue-400 animate-dot-2" />
              <span className="w-2 h-2 rounded-full bg-blue-500 dark:bg-blue-400 animate-dot-3" />
            </div>
            <span className="text-[11px] font-medium text-gray-500 dark:text-gray-400 italic animate-pulse">
              Sam sedang berpikir...
            </span>
          </div>
        ) : (
          <div className="flex items-start justify-between gap-2">
            <div className="leading-relaxed select-text cursor-text flex-1">
              {isTool ? (
                <div className="whitespace-pre-wrap font-mono text-[11px]">{displayText}</div>
              ) : (
                <MarkdownRenderer
                  content={displayText}
                  className={
                    isUser
                      ? 'text-white [&_strong]:text-white [&_h1]:text-white [&_h2]:text-white [&_h3]:text-white [&_h4]:text-white [&_code]:bg-blue-700/60 [&_code]:text-white'
                      : ''
                  }
                />
              )}
            </div>

            {displayText && (
              <div className="flex items-center gap-0.5 shrink-0">
                {!isUser && (
                  <button
                    type="button"
                    onClick={handleDownload}
                    title="Unduh sebagai file Markdown (.md)"
                    aria-label="Unduh file"
                    className="p-1 rounded transition-opacity opacity-70 hover:opacity-100 hover:bg-gray-100 dark:hover:bg-gray-700 text-gray-400 hover:text-gray-600 dark:hover:text-gray-200"
                  >
                    <Download className="w-3.5 h-3.5" />
                  </button>
                )}
                <button
                  type="button"
                  onClick={handleCopy}
                  title={copied ? "Tersalin!" : "Salin pesan"}
                  aria-label={copied ? "Tersalin" : "Salin pesan"}
                  className={`p-1 rounded transition-opacity opacity-70 hover:opacity-100 ${
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
              </div>
            )}
          </div>
        )}

        {message.error && (
          <div className="mt-1 text-red-500 font-medium text-[11px] select-text">{message.error}</div>
        )}

        {message.pendingAction && (
          <ActionConfirmationCard
            proposal={message.pendingAction}
            onConfirm={onConfirmAction || (() => {})}
            onCancel={onCancelAction || (() => {})}
            disabled={isExecuting}
          />
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
