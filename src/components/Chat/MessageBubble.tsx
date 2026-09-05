import React from 'react';
import { ChatMessage, ToolCall } from '../../types';
import { ActionCard } from '../Actions/ActionCard';
import { Bot, User } from 'lucide-react';

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

  return (
    <div className={`flex gap-2 mb-3 ${isUser ? 'justify-end' : 'justify-start'}`}>
      {!isUser && (
        <div className="w-6 h-6 rounded-full bg-blue-600 text-white flex items-center justify-center shrink-0 mt-1">
          <Bot className="w-3.5 h-3.5" />
        </div>
      )}

      <div
        className={`max-w-[85%] rounded-xl px-3.5 py-2.5 text-xs ${
          isUser
            ? 'bg-blue-600 text-white rounded-tr-none'
            : 'bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 text-gray-800 dark:text-gray-200 rounded-tl-none shadow-sm'
        }`}
      >
        <div className="whitespace-pre-wrap leading-relaxed">{message.content}</div>

        {message.error && (
          <div className="mt-1 text-red-500 font-medium text-[11px]">{message.error}</div>
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
