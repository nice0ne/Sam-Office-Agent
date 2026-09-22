import React, { useRef, useEffect } from 'react';
import { ChatMessage, ToolCall } from '../../types';
import { MessageBubble } from './MessageBubble';

export interface ChatContainerProps {
  messages: ChatMessage[];
  onApplyToolCall: (toolCall: ToolCall) => void;
  isExecuting?: boolean;
  onConfirmAction?: (proposalId: string) => void;
  onCancelAction?: (proposalId: string) => void;
}

export const ChatContainer: React.FC<ChatContainerProps> = ({
  messages,
  onApplyToolCall,
  isExecuting,
  onConfirmAction,
  onCancelAction,
}) => {
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView?.({ behavior: 'smooth' });
  }, [messages]);

  return (
    <div className="flex-1 overflow-y-auto px-3 py-2.5 space-y-2 select-text">
      {messages.map(m => (
        <MessageBubble
          key={m.id}
          message={m}
          onApplyToolCall={onApplyToolCall}
          isExecuting={isExecuting}
          onConfirmAction={onConfirmAction}
          onCancelAction={onCancelAction}
        />
      ))}
      <div ref={bottomRef} />
    </div>
  );
};
