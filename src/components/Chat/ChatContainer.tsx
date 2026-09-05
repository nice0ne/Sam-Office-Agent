import React, { useRef, useEffect } from 'react';
import { ChatMessage, ToolCall } from '../../types';
import { MessageBubble } from './MessageBubble';

export interface ChatContainerProps {
  messages: ChatMessage[];
  onApplyToolCall: (toolCall: ToolCall) => void;
  isExecuting?: boolean;
}

export const ChatContainer: React.FC<ChatContainerProps> = ({
  messages,
  onApplyToolCall,
  isExecuting,
}) => {
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView?.({ behavior: 'smooth' });
  }, [messages]);

  return (
    <div className="flex-1 overflow-y-auto p-4 space-y-2">
      {messages.map(m => (
        <MessageBubble
          key={m.id}
          message={m}
          onApplyToolCall={onApplyToolCall}
          isExecuting={isExecuting}
        />
      ))}
      <div ref={bottomRef} />
    </div>
  );
};
