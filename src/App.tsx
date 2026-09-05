import React, { useState, useEffect } from 'react';
import { Header } from './components/Common/Header';
import { ChatContainer } from './components/Chat/ChatContainer';
import { InputBar } from './components/Chat/InputBar';
import { SettingsModal } from './components/Settings/SettingsModal';
import { ChatMessage, ExecutionMode, HostType, ToolCall } from './types';
import { getActiveProvider, getExecutionMode, setExecutionMode } from './services/storage/settingsStorage';
import { SamCoordinator } from './agents/coordinator/samCoordinator';
import { getLLMProvider } from './services/llm/factory';
import { getOfficeDriver } from './services/office';

export const App: React.FC<{ initialHost?: HostType }> = ({ initialHost = 'Excel' }) => {
  const [host, setHost] = useState<HostType>(initialHost);
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: 'welcome',
      role: 'assistant',
      content: `Halo! Saya Sam, asisten AI untuk Microsoft Office (${initialHost}). Apa yang bisa saya bantu hari ini?`,
      timestamp: Date.now(),
    },
  ]);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [executionMode, setMode] = useState<ExecutionMode>(getExecutionMode());
  const [coordinator] = useState(() => new SamCoordinator());
  const [isBusy, setIsBusy] = useState(false);

  useEffect(() => {
    const office = typeof window !== 'undefined' ? (window as any).Office : (globalThis as any).Office;
    if (office && office.context?.host) {
      const h = office.context.host;
      if (h === 'Excel') setHost('Excel');
      else if (h === 'Word') setHost('Word');
      else if (h === 'PowerPoint') setHost('PowerPoint');
    }
    getOfficeDriver(host);
  }, [host]);

  const handleToggleMode = () => {
    const nextMode = executionMode === 'copilot' ? 'autopilot' : 'copilot';
    setMode(nextMode);
    setExecutionMode(nextMode);
  };

  const handleApplyToolCall = async (toolCall: ToolCall) => {
    setIsBusy(true);
    try {
      const specialist = coordinator.getSpecialist(host);
      toolCall.status = 'pending';
      setMessages(prev => [...prev]);
      const result = await specialist.executeTool(toolCall, { host });
      toolCall.status = result.success ? 'applied' : 'failed';
      toolCall.error = result.error;
      setMessages(prev => [...prev]);
    } finally {
      setIsBusy(false);
    }
  };

  const handleSendMessage = async (text: string) => {
    if (!text.trim() || isBusy) return;

    setIsBusy(true);
    const userMsg: ChatMessage = {
      id: `msg_${Date.now()}`,
      role: 'user',
      content: text,
      timestamp: Date.now(),
    };

    const assistantMsgId = `asst_${Date.now()}`;
    const assistantMsg: ChatMessage = {
      id: assistantMsgId,
      role: 'assistant',
      content: '',
      timestamp: Date.now(),
      toolCalls: [],
    };

    setMessages(prev => [...prev, userMsg, assistantMsg]);

    const providerConfig = getActiveProvider();
    const provider = getLLMProvider(providerConfig.id);
    const specialist = coordinator.getSpecialist(host);
    const systemPrompt = coordinator.buildSystemPrompt(host, { host });
    const tools = specialist.getTools();

    try {
      const stream = provider.sendMessage(
        {
          messages: [...messages, userMsg],
          systemPrompt,
          tools,
        },
        providerConfig
      );

      for await (const chunk of stream) {
        if (chunk.type === 'content_delta' && chunk.delta) {
          assistantMsg.content += chunk.delta;
          setMessages(prev => prev.map(m => (m.id === assistantMsgId ? { ...assistantMsg } : m)));
        } else if (chunk.type === 'tool_call' && chunk.toolCall) {
          assistantMsg.toolCalls?.push(chunk.toolCall);
          setMessages(prev => prev.map(m => (m.id === assistantMsgId ? { ...assistantMsg } : m)));

          if (executionMode === 'autopilot') {
            await handleApplyToolCall(chunk.toolCall);
            setIsBusy(true);
          }
        } else if (chunk.type === 'error') {
          assistantMsg.content += `\n[Error: ${chunk.error}]`;
          setMessages(prev => prev.map(m => (m.id === assistantMsgId ? { ...assistantMsg } : m)));
        }
      }
    } catch (e: any) {
      assistantMsg.content += `\n[Error: ${e.message}]`;
      setMessages(prev => prev.map(m => (m.id === assistantMsgId ? { ...assistantMsg } : m)));
    } finally {
      setIsBusy(false);
    }
  };

  return (
    <div className="flex flex-col h-screen bg-gray-50 dark:bg-gray-900 text-gray-900 dark:text-gray-100">
      <Header
        host={host}
        activeProviderId={getActiveProvider().id}
        executionMode={executionMode}
        onToggleMode={handleToggleMode}
        onOpenSettings={() => setIsSettingsOpen(true)}
      />

      <ChatContainer messages={messages} onApplyToolCall={handleApplyToolCall} />

      <InputBar onSendMessage={handleSendMessage} disabled={isBusy} />

      <SettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        onSaved={() => setMessages(prev => [...prev])}
      />
    </div>
  );
};
