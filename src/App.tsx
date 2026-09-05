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
import { ThemeMode, getStoredThemeMode, setStoredThemeMode, applyTheme } from './utils/theme';
import { compressTableContext } from './utils/contextCompressor';
import { AgentContext } from './agents/types';

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
  const [themeMode, setThemeMode] = useState<ThemeMode>(getStoredThemeMode);

  useEffect(() => {
    applyTheme(themeMode);
  }, [themeMode]);

  const handleToggleTheme = () => {
    const nextTheme: ThemeMode =
      themeMode === 'auto' ? 'light' : themeMode === 'light' ? 'dark' : 'auto';
    setThemeMode(nextTheme);
    setStoredThemeMode(nextTheme);
    applyTheme(nextTheme);
  };

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

    let activeContext: AgentContext = { host };
    try {
      const driver = getOfficeDriver(host);
      if (host === 'Excel') {
        const sheetData = driver.readActiveSheetData ? await driver.readActiveSheetData() : await driver.readActiveRange();
        if (sheetData && sheetData.values && sheetData.values.length > 0) {
          const summary = compressTableContext(sheetData.values, 30);
          activeContext = {
            host,
            activeCellOrRange: `${(sheetData as any).sheetName || 'Sheet'}!${sheetData.address}`,
            documentSummary: `Data Lembar Kerja (${(sheetData as any).sheetName || 'Sheet'}!${sheetData.address}):\n${summary.summaryText}\nNilai Data Terbaca:\n${JSON.stringify(sheetData.values.slice(0, 30))}`,
          };
        }
      } else if (host === 'Word') {
        const outline = await driver.getWordOutline?.();
        if (outline && outline !== 'Dokumen Word Kosong.') {
          activeContext = {
            host,
            documentSummary: `Isi Dokumen Word:\n${outline.slice(0, 3000)}`,
          };
        }
      } else if (host === 'PowerPoint') {
        const slide = await driver.getSlideContext?.();
        if (slide) {
          activeContext = {
            host,
            documentSummary: `Slide #${slide.slideNumber}: "${slide.title}"\n${slide.textContent}`,
          };
        }
      }
    } catch (e) {
      console.warn('Gagal membaca konteks dokumen:', e);
    }

    const systemPrompt = coordinator.buildSystemPrompt(host, activeContext);
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
        themeMode={themeMode}
        onToggleTheme={handleToggleTheme}
      />

      <ChatContainer
        messages={messages}
        onApplyToolCall={handleApplyToolCall}
        isExecuting={isBusy}
      />

      <InputBar onSendMessage={handleSendMessage} disabled={isBusy} />

      <SettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        onSaved={() => setMessages(prev => [...prev])}
      />
    </div>
  );
};
