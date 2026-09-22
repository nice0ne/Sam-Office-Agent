import React, { useState, useEffect } from 'react';
import { Header } from './components/Common/Header';
import { ChatContainer } from './components/Chat/ChatContainer';
import { InputBar } from './components/Chat/InputBar';
import { SettingsModal } from './components/Settings/SettingsModal';
import { ChatMessage, ExecutionMode, HostType, ToolCall } from './types';
import { getActiveProvider, setActiveProvider, getExecutionMode, setExecutionMode } from './services/storage/settingsStorage';
import { SamCoordinator } from './agents/coordinator/samCoordinator';
import { ReActExecutionEngine } from './agents/coordinator/reactEngine';
import { documentContextCache } from './services/office/contextCache';
import { getLLMProvider } from './services/llm/factory';
import { getOfficeDriver } from './services/office';
import { ThemeMode, getStoredThemeMode, setStoredThemeMode, applyTheme, initThemeListener } from './utils/theme';
import { compressTableContext } from './utils/contextCompressor';
import { AgentContext } from './agents/types';

const isWritingTool = (name: string): boolean => {
  const readOnly = [
    'read_sheet',
    'read_active_range',
    'get_document_outline',
    'read_slides',
    'get_slide_context',
  ];
  return !readOnly.includes(name) && !name.startsWith('read_') && !name.startsWith('get_');
};

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
  const [reactEngine] = useState(() => new ReActExecutionEngine());
  const [isBusy, setIsBusy] = useState(false);
  const [themeMode, setThemeMode] = useState<ThemeMode>(getStoredThemeMode);

  useEffect(() => {
    applyTheme(themeMode);
    const unsubscribe = initThemeListener(() => themeMode);
    return () => unsubscribe();
  }, [themeMode]);

  const handleSelectTheme = (nextTheme: ThemeMode) => {
    setThemeMode(nextTheme);
    setStoredThemeMode(nextTheme);
    applyTheme(nextTheme);
  };

  const handleToggleTheme = () => {
    const nextTheme: ThemeMode =
      themeMode === 'auto' ? 'light' : themeMode === 'light' ? 'dark' : 'auto';
    handleSelectTheme(nextTheme);
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
    let execResult: { success: boolean; result?: any; error?: string } | null = null;
    try {
      const specialist = coordinator.getSpecialist(host);
      toolCall.status = 'pending';
      setMessages(prev => [...prev]);
      execResult = await specialist.executeTool(toolCall, { host });
      toolCall.status = execResult.success ? 'applied' : 'failed';
      toolCall.error = execResult.error;

      // Invalidate document context snapshot cache if writing tool succeeds
      if (execResult.success && isWritingTool(toolCall.name)) {
        documentContextCache.invalidate(host);
      }
      setMessages(prev => [...prev]);
    } finally {
      setIsBusy(false);
    }

    // Agentic continuation: if read_slides was executed, automatically continue to generate summary/answer
    if (execResult?.success && toolCall.name === 'read_slides' && execResult.result) {
      const readContent = typeof execResult.result === 'string' ? execResult.result : JSON.stringify(execResult.result);
      const lastUserMsg = [...messages].reverse().find(m => m.role === 'user');
      const userGoal = lastUserMsg?.content || 'Buatkan ringkasan materi slide ini';
      const prompt = `${readContent}\n\nIsi materi slide di atas telah berhasil dibaca langsung dari PowerPoint. Silakan selesaikan instruksi saya berdasarkan isi slide tersebut:\n"${userGoal}"`;
      const displayText = `📋 Materi slide berhasil dibaca. Sedang merangkum isi slide...`;

      setTimeout(() => {
        handleSendMessage(prompt, { displayText, isContinuation: true });
      }, 80);
    }

    // Agentic continuation: if read_sheet was executed, automatically continue
    if (execResult?.success && toolCall.name === 'read_sheet' && execResult.result?.values) {
      const readData = execResult.result;
      const rowCount = readData.rowCount || readData.values.length;
      const rangeAddr = readData.address || 'sheet';
      const lastUserMsg = [...messages].reverse().find(m => m.role === 'user');
      const userGoal = lastUserMsg?.content || 'Tulis ringkasan data dan buat visualisasinya';
      const prompt = `[Data Hasil Pembacaan ${readData.sheetName || 'Sheet'}!${rangeAddr} (${rowCount} baris)]:\n${JSON.stringify(readData.values.slice(0, 250))}\n\nData tabel di atas telah berhasil dibaca. Silakan selesaikan instruksi saya:\n"${userGoal}"`;
      const displayText = `📊 Data dari ${readData.sheetName || 'Sheet'}!${rangeAddr} (${rowCount} baris) berhasil dibaca. Memproses...`;

      setTimeout(() => {
        handleSendMessage(prompt, { displayText, isContinuation: true });
      }, 80);
    }
  };

  const handleConfirmAction = async (proposalId: string) => {
    const targetMsg = messages.find(m => m.pendingAction?.id === proposalId);
    if (!targetMsg || !targetMsg.pendingAction) return;

    const proposal = targetMsg.pendingAction;
    setIsBusy(true);

    try {
      const driver = getOfficeDriver(host);
      if (proposal.actionType === 'modify_cells' && proposal.payload) {
        const range = proposal.targetRange || proposal.payload.range || 'A1';
        const { values, formulas } = proposal.payload;
        await driver.writeCells(range, values, formulas);
      } else if (proposal.actionType === 'clean_data') {
        if (driver.cleanData) {
          await driver.cleanData(proposal.payload || {});
        }
      } else if (proposal.actionType === 'format_cells') {
        if (driver.formatRange) {
          const range = proposal.targetRange || proposal.payload?.range || 'A1';
          const styles = proposal.payload?.styles || proposal.payload || {};
          await driver.formatRange(range, styles);
        }
      }

      documentContextCache.invalidate(host);

      setMessages(prev =>
        prev.map(m => {
          if (m.pendingAction?.id === proposalId) {
            const confirmedNote = `\n\n✅ **Aksi Dikonfirmasi & Diterapkan:** ${proposal.description}`;
            return {
              ...m,
              pendingAction: undefined,
              content: m.content ? `${m.content}${confirmedNote}` : confirmedNote.trim(),
            };
          }
          return m;
        })
      );
    } catch (err: any) {
      setMessages(prev =>
        prev.map(m => {
          if (m.pendingAction?.id === proposalId) {
            return {
              ...m,
              pendingAction: undefined,
              error: `Gagal menerapkan aksi: ${err.message || String(err)}`,
            };
          }
          return m;
        })
      );
    } finally {
      setIsBusy(false);
    }
  };

  const handleCancelAction = (proposalId: string) => {
    setMessages(prev =>
      prev.map(m => {
        if (m.pendingAction?.id === proposalId) {
          const desc = m.pendingAction.description;
          const cancelNote = `\n\n❌ **Aksi Dibatalkan:** ${desc}`;
          return {
            ...m,
            pendingAction: undefined,
            content: m.content ? `${m.content}${cancelNote}` : cancelNote.trim(),
          };
        }
        return m;
      })
    );
  };

  const handleSendMessage = async (
    text: string,
    options?: { displayText?: string; isContinuation?: boolean }
  ) => {
    if (!text.trim() || (!options?.isContinuation && isBusy)) return;

    setIsBusy(true);
    const userMsg: ChatMessage = {
      id: `msg_${Date.now()}`,
      role: 'user',
      content: text,
      displayContent: options?.displayText,
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

    const providerConfig = getActiveProvider();
    const provider = getLLMProvider(providerConfig.id);
    const specialist = coordinator.getSpecialist(host);

    let activeContext: AgentContext | null = documentContextCache.get(host);
    if (!activeContext) {
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
            let summaryText = '';
            if (slide.slides && slide.slides.length > 0) {
              const slideDeckDetails = slide.slides
                .map((s) => `[Slide ${s.slideIndex}]: "${s.title}"\n${s.textContent}`)
                .join('\n\n');
              summaryText = `Slide Aktif: Slide #${slide.slideNumber} ("${slide.title}")\n${slide.textContent}\n\n--- DAFTAR SELURUH SLIDE PRESENTASI (${slide.slides.length} SLIDE) ---\n${slideDeckDetails}`;
            } else {
              summaryText = `Slide #${slide.slideNumber}: "${slide.title}"\n${slide.textContent}`;
            }

            activeContext = {
              host,
              documentSummary: summaryText,
            };
          }
        }
        if (activeContext) {
          documentContextCache.set(host, activeContext);
        }
      } catch (e) {
        console.warn('Gagal membaca konteks dokumen:', e);
      }
    }

    if (!activeContext) {
      activeContext = { host };
    }

    const systemPrompt = coordinator.buildSystemPrompt(host, activeContext);
    const tools = specialist.getTools();

    // In Autopilot Mode: Execute through Autonomous ReAct Execution Engine
    if (executionMode === 'autopilot') {
      setMessages(prev => [...prev, userMsg, assistantMsg]);
      try {
        const reactResult = await reactEngine.runLoop(
          {
            provider,
            providerConfig,
            specialist,
            initialMessages: [...messages, userMsg],
            systemPrompt,
            context: activeContext,
            maxIterations: 6,
          },
          {
            onStepProgress: (step, maxSteps, description) => {
              setMessages(prev => {
                const copy = [...prev];
                const last = copy[copy.length - 1];
                if (last && last.role === 'assistant') {
                  last.stepProgress = { step, maxSteps, description };
                }
                return copy;
              });
            },
            onContentDelta: delta => {
              assistantMsg.content += delta;
              setMessages(prev =>
                prev.map(m => (m.id === assistantMsgId ? { ...assistantMsg } : m))
              );
            },
            onToolExecuted: (toolCall, success) => {
              if (success && isWritingTool(toolCall.name)) {
                documentContextCache.invalidate(host);
              }
              // Reflect executed tool state in assistant message
              setMessages(prev =>
                prev.map(m => {
                  if (m.id === assistantMsgId && m.toolCalls) {
                    const match = m.toolCalls.find(tc => tc.id === toolCall.id);
                    if (match) {
                      match.status = toolCall.status;
                      match.error = toolCall.error;
                    }
                  }
                  return m;
                })
              );
            },
          }
        );

        // Update full message history with ReAct conversation
        setMessages(reactResult.allMessages);
      } catch (e: any) {
        assistantMsg.content += `\n[Error: ${e.message}]`;
        setMessages(prev => prev.map(m => (m.id === assistantMsgId ? { ...assistantMsg } : m)));
      } finally {
        setIsBusy(false);
      }
      return;
    }

    // In Copilot Mode: Human-In-The-Loop interactive stream with tool previews
    setMessages(prev => [...prev, userMsg, assistantMsg]);

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
        onSelectProvider={(id) => {
          setActiveProvider(id);
          setMessages(prev => [...prev]);
        }}
      />

      <ChatContainer
        messages={messages}
        onApplyToolCall={handleApplyToolCall}
        isExecuting={isBusy}
        onConfirmAction={handleConfirmAction}
        onCancelAction={handleCancelAction}
      />

      <InputBar onSendMessage={handleSendMessage} disabled={isBusy} host={host} />

      <SettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        onSaved={() => setMessages(prev => [...prev])}
        currentTheme={themeMode}
        onThemeChange={handleSelectTheme}
      />
    </div>
  );
};
