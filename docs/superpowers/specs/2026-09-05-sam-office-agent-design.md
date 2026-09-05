# Design Specification: Sam-Office-Agent

- **Date**: 2026-09-05
- **Topic**: Universal Microsoft Office Multi-Agent Assistant (Excel, Word, PowerPoint) with BYOK
- **Status**: Approved
- **Repository**: `D:/VIBE-CODING/Sam-Office-Agent`

[English](2026-09-05-sam-office-agent-design.en.md) | Bahasa Indonesia

---

## 1. Executive Summary & Objective

**Sam-Office-Agent** adalah asisten AI produktivitas kantor berbasis **Microsoft Office Web Add-in** yang berjalan di panel samping (*Taskpane*) aplikasi Microsoft Office (**Excel, Word, PowerPoint**) baik versi Desktop Office 365 maupun Office Online (Web).

Sistem mengadopsi prinsip **BYOK (Bring Your Own Key)** yang memungkinkan pengguna memasukkan API Key LLM mereka sendiri (OpenAI, Google Gemini, Anthropic Claude, GLM Coding Global / Zhipu AI, OpenRouter/DeepSeek, dan Local Ollama) tanpa server perantara pihak ketiga. Seluruh pemrosesan berjalan langsung di sisi peramban/taskpane (*client-side sandbox*), menjadikannya sangat ringan (RAM 50–120 MB), aman, cepat, dan terisolasi.

---

## 2. Core Architecture & Tech Stack

### 2.1 Technology Stack
* **Framework**: React 18+ dengan TypeScript.
* **Build Tool**: Vite (mendukung HTTPS dev server lokal untuk Office Add-in).
* **Styling**: Tailwind CSS + Shadcn UI / Radix UI + Lucide Icons (mendukung Light dan Dark theme modern).
* **Office Integration SDK**: `@types/office-js` dan `Office.js` resmi dari Microsoft CDN.
* **Testing**: Vitest untuk unit test (LLM adapter, validator logika) dan Mock Office Driver untuk pengujian di browser biasa.

### 2.2 Host Auto-Detection
Add-in menggunakan satu file `manifest.xml` universal yang didaftarkan ke Word, Excel, dan PowerPoint. Saat taskpane terbuka:
1. `Office.onReady()` mengecek `Office.context.host`.
2. UI dan modul agen otomatis menyesuaikan:
   * Jika di **Excel**: Mengaktifkan `ExcelAgent`, badge hijau Excel, dan toolbar formula/tabel.
   * Jika di **Word**: Mengaktifkan `WordAgent`, badge biru Word, dan toolbar teks/dokumen.
   * Jika di **PowerPoint**: Mengaktifkan `PPTAgent`, badge jingga PPT, dan toolbar slide/outline.
   * Jika di **Browser biasa (Dev)**: Mengaktifkan `MockOfficeDriver` untuk simulasi lokal tanpa Office.

### 2.3 File & Directory Structure
```text
Sam-Office-Agent/
├── manifest.xml                 # Office Add-in manifest mendaftarkan taskpane untuk Excel, Word, PPT
├── package.json                 # Project dependencies & scripts
├── vite.config.ts               # Vite configuration (HTTPS certs, path aliases)
├── tsconfig.json                # TypeScript strict configuration
├── docs/
│   └── superpowers/
│       └── specs/
│           └── 2026-09-05-sam-office-agent-design.md
├── src/
│   ├── index.html               # Entry point HTML yang memuat script Office.js
│   ├── main.tsx                 # Office.onReady() bootstrap & React root
│   ├── components/
│   │   ├── Chat/                # Chat container, MessageBubble, StreamingCursor, InputBar
│   │   ├── Actions/             # ActionCard (Diff preview, step status, Apply Button)
│   │   ├── Settings/            # Modal konfigurasi BYOK, Test Connection, Model selector
│   │   └── Common/              # Header, HostBadge, ThemeToggle, ModeToggle (Copilot/Autopilot)
│   ├── agents/
│   │   ├── types.ts             # Definisi Agent, ToolDefinition, ToolCall, ExecutionContext
│   │   ├── coordinator/         # Sam Coordinator (Intent parser, task decomposition, memory)
│   │   ├── excel/               # Excel Specialist Agent & tool implementations
│   │   ├── word/                # Word Specialist Agent & tool implementations
│   │   └── powerpoint/          # PowerPoint Specialist Agent & tool implementations
│   ├── services/
│   │   ├── llm/                 # Unified BYOK Adapter
│   │   │   ├── types.ts         # ILLMProvider interface
│   │   │   ├── openai.ts        # OpenAI adapter (GPT-4o, mini, o3-mini)
│   │   │   ├── gemini.ts        # Gemini adapter (Gemini 2.0 Flash, 1.5 Pro)
│   │   │   ├── anthropic.ts     # Claude adapter (Claude 3.5 Sonnet, Haiku)
│   │   │   ├── glm.ts           # GLM Coding Global (Zhipu BigModel API)
│   │   │   ├── openrouter.ts    # OpenRouter & DeepSeek API
│   │   │   └── ollama.ts        # Local Ollama / Custom Base URL adapter
│   │   ├── office/              # Office.js Safe Driver (Batch sync executor)
│   │   │   ├── excelDriver.ts   # Excel.run batching
│   │   │   ├── wordDriver.ts    # Word.run batching
│   │   │   ├── pptDriver.ts     # PowerPoint.run batching
│   │   │   └── mockDriver.ts    # Browser standalone mock driver
│   │   └── storage/             # localStorage helper untuk API Keys & chat history
│   └── utils/
│       ├── formulaValidator.ts  # Heuristic formula syntax checker
│       └── contextCompressor.ts # Sampling data tabel/dokumen untuk efisiensi token
```

---

## 3. Multi-Agent Specialization & Tool Calling

### 3.1 Coordinator Agent ("Sam")
* **Fokus**: Antarmuka ramah pengguna, pengenalan intent, pembagian tugas, dan memori obrolan.
* **Alur**:
  1. Menerima instruksi pengguna.
  2. Mengidentifikasi apakah instruksi adalah pertanyaan konsultasi atau perintah manipulasi dokumen.
  3. Mengidentifikasi kompleksitas tugas:
     * **Single-step**: Langsung memanggil agen spesialis terkait.
     * **Multi-step**: Menyusun `Action Queue` (daftar tugas berantai dengan status: *Pending*, *In Progress*, *Completed*).

### 3.2 Agen Spesialis & Definisi Tools

#### A. Excel Specialist Agent
* `get_active_range()`: Membaca alamat sel aktif, ukuran tabel, nama header kolom, dan sampel 5 baris pertama data.
* `write_cells(range: string, values?: any[][], formulas?: string[][])`: Menulis data atau rumus Excel (misal: `=SUM(B2:D2)`, `=XLOOKUP(...)`).
* `format_range(range: string, styles: CellStyle)`: Mengatur gaya sel (zebra banding), font tebal, border, dan format angka (IDR/USD/Persentase).
* `create_chart(chartType: string, dataRange: string, title?: string)`: Membuat grafik batang, garis, pie, atau kolom otomatis.
* `manage_sheet(action: 'add' | 'rename' | 'activate', name: string)`: Manajemen lembar kerja.

#### B. Word Specialist Agent
* `get_selection_or_outline()`: Membaca teks yang sedang diseleksi atau ringkasan struktur dokumen (heading H1/H2).
* `insert_content(position: 'start' | 'end' | 'cursor', text: string, type: 'paragraph' | 'h1' | 'h2' | 'bullet')`: Menyisipkan teks dengan format terstruktur.
* `replace_selection(newText: string)`: Mengganti seleksi teks dengan versi yang telah direvisi/diterjemahkan/dirapikan.
* `insert_table(rows: number, cols: number, data: string[][])`: Membuat dan mengisi tabel formal di Word.
* `format_text(styles: TextStyle)`: Mengubah format teks (bold, italic, font size, warna).

#### C. PowerPoint Specialist Agent
* `get_slide_context()`: Membaca judul, nomor slide aktif, dan teks di dalam slide yang sedang dibuka.
* `add_slide(layout: 'TitleOnly' | 'TitleAndContent' | 'TwoContent')`: Menambahkan slide presentasi baru.
* `insert_slide_content(title: string, bullets: string[])`: Mengisi judul dan butir-butir materi presentasi.
* `set_speaker_notes(notes: string)`: Menambahkan catatan pembicara (*speaker notes*) untuk pemaparan materi.

---

## 4. BYOK Provider Engine & Adapter

### 4.1 Antarmuka Terpadu (`ILLMProvider`)
Semua provider mengimplementasikan interface:
```typescript
export interface ChatRequest {
  messages: Array<{ role: 'system' | 'user' | 'assistant'; content: string }>;
  tools?: Array<ToolDefinition>;
  temperature?: number;
}

export interface StreamEvent {
  type: 'content_delta' | 'tool_call' | 'done' | 'error';
  delta?: string;
  toolCall?: { id: string; name: string; arguments: Record<string, any> };
  error?: string;
}

export interface ILLMProvider {
  id: string;
  name: string;
  sendMessage(req: ChatRequest, config: ProviderConfig): AsyncIterable<StreamEvent>;
  testConnection(config: ProviderConfig): Promise<{ success: boolean; message: string }>;
}
```

### 4.2 Dukungan Provider
1. **Google Gemini**: Model `gemini-2.0-flash`, `gemini-1.5-pro` (Endpoint: `generativelanguage.googleapis.com`).
2. **OpenAI**: Model `gpt-4o`, `gpt-4o-mini`, `o3-mini` (Endpoint: `api.openai.com/v1`).
3. **Anthropic Claude**: Model `claude-3-5-sonnet`, `claude-3-5-haiku` (Header: `anthropic-dangerous-direct-browser-access: true`).
4. **GLM Coding Global (Zhipu AI)**: Model `glm-4-plus`, `glm-4-air`, `codegeex-4` (Endpoint: `open.bigmodel.cn/api/paas/v4`).
5. **OpenRouter / DeepSeek**: Model `deepseek/deepseek-chat`, `deepseek/deepseek-r1` (Endpoint: `openrouter.ai/api/v1`).
6. **Local LLM / Custom**: Mendukung Ollama (`localhost:11434`), LM Studio, dan Custom Base URL / Reverse Proxy.

---

## 5. Performance, Execution, & Security

### 5.1 Mode Eksekusi (Copilot vs Autopilot)
* **Mode Copilot (Default)**: Agen menampilkan kartu preview tindakan (*Action Card*) di panel chat yang berisi deskripsi perubahan, cuplikan diff, dan tombol **"Terapkan ke Dokumen"**.
* **Mode Autopilot**: Agen langsung mengeksekusi tindakan ke dokumen tanpa menunggu klik manual (ideal untuk otomatisasi pemformatan atau pengisian formula massal).
* Pengguna dapat mengganti mode secara instan lewat toggle di bagian atas taskpane.

### 5.2 Optimasi Kecepatan (Batch `context.sync()`)
Untuk mencegah *lag* atau *freezing* pada aplikasi Office:
* Operasi manipulasi sel/paragraf dikelompokkan ke dalam 1 memory transaction.
* Pemanggilan `await context.sync()` hanya dilakukan **1 kali per batch eksekusi**.

### 5.3 Keamanan & Privasi
* **Zero Middleman**: Seluruh request LLM dikirim langsung dari browser WebView2 ke provider via HTTPS / TLS 1.3.
* **Local Storage Only**: API Key hanya disimpan di `localStorage` peramban komputer lokal.
* **Sandbox Office**: Add-in terisolasi penuh di dalam sandbox Microsoft Office, tidak memiliki izin untuk mengakses sistem file disk OS Windows.
* **Native Undo**: Setiap aksi batch mendaftarkan state undo sehingga pengguna bisa menekan `Ctrl+Z` kapan saja.

---

## 6. Testing & Quality Assurance Plan

1. **Vitest Unit Tests**:
   * Pengujian serialisasi request & response parser untuk setiap provider LLM.
   * Pengujian heuristic formula validator (mendeteksi formula yang salah sintaks atau kurung tidak seimbang).
   * Pengujian router host context (Word, Excel, PowerPoint).
2. **Mock Office Environment**:
   * Simulasi state lembar kerja Excel (array 2D cell), dokumen Word (array paragraph), dan slide PPT di browser biasa untuk kecepatan development harian.
3. **Office Manifest Validation**:
   * Validasi XML manifest via CLI Microsoft (`npx office-addin-manifest validate manifest.xml`) untuk memastikan kompatibilitas penuh sebelum instalasi/sideload.
