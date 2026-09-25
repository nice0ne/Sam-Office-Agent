import React, { useState } from 'react';
import { Copy, Check } from 'lucide-react';

export interface MarkdownRendererProps {
  content: string;
  className?: string;
}

export const MarkdownRenderer: React.FC<MarkdownRendererProps> = ({ content, className = '' }) => {
  if (!content) return null;

  const blocks = parseMarkdownBlocks(content);

  return (
    <div className={`space-y-1.5 leading-relaxed break-words text-xs ${className}`}>
      {blocks.map((block, idx) => (
        <React.Fragment key={idx}>{renderBlock(block, idx)}</React.Fragment>
      ))}
    </div>
  );
};

interface CodeBlockProps {
  language: string;
  code: string;
}

const CodeBlock: React.FC<CodeBlockProps> = ({ language, code }) => {
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    try {
      if (typeof navigator !== 'undefined' && navigator.clipboard) {
        await navigator.clipboard.writeText(code);
      }
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.warn('Gagal menyalin kode:', err);
    }
  };

  return (
    <div className="my-2 rounded-lg overflow-hidden border border-gray-200 dark:border-gray-700 bg-gray-900 text-gray-100 text-[11px] font-mono shadow-sm">
      <div className="flex items-center justify-between px-3 py-1.5 bg-gray-800/90 border-b border-gray-700 text-gray-400 text-[10px]">
        <span className="font-semibold uppercase tracking-wider">{language || 'text'}</span>
        <button
          type="button"
          onClick={handleCopy}
          className="flex items-center gap-1 hover:text-white transition-colors"
          title="Salin kode"
        >
          {copied ? <Check className="w-3 h-3 text-emerald-400" /> : <Copy className="w-3 h-3" />}
          <span>{copied ? 'Tersalin' : 'Salin'}</span>
        </button>
      </div>
      <pre className="p-3 overflow-x-auto whitespace-pre leading-normal">
        <code>{code}</code>
      </pre>
    </div>
  );
};

type BlockToken =
  | { type: 'code'; language: string; content: string }
  | { type: 'table'; headers: string[]; rows: string[][] }
  | { type: 'heading'; level: number; text: string }
  | { type: 'blockquote'; text: string }
  | { type: 'ul'; items: string[] }
  | { type: 'ol'; items: string[] }
  | { type: 'hr' }
  | { type: 'paragraph'; text: string };

function parseMarkdownBlocks(text: string): BlockToken[] {
  const lines = text.split(/\r?\n/);
  const blocks: BlockToken[] = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // 1. Fenced Code Block
    if (line.trim().startsWith('```')) {
      const language = line.trim().replace(/^```/, '').trim();
      const codeLines: string[] = [];
      i++;
      while (i < lines.length && !lines[i].trim().startsWith('```')) {
        codeLines.push(lines[i]);
        i++;
      }
      i++; // Skip closing ```
      blocks.push({
        type: 'code',
        language,
        content: codeLines.join('\n'),
      });
      continue;
    }

    // 2. Table (| Col 1 | Col 2 |)
    if (
      line.trim().startsWith('|') &&
      i + 1 < lines.length &&
      lines[i + 1].trim().startsWith('|') &&
      lines[i + 1].includes('---')
    ) {
      const parseRow = (rowLine: string) =>
        rowLine
          .split('|')
          .slice(1, -1)
          .map(cell => cell.trim());

      const headers = parseRow(line);
      i += 2; // Skip header and separator line
      const rows: string[][] = [];

      while (i < lines.length && lines[i].trim().startsWith('|')) {
        rows.push(parseRow(lines[i]));
        i++;
      }

      blocks.push({ type: 'table', headers, rows });
      continue;
    }

    // 3. Headings (# H1, ## H2, ### H3, #### H4)
    const headingMatch = line.match(/^(#{1,4})\s+(.+)$/);
    if (headingMatch) {
      blocks.push({
        type: 'heading',
        level: headingMatch[1].length,
        text: headingMatch[2].trim(),
      });
      i++;
      continue;
    }

    // 4. Horizontal Rule (--- or ***)
    if (/^(\-{3,}|\*{3,})$/.test(line.trim())) {
      blocks.push({ type: 'hr' });
      i++;
      continue;
    }

    // 5. Blockquote (> quote)
    if (line.trim().startsWith('>')) {
      const quoteLines: string[] = [];
      while (i < lines.length && lines[i].trim().startsWith('>')) {
        quoteLines.push(lines[i].replace(/^>\s?/, ''));
        i++;
      }
      blocks.push({
        type: 'blockquote',
        text: quoteLines.join('\n'),
      });
      continue;
    }

    // 6. Unordered List (- item or * item)
    if (/^[\*\-]\s+(.+)$/.test(line.trim())) {
      const items: string[] = [];
      while (i < lines.length && /^[\*\-]\s+(.+)$/.test(lines[i].trim())) {
        items.push(lines[i].trim().replace(/^[\*\-]\s+/, ''));
        i++;
      }
      blocks.push({ type: 'ul', items });
      continue;
    }

    // 7. Ordered List (1. item)
    if (/^\d+\.\s+(.+)$/.test(line.trim())) {
      const items: string[] = [];
      while (i < lines.length && /^\d+\.\s+(.+)$/.test(lines[i].trim())) {
        items.push(lines[i].trim().replace(/^\d+\.\s+/, ''));
        i++;
      }
      blocks.push({ type: 'ol', items });
      continue;
    }

    // 8. Empty lines
    if (!line.trim()) {
      i++;
      continue;
    }

    // 9. Regular Paragraph
    const paraLines: string[] = [line];
    i++;
    while (
      i < lines.length &&
      lines[i].trim() &&
      !lines[i].trim().startsWith('```') &&
      !lines[i].trim().startsWith('|') &&
      !lines[i].trim().startsWith('#') &&
      !lines[i].trim().startsWith('>') &&
      !/^[\*\-]\s+/.test(lines[i].trim()) &&
      !/^\d+\.\s+/.test(lines[i].trim()) &&
      !/^(\-{3,}|\*{3,})$/.test(lines[i].trim())
    ) {
      paraLines.push(lines[i]);
      i++;
    }

    blocks.push({
      type: 'paragraph',
      text: paraLines.join('\n'),
    });
  }

  return blocks;
}

function renderBlock(block: BlockToken, key: number): React.ReactNode {
  switch (block.type) {
    case 'code':
      return <CodeBlock key={key} language={block.language} code={block.content} />;

    case 'table':
      return (
        <div key={key} className="my-2 overflow-x-auto rounded-lg border border-gray-200 dark:border-gray-700 shadow-xs">
          <table className="min-w-full divide-y divide-gray-200 dark:divide-gray-700 text-[11px]">
            <thead className="bg-gray-100/80 dark:bg-gray-800 text-gray-700 dark:text-gray-200 font-semibold">
              <tr>
                {block.headers.map((h, hIdx) => (
                  <th key={hIdx} className="px-3 py-1.5 text-left whitespace-nowrap">
                    {renderInline(h)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100 dark:divide-gray-800 bg-white dark:bg-gray-850">
              {block.rows.map((row, rIdx) => (
                <tr key={rIdx} className="even:bg-gray-50/50 dark:even:bg-gray-800/40 hover:bg-blue-50/30 dark:hover:bg-blue-900/20">
                  {row.map((cell, cIdx) => (
                    <td key={cIdx} className="px-3 py-1.5 text-gray-800 dark:text-gray-200">
                      {renderInline(cell)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );

    case 'heading': {
      const headingClasses = [
        '',
        'text-sm font-bold text-gray-900 dark:text-gray-100 mt-2 mb-1 border-b border-gray-200 dark:border-gray-700 pb-0.5',
        'text-xs font-bold text-gray-900 dark:text-gray-100 mt-2 mb-0.5',
        'text-xs font-semibold text-gray-800 dark:text-gray-200 mt-1.5 mb-0.5',
        'text-[11px] font-semibold text-gray-700 dark:text-gray-300 mt-1 mb-0.5',
      ][block.level] || 'text-xs font-bold';

      const Tag = (`h${block.level}` as keyof JSX.IntrinsicElements) || 'h3';
      return (
        <Tag key={key} className={headingClasses}>
          {renderInline(block.text)}
        </Tag>
      );
    }

    case 'blockquote':
      return (
        <blockquote
          key={key}
          className="border-l-2 border-blue-500 pl-2.5 py-0.5 my-1.5 italic text-gray-600 dark:text-gray-300 bg-blue-50/40 dark:bg-blue-950/20 rounded-r text-[11px]"
        >
          {renderInline(block.text)}
        </blockquote>
      );

    case 'ul':
      return (
        <ul key={key} className="list-disc list-outside pl-4 space-y-0.5 my-1 text-gray-800 dark:text-gray-200">
          {block.items.map((it, idx) => (
            <li key={idx}>{renderInline(it)}</li>
          ))}
        </ul>
      );

    case 'ol':
      return (
        <ol key={key} className="list-decimal list-outside pl-4 space-y-0.5 my-1 text-gray-800 dark:text-gray-200">
          {block.items.map((it, idx) => (
            <li key={idx}>{renderInline(it)}</li>
          ))}
        </ol>
      );

    case 'hr':
      return <hr key={key} className="my-2 border-gray-200 dark:border-gray-700" />;

    case 'paragraph':
    default:
      return (
        <p key={key} className="leading-relaxed my-0.5 whitespace-pre-wrap">
          {renderInline(block.text)}
        </p>
      );
  }
}

/**
 * Parses inline elements: `code`, **bold**, *italic*, ~~strikethrough~~, [link](url)
 */
function renderInline(text: string): React.ReactNode {
  if (!text) return null;

  // Split by inline markdown patterns
  // Matches: `inline code`, **bold**, *italic*, ~~strikethrough~~, [link](url)
  const regex = /(`[^`]+`|\*\*[^*]+\*\*|\*[^*]+\*|~~[^~]+~~|\[[^\]]+\]\([^)]+\))/g;
  const parts = text.split(regex);

  return parts.map((part, index) => {
    if (!part) return null;

    // Inline Code
    if (part.startsWith('`') && part.endsWith('`') && part.length >= 2) {
      return (
        <code
          key={index}
          className="px-1 py-0.5 rounded bg-gray-100 dark:bg-gray-800 text-pink-600 dark:text-pink-400 font-mono text-[11px]"
        >
          {part.slice(1, -1)}
        </code>
      );
    }

    // Bold
    if (part.startsWith('**') && part.endsWith('**') && part.length >= 4) {
      return (
        <strong key={index} className="font-semibold text-gray-900 dark:text-white">
          {part.slice(2, -2)}
        </strong>
      );
    }

    // Italic
    if (part.startsWith('*') && part.endsWith('*') && part.length >= 2) {
      return (
        <em key={index} className="italic">
          {part.slice(1, -1)}
        </em>
      );
    }

    // Strikethrough
    if (part.startsWith('~~') && part.endsWith('~~') && part.length >= 4) {
      return (
        <del key={index} className="line-through opacity-75">
          {part.slice(2, -2)}
        </del>
      );
    }

    // Link
    const linkMatch = part.match(/^\[([^\]]+)\]\(([^)]+)\)$/);
    if (linkMatch) {
      return (
        <a
          key={index}
          href={linkMatch[2]}
          target="_blank"
          rel="noopener noreferrer"
          className="text-blue-600 dark:text-blue-400 underline underline-offset-2 hover:opacity-80 transition-opacity"
        >
          {linkMatch[1]}
        </a>
      );
    }

    return <React.Fragment key={index}>{part}</React.Fragment>;
  });
}
