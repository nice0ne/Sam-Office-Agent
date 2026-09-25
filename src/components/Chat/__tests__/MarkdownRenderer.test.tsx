import React from 'react';
import { renderToString } from 'react-dom/server';
import { describe, it, expect } from 'vitest';
import { MarkdownRenderer } from '../MarkdownRenderer';

function render(ui: React.ReactElement): string {
  return renderToString(ui).replace(/<!--.*?-->/g, '');
}

describe('MarkdownRenderer', () => {
  it('renders bold, italic, and strikethrough text', () => {
    const content = 'Teks ini **tebal**, teks ini *miring*, dan ini ~~dicoret~~.';
    const html = render(<MarkdownRenderer content={content} />);

    expect(html).toContain('<strong');
    expect(html).toContain('tebal</strong>');
    expect(html).toContain('<em');
    expect(html).toContain('miring</em>');
    expect(html).toContain('<del');
    expect(html).toContain('dicoret</del>');
  });

  it('renders inline code and fenced code blocks with language', () => {
    const content = 'Gunakan fungsi `calculateTotal()` untuk memproses.\n\n```typescript\nconst sum = 100 + 200;\nreturn sum;\n```';
    const html = render(<MarkdownRenderer content={content} />);

    expect(html).toContain('<code');
    expect(html).toContain('calculateTotal()</code>');
    expect(html).toContain('const sum = 100 + 200;');
    expect(html).toContain('typescript');
  });

  it('renders headings level 1 to 4', () => {
    const content = '# Judul Besar\n## Sub Judul\n### Seksi Tiga';
    const html = render(<MarkdownRenderer content={content} />);

    expect(html).toContain('<h1');
    expect(html).toContain('Judul Besar');
    expect(html).toContain('<h2');
    expect(html).toContain('Sub Judul');
    expect(html).toContain('<h3');
    expect(html).toContain('Seksi Tiga');
  });

  it('renders unordered and ordered lists', () => {
    const content = '- Item pertama\n- Item kedua\n\n1. Tahap Satu\n2. Tahap Dua';
    const html = render(<MarkdownRenderer content={content} />);

    expect(html).toContain('<ul');
    expect(html).toContain('Item pertama');
    expect(html).toContain('Item kedua');
    expect(html).toContain('<ol');
    expect(html).toContain('Tahap Satu');
    expect(html).toContain('Tahap Dua');
  });

  it('renders blockquote', () => {
    const content = '> Ini adalah kutipan penting dari dokumen.';
    const html = render(<MarkdownRenderer content={content} />);

    expect(html).toContain('<blockquote');
    expect(html).toContain('Ini adalah kutipan penting');
  });

  it('renders markdown tables correctly', () => {
    const content = '| Nama | Peran | Skor |\n|---|---|---|\n| Alice | Lead | 95 |\n| Bob | Dev | 88 |';
    const html = render(<MarkdownRenderer content={content} />);

    expect(html).toContain('<table');
    expect(html).toContain('<thead');
    expect(html).toContain('<tbody');
    expect(html).toContain('Alice');
    expect(html).toContain('Lead');
    expect(html).toContain('95');
    expect(html).toContain('Bob');
  });

  it('renders markdown links', () => {
    const content = 'Kunjungi [Portal Resmi](https://example.com) untuk info.';
    const html = render(<MarkdownRenderer content={content} />);

    expect(html).toContain('<a');
    expect(html).toContain('href="https://example.com"');
    expect(html).toContain('Portal Resmi');
  });
});
