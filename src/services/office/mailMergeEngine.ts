import { MailMergeOptions, MailMergeResult } from './types';

/**
 * Extracts unique placeholder tags matching {{placeholder}} from a template string.
 */
export function extractPlaceholders(templateText: string): string[] {
  if (!templateText) return [];
  const regex = /\{\{\s*([a-zA-Z0-9_\s-]+?)\s*\}\}/g;
  const matches = new Set<string>();
  let match: RegExpExecArray | null;

  while ((match = regex.exec(templateText)) !== null) {
    if (match[1]) {
      matches.add(match[1].trim());
    }
  }

  return Array.from(matches);
}

/**
 * Substitutes {{placeholder}} tags with values from record object case-insensitively.
 * Falls back to '-' if a field is missing or undefined.
 */
export function substitutePlaceholders(templateText: string, record: Record<string, any>): string {
  if (!templateText) return '';
  if (!record) return templateText;

  // Build lowercased lookup map
  const normalizedRecord = new Map<string, any>();
  for (const [key, value] of Object.entries(record)) {
    normalizedRecord.set(key.toLowerCase().trim(), value);
  }

  const regex = /\{\{\s*([a-zA-Z0-9_\s-]+?)\s*\}\}/g;
  return templateText.replace(regex, (_match, placeholder) => {
    const cleanKey = placeholder.trim().toLowerCase();
    if (normalizedRecord.has(cleanKey)) {
      const val = normalizedRecord.get(cleanKey);
      if (val !== null && val !== undefined && val !== '') {
        return String(val);
      }
    }
    return '-';
  });
}

/**
 * Converts tabular headers and rows into an array of structured key-value record objects.
 */
export function mapTableRowsToRecords(
  headers: string[],
  rows: (string | number)[][]
): Array<Record<string, any>> {
  if (!headers || !rows) return [];

  return rows.map((row) => {
    const record: Record<string, any> = {};
    headers.forEach((header, colIdx) => {
      const cleanHeader = header ? header.trim() : `Column${colIdx + 1}`;
      record[cleanHeader] = row[colIdx] !== undefined ? row[colIdx] : '';
    });
    return record;
  });
}

/**
 * Composes multi-document mail merge pages and summary result metrics.
 */
export function composeMailMergeBatch(
  options: MailMergeOptions,
  tableData: { headers: string[]; rows: (string | number)[][] }
): { pages: string[]; result: MailMergeResult } {
  const records = mapTableRowsToRecords(tableData.headers, tableData.rows);
  const pages: string[] = [];
  const recipientNames: string[] = [];

  for (let i = 0; i < records.length; i++) {
    const record = records[i];
    const pageText = substitutePlaceholders(options.templateText, record);
    pages.push(pageText);

    // Resolve recipient name
    let recipient = '';
    if (options.recipientColumnName && record[options.recipientColumnName]) {
      recipient = String(record[options.recipientColumnName]);
    } else {
      // Look for common recipient key names
      const candidateKeys = ['nama', 'name', 'recipient', 'penerima', 'kepada', 'peserta', 'karyawan'];
      for (const [k, v] of Object.entries(record)) {
        if (candidateKeys.includes(k.toLowerCase().trim()) && v) {
          recipient = String(v);
          break;
        }
      }
    }

    if (!recipient) {
      recipient = record[tableData.headers[0]] ? String(record[tableData.headers[0]]) : `Penerima ${i + 1}`;
    }
    recipientNames.push(recipient);
  }

  const result: MailMergeResult = {
    totalRecordsProcessed: records.length,
    totalPagesCreated: pages.length,
    recipientNames,
    summaryMessage: `Berhasil memproses ${records.length} dokumen Mail Merge untuk ${recipientNames.length} penerima.`,
  };

  return { pages, result };
}
