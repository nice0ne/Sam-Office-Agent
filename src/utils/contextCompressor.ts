export interface TableSummary {
  headers: string[];
  sampledRows: any[][];
  totalRows: number;
  totalColumns: number;
  summaryText: string;
}

export function compressTableContext(data: any[][], maxSampleRows = 5): TableSummary {
  if (!data || data.length === 0) {
    return {
      headers: [],
      sampledRows: [],
      totalRows: 0,
      totalColumns: 0,
      summaryText: 'Tabel kosong.',
    };
  }

  const headers = data[0].map(h => String(h || ''));
  const rows = data.slice(1);
  const totalRows = rows.length;
  const totalColumns = headers.length;
  const sampledRows = rows.slice(0, maxSampleRows);

  const summaryText = `Tabel berukuran ${totalColumns} kolom x ${totalRows} baris data. (Total baris: ${totalRows}). Header: [${headers.join(', ')}].`;

  return {
    headers,
    sampledRows,
    totalRows,
    totalColumns,
    summaryText,
  };
}
