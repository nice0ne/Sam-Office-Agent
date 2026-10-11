import {
  ReconcileOptions,
  ReconcileSummary,
  ReconcileItemDiscrepancy,
  ReconcileItemUnmatched,
} from './types';

/**
 * Parsing angka yang toleran terhadap format mata uang Indonesia,
 * format akuntansi tanda kurung (negatif), pemisah titik/koma, dan simbol mata uang.
 */
export function parseFlexibleNumber(val: any): number {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') {
    return isNaN(val) ? 0 : val;
  }

  let str = String(val).trim();
  if (!str || str === '-' || str.toLowerCase() === 'n/a') return 0;

  let isNegative = false;
  // Accounting parentheses negative: (100.000) or (Rp 100.000,00)
  const parenMatch = str.match(/^\((.*)\)$/);
  if (parenMatch) {
    isNegative = true;
    str = parenMatch[1].trim();
  } else if (str.startsWith('-')) {
    isNegative = true;
    str = str.substring(1).trim();
  }

  // Hapus simbol mata uang dan spasi
  str = str.replace(/^(rp\.?|idr|\$|€|£)\s*/i, '').replace(/\s+/g, '');

  // Deteksi pemisah desimal dan ribuan
  const hasDot = str.includes('.');
  const hasComma = str.includes(',');

  if (hasDot && hasComma) {
    if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
      // Format Indonesia / Eropa: 1.500.000,50
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      // Format US / International: 1,500,000.50
      str = str.replace(/,/g, '');
    }
  } else if (hasDot && !hasComma) {
    // Hanya ada titik. Misal "1.500.000" atau "10.000" (ribuan) vs "15.5" (desimal)
    const parts = str.split('.');
    if (parts.length > 2) {
      // Banyak titik -> pasti pemisah ribuan
      str = str.replace(/\./g, '');
    } else if (parts.length === 2 && parts[1].length === 3 && parts[0].length >= 1) {
      // Tepat 3 angka setelah titik dan bukan angka kecil -> kemungkinan besar ribuan (misal: 10.000 atau 500.000)
      str = str.replace(/\./g, '');
    }
  } else if (hasComma && !hasDot) {
    // Hanya ada koma. Misal "1500,50" atau "1,500,000"
    const parts = str.split(',');
    if (parts.length > 2) {
      str = str.replace(/,/g, '');
    } else if (parts.length === 2 && parts[1].length !== 3) {
      // Desimal koma Indonesia: 1500,50 -> 1500.50
      str = str.replace(',', '.');
    } else if (parts.length === 2 && parts[1].length === 3) {
      // Format ribuan koma
      str = str.replace(/,/g, '');
    }
  }

  const parsed = parseFloat(str);
  if (isNaN(parsed)) return 0;
  return isNegative ? -parsed : parsed;
}

/**
 * Deteksi otomatis kolom ID/Kunci dan kolom Nominal berdasarkan nama header
 */
export function detectKeyAndAmountColumns(headers: string[]): { keyIndex: number; amountIndex: number; descIndex?: number } {
  let keyIndex = -1;
  let amountIndex = -1;
  let descIndex = -1;

  const keyRegex = /(\bno\b|\bref\b|nomor|kode|id\b|invoice|inv|faktur|bukti|transaksi|trx|tiket|nik)/i;
  const amountRegex = /(\bnominal\b|jumlah|total|debit|kredit|saldo|nilai|amount|rupiah|\brp\b|bayar|tagihan|biaya|harga)/i;
  const descRegex = /(ket|keterangan|deskripsi|uraian|nama|rincian|item|barang|uraian)/i;

  headers.forEach((h, idx) => {
    const text = String(h || '').trim();
    if (amountIndex === -1 && amountRegex.test(text)) {
      amountIndex = idx;
    } else if (keyIndex === -1 && keyRegex.test(text)) {
      keyIndex = idx;
    } else if (descIndex === -1 && descRegex.test(text)) {
      descIndex = idx;
    }
  });

  // Fallbacks jika tidak terdeteksi
  if (keyIndex === -1) keyIndex = 0;
  if (amountIndex === -1) {
    amountIndex = headers.length > 1 ? (keyIndex === 0 ? 1 : 0) : 0;
  }

  return { keyIndex, amountIndex, descIndex: descIndex !== -1 ? descIndex : undefined };
}

/**
 * Normalisasi baris data menjadi array objek siap banding
 */
function extractRows(
  table: any[][],
  specifiedKeyCol?: string | number,
  specifiedAmountCol?: string | number
) {
  if (!table || table.length <= 1) return { rows: [], keyIdx: 0, amountIdx: 1, descIdx: undefined };

  const headers = table[0].map(h => String(h || '').trim());
  const detected = detectKeyAndAmountColumns(headers);

  let keyIdx = detected.keyIndex;
  let amountIdx = detected.amountIndex;
  let descIdx = detected.descIndex;

  if (specifiedKeyCol !== undefined) {
    if (typeof specifiedKeyCol === 'number') keyIdx = specifiedKeyCol;
    else {
      const found = headers.findIndex(h => h.toLowerCase() === specifiedKeyCol.toLowerCase());
      if (found !== -1) keyIdx = found;
    }
  }

  if (specifiedAmountCol !== undefined) {
    if (typeof specifiedAmountCol === 'number') amountIdx = specifiedAmountCol;
    else {
      const found = headers.findIndex(h => h.toLowerCase() === specifiedAmountCol.toLowerCase());
      if (found !== -1) amountIdx = found;
    }
  }

  const rows: Array<{ key: string; amount: number; desc?: string; originalRow: any[] }> = [];

  for (let i = 1; i < table.length; i++) {
    const row = table[i];
    if (!row || row.length === 0 || row.every(c => c === '' || c === null || c === undefined)) {
      continue;
    }
    const rawKey = String(row[keyIdx] || '').trim();
    if (!rawKey) continue;

    const amount = parseFlexibleNumber(row[amountIdx]);
    const desc = descIdx !== undefined && row[descIdx] ? String(row[descIdx]).trim() : undefined;

    rows.push({
      key: rawKey,
      amount,
      desc,
      originalRow: row,
    });
  }

  return { rows, keyIdx, amountIdx, descIdx };
}

/**
 * Mesin kalkulasi rekonsiliasi data dua arah (bidirectional diff)
 */
export function reconcileDataSets(
  tableA: any[][],
  tableB: any[][],
  options?: ReconcileOptions
): ReconcileSummary {
  const tolerance = options?.tolerance || 0;

  const datasetA = extractRows(tableA, options?.keyColumnA, options?.amountColumnA);
  const datasetB = extractRows(tableB, options?.keyColumnB, options?.amountColumnB);

  const mapB = new Map<string, { key: string; amount: number; desc?: string }>();
  datasetB.rows.forEach(item => {
    mapB.set(item.key.toLowerCase(), item);
  });

  const matchedKeys = new Set<string>();
  const discrepancies: ReconcileItemDiscrepancy[] = [];
  const unmatchedItems: ReconcileItemUnmatched[] = [];

  let matchedCount = 0;
  let matchedTotalAmount = 0;
  let netDiscrepancyAmount = 0;

  // Analisis Sumber A terhadap B
  for (const itemA of datasetA.rows) {
    const keyLower = itemA.key.toLowerCase();
    const itemB = mapB.get(keyLower);

    if (itemB) {
      matchedKeys.add(keyLower);
      const diff = itemA.amount - itemB.amount;
      const absDiff = Math.abs(diff);

      if (absDiff <= tolerance) {
        matchedCount++;
        matchedTotalAmount += itemA.amount;
      } else {
        discrepancies.push({
          key: itemA.key,
          descriptionA: itemA.desc,
          descriptionB: itemB.desc,
          amountA: itemA.amount,
          amountB: itemB.amount,
          diff,
        });
        netDiscrepancyAmount += diff;
      }
    } else {
      unmatchedItems.push({
        key: itemA.key,
        description: itemA.desc,
        amount: itemA.amount,
        source: 'A',
      });
    }
  }

  // Analisis Sumber B yang tidak ada di A
  for (const itemB of datasetB.rows) {
    const keyLower = itemB.key.toLowerCase();
    if (!matchedKeys.has(keyLower)) {
      unmatchedItems.push({
        key: itemB.key,
        description: itemB.desc,
        amount: itemB.amount,
        source: 'B',
      });
    }
  }

  const unmatchedCountA = unmatchedItems.filter(i => i.source === 'A').length;
  const unmatchedCountB = unmatchedItems.filter(i => i.source === 'B').length;

  // Bangun narasi ringkasan eksekutif
  const formatRp = (num: number) => `Rp ${Math.abs(num).toLocaleString('id-ID')}`;
  let md = `### 📊 Laporan Rekonsiliasi Data & Kas\n\n`;
  md += `- **Total Baris Data:** Sumber A (${datasetA.rows.length}) vs Sumber B (${datasetB.rows.length})\n`;
  md += `- **Cocok Sempurna (Matched):** ✅ ${matchedCount} transaksi (${formatRp(matchedTotalAmount)})\n`;
  md += `- **Selisih Nominal (Discrepancy):** ${discrepancies.length > 0 ? '⚠️' : '✅'} ${discrepancies.length} transaksi (Net Selisih: ${netDiscrepancyAmount >= 0 ? '+' : '-'}${formatRp(netDiscrepancyAmount)})\n`;
  md += `- **Hanya di Sumber A:** ${unmatchedCountA} transaksi\n`;
  md += `- **Hanya di Sumber B:** ${unmatchedCountB} transaksi\n\n`;

  if (discrepancies.length > 0) {
    md += `#### ⚠️ Rincian Selisih Kunci:\n`;
    for (const d of discrepancies.slice(0, 5)) {
      md += `- **${d.key}**: Sumber A (${formatRp(d.amountA)}) vs Sumber B (${formatRp(d.amountB)}) ➔ Selisih: ${d.diff > 0 ? '+' : '-'}${formatRp(d.diff)}\n`;
    }
    if (discrepancies.length > 5) {
      md += `*...dan ${discrepancies.length - 5} selisih lainnya tercantum pada lembar kerja Hasil_Rekonsiliasi.*\n`;
    }
  }

  return {
    totalRowsA: datasetA.rows.length,
    totalRowsB: datasetB.rows.length,
    matchedCount,
    matchedTotalAmount,
    discrepancyCount: discrepancies.length,
    netDiscrepancyAmount,
    unmatchedCountA,
    unmatchedCountB,
    discrepancies,
    unmatchedItems,
    markdownSummary: md,
  };
}

/**
 * Menyusun data array 2D untuk lembar kerja "Hasil_Rekonsiliasi"
 */
export function buildReconciliationSheetGrid(summary: ReconcileSummary): { summaryRows: any[][]; detailRows: any[][] } {
  const dateStr = new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' });
  const statusOverall = summary.discrepancyCount === 0 && summary.unmatchedCountA === 0 && summary.unmatchedCountB === 0
    ? 'SEIMBANG (RECONCILED)'
    : 'ADA SELISIH (DISCREPANCIES DETECTED)';

  const totalUnmatchedAAmount = summary.unmatchedItems.filter(i => i.source === 'A').reduce((acc, i) => acc + i.amount, 0);
  const totalUnmatchedBAmount = summary.unmatchedItems.filter(i => i.source === 'B').reduce((acc, i) => acc + i.amount, 0);

  const summaryRows = [
    ['LAPORAN REKONSILIASI DATA & KAS'],
    ['Tanggal Eksekusi:', dateStr, 'Status Global:', statusOverall],
    [],
    ['KATEGORI REKONSILIASI', 'JUMLAH TRANSAKSI', 'TOTAL NOMINAL', 'STATUS'],
    ['Cocok Sempurna (Matched)', summary.matchedCount, summary.matchedTotalAmount, 'OK (Sesuai)'],
    ['Selisih Nilai (Discrepancy)', summary.discrepancyCount, summary.netDiscrepancyAmount, summary.discrepancyCount > 0 ? 'Perlu Investigasi' : 'OK'],
    ['Hanya di Sumber A', summary.unmatchedCountA, totalUnmatchedAAmount, summary.unmatchedCountA > 0 ? 'Belum Ada di B' : 'OK'],
    ['Hanya di Sumber B', summary.unmatchedCountB, totalUnmatchedBAmount, summary.unmatchedCountB > 0 ? 'Belum Tercatat di A' : 'OK'],
    [],
  ];

  const detailRows = [
    ['Kunci / No. Ref', 'Deskripsi Sumber A', 'Deskripsi Sumber B', 'Nilai Sumber A', 'Nilai Sumber B', 'Selisih (Diff)', 'Status'],
  ];

  for (const d of summary.discrepancies) {
    detailRows.push([
      d.key,
      d.descriptionA || '-',
      d.descriptionB || '-',
      d.amountA,
      d.amountB,
      d.diff,
      'SELISIH NOMINAL',
    ]);
  }

  for (const u of summary.unmatchedItems) {
    if (u.source === 'A') {
      detailRows.push([
        u.key,
        u.description || '-',
        '-',
        u.amount,
        0,
        u.amount,
        'HANYA DI SUMBER A',
      ]);
    } else {
      detailRows.push([
        u.key,
        '-',
        u.description || '-',
        0,
        u.amount,
        -u.amount,
        'HANYA DI SUMBER B',
      ]);
    }
  }

  return { summaryRows, detailRows };
}
