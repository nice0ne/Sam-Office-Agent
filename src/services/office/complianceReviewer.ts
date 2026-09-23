import { ComplianceClause, ComplianceReviewOptions, ComplianceReviewResult } from './types';

export function analyzeContractCompliance(
  text: string,
  options?: ComplianceReviewOptions
): ComplianceReviewResult {
  const contractType = options?.contractType || 'general';
  const paragraphs = text
    .split(/\r?\n/)
    .map((p) => p.trim())
    .filter((p) => p.length > 0);

  const identifiedClauses: ComplianceClause[] = [];

  // 1. Liability & Indemnity
  const liabilityRegex = /(ganti rugi|tanggung jawab|liability|indemnitas|indemnity|tuntutan|kerugian)/i;
  const liabilityMatch = paragraphs.find((p) => liabilityRegex.test(p));
  if (liabilityMatch) {
    const isUnlimited = /(tanpa batas|unlimited|seluruh kerugian|semua kerugian|segala tuntutan|tidak terbatas)/i.test(liabilityMatch);
    identifiedClauses.push({
      category: 'liability_indemnity',
      excerpt: liabilityMatch,
      status: isUnlimited ? 'high_risk' : 'compliant',
      analysis: isUnlimited
        ? 'Ditemukan klausul pertanggungjawaban/ganti rugi tanpa batasan maksimal (unlimited liability).'
        : 'Klausul pertanggungjawaban dan ganti rugi teridentifikasi dengan batasan wajar.',
      recommendation: isUnlimited
        ? 'Batasi tanggung jawab maksimal setara nilai kontrak (liability cap).'
        : undefined,
    });
  }

  // 2. Payment Terms
  const paymentRegex = /(pembayaran|termin|invoice|jatuh tempo|net 30|30 hari|rekening|transfer)/i;
  const paymentMatch = paragraphs.find((p) => paymentRegex.test(p));
  if (paymentMatch) {
    const isVague = /(sewaktu-waktu|tanpa jadwal|bila disepakati)/i.test(paymentMatch);
    identifiedClauses.push({
      category: 'payment_terms',
      excerpt: paymentMatch,
      status: isVague ? 'warning' : 'compliant',
      analysis: isVague
        ? 'Ketentuan termin pembayaran terindikasi tidak spesifik.'
        : 'Ketentuan termin pembayaran dan jatuh tempo terdefinisi jelas.',
      recommendation: isVague ? 'Tetapkan jangka waktu pasti jatuh tempo pembayaran (misal: 30 hari kalender).' : undefined,
    });
  }

  // 3. SLA & Performance Penalties
  const slaRegex = /(sla|service level|denda|penalti|keterlambatan|uptime)/i;
  const slaMatch = paragraphs.find((p) => slaRegex.test(p));
  if (slaMatch) {
    const isDisproportionate = /(tanpa batas|tanpa maksimal|5% per hari|10% per hari)/i.test(slaMatch);
    identifiedClauses.push({
      category: 'sla_performance',
      excerpt: slaMatch,
      status: isDisproportionate ? 'high_risk' : 'compliant',
      analysis: isDisproportionate
        ? 'Denda keterlambatan/penalti SLA berisiko disproportionate atau tanpa batas akumulasi.'
        : 'Metrik SLA dan mekanisme denda keterlambatan tertera secara proporsional.',
      recommendation: isDisproportionate ? 'Terapkan batas maksimal akumulasi penalti (misal: maks 5% - 10% dari nilai kontrak).' : undefined,
    });
  }

  // 4. Termination & Cancellation
  const termRegex = /(pemutusan|pengakhiran|terminasi|pembatalan)/i;
  const termMatch = paragraphs.find((p) => termRegex.test(p));
  if (termMatch) {
    const isUnilateral = /(sepihak|tanpa pemberitahuan|tanpa alasan|seketika)/i.test(termMatch);
    identifiedClauses.push({
      category: 'termination_cancellation',
      excerpt: termMatch,
      status: isUnilateral ? 'high_risk' : 'compliant',
      analysis: isUnilateral
        ? 'Risiko pemutusan kontrak sepihak tanpa mekanisme pemberitahuan tertulis sebelumnya.'
        : 'Mekanisme pengakhiran kontrak memiliki prosedur pemberitahuan yang teratur.',
      recommendation: isUnilateral ? 'Wajibkan pemberitahuan tertulis minimal 14 atau 30 hari kalender sebelum pengakhiran.' : undefined,
    });
  }

  // 5. Confidentiality (NDA)
  const confRegex = /(kerahasiaan|rahasia|non-disclosure|nda|informasi rahasia|confidential)/i;
  const confMatch = paragraphs.find((p) => confRegex.test(p));
  if (confMatch) {
    identifiedClauses.push({
      category: 'confidentiality_nda',
      excerpt: confMatch,
      status: 'compliant',
      analysis: 'Klausul perlindungan kerahasiaan data dan informasi usaha terakomodasi.',
    });
  }

  // 6. Dispute Resolution
  const disputeRegex = /(sengketa|perselisihan|musyawarah|arbitrase|bani|pengadilan|yurisdiksi)/i;
  const disputeMatch = paragraphs.find((p) => disputeRegex.test(p));
  if (disputeMatch) {
    identifiedClauses.push({
      category: 'dispute_resolution',
      excerpt: disputeMatch,
      status: 'compliant',
      analysis: 'Mekanisme penyelesaian sengketa telah ditentukan secara eksplisit.',
    });
  }

  // 7. Force Majeure
  const fmRegex = /(force majeure|keadaan kahar|bencana alam|darurat)/i;
  const fmMatch = paragraphs.find((p) => fmRegex.test(p));
  if (fmMatch) {
    identifiedClauses.push({
      category: 'force_majeure',
      excerpt: fmMatch,
      status: 'compliant',
      analysis: 'Klausul keadaan kahar (force majeure) telah diatur untuk melindungi kedua belah pihak.',
    });
  }

  // Check missing critical clauses
  const missingCriticalClauses: string[] = [];
  if (!fmMatch) {
    missingCriticalClauses.push('force_majeure');
  }
  if (!confMatch && contractType !== 'employment') {
    missingCriticalClauses.push('confidentiality_nda');
  }
  if (!disputeMatch) {
    missingCriticalClauses.push('dispute_resolution');
  }

  // Calculate overall risk
  let overallRiskLevel: 'low' | 'medium' | 'high' = 'low';
  const hasHighRisk = identifiedClauses.some((c) => c.status === 'high_risk');
  const hasWarning = identifiedClauses.some((c) => c.status === 'warning');

  if (hasHighRisk) {
    overallRiskLevel = 'high';
  } else if (hasWarning || missingCriticalClauses.length >= 2) {
    overallRiskLevel = 'medium';
  }

  const actionableRecommendations: string[] = [];
  for (const clause of identifiedClauses) {
    if (clause.recommendation) {
      actionableRecommendations.push(clause.recommendation);
    }
  }
  if (missingCriticalClauses.includes('force_majeure')) {
    actionableRecommendations.push('Tambahkan klausul Force Majeure untuk perlindungan terhadap kejadian luar biasa di luar kendali para pihak.');
  }

  const executiveSummary = overallRiskLevel === 'high'
    ? `Ditemukan potensi risiko tinggi (${identifiedClauses.filter((c) => c.status === 'high_risk').length} klausul kritis) dan ${missingCriticalClauses.length} klausul esensial yang belum tercantum.`
    : overallRiskLevel === 'medium'
    ? `Dokumen tergolong risiko moderat dengan beberapa catatan penyempurnaan klausul.`
    : `Dokumen memenuhi standar umum kepatuhan dengan tingkat risiko rendah.`;

  return {
    contractType,
    overallRiskLevel,
    clausesReviewedCount: identifiedClauses.length,
    identifiedClauses,
    missingCriticalClauses,
    executiveSummary,
    actionableRecommendations,
  };
}
