export interface ChartRenderOptions {
  chartType: 'bar' | 'line' | 'pie';
  labels: string[];
  data: number[];
  title?: string;
  seriesName?: string;
}

const COLOR_PALETTE = [
  '#2563EB', // Blue
  '#10B981', // Emerald
  '#F59E0B', // Amber
  '#EF4444', // Red
  '#8B5CF6', // Purple
  '#06B6D4', // Cyan
  '#EC4899', // Pink
  '#6366F1', // Indigo
];

// Fallback 1x1 transparent PNG Base64 string for headless environments without canvas
const FALLBACK_1X1_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

/**
 * Renders a 2D chart (bar, line, pie) to a Base64-encoded PNG string.
 * Compatible with Word's `insertInlinePictureFromBase64` API.
 */
export function renderChartToBase64Png(options: ChartRenderOptions): string {
  const { chartType, labels, data, title, seriesName } = options;

  if (typeof document === 'undefined' || typeof document.createElement !== 'function') {
    return FALLBACK_1X1_PNG;
  }

  let canvas: HTMLCanvasElement;
  try {
    canvas = document.createElement('canvas');
  } catch {
    return FALLBACK_1X1_PNG;
  }

  const width = 640;
  const height = 380;
  canvas.width = width;
  canvas.height = height;

  const ctx = canvas.getContext('2d');
  if (!ctx) {
    return FALLBACK_1X1_PNG;
  }

  // Background
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, width, height);

  // Border frame
  ctx.strokeStyle = '#E2E8F0';
  ctx.lineWidth = 1.5;
  ctx.strokeRect(1, 1, width - 2, height - 2);

  // Title
  const chartTitle = title || (seriesName ? `Grafik ${seriesName}` : 'Grafik Data');
  ctx.fillStyle = '#1E293B';
  ctx.font = 'bold 18px "Segoe UI", Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText(chartTitle, width / 2, 38);

  const numItems = Math.min(labels.length, data.length);
  if (numItems === 0) {
    ctx.font = '14px sans-serif';
    ctx.fillStyle = '#64748B';
    ctx.fillText('Tidak ada data yang dapat ditampilkan.', width / 2, height / 2);
    return canvasToCleanBase64(canvas);
  }

  if (chartType === 'pie') {
    renderPieChart(ctx, labels, data, numItems, width, height);
  } else if (chartType === 'line') {
    renderLineChart(ctx, labels, data, numItems, width, height);
  } else {
    // Default to Bar chart
    renderBarChart(ctx, labels, data, numItems, width, height);
  }

  return canvasToCleanBase64(canvas);
}

function renderBarChart(
  ctx: CanvasRenderingContext2D,
  labels: string[],
  data: number[],
  numItems: number,
  width: number,
  height: number
) {
  const leftMargin = 65;
  const rightMargin = 40;
  const topMargin = 70;
  const bottomMargin = 60;

  const plotWidth = width - leftMargin - rightMargin;
  const plotHeight = height - topMargin - bottomMargin;

  const rawMax = Math.max(...data.slice(0, numItems), 0);
  const maxVal = rawMax === 0 ? 10 : Math.ceil(rawMax * 1.15);

  // Y-axis grid and scale labels
  const steps = 4;
  ctx.strokeStyle = '#F1F5F9';
  ctx.lineWidth = 1;
  ctx.textAlign = 'right';
  ctx.font = '11px sans-serif';
  ctx.fillStyle = '#64748B';

  for (let i = 0; i <= steps; i++) {
    const yVal = Math.round((maxVal / steps) * i);
    const y = topMargin + plotHeight - (plotHeight / steps) * i;

    ctx.beginPath();
    ctx.moveTo(leftMargin, y);
    ctx.lineTo(width - rightMargin, y);
    ctx.stroke();

    ctx.fillText(yVal.toLocaleString(), leftMargin - 8, y + 4);
  }

  // X/Y main axes lines
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.moveTo(leftMargin, topMargin);
  ctx.lineTo(leftMargin, topMargin + plotHeight);
  ctx.lineTo(width - rightMargin, topMargin + plotHeight);
  ctx.stroke();

  // Draw Bars
  const slotWidth = plotWidth / numItems;
  const barWidth = Math.min(Math.max(slotWidth * 0.55, 18), 65);

  for (let i = 0; i < numItems; i++) {
    const val = data[i] || 0;
    const barHeight = Math.max((val / maxVal) * plotHeight, 0);
    const x = leftMargin + i * slotWidth + (slotWidth - barWidth) / 2;
    const y = topMargin + plotHeight - barHeight;

    // Fill bar
    ctx.fillStyle = COLOR_PALETTE[i % COLOR_PALETTE.length];
    ctx.fillRect(x, y, barWidth, barHeight);

    // Value label above bar
    ctx.textAlign = 'center';
    ctx.font = 'bold 11px sans-serif';
    ctx.fillStyle = '#0F172A';
    ctx.fillText(val.toLocaleString(), x + barWidth / 2, y - 6);

    // X-axis Category label
    ctx.font = '11px sans-serif';
    ctx.fillStyle = '#334155';
    const label = labels[i] || `Item ${i + 1}`;
    const truncatedLabel = label.length > 10 ? `${label.slice(0, 8)}...` : label;
    ctx.fillText(truncatedLabel, x + barWidth / 2, topMargin + plotHeight + 20);
  }
}

function renderLineChart(
  ctx: CanvasRenderingContext2D,
  labels: string[],
  data: number[],
  numItems: number,
  width: number,
  height: number
) {
  const leftMargin = 65;
  const rightMargin = 40;
  const topMargin = 70;
  const bottomMargin = 60;

  const plotWidth = width - leftMargin - rightMargin;
  const plotHeight = height - topMargin - bottomMargin;

  const rawMax = Math.max(...data.slice(0, numItems), 0);
  const maxVal = rawMax === 0 ? 10 : Math.ceil(rawMax * 1.15);

  // Y-axis grid
  const steps = 4;
  ctx.strokeStyle = '#F1F5F9';
  ctx.lineWidth = 1;
  ctx.textAlign = 'right';
  ctx.font = '11px sans-serif';
  ctx.fillStyle = '#64748B';

  for (let i = 0; i <= steps; i++) {
    const yVal = Math.round((maxVal / steps) * i);
    const y = topMargin + plotHeight - (plotHeight / steps) * i;

    ctx.beginPath();
    ctx.moveTo(leftMargin, y);
    ctx.lineTo(width - rightMargin, y);
    ctx.stroke();

    ctx.fillText(yVal.toLocaleString(), leftMargin - 8, y + 4);
  }

  // Main axes
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.moveTo(leftMargin, topMargin);
  ctx.lineTo(leftMargin, topMargin + plotHeight);
  ctx.lineTo(width - rightMargin, topMargin + plotHeight);
  ctx.stroke();

  // Points coordinates
  const slotWidth = numItems > 1 ? plotWidth / (numItems - 1) : plotWidth;
  const points: { x: number; y: number; val: number; label: string }[] = [];

  for (let i = 0; i < numItems; i++) {
    const val = data[i] || 0;
    const x = numItems > 1 ? leftMargin + i * slotWidth : leftMargin + plotWidth / 2;
    const y = topMargin + plotHeight - Math.max((val / maxVal) * plotHeight, 0);
    points.push({ x, y, val, label: labels[i] || `Item ${i + 1}` });
  }

  // Draw line
  ctx.strokeStyle = '#2563EB';
  ctx.lineWidth = 3;
  ctx.beginPath();
  points.forEach((pt, idx) => {
    if (idx === 0) ctx.moveTo(pt.x, pt.y);
    else ctx.lineTo(pt.x, pt.y);
  });
  ctx.stroke();

  // Draw points & labels
  points.forEach(pt => {
    // Dot
    ctx.fillStyle = '#FFFFFF';
    ctx.beginPath();
    ctx.arc(pt.x, pt.y, 5, 0, Math.PI * 2);
    ctx.fill();

    ctx.strokeStyle = '#2563EB';
    ctx.lineWidth = 2.5;
    ctx.beginPath();
    ctx.arc(pt.x, pt.y, 5, 0, Math.PI * 2);
    ctx.stroke();

    // Value
    ctx.textAlign = 'center';
    ctx.font = 'bold 11px sans-serif';
    ctx.fillStyle = '#0F172A';
    ctx.fillText(pt.val.toLocaleString(), pt.x, pt.y - 10);

    // X-label
    ctx.font = '11px sans-serif';
    ctx.fillStyle = '#334155';
    const truncatedLabel = pt.label.length > 10 ? `${pt.label.slice(0, 8)}...` : labelOrFallback(pt.label);
    ctx.fillText(truncatedLabel, pt.x, topMargin + plotHeight + 20);
  });
}

function labelOrFallback(label: string): string {
  return label.length > 10 ? `${label.slice(0, 8)}...` : label;
}

function renderPieChart(
  ctx: CanvasRenderingContext2D,
  labels: string[],
  data: number[],
  numItems: number,
  width: number,
  height: number
) {
  const total = data.slice(0, numItems).reduce((sum, v) => sum + (v > 0 ? v : 0), 0);
  const centerX = width * 0.38;
  const centerY = height * 0.54;
  const radius = Math.min(width, height) * 0.32;

  let currentAngle = -Math.PI / 2;

  for (let i = 0; i < numItems; i++) {
    const val = data[i] > 0 ? data[i] : 0;
    const sliceAngle = total > 0 ? (val / total) * Math.PI * 2 : (Math.PI * 2) / numItems;

    ctx.fillStyle = COLOR_PALETTE[i % COLOR_PALETTE.length];
    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.arc(centerX, centerY, radius, currentAngle, currentAngle + sliceAngle);
    ctx.closePath();
    ctx.fill();

    ctx.strokeStyle = '#FFFFFF';
    ctx.lineWidth = 2;
    ctx.stroke();

    currentAngle += sliceAngle;
  }

  // Legend on right side
  const legendX = width * 0.68;
  let legendY = 85;
  ctx.textAlign = 'left';

  for (let i = 0; i < numItems; i++) {
    const val = data[i] || 0;
    const pct = total > 0 ? ((val / total) * 100).toFixed(1) : '0';
    const label = labels[i] || `Item ${i + 1}`;

    // Color box
    ctx.fillStyle = COLOR_PALETTE[i % COLOR_PALETTE.length];
    ctx.fillRect(legendX, legendY - 10, 14, 14);

    // Label text
    ctx.fillStyle = '#1E293B';
    ctx.font = '12px sans-serif';
    ctx.fillText(`${label} (${pct}%)`, legendX + 22, legendY + 2);

    legendY += 26;
  }
}

function canvasToCleanBase64(canvas: HTMLCanvasElement): string {
  try {
    const dataUrl = canvas.toDataURL('image/png');
    return dataUrl.replace(/^data:image\/png;base64,/, '');
  } catch {
    return FALLBACK_1X1_PNG;
  }
}
