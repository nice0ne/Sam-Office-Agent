import {
  FlowchartDefinition,
  FlowchartEdge,
  FlowchartNode,
  FlowchartNodeType,
  FlowchartOptions,
} from '../services/office/types';

// Fallback 1x1 transparent PNG Base64 string for headless environments without canvas
const FALLBACK_1X1_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

export interface DiagramThemePalette {
  primary: string;
  accent: string;
  bg: string;
  border: string;
  text: string;
  nodeBg: string;
  decisionBg: string;
  decisionText: string;
  documentBg: string;
  arrowColor: string;
}

export const DIAGRAM_THEMES: Record<string, DiagramThemePalette> = {
  corporate_navy: {
    primary: '#1E3A8A',
    accent: '#3B82F6',
    bg: '#F8FAFC',
    border: '#94A3B8',
    text: '#0F172A',
    nodeBg: '#FFFFFF',
    decisionBg: '#EFF6FF',
    decisionText: '#1E3A8A',
    documentBg: '#F1F5F9',
    arrowColor: '#475569',
  },
  emerald_executive: {
    primary: '#065F46',
    accent: '#10B981',
    bg: '#F0FDF4',
    border: '#A7F3D0',
    text: '#064E3B',
    nodeBg: '#FFFFFF',
    decisionBg: '#ECFDF5',
    decisionText: '#065F46',
    documentBg: '#E6F4EA',
    arrowColor: '#047857',
  },
  modern_dark: {
    primary: '#0F172A',
    accent: '#64748B',
    bg: '#1E293B',
    border: '#475569',
    text: '#F8FAFC',
    nodeBg: '#334155',
    decisionBg: '#1E293B',
    decisionText: '#93C5FD',
    documentBg: '#2D3748',
    arrowColor: '#94A3B8',
  },
  amber_warm: {
    primary: '#92400E',
    accent: '#F59E0B',
    bg: '#FFFBEB',
    border: '#FDE68A',
    text: '#78350F',
    nodeBg: '#FFFFFF',
    decisionBg: '#FEF3C7',
    decisionText: '#92400E',
    documentBg: '#FDF6E2',
    arrowColor: '#B45309',
  },
};

/**
 * Synthesizes a structured FlowchartDefinition from natural language SOP or process descriptions.
 */
export function synthesizeFlowchartFromText(
  rawText?: string,
  options?: FlowchartOptions
): FlowchartDefinition {
  const title = options?.title || 'Diagram Alur Proses';
  const theme = options?.theme || 'corporate_navy';
  const direction = options?.direction || 'TD';

  if (!rawText || rawText.trim().length === 0) {
    return {
      title,
      direction,
      theme,
      nodes: [
        { id: 'step-1', label: 'Mulai', type: 'start', subText: 'Inisiasi Proses' },
        { id: 'step-2', label: 'Proses Kerja Utama', type: 'process', subText: 'Pelaksanaan tugas' },
        { id: 'step-3', label: 'Verifikasi & Approval', type: 'decision', subText: 'Pemeriksaan standar' },
        { id: 'step-4', label: 'Dokumen / Laporan', type: 'document', subText: 'Pengarsipan berkas' },
        { id: 'step-5', label: 'Selesai', type: 'end', subText: 'Proses tuntas' },
      ],
      edges: [
        { from: 'step-1', to: 'step-2' },
        { from: 'step-2', to: 'step-3' },
        { from: 'step-3', to: 'step-4', label: 'Ya' },
        { from: 'step-4', to: 'step-5' },
      ],
    };
  }

  // 1. Split text into individual process steps
  let rawSteps: string[] = [];
  if (rawText.includes('-->') || rawText.includes('->') || rawText.includes('=>')) {
    rawSteps = rawText.split(/\s*(?:-->|->|=>)\s*/);
  } else if (rawText.includes('\n')) {
    rawSteps = rawText.split(/\r?\n/);
  } else {
    const numberedMatches = rawText.split(/(?=\b\d+[\.\)])/);
    if (numberedMatches.length > 1) {
      rawSteps = numberedMatches;
    } else {
      rawSteps = [rawText];
    }
  }

  const cleanedSteps = rawSteps
    .map(s => s.trim())
    .filter(s => s.length > 0)
    .map(s => s.replace(/^(\d+[\.\)]|[-*•]|\bstep\s*\d+[:\.]?)\s*/i, '').trim())
    .filter(s => s.length > 0);

  if (cleanedSteps.length === 0) {
    return synthesizeFlowchartFromText('', options);
  }

  const total = cleanedSteps.length;
  const nodes: FlowchartNode[] = [];

  cleanedSteps.forEach((stepText, idx) => {
    const lower = stepText.toLowerCase();
    let type: FlowchartNodeType = 'process';

    // Categorize node based on keywords and sequence position
    if (idx === 0 && (/^(mulai|start|awal|inisiasi)\b/i.test(lower) || /\b(mulai|start)\b/i.test(lower))) {
      type = 'start';
    } else if (idx === total - 1 && (/^(selesai|end|finish|tuntas)\b/i.test(lower) || /\b(selesai|finish|tuntas)\b/i.test(lower))) {
      type = 'end';
    } else if (/\?|\b(apakah|apabila|jika|keputusan|decision)\b/i.test(lower)) {
      type = 'decision';
    } else if (/\b(dokumen|berkas|formulir|form|laporan|surat|faktur|invoice)\b/i.test(lower)) {
      type = 'document';
    } else if (/\b(subproses|sub-proses|subroutine|sub-routine)\b/i.test(lower)) {
      type = 'subroutine';
    } else if (idx === 0) {
      type = 'start';
    } else if (idx === total - 1 && total > 2) {
      type = 'end';
    } else {
      type = 'process';
    }

    // Split label and subText for optimal visual clarity
    let label = stepText;
    let subText: string | undefined = undefined;

    const colonIdx = stepText.indexOf(':');
    if (colonIdx > 0 && colonIdx <= 25) {
      label = stepText.slice(0, colonIdx).trim();
      subText = stepText.slice(colonIdx + 1).trim();
    } else if (stepText.includes('?')) {
      const qIdx = stepText.indexOf('?');
      label = stepText.slice(0, qIdx + 1).trim();
      const rem = stepText.slice(qIdx + 1).trim();
      if (rem.length > 0) subText = rem;
    } else if (stepText.length > 35) {
      label = stepText.slice(0, 32).trim() + '...';
      subText = stepText;
    }

    nodes.push({
      id: `step-${idx + 1}`,
      label,
      type,
      subText,
    });
  });

  // 2. Generate sequential edges with conditional labels for decisions
  const edges: FlowchartEdge[] = [];
  for (let i = 0; i < nodes.length - 1; i++) {
    const currentNode = nodes[i];
    const nextNode = nodes[i + 1];
    const edge: FlowchartEdge = {
      from: currentNode.id,
      to: nextNode.id,
    };

    if (currentNode.type === 'decision') {
      edge.label = 'Ya';
    }

    edges.push(edge);
  }

  return {
    title,
    direction,
    theme,
    nodes,
    edges,
  };
}

interface NodeLayout {
  node: FlowchartNode;
  x: number;
  y: number;
  w: number;
  h: number;
  cx: number;
  cy: number;
}

/**
 * Renders a FlowchartDefinition to a clean Base64-encoded PNG string (without data URL prefix).
 * Ready for direct insertion via Word inline picture or PowerPoint shape APIs.
 */
export function renderFlowchartToPngBase64(
  def: FlowchartDefinition,
  options?: FlowchartOptions
): string {
  if (typeof document === 'undefined' || typeof document.createElement !== 'function') {
    return FALLBACK_1X1_PNG;
  }

  let canvas: HTMLCanvasElement;
  try {
    canvas = document.createElement('canvas');
  } catch {
    return FALLBACK_1X1_PNG;
  }

  if (!canvas || typeof canvas.getContext !== 'function') {
    return FALLBACK_1X1_PNG;
  }

  const ctx = canvas.getContext('2d');
  if (!ctx) {
    return FALLBACK_1X1_PNG;
  }

  const themeKey = options?.theme || def.theme || 'corporate_navy';
  const theme = DIAGRAM_THEMES[themeKey] || DIAGRAM_THEMES.corporate_navy;
  const direction = options?.direction || def.direction || 'TD';
  const scale = options?.scale || 2;
  const numNodes = def.nodes.length;

  let canvasWidth = 800;
  let canvasHeight = 600;
  const nodeLayouts = new Map<string, NodeLayout>();

  if (direction === 'LR') {
    const nodeWidth = 150;
    const nodeHeight = 70;
    const gapX = 60;
    const leftMargin = 50;
    const rightMargin = 50;
    const topMargin = def.title ? 85 : 50;
    const bottomMargin = 50;

    canvasWidth = Math.max(leftMargin + numNodes * nodeWidth + (numNodes - 1) * gapX + rightMargin, 600);
    canvasHeight = topMargin + nodeHeight + bottomMargin;

    def.nodes.forEach((node, idx) => {
      const x = leftMargin + idx * (nodeWidth + gapX);
      const y = topMargin + 10;
      nodeLayouts.set(node.id, {
        node,
        x,
        y,
        w: nodeWidth,
        h: nodeHeight,
        cx: x + nodeWidth / 2,
        cy: y + nodeHeight / 2,
      });
    });
  } else {
    // Default: Top-Down (TD)
    const nodeWidth = 220;
    const nodeHeight = 64;
    const gapY = 50;
    const topMargin = def.title ? 85 : 45;
    const bottomMargin = 45;

    canvasWidth = 720;
    canvasHeight = topMargin + numNodes * nodeHeight + Math.max(numNodes - 1, 0) * gapY + bottomMargin;

    def.nodes.forEach((node, idx) => {
      const x = (canvasWidth - nodeWidth) / 2;
      const y = topMargin + idx * (nodeHeight + gapY);
      nodeLayouts.set(node.id, {
        node,
        x,
        y,
        w: nodeWidth,
        h: nodeHeight,
        cx: x + nodeWidth / 2,
        cy: y + nodeHeight / 2,
      });
    });
  }

  // High-DPI scaling
  canvas.width = canvasWidth * scale;
  canvas.height = canvasHeight * scale;
  ctx.scale(scale, scale);

  // 1. Draw Background
  ctx.fillStyle = theme.bg;
  ctx.fillRect(0, 0, canvasWidth, canvasHeight);

  // 2. Draw Frame Border
  ctx.strokeStyle = theme.border;
  ctx.lineWidth = 1;
  ctx.strokeRect(1, 1, canvasWidth - 2, canvasHeight - 2);

  // 3. Draw Title (if present)
  if (def.title) {
    ctx.font = 'bold 18px "Segoe UI", Arial, sans-serif';
    ctx.fillStyle = theme.text;
    ctx.textAlign = 'center';
    ctx.fillText(def.title, canvasWidth / 2, 40);

    ctx.strokeStyle = theme.accent;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(canvasWidth / 2 - 50, 48);
    ctx.lineTo(canvasWidth / 2 + 50, 48);
    ctx.stroke();
  }

  // 4. Draw Connecting Edges with Arrowheads & Labels
  def.edges.forEach(edge => {
    const source = nodeLayouts.get(edge.from);
    const target = nodeLayouts.get(edge.to);
    if (!source || !target) return;

    ctx.save();
    ctx.strokeStyle = theme.arrowColor;
    ctx.lineWidth = 2;
    if (edge.style === 'dashed') {
      ctx.setLineDash([5, 5]);
    } else {
      ctx.setLineDash([]);
    }

    let startX: number;
    let startY: number;
    let endX: number;
    let endY: number;

    if (direction === 'LR') {
      startX = source.x + source.w;
      startY = source.cy;
      endX = target.x;
      endY = target.cy;
    } else {
      startX = source.cx;
      startY = source.y + source.h;
      endX = target.cx;
      endY = target.y;
    }

    // Line
    ctx.beginPath();
    ctx.moveTo(startX, startY);
    ctx.lineTo(endX, endY);
    ctx.stroke();

    // Arrowhead
    const arrowLength = 8;
    const angle = Math.atan2(endY - startY, endX - startX);
    ctx.fillStyle = theme.arrowColor;
    ctx.beginPath();
    ctx.moveTo(endX, endY);
    ctx.lineTo(
      endX - arrowLength * Math.cos(angle - Math.PI / 6),
      endY - arrowLength * Math.sin(angle - Math.PI / 6)
    );
    ctx.lineTo(
      endX - arrowLength * Math.cos(angle + Math.PI / 6),
      endY - arrowLength * Math.sin(angle + Math.PI / 6)
    );
    ctx.closePath();
    ctx.fill();

    // Edge Label
    if (edge.label) {
      const midX = (startX + endX) / 2;
      const midY = (startY + endY) / 2;
      ctx.font = 'bold 11px "Segoe UI", Arial, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';

      const labelWidth = ctx.measureText(edge.label).width + 12;
      const labelHeight = 16;

      ctx.fillStyle = theme.bg;
      ctx.fillRect(midX - labelWidth / 2, midY - labelHeight / 2, labelWidth, labelHeight);
      ctx.strokeStyle = theme.border;
      ctx.lineWidth = 1;
      ctx.strokeRect(midX - labelWidth / 2, midY - labelHeight / 2, labelWidth, labelHeight);

      ctx.fillStyle = theme.primary;
      ctx.fillText(edge.label, midX, midY);
    }
    ctx.restore();
  });

  // 5. Draw Nodes & Shapes
  nodeLayouts.forEach(({ node, x, y, w, h, cx, cy }) => {
    ctx.save();

    let fillColor = theme.nodeBg;
    let strokeColor = theme.border;
    let textColor = theme.text;
    const lineWidth = 2;

    switch (node.type) {
      case 'start':
        fillColor = theme.primary;
        strokeColor = theme.primary;
        textColor = '#FFFFFF';
        break;
      case 'end':
        fillColor = theme.accent;
        strokeColor = theme.primary;
        textColor = '#FFFFFF';
        break;
      case 'decision':
        fillColor = theme.decisionBg;
        strokeColor = theme.accent;
        textColor = theme.decisionText;
        break;
      case 'document':
        fillColor = theme.documentBg;
        strokeColor = theme.border;
        textColor = theme.text;
        break;
      case 'subroutine':
        fillColor = theme.nodeBg;
        strokeColor = theme.primary;
        textColor = theme.text;
        break;
      case 'process':
      default:
        fillColor = theme.nodeBg;
        strokeColor = theme.border;
        textColor = theme.text;
        break;
    }

    if (node.color) {
      fillColor = node.color;
    }

    ctx.fillStyle = fillColor;
    ctx.strokeStyle = strokeColor;
    ctx.lineWidth = lineWidth;

    if (node.type === 'start' || node.type === 'end') {
      // Rounded pill / stadium shape
      const r = h / 2;
      ctx.beginPath();
      ctx.moveTo(x + r, y);
      ctx.lineTo(x + w - r, y);
      ctx.arc(x + w - r, y + r, r, -Math.PI / 2, Math.PI / 2);
      ctx.lineTo(x + r, y + h);
      ctx.arc(x + r, y + r, r, Math.PI / 2, (3 * Math.PI) / 2);
      ctx.closePath();
      ctx.fill();
      ctx.stroke();
    } else if (node.type === 'decision') {
      // Diamond polygon
      ctx.beginPath();
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, cy);
      ctx.closePath();
      ctx.fill();
      ctx.stroke();
    } else if (node.type === 'document') {
      // Box with curved wave bottom
      const waveHeight = 8;
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - waveHeight);
      ctx.bezierCurveTo(
        x + (w * 3) / 4,
        y + h - waveHeight - 5,
        x + w / 2,
        y + h + waveHeight,
        x,
        y + h - waveHeight
      );
      ctx.closePath();
      ctx.fill();
      ctx.stroke();
    } else if (node.type === 'subroutine') {
      // Rounded rect with double vertical borders
      drawRoundRect(ctx, x, y, w, h, 6);
      ctx.fill();
      ctx.stroke();
      const inset = 12;
      ctx.beginPath();
      ctx.moveTo(x + inset, y);
      ctx.lineTo(x + inset, y + h);
      ctx.moveTo(x + w - inset, y);
      ctx.lineTo(x + w - inset, y + h);
      ctx.stroke();
    } else {
      // Process: rounded rectangle
      drawRoundRect(ctx, x, y, w, h, 8);
      ctx.fill();
      ctx.stroke();
    }

    // Centered label text inside node
    ctx.fillStyle = textColor;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    if (node.subText) {
      ctx.font = 'bold 12px "Segoe UI", Arial, sans-serif';
      ctx.fillText(truncateText(ctx, node.label, w - 24), cx, cy - 8);
      ctx.font = '10px "Segoe UI", Arial, sans-serif';
      ctx.fillText(truncateText(ctx, node.subText, w - 28), cx, cy + 10);
    } else {
      ctx.font = 'bold 12px "Segoe UI", Arial, sans-serif';
      ctx.fillText(truncateText(ctx, node.label, w - 20), cx, cy);
    }

    ctx.restore();
  });

  return canvasToCleanBase64(canvas);
}

function drawRoundRect(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  r: number
): void {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

function truncateText(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string {
  if (ctx.measureText(text).width <= maxWidth) {
    return text;
  }
  let truncated = text;
  while (truncated.length > 3 && ctx.measureText(truncated + '...').width > maxWidth) {
    truncated = truncated.slice(0, -1);
  }
  return truncated + '...';
}

function canvasToCleanBase64(canvas: HTMLCanvasElement): string {
  try {
    const dataUrl = canvas.toDataURL('image/png');
    return dataUrl.replace(/^data:image\/png;base64,/, '');
  } catch {
    return FALLBACK_1X1_PNG;
  }
}
