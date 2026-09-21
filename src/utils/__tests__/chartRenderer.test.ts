import { describe, it, expect } from 'vitest';
import { renderChartToBase64Png } from '../chartRenderer';

describe('chartRenderer', () => {
  it('returns valid base64 PNG string for bar chart', () => {
    const base64 = renderChartToBase64Png({
      chartType: 'bar',
      labels: ['Jan', 'Feb', 'Mar'],
      data: [100, 200, 150],
      title: 'Penjualan Q1',
    });

    expect(typeof base64).toBe('string');
    expect(base64.length).toBeGreaterThan(20);
    // Ensure dataUrl prefix is stripped for Word inline picture API
    expect(base64).not.toContain('data:image/png;base64,');
  });

  it('handles empty data gracefully without throwing', () => {
    const base64 = renderChartToBase64Png({
      chartType: 'line',
      labels: [],
      data: [],
    });

    expect(typeof base64).toBe('string');
    expect(base64.length).toBeGreaterThan(0);
  });

  it('renders line and pie charts without errors', () => {
    const lineBase64 = renderChartToBase64Png({
      chartType: 'line',
      labels: ['A', 'B'],
      data: [10, 20],
      title: 'Tren Garis',
    });
    expect(lineBase64).toBeDefined();

    const pieBase64 = renderChartToBase64Png({
      chartType: 'pie',
      labels: ['Produk A', 'Produk B', 'Produk C'],
      data: [40, 35, 25],
      title: 'Pangsa Pasar',
    });
    expect(pieBase64).toBeDefined();
  });
});
