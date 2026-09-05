import { describe, it, expect } from 'vitest';
import { validateExcelFormula } from '../formulaValidator';

describe('validateExcelFormula', () => {
  it('validates correct Excel formulas', () => {
    expect(validateExcelFormula('=SUM(A1:A10)').isValid).toBe(true);
    expect(validateExcelFormula('=XLOOKUP(A2, Sheet2!A:A, Sheet2!B:B, "N/A")').isValid).toBe(true);
    expect(validateExcelFormula('=IF(B2>100, B2*0.1, 0)').isValid).toBe(true);
  });

  it('flags unclosed parentheses', () => {
    const res = validateExcelFormula('=SUM(A1:A10');
    expect(res.isValid).toBe(false);
    expect(res.error).toContain('kurung');
  });

  it('flags formula missing leading equal sign', () => {
    const res = validateExcelFormula('SUM(A1:A10)');
    expect(res.isValid).toBe(false);
    expect(res.error).toContain('=');
  });

  it('flags excess closing parentheses', () => {
    const res = validateExcelFormula('=SUM(A1:A10))');
    expect(res.isValid).toBe(false);
    expect(res.error).toContain('berlebih');
  });

  it('handles parentheses inside string literals correctly', () => {
    expect(validateExcelFormula('=IF(A1="test (1)", "ok )", "fail")').isValid).toBe(true);
  });

  it('handles Excel doubled quotes "" inside string literals correctly', () => {
    expect(validateExcelFormula('=IF(A1="He said ""Hello (world)""", 1, 0)').isValid).toBe(true);
    expect(validateExcelFormula('="Quotes: ""("" and "")"""').isValid).toBe(true);
  });
});
