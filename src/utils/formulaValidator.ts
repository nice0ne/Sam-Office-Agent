export interface FormulaValidationResult {
  isValid: boolean;
  error?: string;
  sanitizedFormula?: string;
}

export function validateExcelFormula(formula: string): FormulaValidationResult {
  const trimmed = formula.trim();
  if (!trimmed.startsWith('=')) {
    return {
      isValid: false,
      error: 'Formula Excel harus diawali dengan tanda sama dengan (=).',
    };
  }

  let openParenCount = 0;
  let inQuotes = false;

  for (let i = 0; i < trimmed.length; i++) {
    const char = trimmed[i];
    if (char === '"' && (i === 0 || trimmed[i - 1] !== '\\')) {
      inQuotes = !inQuotes;
    } else if (!inQuotes) {
      if (char === '(') openParenCount++;
      else if (char === ')') openParenCount--;
    }

    if (openParenCount < 0) {
      return {
        isValid: false,
        error: 'Tanda kurung tutup ")" berlebih tanpa kurung buka yang sesuai.',
      };
    }
  }

  if (openParenCount > 0) {
    return {
      isValid: false,
      error: `Tanda kurung tidak seimbang: ada ${openParenCount} kurung buka "(" yang belum ditutup.`,
    };
  }

  return {
    isValid: true,
    sanitizedFormula: trimmed,
  };
}
