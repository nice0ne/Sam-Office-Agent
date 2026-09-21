import { CleanDataOptions, CleanDataResult, ConditionalFormattingOptions, IDocumentDriver } from './types';

declare const Excel: any;

export class ExcelDriver implements Partial<IDocumentDriver> {
  hostType = 'Excel';

  async readActiveRange() {
    return await Excel.run(async (context: any) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'values', 'formulas']);
      await context.sync();
      return {
        address: range.address,
        values: range.values,
        formulas: range.formulas,
      };
    });
  }

  async readActiveSheetData(rangeAddress?: string) {
    return await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load(['name']);
      let range: any;
      if (rangeAddress) {
        range = sheet.getRange(rangeAddress);
      } else {
        range = typeof sheet.getUsedRangeOrNullObject === 'function'
          ? sheet.getUsedRangeOrNullObject(true)
          : sheet.getUsedRange(true);
      }
      range.load(['address', 'values', 'formulas', 'rowCount', 'columnCount']);
      await context.sync();

      if (range.isNullObject || !range.values) {
        return {
          sheetName: sheet.name,
          address: 'Empty',
          rowCount: 0,
          columnCount: 0,
          values: [],
          formulas: [],
        };
      }

      return {
        sheetName: sheet.name,
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
        formulas: range.formulas,
      };
    });
  }

  async writeCells(rangeAddress: string, values?: any[][], formulas?: string[][]) {
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      if (values) range.values = values;
      if (formulas) range.formulas = formulas;
      await context.sync();
    });
  }

  async formatRange(rangeAddress: string, styles: Record<string, any>) {
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      if (styles.bold !== undefined) range.format.font.bold = styles.bold;
      if (styles.color) range.format.font.color = styles.color;
      if (styles.fillColor) range.format.fill.color = styles.fillColor;
      if (styles.numberFormat) range.numberFormat = [[styles.numberFormat]];
      await context.sync();
    });
  }

  async createChart(type: string, dataRange: string, title?: string) {
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(dataRange);

      let chartType = Excel.ChartType?.columnClustered || 'ColumnClustered';
      if (type && Excel.ChartType) {
        const normalized = type.toLowerCase();
        const foundKey = Object.keys(Excel.ChartType).find(
          k => k.toLowerCase() === normalized || String(Excel.ChartType[k]).toLowerCase() === normalized
        );
        if (foundKey) {
          chartType = Excel.ChartType[foundKey];
        } else {
          chartType = type;
        }
      } else if (type) {
        chartType = type;
      }

      const chart = sheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);
      if (title && chart.title) {
        chart.title.text = title;
        chart.title.visible = true;
      }
      await context.sync();
    });
  }

  async addWorksheet(sheetName: string) {
    return await Excel.run(async (context: any) => {
      const sheets = context.workbook.worksheets;
      const sheet = sheets.add(sheetName);
      sheet.activate();
      sheet.load('name');
      await context.sync();
      return { sheetName: sheet.name };
    });
  }

  async autoFitColumns(rangeAddress?: string) {
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = rangeAddress
        ? sheet.getRange(rangeAddress)
        : (typeof sheet.getUsedRangeOrNullObject === 'function'
            ? sheet.getUsedRangeOrNullObject(true)
            : sheet.getUsedRange(true));
      if (range && !range.isNullObject) {
        range.format.autofitColumns();
      }
      await context.sync();
    });
  }

  async sortRange(
    rangeAddress: string,
    columnIndex: number,
    ascending: boolean = true,
    hasHeaders: boolean = true,
    enableAutoFilter: boolean = false
  ) {
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      range.sort.apply([{ key: columnIndex, ascending }], false, hasHeaders);
      if (enableAutoFilter && typeof sheet.autoFilter?.apply === 'function') {
        sheet.autoFilter.apply(range);
      }
      await context.sync();
    });
  }

  async cleanData(options?: CleanDataOptions): Promise<CleanDataResult> {
    return await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let range: any;
      if (options?.range) {
        range = sheet.getRange(options.range);
      } else {
        range = typeof sheet.getUsedRangeOrNullObject === 'function'
          ? sheet.getUsedRangeOrNullObject(true)
          : sheet.getUsedRange(true);
      }
      range.load(['address', 'values', 'rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
      await context.sync();

      if (range.isNullObject || !range.values || range.values.length === 0) {
        return {
          cleanedRows: 0,
          removedDuplicatesCount: 0,
          trimmedCellsCount: 0,
          filledCellsCount: 0,
        };
      }

      let values: any[][] = range.values.map((row: any[]) => [...row]);
      let trimmedCellsCount = 0;
      let filledCellsCount = 0;
      let removedDuplicatesCount = 0;

      // Trim whitespace
      if (options?.trimWhitespace) {
        for (let r = 0; r < values.length; r++) {
          for (let c = 0; c < values[r].length; c++) {
            const cell = values[r][c];
            if (typeof cell === 'string') {
              const trimmed = cell.trim();
              if (trimmed !== cell) {
                values[r][c] = trimmed;
                trimmedCellsCount++;
              }
            }
          }
        }
      }

      // Fill empty values
      if (options?.fillEmptyValues !== undefined) {
        for (let r = 0; r < values.length; r++) {
          for (let c = 0; c < values[r].length; c++) {
            const cell = values[r][c];
            if (cell === null || cell === undefined || (typeof cell === 'string' && cell.trim() === '') || cell === '') {
              values[r][c] = options.fillEmptyValues;
              filledCellsCount++;
            }
          }
        }
      }

      // Remove duplicates
      if (options?.removeDuplicates) {
        const seen = new Set<string>();
        const unique: any[][] = [];
        for (const row of values) {
          const key = JSON.stringify(row);
          if (seen.has(key)) {
            removedDuplicatesCount++;
          } else {
            seen.add(key);
            unique.push(row);
          }
        }
        values = unique;
      }

      // Batch write changes in a single sync
      if (removedDuplicatesCount > 0) {
        range.clear();
        if (values.length > 0) {
          const newRange = sheet.getRangeByIndexes(range.rowIndex, range.columnIndex, values.length, range.columnCount);
          newRange.values = values;
        }
      } else {
        range.values = values;
      }

      await context.sync();

      return {
        cleanedRows: values.length,
        removedDuplicatesCount,
        trimmedCellsCount,
        filledCellsCount,
      };
    });
  }

  async applyConditionalFormatting(options: ConditionalFormattingOptions): Promise<{ success: boolean; rule: string }> {
    return await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(options.range);

      if (range.conditionalFormats && typeof range.conditionalFormats.add === 'function') {
        if (options.type === 'color_scale') {
          range.conditionalFormats.add(Excel.ConditionalFormatType?.colorScale || 'ColorScale');
        } else if (options.type === 'data_bar') {
          const cf = range.conditionalFormats.add(Excel.ConditionalFormatType?.dataBar || 'DataBar');
          if (options.color && cf.dataBar) {
            cf.dataBar.barColor = options.color;
          }
        } else if (options.type === 'highlight_threshold') {
          const cf = range.conditionalFormats.add(Excel.ConditionalFormatType?.cellValue || 'CellValue');
          if (cf.cellValue) {
            cf.cellValue.operator = Excel.ConditionalCellValueOperator?.greaterThan || 'GreaterThan';
            cf.cellValue.formula1 = String(options.thresholdValue ?? 0);
            if (options.color) {
              cf.cellValue.format.fill.color = options.color;
            } else {
              cf.cellValue.format.fill.color = '#FEE2E2';
              cf.cellValue.format.font.color = '#991B1B';
            }
          }
        }
      }

      await context.sync();
      return {
        success: true,
        rule: `${options.type} on ${options.range}`,
      };
    });
  }
}
