import { IDocumentDriver } from './types';

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
      const chartType = Excel.ChartType[type] || Excel.ChartType.columnClustered;
      const chart = sheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);
      if (title) chart.title.text = title;
      await context.sync();
    });
  }
}
