import 'regenerator-runtime/runtime';

import { expect } from 'chai';
import ExcelJSRuntime from '../../index';
// index.ts re-exports the untyped dist build, so check against index.d.ts explicitly
import * as ExcelJSTypes from '../../index.d';

const ExcelJS = ExcelJSRuntime as unknown as typeof ExcelJSTypes;

describe('typescript fork API', () => {
  it('types addPivotTable', async () => {
    const wb = new ExcelJS.Workbook();
    const source = wb.addWorksheet('Source');
    source.addRows([
      ['Region', 'Product', 'Sales'],
      ['North', 'Apple', 10],
      ['South', 'Pear', 20],
    ]);
    const options: ExcelJSTypes.PivotTableOptions = {
      sourceSheet: source,
      rows: ['Region'],
      columns: ['Product'],
      values: ['Sales'],
      metric: 'count',
    };
    const pivotTable: ExcelJSTypes.PivotTable = wb.addWorksheet('Pivot').addPivotTable(options);
    expect(pivotTable.cacheFields.map(field => field.name)).to.deep.equal([
      'Region',
      'Product',
      'Sales',
    ]);

    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    expect(wb2.getWorksheet('Source').getCell('C3').value).to.equal(20);
  });

  it('types form checkboxes', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form');
    const checkbox: ExcelJSTypes.FormCheckbox = ws.addFormCheckbox('B2:D3', {
      text: 'Accept terms',
      link: 'A2',
    });
    checkbox.checked = true;
    ws.addFormCheckbox({ startCol: 1, startRow: 4, endCol: 3, endRow: 5 });
    expect(ws.getFormCheckboxes()).to.have.length(2);
    expect(checkbox.link).to.equal('$A$2');
  });

  it('types charts', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sales Data');
    ws.addRows([
      ['Month', 'East', 'West'],
      ['Jan', 10, 30],
      ['Feb', 40, 60],
    ]);
    const options: ExcelJSTypes.ChartOptions = {
      type: 'bar',
      title: 'Sales',
      series: [
        {
          name: { sheet: ws, ref: 'B1' },
          categories: { sheet: 'Sales Data', ref: 'A2:A3' },
          values: "'Sales Data'!$B$2:$B$3",
        },
      ],
      axes: { y: { min: 0, numFmt: '#,##0' } },
      legend: { position: 'b' },
      range: 'E2:L17',
    };
    const chart: ExcelJSTypes.Chart = ws.addChart(options);
    ws.addChart({
      type: 'line',
      data: { sheet: ws, ref: 'A1:C3' },
      range: { tl: { col: 4, row: 18 }, ext: { width: 480, height: 288 } },
    });
    expect(chart.type).to.equal('bar');
    expect(ws.getCharts()).to.have.length(2);
  });

  it('types table edits', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');
    const table: ExcelJSTypes.Table = ws.addTable({
      name: 'Sales',
      ref: 'A1',
      columns: [
        { name: 'Qty' },
        { name: 'Double', calculatedColumnFormula: 'Sales[[#This Row],[Qty]]*2' },
      ],
      rows: [[1, { formula: 'Sales[[#This Row],[Qty]]*2' }]],
    });
    table.addRow([2]);
    table.name = 'Orders';
    table.theme = 'TableStyleLight9';
    ws.getCell('D1').value = { formula: 'SUM(Orders[Qty])', date1904: false };
    ws.removeTable('Orders');
    expect(ws.getCell('D1').formula).to.equal('SUM(Data!$A$2:$A$3)');
  });

  it('types numFmtLimit, addNumberFormat and NumberFormatLimitError', () => {
    const wb = new ExcelJS.Workbook({ numFmtLimit: 1 });
    expect(wb.numFmtLimit).to.equal(1);
    const id: number = wb.addNumberFormat('0.000');
    expect(id).to.be.a('number');
    try {
      wb.addNumberFormat('0.0000');
      expect.fail('expected NumberFormatLimitError');
    } catch (error) {
      expect(error).to.be.instanceOf(ExcelJS.NumberFormatLimitError);
      const limitError = error as ExcelJSTypes.NumberFormatLimitError;
      expect(limitError.limit).to.equal(1);
      expect(limitError.formatCode).to.equal('0.0000');
    }
  });
});
