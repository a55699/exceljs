const ExcelJS = verquire('exceljs');

// More unique values than fit in one function call's arguments
// (Math.min(...values) throws "Maximum call stack size exceeded").
const ROW_COUNT = 150000;

describe('Workbook', () => {
  describe('Pivot Tables with large source data', () => {
    it('writes a pivot table over 150k rows of unique numbers', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRow(['Group', 'Type', 'Value']);
      for (let i = 0; i < ROW_COUNT; i++) {
        worksheet1.addRow([`g${i % 10}`, `t${i % 2}`, i]);
      }
      workbook.addWorksheet('Sheet2').addPivotTable({
        sourceSheet: worksheet1,
        rows: ['Group'],
        columns: ['Type'],
        values: ['Value'],
        metric: 'sum',
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      const lastRow = ROW_COUNT + 1;
      expect(workbook2.getWorksheet('Sheet1').getCell(`C${lastRow}`).value).to.equal(ROW_COUNT - 1);
    }).timeout(60000);
  });
});
