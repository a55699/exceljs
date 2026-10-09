const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

async function writeStream(addTables, options = {}) {
  const stream = new ExcelJS.stream.xlsx.WorkbookWriter({useStyles: true, ...options});
  addTables(stream);
  await stream.commit();
  return stream.stream.read();
}

describe('WorkbookWriter', () => {
  describe('tables', () => {
    it('writes tables of two worksheets', async () => {
      const buffer = await writeStream(workbook => {
        const data = workbook.addWorksheet('Data');
        data.addTable({
          name: 'Sales',
          ref: 'A1',
          totalsRow: true,
          columns: [
            {name: 'Month'},
            {name: 'Qty', totalsRowFunction: 'sum', style: {numFmt: '0.00'}},
          ],
          rows: [
            ['Jan', 1],
            ['Feb', 2],
          ],
        });
        data.getCell('A1').note = 'a note';
        data.commit();
        const other = workbook.addWorksheet('Other');
        other.addTable({name: 'Codes', ref: 'B2', columns: [{name: 'Code'}], rows: [['x']]});
        other.commit();
      });

      const zip = await JSZip.loadAsync(buffer);
      const sales = await zip.file('xl/tables/table1.xml').async('string');
      const codes = await zip.file('xl/tables/table2.xml').async('string');
      expect(sales).to.include('name="Sales"').and.to.include('ref="A1:B4"');
      expect(codes).to.include('name="Codes"').and.to.include('ref="B2:B3"');
      const contentTypes = await zip.file('[Content_Types].xml').async('string');
      expect(contentTypes).to.include('PartName="/xl/tables/table1.xml"');
      expect(contentTypes).to.include('PartName="/xl/tables/table2.xml"');
      const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
      // the table part comes after the drawing of the note
      expect(sheet).to.match(
        /<legacyDrawing [^>]*\/><tableParts count="1"><tablePart r:id="rId\d+"\/><\/tableParts><\/worksheet>/
      );

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const worksheet = workbook.getWorksheet('Data');
      expect(worksheet.getCell('B4').value.formula).to.equal('SUBTOTAL(109,Sales[Qty])');
      expect(worksheet.getTable('Sales')).to.be.ok();
    });

    it('writes a table without styles', async () => {
      const buffer = await writeStream(
        workbook => {
          const worksheet = workbook.addWorksheet('Data');
          worksheet.addTable({
            name: 'Plain',
            ref: 'A1',
            columns: [{name: 'a', style: {font: {bold: true}}}],
            rows: [[1]],
          });
          worksheet.commit();
        },
        {useStyles: false}
      );
      const zip = await JSZip.loadAsync(buffer);
      expect(await zip.file('xl/tables/table1.xml').async('string')).to.include('name="Plain"');
    });

    it('cannot add a table on rows already committed', () => {
      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({});
      const worksheet = workbook.addWorksheet('Data');
      worksheet.addRow(['a']).commit();
      expect(() =>
        worksheet.addTable({name: 'Late', ref: 'A1', columns: [{name: 'a'}], rows: [[1]]})
      ).to.throw(/first row has been committed/);
    });
  });
});
