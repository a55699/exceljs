const JSZip = require('jszip');
const {SaxesParser} = require('saxes');

const ExcelJS = verquire('exceljs');

const PIVOT_XML_FILEPATHS = [
  'xl/pivotCache/pivotCacheDefinition1.xml',
  'xl/pivotCache/pivotCacheRecords1.xml',
  'xl/pivotTables/pivotTable1.xml',
];

function expectWellFormedXml(xml) {
  const parser = new SaxesParser();
  let error = null;
  parser.on('error', err => {
    error = err;
  });
  parser.write(xml).close();
  expect(error).to.be.null();
}

function addSourceSheet(workbook, rows) {
  const worksheet = workbook.addWorksheet('Sheet1');
  worksheet.addRows(rows);
  return worksheet;
}

describe('Workbook', () => {
  describe('Pivot Tables with XML special characters', () => {
    it('escapes field names and values in the pivot XML', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet1 = addSourceSheet(workbook, [
        ['Name <&>', 'Group "Q"', 'R&D <x>'],
        ['a & b', 'g1', 1],
        ['<c>', 'g2', 2],
      ]);
      workbook.addWorksheet('Sheet2').addPivotTable({
        sourceSheet: worksheet1,
        rows: ['Name <&>'],
        columns: ['Group "Q"'],
        values: ['R&D <x>'],
        metric: 'sum',
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const zip = await JSZip.loadAsync(buffer);
      const readXml = filepath => zip.file(filepath).async('string');
      const xmls = await Promise.all(PIVOT_XML_FILEPATHS.map(readXml));
      xmls.forEach(expectWellFormedXml);
      const pivotXml = await zip.file('xl/pivotTables/pivotTable1.xml').async('string');
      expect(pivotXml).to.include('name="Sum of R&amp;D &lt;x&gt;"');

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      expect(workbook2.getWorksheet('Sheet1').getCell('A2').value).to.equal('a & b');
    });

    it('throws a clear error for an empty header', () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet1 = addSourceSheet(workbook, [
        ['A', null, 'B', 'C'],
        ['a1', 'x', 'b1', 1],
      ]);
      const addPivotTable = () => {
        workbook.addWorksheet('Sheet2').addPivotTable({
          sourceSheet: worksheet1,
          rows: ['A'],
          columns: ['B'],
          values: ['C'],
        });
      };
      expect(addPivotTable).to.throw('The header in column 2 of Sheet1 is empty.');
    });

    it('throws a clear error for data to the right of the last header', () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet1 = addSourceSheet(workbook, [
        ['A', 'B', 'C'],
        ['a1', 'b1', 1, 'extra'],
      ]);
      const addPivotTable = () => {
        workbook.addWorksheet('Sheet2').addPivotTable({
          sourceSheet: worksheet1,
          rows: ['A'],
          columns: ['B'],
          values: ['C'],
        });
      };
      expect(addPivotTable).to.throw('The header in column 4 of Sheet1 is empty.');
    });

    it('throws a clear error for a duplicate header', () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet1 = addSourceSheet(workbook, [
        ['A', 'B', 'C', 'a'],
        ['a1', 'b1', 1, 'q'],
      ]);
      const addPivotTable = () => {
        workbook.addWorksheet('Sheet2').addPivotTable({
          sourceSheet: worksheet1,
          rows: ['B'],
          columns: ['A'],
          values: ['C'],
        });
      };
      expect(addPivotTable).to.throw('The header name "a" is duplicated in Sheet1.');
    });
  });
});
