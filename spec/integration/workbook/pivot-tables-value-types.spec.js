const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const CACHE_DEFINITION = 'xl/pivotCache/pivotCacheDefinition1.xml';
const CACHE_RECORDS = 'xl/pivotCache/pivotCacheRecords1.xml';

async function writePivot(rows) {
  const workbook = new ExcelJS.Workbook();
  const worksheet1 = workbook.addWorksheet('Sheet1');
  worksheet1.addRows(rows);
  workbook.addWorksheet('Sheet2').addPivotTable({
    sourceSheet: worksheet1,
    rows: ['A'],
    columns: ['B'],
    values: ['C'],
    metric: 'sum',
  });
  const buffer = await workbook.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const definition = await zip.file(CACHE_DEFINITION).async('string');
  const records = await zip.file(CACHE_RECORDS).async('string');
  return {buffer, definition, records};
}

describe('Workbook', () => {
  describe('Pivot Tables with non-primitive cell values', () => {
    it('uses formula results instead of throwing', async () => {
      const {buffer, definition} = await writePivot([
        ['A', 'B', 'C'],
        ['a1', 'b1', {formula: '1+1', result: 2}],
        ['a2', 'b2', {formula: '2+2', result: 4}],
        [{formula: '"a"&"1"', result: 'a1'}, 'b1', 5],
      ]);
      expect(definition).to.include('<n v="2" /><n v="4" /><n v="5" />');
      expect(definition).to.include('<s v="a1" /><s v="a2" />');

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      expect(workbook2.getWorksheet('Sheet1').getCell('C2').value).to.deep.equal({
        formula: '1+1',
        result: 2,
      });
    });

    it('uses the text of rich text and hyperlink cells', async () => {
      const {definition} = await writePivot([
        ['A', 'B', 'C'],
        [{richText: [{text: 'rich'}, {font: {bold: true}, text: ' text'}]}, 'b1', 1],
        [{text: 'link', hyperlink: 'https://example.com'}, 'b2', 2],
      ]);
      expect(definition).to.include('<s v="link" /><s v="rich text" />');
    });

    it('writes dates as date items and merges equal dates', async () => {
      const {definition, records} = await writePivot([
        ['A', 'B', 'C'],
        [new Date(Date.UTC(2024, 1, 1)), 'b1', 1],
        [new Date(Date.UTC(2024, 0, 1)), 'b2', 2],
        [new Date(Date.UTC(2024, 1, 1)), 'b2', 3],
      ]);
      expect(definition).to.include('containsDate="1"');
      const dates = '<d v="2024-01-01T00:00:00" /><d v="2024-02-01T00:00:00" />';
      expect(definition).to.include(dates);
      const firstFieldIndexes = [...records.matchAll(/<r>\s*<x v="(\d+)" \/>/g)].map(m => m[1]);
      expect(firstFieldIndexes).to.deep.equal(['1', '0', '1']);
    });

    it('does not write non-decimal strings as numbers', async () => {
      const {definition} = await writePivot([
        ['A', 'B', 'C'],
        ['0x10', 'b1', 1],
        ['12', 'b2', 2],
      ]);
      expect(definition).to.include('<s v="0x10" />');
      expect(definition).to.not.include('<n v="0x10" />');
    });

    it('sorts mixed numbers and strings with numbers first, numerically', async () => {
      const {definition} = await writePivot([
        ['A', 'B', 'C'],
        [10, 'b1', 1],
        ['x', 'b2', 2],
        [9, 'b1', 3],
      ]);
      const attributes = 'containsMixedTypes="1" containsNumber="1" containsInteger="1"';
      expect(definition).to.include(`${attributes} minValue="9" maxValue="10"`);
      expect(definition).to.include('<n v="9" /><n v="10" /><s v="x" />');
    });

    it('describes mixed numbers and dates with a date range only, like Excel', async () => {
      const {definition} = await writePivot([
        ['A', 'B', 'C'],
        [10, 'b1', 1],
        [new Date(Date.UTC(2024, 0, 1)), 'b2', 2],
      ]);
      const attributes = 'containsSemiMixedTypes="0" containsString="0" containsDate="1"';
      expect(definition).to.include(`${attributes} containsMixedTypes="1"`);
      expect(definition).to.include('minDate="2024-01-01T00:00:00"');
      expect(definition).to.not.include('minValue="10"');
    });
  });
});
