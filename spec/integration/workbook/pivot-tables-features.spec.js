const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const SOURCE_ROWS = [
  ['Region', 'Product', 'Sales', 'Latest'],
  ['North', 'Apple', 10, 1],
  ['North', 'Pear', 20, 0],
  ['South', 'Apple', 30, 1],
  ['South', 'Pear', 40, 0],
];

function addSourceSheet(workbook, rows = SOURCE_ROWS) {
  const worksheet = workbook.addWorksheet('Source');
  worksheet.addRows(rows);
  return worksheet;
}

async function writeZip(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  return {buffer, zip: await JSZip.loadAsync(buffer)};
}

function readFile(zip, name) {
  return zip.file(name).async('string');
}

describe('Workbook', () => {
  describe('Pivot Tables features', () => {
    it('groups values that differ only in case into one item', async () => {
      const workbook = new ExcelJS.Workbook();
      const source = addSourceSheet(workbook, [
        ['Fruit', 'Store', 'Sales'],
        ['Apple', 's1', 1],
        ['apple', 's2', 2],
        ['APPLE', 's1', 3],
        ['Pear', 's2', 4],
      ]);
      workbook.addWorksheet('Pivot').addPivotTable({
        sourceSheet: source,
        rows: ['Fruit'],
        columns: ['Store'],
        values: ['Sales'],
        metric: 'sum',
      });

      const {zip} = await writeZip(workbook);
      const definition = await readFile(zip, 'xl/pivotCache/pivotCacheDefinition1.xml');
      const fruitField = definition.match(/<cacheField name="Fruit"[\s\S]*?<\/cacheField>/)[0];
      expect(fruitField.match(/<s v="[^"]*" \/>/g)).to.deep.equal([
        '<s v="Apple" />',
        '<s v="Pear" />',
      ]);
      const records = await readFile(zip, 'xl/pivotCache/pivotCacheRecords1.xml');
      const fruitIndexes = [...records.matchAll(/<r>\s*<x v="(\d+)" \/>/g)].map(m => m[1]);
      expect(fruitIndexes).to.deep.equal(['0', '0', '0', '1']);
    });

    it('writes one cache per pivot table built from the same source', async () => {
      const workbook = new ExcelJS.Workbook();
      const source = addSourceSheet(workbook);
      ['ByRegion', 'ByProduct', 'ByLatest'].forEach((name, index) => {
        workbook.addWorksheet(name).addPivotTable({
          sourceSheet: source,
          rows: [['Region', 'Product', 'Latest'][index]],
          columns: [['Product', 'Region', 'Region'][index]],
          values: ['Sales'],
          metric: 'sum',
        });
      });

      const {buffer, zip} = await writeZip(workbook);
      const workbookXml = await readFile(zip, 'xl/workbook.xml');
      const cacheIds = [...workbookXml.matchAll(/<pivotCache cacheId="(\d+)"/g)].map(m => m[1]);
      expect(cacheIds).to.deep.equal(['10', '11', '12']);
      [1, 2, 3].forEach(n => {
        expect(zip.file(`xl/pivotTables/pivotTable${n}.xml`)).to.not.be.null();
        expect(zip.file(`xl/pivotCache/pivotCacheDefinition${n}.xml`)).to.not.be.null();
      });
      const pivotTable2 = await readFile(zip, 'xl/pivotTables/pivotTable2.xml');
      expect(pivotTable2).to.include('cacheId="11"');

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      expect(workbook2.worksheets.map(worksheet => worksheet.name)).to.deep.equal([
        'Source',
        'ByRegion',
        'ByProduct',
        'ByLatest',
      ]);
    });

    it('writes page fields and hides all but the default item', async () => {
      const workbook = new ExcelJS.Workbook();
      const source = addSourceSheet(workbook);
      workbook.addWorksheet('WithDefault').addPivotTable({
        sourceSheet: source,
        rows: ['Region'],
        columns: ['Product'],
        values: ['Sales'],
        pages: ['Latest'],
        pageDefaults: {Latest: 1},
        metric: 'sum',
      });
      workbook.addWorksheet('NoDefault').addPivotTable({
        sourceSheet: source,
        rows: ['Region'],
        columns: ['Product'],
        values: ['Sales'],
        pages: ['Latest'],
        metric: 'sum',
      });

      const {zip} = await writeZip(workbook);
      const withDefault = await readFile(zip, 'xl/pivotTables/pivotTable1.xml');
      expect(withDefault).to.include('<pageField fld="3" hier="-1" />');
      const pageField = withDefault.match(/<pivotField axis="axisPage"[\s\S]*?<\/pivotField>/)[0];
      // Latest has items [0, 1]; default 1 -> item 0 hidden
      expect(pageField).to.include('<item h="1" x="0" />');
      expect(pageField).to.include('<item x="1" />');

      const noDefault = await readFile(zip, 'xl/pivotTables/pivotTable2.xml');
      expect(noDefault).to.include('axis="axisPage"');
      expect(noDefault).to.not.include('h="1"');
    });

    it('rejects a page field that is also a row, column or value field', () => {
      const workbook = new ExcelJS.Workbook();
      const source = addSourceSheet(workbook);
      const addPivotTable = () => {
        workbook.addWorksheet('Pivot').addPivotTable({
          sourceSheet: source,
          rows: ['Region'],
          columns: ['Product'],
          values: ['Sales'],
          pages: ['Region'],
        });
      };
      const message = 'Page field "Region" cannot also be used as a row, column, or value field.';
      expect(addPivotTable).to.throw(message);
    });

    it('rejects pageDefaults for a field that is not a page field', () => {
      const workbook = new ExcelJS.Workbook();
      const source = addSourceSheet(workbook);
      const addPivotTable = () => {
        workbook.addWorksheet('Pivot').addPivotTable({
          sourceSheet: source,
          rows: ['Region'],
          columns: ['Product'],
          values: ['Sales'],
          pages: ['Latest'],
          pageDefaults: {Region: 'North'},
        });
      };
      expect(addPivotTable).to.throw('pageDefaults field "Region" is not in the pages array.');
    });
  });
});
