const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

function makeWorkbook(options = {}) {
  const workbook = new ExcelJS.Workbook();
  const source = workbook.addWorksheet('Source');
  source.addRows([
    ['Region', 'Product', 'Sales', 'Year'],
    ['North', 'Apple', 10, 2024],
    ['South', 'Pear', 20, 2024],
  ]);
  const pivotTable = workbook.addWorksheet('Pivot').addPivotTable({
    sourceSheet: source,
    rows: ['Region'],
    columns: ['Product'],
    values: ['Sales'],
    metric: 'sum',
    ...options,
  });
  return {workbook, source, pivotTable};
}

async function readPivotParts(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const read = name => zip.file(name).async('string');
  return {
    buffer,
    definition: await read('xl/pivotCache/pivotCacheDefinition1.xml'),
    pivotTable: await read('xl/pivotTables/pivotTable1.xml'),
  };
}

describe('Workbook', () => {
  describe('Pivot Tables whose source data changes after addPivotTable', () => {
    it('writes rows added after the pivot table was created', async () => {
      const {workbook, source, pivotTable} = makeWorkbook();
      source.addRow(['East', 'Plum', 30, 2025]);
      source.spliceRows(2, 0, ['West', 'Fig', 40, 2025]);

      const {buffer, definition} = await readPivotParts(workbook);
      expect(definition).to.include('recordCount="4"');
      const regionField = definition.match(/<cacheField name="Region"[\s\S]*?<\/cacheField>/)[0];
      expect(regionField.match(/<s v="[^"]*" \/>/g)).to.deep.equal([
        '<s v="East" />',
        '<s v="North" />',
        '<s v="South" />',
        '<s v="West" />',
      ]);
      expect(pivotTable.cacheFields[0].sharedItems).to.have.length(4);

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      expect(workbook2.getWorksheet('Source').getCell('A2').value).to.equal('West');
    });

    it('resolves pageDefaults against the data at write time', async () => {
      const {workbook, source} = makeWorkbook({pages: ['Year'], pageDefaults: {Year: 2025}});
      source.addRow(['East', 'Plum', 30, 2025]);

      const {pivotTable} = await readPivotParts(workbook);
      const pageField = pivotTable.match(/<pivotField axis="axisPage"[\s\S]*?<\/pivotField>/)[0];
      // Year items [2024, 2025]; default 2025 -> item 0 hidden
      expect(pageField).to.include('<item h="1" x="0" />');
      expect(pageField).to.include('<item x="1" />');
    });

    it('throws a clear error when a used header was changed', async () => {
      const {workbook, source} = makeWorkbook();
      source.getCell('B1').value = 'Item';

      let error;
      try {
        await workbook.xlsx.writeBuffer();
      } catch (err) {
        error = err;
      }
      expect(error && error.message).to.equal('The header name "Product" was not found in Source.');
    });
  });
});
