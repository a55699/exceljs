const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

async function writePivot(rows) {
  const workbook = new ExcelJS.Workbook();
  const worksheet1 = workbook.addWorksheet('Sheet1');
  rows.forEach((values, index) => {
    if (values) {
      worksheet1.getRow(index + 1).values = values;
    }
  });
  workbook.addWorksheet('Sheet2').addPivotTable({
    sourceSheet: worksheet1,
    rows: ['A'],
    columns: ['B'],
    values: ['C'],
    metric: 'sum',
  });
  const buffer = await workbook.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const read = filepath => zip.file(filepath).async('string');
  return {
    buffer,
    definition: await read('xl/pivotCache/pivotCacheDefinition1.xml'),
    records: await read('xl/pivotCache/pivotCacheRecords1.xml'),
    pivotTable: await read('xl/pivotTables/pivotTable1.xml'),
  };
}

function countMatches(xml, regex) {
  return (xml.match(regex) || []).length;
}

describe('Workbook', () => {
  describe('Pivot Tables counts', () => {
    it('writes recordCount and records count as the number of data rows', async () => {
      const {definition, records} = await writePivot([
        ['A', 'B', 'C'],
        ['a1', 'b1', 1],
        null,
        ['a2', 'b2', 2],
      ]);
      expect(definition).to.include('recordCount="2"');
      expect(records).to.match(/<pivotCacheRecords [^>]*count="2"/);
      expect(countMatches(records, /<r>/g)).to.equal(2);
    });

    it('writes items count equal to the number of items', async () => {
      const {pivotTable} = await writePivot([
        ['A', 'B', 'C'],
        ['a1', 'b1', 1],
        ['a2', 'b2', 2],
        ['a3', 'b1', 3],
      ]);
      const itemsCounts = [...pivotTable.matchAll(/<items count="(\d+)">([\s\S]*?)<\/items>/g)];
      expect(itemsCounts).to.have.length(2);
      for (const [, count, items] of itemsCounts) {
        expect(countMatches(items, /<item /g)).to.equal(Number(count));
      }
    });

    it('writes one value per field for rows shorter than the header', async () => {
      const {records} = await writePivot([
        ['A', 'B', 'C', 'D'],
        ['a1', 'b1', 1, 'x'],
        ['a2', 'b2', 2],
      ]);
      const recordList = records.match(/<r>[\s\S]*?<\/r>/g);
      expect(recordList).to.have.length(2);
      for (const record of recordList) {
        expect(countMatches(record, /<[xnsmdb] /g)).to.equal(4);
      }
    });

    it('writes an empty column like Excel instead of an Infinity range', async () => {
      const {buffer, definition} = await writePivot([
        ['A', 'B', 'C', 'D'],
        ['a1', 'b1', 1],
        ['a2', 'b2', 2],
      ]);
      expect(definition).to.not.include('Infinity');
      expect(definition).to.include('containsBlank="1"');

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      expect(workbook2.getWorksheet('Sheet1').getCell('D1').value).to.equal('D');
    });
  });
});
