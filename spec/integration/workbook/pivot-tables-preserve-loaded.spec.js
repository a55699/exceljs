const path = require('path');
const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// Made in Excel: pivot tables "ByRegion" and "ByProduct" share one pivot
// cache (pivotTable1 and pivotTable2 both use pivotCacheDefinition1).
const SHARED_CACHE_FILE = './spec/integration/data/pivot-shared-cache.xlsx';

function addPivot(workbook) {
  const worksheet = workbook.addWorksheet(`New${workbook.worksheets.length}`);
  worksheet.addPivotTable({
    sourceSheet: workbook.getWorksheet('Source'),
    rows: ['Region'],
    columns: ['Product'],
    values: ['Sales'],
    metric: 'sum',
  });
}

async function readZip(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  return {buffer, zip: await JSZip.loadAsync(buffer)};
}

const PIVOT_RELS = /^xl\/pivot(Tables|Cache)\/_rels\/.*\.rels$/;

// Every pivot relationship and pivot content type must point at a part
// that exists in the package.
async function expectPivotPartsResolve(zip) {
  const relsNames = Object.keys(zip.files).filter(name => PIVOT_RELS.test(name));
  const relsXmls = await Promise.all(relsNames.map(name => zip.file(name).async('string')));
  relsNames.forEach((relsName, index) => {
    const baseDir = path.posix.dirname(path.posix.dirname(relsName));
    for (const [, target] of relsXmls[index].matchAll(/Target="([^"]+)"/g)) {
      const partName = path.posix.normalize(`${baseDir}/${target}`);
      expect(zip.file(partName), `${relsName} -> ${partName}`).to.not.be.null();
    }
  });

  const contentTypes = await zip.file('[Content_Types].xml').async('string');
  for (const [, partName] of contentTypes.matchAll(/PartName="\/(xl\/pivot[^"]+)"/g)) {
    expect(zip.file(partName), partName).to.not.be.null();
  }
}

async function getCacheIds(zip) {
  const workbookXml = await zip.file('xl/workbook.xml').async('string');
  return [...workbookXml.matchAll(/<pivotCache cacheId="(\d+)"/g)].map(match => match[1]);
}

describe('Workbook', () => {
  describe('Pivot Tables in a loaded workbook', () => {
    it('keeps pivot tables that share a cache pointing at that cache', async () => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(SHARED_CACHE_FILE);

      const {buffer, zip} = await readZip(workbook);
      await expectPivotPartsResolve(zip);
      const rels = await zip.file('xl/pivotTables/_rels/pivotTable2.xml.rels').async('string');
      expect(rels).to.include('Target="../pivotCache/pivotCacheDefinition1.xml"');
      expect(zip.file('xl/pivotCache/pivotCacheDefinition2.xml')).to.be.null();

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      expect(workbook2.getWorksheet('Source').getCell('C5').value).to.equal(40);
    });

    it('numbers new pivot tables after the loaded ones', async () => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(SHARED_CACHE_FILE);
      addPivot(workbook);
      addPivot(workbook);

      const {zip} = await readZip(workbook);
      await expectPivotPartsResolve(zip);
      for (const partName of [
        'xl/pivotTables/pivotTable3.xml',
        'xl/pivotTables/pivotTable4.xml',
        'xl/pivotCache/pivotCacheDefinition2.xml',
        'xl/pivotCache/pivotCacheDefinition3.xml',
      ]) {
        expect(zip.file(partName), partName).to.not.be.null();
      }
      const cacheIds = await getCacheIds(zip);
      expect(cacheIds).to.have.length(3);
      expect(new Set(cacheIds).size).to.equal(3);
    });

    it('does not reuse the cache id of a loaded ExcelJS pivot table', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Source');
      worksheet.addRows([
        ['Region', 'Product', 'Sales'],
        ['North', 'Apple', 10],
        ['South', 'Pear', 20],
      ]);
      addPivot(workbook);

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(await workbook.xlsx.writeBuffer());
      addPivot(workbook2);

      const {zip} = await readZip(workbook2);
      await expectPivotPartsResolve(zip);
      const cacheIds = await getCacheIds(zip);
      expect(cacheIds).to.have.length(2);
      expect(new Set(cacheIds).size).to.equal(2);
      const newPivotXml = await zip.file('xl/pivotTables/pivotTable2.xml').async('string');
      const newCacheId = newPivotXml.match(/cacheId="(\d+)"/)[1];
      expect(cacheIds).to.include(newCacheId);
    });
  });
});
