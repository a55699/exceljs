const path = require('path');
const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// Saved by Excel. Data: a chart (with a text box drawn over it, kept in its
// own drawing part), a picture and a rectangle with text. Second: a chart.
const MIXED_FILE = './spec/integration/data/drawings-mixed.xlsx';
const IMAGE_FILE = path.join(__dirname, '../data/image.png');

async function loadFile() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(MIXED_FILE);
  return workbook;
}

async function writeZip(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  return {buffer, zip: await JSZip.loadAsync(buffer)};
}

function readFile(zip, name) {
  return zip.file(name).async('string');
}

function relsOf(xml) {
  return [...xml.matchAll(/<Relationship ([^>]*)\/>/g)].map(match => {
    const attr = name => (match[1].match(new RegExp(`${name}="([^"]*)"`)) || [])[1];
    return {Id: attr('Id'), Target: attr('Target'), TargetMode: attr('TargetMode')};
  });
}

// every internal relationship target and every content type override is a
// part of the package, and no part has two overrides
async function checkPackage(zip) {
  const names = Object.keys(zip.files);
  const relsNames = names.filter(name => name.endsWith('.rels'));
  const relsXml = await Promise.all(relsNames.map(relsName => readFile(zip, relsName)));
  relsNames.forEach((relsName, index) => {
    const baseDir = path.posix.dirname(path.posix.dirname(relsName));
    relsOf(relsXml[index])
      .filter(rel => rel.TargetMode !== 'External')
      .forEach(rel => {
        const target = rel.Target.startsWith('/')
          ? rel.Target.slice(1)
          : path.posix.normalize(`${baseDir}/${rel.Target}`);
        expect(names, `${relsName} ${rel.Id}`).to.include(target);
      });
  });
  const contentTypes = await readFile(zip, '[Content_Types].xml');
  const parts = [...contentTypes.matchAll(/PartName="\/([^"]+)"/g)].map(match => match[1]);
  expect(new Set(parts).size, 'duplicate overrides').to.equal(parts.length);
  parts.forEach(part => expect(names, 'override').to.include(part));
}

// the drawing used by a worksheet, with its relationships
async function sheetDrawing(zip, sheetFile) {
  const sheetRels = relsOf(await readFile(zip, `xl/worksheets/_rels/${sheetFile}.rels`));
  const drawingRel = sheetRels.find(r => r.Target.includes('/drawings/drawing'));
  const name = drawingRel.Target.match(/drawings\/(drawing\d+)\.xml/)[1];
  const xml = await readFile(zip, `xl/drawings/${name}.xml`);
  const rels = relsOf(await readFile(zip, `xl/drawings/_rels/${name}.xml.rels`));
  const ids = rels.map(rel => rel.Id);
  const used = [...xml.matchAll(/ r:(?:id|embed)="([^"]+)"/g)].map(match => match[1]);
  used.forEach(id => expect(ids, `${name} ${id}`).to.include(id));
  return {name, xml, rels};
}

const targetsOf = rels => rels.map(rel => rel.Target).sort();

describe('Workbook', () => {
  describe('Drawings with charts and shapes', () => {
    it('writes the charts, shapes and pictures of a loaded workbook back', async () => {
      const workbook = await loadFile();
      expect(workbook.getWorksheet('Data').getImages()).to.have.length(1);

      const {zip} = await writeZip(workbook);
      await checkPackage(zip);

      const data = await sheetDrawing(zip, 'sheet1.xml');
      expect(data.xml).to.include('<c:chart ');
      expect(data.xml).to.include('Note &amp; &lt;box&gt;');
      expect(data.xml).to.include('<xdr:pic>');
      expect(targetsOf(data.rels)).to.deep.equal(['../charts/chart1.xml', '../media/image1.png']);
      const second = await sheetDrawing(zip, 'sheet2.xml');
      expect(targetsOf(second.rels)).to.deep.equal(['../charts/chart2.xml']);
      expect(second.name).not.to.equal(data.name);

      // the text box drawn over the first chart keeps its part and content type
      const chartRels = relsOf(await readFile(zip, 'xl/charts/_rels/chart1.xml.rels'));
      expect(chartRels.map(rel => rel.Target)).to.deep.equal(['../drawings/drawing2.xml']);
      expect(await readFile(zip, 'xl/drawings/drawing2.xml')).to.include('<c:userShapes');
      expect(await readFile(zip, '[Content_Types].xml')).to.include(
        '<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"/>'
      );
    });

    it('adds an image to a worksheet that has a chart', async () => {
      const workbook = await loadFile();
      const imageId = workbook.addImage({filename: IMAGE_FILE, extension: 'png'});
      workbook.getWorksheet('Data').addImage(imageId, 'H20:J26');

      const {buffer, zip} = await writeZip(workbook);
      await checkPackage(zip);
      const data = await sheetDrawing(zip, 'sheet1.xml');
      expect(data.xml).to.include('<c:chart ');
      expect(data.xml.match(/<xdr:pic>/g)).to.have.length(2);
      expect(targetsOf(data.rels)).to.deep.equal([
        '../charts/chart1.xml',
        '../media/image1.png',
        '../media/image2.png',
      ]);

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      expect(workbook2.getWorksheet('Data').getImages()).to.have.length(2);
      const {zip: zip2} = await writeZip(workbook2);
      await checkPackage(zip2);
      expect((await sheetDrawing(zip2, 'sheet1.xml')).xml).to.include('<c:chart ');
    });

    it('adds an image to another worksheet without replacing a chart drawing', async () => {
      const workbook = await loadFile();
      const imageId = workbook.addImage({filename: IMAGE_FILE, extension: 'png'});
      const imageId2 = workbook.addImage({filename: IMAGE_FILE, extension: 'png'});
      workbook.getWorksheet('Second').addImage(imageId, 'H20:J26');
      workbook.addWorksheet('New').addImage(imageId2, 'B2:D6');

      const {zip} = await writeZip(workbook);
      await checkPackage(zip);
      const drawings = [
        await sheetDrawing(zip, 'sheet1.xml'),
        await sheetDrawing(zip, 'sheet2.xml'),
        await sheetDrawing(zip, 'sheet3.xml'),
      ];
      expect(new Set(drawings.map(drawing => drawing.name)).size).to.equal(3);
      expect(drawings.map(drawing => drawing.name)).not.to.include('drawing2');
      expect(targetsOf(drawings[1].rels)).to.deep.equal([
        '../charts/chart2.xml',
        '../media/image2.png',
      ]);
      expect(targetsOf(drawings[2].rels)).to.deep.equal(['../media/image3.png']);
    });

    it('reads drawing relationships stored before the drawing', async () => {
      // the order of the parts in the package is free; put every rels part first
      const source = await JSZip.loadAsync(await require('fs').promises.readFile(MIXED_FILE));
      const names = Object.keys(source.files).filter(name => !source.files[name].dir);
      const relsFirst = new JSZip();
      const ordered = [
        ...names.filter(name => name.endsWith('.rels')),
        ...names.filter(name => !name.endsWith('.rels')),
      ];
      const contents = await Promise.all(
        ordered.map(name => source.file(name).async('nodebuffer'))
      );
      ordered.forEach((name, index) => relsFirst.file(name, contents[index]));
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await relsFirst.generateAsync({type: 'nodebuffer'}));

      const {zip} = await writeZip(workbook);
      await checkPackage(zip);
      const data = await sheetDrawing(zip, 'sheet1.xml');
      expect(targetsOf(data.rels)).to.deep.equal(['../charts/chart1.xml', '../media/image1.png']);
    });
  });
});
