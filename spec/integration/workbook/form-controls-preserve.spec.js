const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// Saved by Excel. Controls: a drop-down, a button, a spinner, a group box with
// two option buttons, a checkbox and a note. Buttons: one button only.
const CONTROLS_FILE = './spec/integration/data/form-controls.xlsx';

async function loadFile() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(CONTROLS_FILE);
  return workbook;
}

async function writeZip(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  return {buffer, zip: await JSZip.loadAsync(buffer)};
}

function objectTypes(vml) {
  return [...vml.matchAll(/<x:ClientData ObjectType="([^"]+)"/g)].map(match => match[1]);
}

function shapeIds(vml) {
  return [...vml.matchAll(/<v:shape id="([^"]+)"/g)].map(match => match[1]);
}

// the <v:shape> element whose ClientData has this object type
function shapeOf(vml, objectType) {
  const shapes = vml.match(/<v:shape [\s\S]*?<\/v:shape>/g);
  return shapes.find(shape => shape.includes(`ObjectType="${objectType}"`));
}

describe('Workbook', () => {
  describe('Form controls other than checkboxes', () => {
    it('writes them back after a workbook is loaded', async () => {
      const workbook = await loadFile();

      const {zip} = await writeZip(workbook);
      const vml = await zip.file('xl/drawings/vmlDrawing1.vml').async('string');
      expect(objectTypes(vml)).to.deep.equal([
        'Note',
        'Checkbox',
        'Drop',
        'Button',
        'Spin',
        'GBox',
        'Radio',
        'Radio',
      ]);
      expect(new Set(shapeIds(vml)).size).to.equal(8);

      const dropDown = shapeOf(vml, 'Drop');
      expect(dropDown).to.include('type="#_x0000_t201"');
      expect(dropDown).to.include('<x:FmlaRange>$A$1:$A$3</x:FmlaRange>');
      expect(dropDown).to.include('<x:FmlaLink>$E$2</x:FmlaLink>');
      expect(dropDown).to.include('<x:Sel>2</x:Sel>');
      expect(dropDown).to.include('<x:NoThreeD2/>');
      expect(shapeOf(vml, 'Button')).to.include('Run &amp; &lt;go&gt;');
      expect(shapeOf(vml, 'Spin')).to.include('<x:Max>10</x:Max>');
    });

    it('writes them on a worksheet without notes or checkboxes', async () => {
      const workbook = await loadFile();

      const {zip} = await writeZip(workbook);
      const sheet = await zip.file('xl/worksheets/sheet2.xml').async('string');
      expect(sheet).to.match(/<legacyDrawing r:id="rId\d+"\/>/);
      const rels = await zip.file('xl/worksheets/_rels/sheet2.xml.rels').async('string');
      expect(rels).to.include('Target="../drawings/vmlDrawing2.vml"');
      const vml = await zip.file('xl/drawings/vmlDrawing2.vml').async('string');
      expect(objectTypes(vml)).to.deep.equal(['Button']);
      expect(shapeIds(vml)).to.deep.equal(['_x0000_s1025']);
      expect(vml).to.include('<v:shapetype id="_x0000_t201"');
      expect(vml).to.include('Only a button');
    });

    it('keeps them through another load and with new checkboxes', async () => {
      const workbook = await loadFile();
      workbook.getWorksheet('Controls').addFormCheckbox('G2', {text: 'New'});

      const {buffer} = await writeZip(workbook);
      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      const worksheet2 = workbook2.getWorksheet('Controls');
      expect(worksheet2.getFormCheckboxes().map(checkbox => checkbox.text)).to.deep.equal([
        'Checked',
        'New',
      ]);

      const {zip} = await writeZip(workbook2);
      const vml = await zip.file('xl/drawings/vmlDrawing1.vml').async('string');
      const kept = objectTypes(vml).filter(type => type !== 'Checkbox' && type !== 'Note');
      expect(kept).to.deep.equal(['Drop', 'Button', 'Spin', 'GBox', 'Radio', 'Radio']);
      // the new checkbox moves the kept shapes along
      expect(shapeIds(vml)).to.deep.equal([
        '_x0000_s1025',
        '_x0000_s1026',
        '_x0000_s1027',
        '_x0000_s1028',
        '_x0000_s1029',
        '_x0000_s1030',
        '_x0000_s1031',
        '_x0000_s1032',
        '_x0000_s1033',
      ]);
      expect(worksheet2.getCell('A18').note.texts[0].text).to.equal('a note');
    });
  });
});
