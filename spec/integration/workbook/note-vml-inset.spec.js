const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const VML = 'xl/drawings/vmlDrawing1.vml';
const INSET = /inset="[^"]*"/g;

async function writeZip(workbook) {
  return JSZip.loadAsync(await workbook.xlsx.writeBuffer());
}

async function load(zip) {
  const buffer = await zip.generateAsync({type: 'nodebuffer'});
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

describe('Workbook', () => {
  describe('Note margins', () => {
    it('writes the default margins for a note loaded without an inset', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Notes');
      worksheet.getCell('A1').value = 'Noted';
      // Excel writes the author in bold, so its notes load as rich text
      worksheet.getCell('A1').note = {texts: [{text: 'a note', font: {bold: true}}]};

      // and <v:textbox> without an inset for notes with default margins
      const zip = await writeZip(workbook);
      const vml = await zip.file(VML).async('string');
      zip.file(VML, vml.replace(INSET, ''));

      const workbook2 = await load(zip);
      expect(workbook2.getWorksheet('Notes').getCell('A1').note.texts[0].text).to.equal('a note');
      const vml2 = await (await writeZip(workbook2)).file(VML).async('string');
      expect(vml2.match(INSET)).to.deep.equal(['inset="1.3mm,1.3mm,2.5mm,2.5mm"']);
    });

    it('keeps the margins of a note', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Notes');
      worksheet.getCell('A1').value = 'Noted';
      worksheet.getCell('A1').note = {
        texts: [{text: 'wide margins', font: {bold: true}}],
        margins: {insetmode: 'custom', inset: [0.25, 0.25, 0.35, 0.35]},
      };

      const workbook2 = await load(await writeZip(workbook));
      const {margins} = workbook2.getWorksheet('Notes').getCell('A1').note;
      expect(margins.inset).to.deep.equal([0.25, 0.25, 0.35, 0.35]);
    });
  });
});
