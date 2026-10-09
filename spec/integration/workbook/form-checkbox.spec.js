const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

async function writeZip(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  return {buffer, zip: await JSZip.loadAsync(buffer)};
}

function readFile(zip, name) {
  return zip.file(name).async('string');
}

const CTRL_PROP_TARGET = /Target="\.\.\/ctrlProps\/([^"]+)"/g;

function ctrlPropTargets(rels) {
  return [...rels.matchAll(CTRL_PROP_TARGET)].map(match => match[1]);
}

// col and row of the top-left and bottom-right corners (0-based), without offsets
function anchorCells(data) {
  const anchor = data.match(/<x:Anchor>([^<]+)<\/x:Anchor>/)[1].split(', ');
  return anchor.filter((value, index) => index % 2 === 0);
}

function clientData(vml) {
  return vml.match(/<x:ClientData ObjectType="Checkbox">[\s\S]*?<\/x:ClientData>/g) || [];
}

describe('Workbook', () => {
  describe('Form checkboxes', () => {
    it('writes control properties and VML for each checkbox', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Form');
      worksheet.addFormCheckbox('B2:D3', {text: 'Accept & <go>', link: 'A2', checked: true});
      worksheet.addFormCheckbox('B5');

      const {zip} = await writeZip(workbook);
      const ctrlProp1 = await readFile(zip, 'xl/ctrlProps/ctrlProp1.xml');
      expect(ctrlProp1).to.include('objectType="CheckBox"');
      expect(ctrlProp1).to.include('checked="1"');
      expect(ctrlProp1).to.include('fmlaLink="$A$2"');
      const ctrlProp2 = await readFile(zip, 'xl/ctrlProps/ctrlProp2.xml');
      expect(ctrlProp2).to.include('checked="0"');
      expect(ctrlProp2).to.not.include('fmlaLink');

      const vml = await readFile(zip, 'xl/drawings/vmlDrawing1.vml');
      const [first, second] = clientData(vml);
      expect(first).to.include('<x:Anchor>1, 15, 1, 3, 3, 29, 2, 20</x:Anchor>');
      expect(first).to.include('<x:FmlaLink>$A$2</x:FmlaLink>');
      expect(first).to.include('<x:Checked>1</x:Checked>');
      expect(second).to.include('<x:Checked>0</x:Checked>');
      expect(second).to.not.include('<x:FmlaLink>');
      expect(vml).to.include('Accept &amp; &lt;go&gt;');
    });

    it('numbers control properties across worksheets and links them', async () => {
      const workbook = new ExcelJS.Workbook();
      workbook.addWorksheet('One').addFormCheckbox('B2');
      workbook.addWorksheet('One more').addFormCheckbox('B2');
      workbook.getWorksheet('One').addFormCheckbox('B4');

      const {zip} = await writeZip(workbook);
      const rels1 = await readFile(zip, 'xl/worksheets/_rels/sheet1.xml.rels');
      const rels2 = await readFile(zip, 'xl/worksheets/_rels/sheet2.xml.rels');
      const allTargets = [...ctrlPropTargets(rels1), ...ctrlPropTargets(rels2)];
      expect(allTargets.sort()).to.deep.equal(['ctrlProp1.xml', 'ctrlProp2.xml', 'ctrlProp3.xml']);
      expect(rels1).to.include('Target="../drawings/vmlDrawing1.vml"');
      expect(rels2).to.include('Target="../drawings/vmlDrawing2.vml"');

      const contentTypes = await readFile(zip, '[Content_Types].xml');
      allTargets.forEach(target => {
        expect(contentTypes).to.include(`PartName="/xl/ctrlProps/${target}"`);
      });
    });

    it('accepts cell, range, start/end and tl/br positions', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Form');
      worksheet.addFormCheckbox('C3');
      worksheet.addFormCheckbox('C5:E6');
      worksheet.addFormCheckbox({startCol: 2, startRow: 7, endCol: 4, endRow: 8});
      worksheet.addFormCheckbox({tl: 'C10', br: 'E11'});

      const {zip} = await writeZip(workbook);
      const vml = await readFile(zip, 'xl/drawings/vmlDrawing1.vml');
      expect(clientData(vml).map(anchorCells)).to.deep.equal([
        ['2', '2', '4', '3'],
        ['2', '4', '4', '5'],
        ['2', '7', '4', '8'],
        ['2', '9', '4', '10'],
      ]);
    });

    it('updates checked, link and text after the checkbox is added', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Form');
      const checkbox = worksheet.addFormCheckbox('B2');
      checkbox.checked = true;
      checkbox.link = 'D4';
      checkbox.text = 'Later';
      expect(checkbox.checked).to.be.true();
      expect(checkbox.link).to.equal('$D$4');
      expect(worksheet.getFormCheckboxes()).to.deep.equal([checkbox]);

      const {zip} = await writeZip(workbook);
      const ctrlProp = await readFile(zip, 'xl/ctrlProps/ctrlProp1.xml');
      expect(ctrlProp).to.include('checked="1"');
      expect(ctrlProp).to.include('fmlaLink="$D$4"');
      expect(await readFile(zip, 'xl/drawings/vmlDrawing1.vml')).to.include('Later');
    });

    it('keeps cell values when a workbook with checkboxes is loaded', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Form');
      worksheet.getCell('A2').value = true;
      worksheet.addFormCheckbox('B2', {link: 'A2', checked: true});

      const {buffer} = await writeZip(workbook);
      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      expect(workbook2.getWorksheet('Form').getCell('A2').value).to.be.true();
    });
  });
});
