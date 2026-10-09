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

function shapeIds(vml) {
  return [...vml.matchAll(/<v:shape id="([^"]+)"/g)].map(match => match[1]);
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

    it('gives notes and checkboxes on a worksheet different shape ids', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Form');
      worksheet.getCell('A1').value = 'Noted';
      worksheet.getCell('A1').note = 'a note';
      worksheet.addFormCheckbox('B2');
      worksheet.addFormCheckbox('B4');

      const {zip} = await writeZip(workbook);
      const vml = await readFile(zip, 'xl/drawings/vmlDrawing1.vml');
      expect(shapeIds(vml)).to.deep.equal(['_x0000_s1025', '_x0000_s1026', '_x0000_s1027']);
    });

    // Excel shows the value of the linked cell, not the checked state saved
    // with the checkbox: an empty linked cell shows an unchecked box
    describe('linked cell', () => {
      it('gets the checked state when it is empty', async () => {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Form');
        worksheet.addFormCheckbox('B2', {link: 'A2', checked: true});
        worksheet.addFormCheckbox('B4', {link: 'A4'});
        expect(worksheet.getCell('A2').value).to.be.true();
        expect(worksheet.getCell('A4').value).to.be.false();

        const {buffer} = await writeZip(workbook);
        const workbook2 = new ExcelJS.Workbook();
        await workbook2.xlsx.load(buffer);
        const worksheet2 = workbook2.getWorksheet('Form');
        expect(worksheet2.getCell('A2').value).to.be.true();
        expect(worksheet2.getCell('A4').value).to.be.false();
      });

      it('keeps a value already in the cell', () => {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Form');
        worksheet.getCell('A2').value = false;
        worksheet.getCell('A4').value = 'yes';
        worksheet.addFormCheckbox('B2', {link: 'A2', checked: true});
        worksheet.addFormCheckbox('B4', {link: 'A4', checked: true});
        expect(worksheet.getCell('A2').value).to.be.false();
        expect(worksheet.getCell('A4').value).to.equal('yes');
      });

      it('follows the checked and link setters', () => {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Form');
        const checkbox = worksheet.addFormCheckbox('B2', {link: 'A2', checked: true});
        checkbox.checked = false;
        expect(worksheet.getCell('A2').value).to.be.false();

        checkbox.link = 'A3';
        expect(worksheet.getCell('A3').value).to.be.false();
        checkbox.checked = true;
        expect(worksheet.getCell('A3').value).to.be.true();
        expect(worksheet.getCell('A2').value).to.be.false();
      });

      it('can be on another worksheet', () => {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Form');
        const other = workbook.addWorksheet('My answers');
        // eslint-disable-next-line quotes
        worksheet.addFormCheckbox('B2', {link: "'My answers'!$C$3", checked: true});
        expect(other.getCell('C3').value).to.be.true();
        expect(worksheet.getCell('C3').value).to.equal(null);
      });

      it('is not changed when a workbook is loaded', async () => {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Form');
        worksheet.addFormCheckbox('B2', {link: 'A2', checked: true});
        worksheet.getCell('A2').value = null;

        const {buffer} = await writeZip(workbook);
        const workbook2 = new ExcelJS.Workbook();
        await workbook2.xlsx.load(buffer);
        const worksheet2 = workbook2.getWorksheet('Form');
        expect(worksheet2.getCell('A2').value).to.equal(null);
        expect(worksheet2.getFormCheckboxes()[0].checked).to.be.true();
      });
    });

    describe('loaded from a file', () => {
      // Saved by Excel: three checkboxes, a button and a note on Form, and
      // one checkbox on Second
      const FORM_FILE = './spec/integration/data/form-checkboxes.xlsx';
      const FORM_CHECKBOXES = [
        {text: 'Ship & <pay>', link: '$B$5', checked: false, cell: {col: 1, row: 1}},
        {text: 'Unlinked on', link: undefined, checked: true, cell: {col: 3, row: 1}},
        {text: 'Off', link: undefined, checked: false, cell: {col: 3, row: 3}},
      ];
      const SECOND_CHECKBOXES = [
        {text: 'Second sheet', link: undefined, checked: true, cell: {col: 2, row: 2}},
      ];

      function summary(worksheet) {
        return worksheet.getFormCheckboxes().map(checkbox => ({
          text: checkbox.text,
          link: checkbox.link,
          checked: checkbox.checked,
          cell: {col: checkbox.model.tl.col, row: checkbox.model.tl.row},
        }));
      }

      it('reads the checkboxes saved by Excel', async () => {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(FORM_FILE);

        expect(summary(workbook.getWorksheet('Form'))).to.deep.equal(FORM_CHECKBOXES);
        expect(summary(workbook.getWorksheet('Second'))).to.deep.equal(SECOND_CHECKBOXES);
        const [checkbox] = workbook.getWorksheet('Form').getFormCheckboxes();
        // VML anchors are in pixels
        expect(checkbox.model.br).to.deep.equal({
          col: 3,
          colOff: 5 * 9525,
          row: 2,
          rowOff: 4 * 9525,
        });
        expect(checkbox.model.noThreeD).to.be.true();
        expect(checkbox.model.print).to.be.true();
      });

      it('keeps the note apart from the checkboxes in the same drawing', async () => {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(FORM_FILE);

        const {note} = workbook.getWorksheet('Form').getCell('A8');
        expect(note.texts[0].text).to.equal('a note');
        // from the note's own shape, which Excel writes after the controls
        expect(note.anchor.trim()).to.equal('1, 15, 6, 10, 3, 31, 10, 9');
      });

      it('writes the loaded checkboxes back', async () => {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(FORM_FILE);

        const {buffer, zip} = await writeZip(workbook);
        const vml = await readFile(zip, 'xl/drawings/vmlDrawing1.vml');
        expect(clientData(vml)).to.have.length(3);
        expect(vml).not.to.include('ObjectType="Button"');
        expect(new Set(shapeIds(vml)).size).to.equal(4);
        const rels = await readFile(zip, 'xl/worksheets/_rels/sheet1.xml.rels');
        expect(ctrlPropTargets(rels)).to.deep.equal([
          'ctrlProp1.xml',
          'ctrlProp2.xml',
          'ctrlProp3.xml',
        ]);

        const workbook2 = new ExcelJS.Workbook();
        await workbook2.xlsx.load(buffer);
        expect(summary(workbook2.getWorksheet('Form'))).to.deep.equal(FORM_CHECKBOXES);
        expect(summary(workbook2.getWorksheet('Second'))).to.deep.equal(SECOND_CHECKBOXES);
        expect(workbook2.getWorksheet('Form').getCell('A8').note.texts[0].text).to.equal('a note');
      });

      it('reads back the checkboxes ExcelJS writes', async () => {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Form');
        worksheet.addFormCheckbox('B2:D3', {text: 'Accept', link: 'A2', checked: true});
        const position = {startCol: 5, startRow: 6, endCol: 7, endRow: 7};
        worksheet.addFormCheckbox(position, {text: 'Flat off', noThreeD: false, print: true});
        // the part numbers are given when writing
        const settings = checkbox => {
          const {ctrlPropId, ctrlPropRelId, ...model} = checkbox.model;
          return model;
        };
        const models = worksheet.getFormCheckboxes().map(settings);

        const {buffer} = await writeZip(workbook);
        const workbook2 = new ExcelJS.Workbook();
        await workbook2.xlsx.load(buffer);
        const loaded = workbook2.getWorksheet('Form').getFormCheckboxes();
        expect(loaded.map(settings)).to.deep.equal(models);

        loaded[0].checked = false;
        loaded[1].text = 'Changed';
        const {buffer: buffer2} = await writeZip(workbook2);
        const workbook3 = new ExcelJS.Workbook();
        await workbook3.xlsx.load(buffer2);
        const reloaded = workbook3.getWorksheet('Form').getFormCheckboxes();
        expect(reloaded.map(checkbox => checkbox.checked)).to.deep.equal([false, false]);
        expect(reloaded.map(checkbox => checkbox.text)).to.deep.equal(['Accept', 'Changed']);
      });
    });
  });
});
