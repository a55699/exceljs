const ExcelJS = verquire('exceljs');

async function roundTrip(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  const workbook2 = new ExcelJS.Workbook();
  await workbook2.xlsx.load(buffer);
  return workbook2;
}

describe('Workbook', () => {
  describe('Notes on cells without a value', () => {
    it('loads a note on an empty cell', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Notes');
      worksheet.getCell('C3').value = 1;
      worksheet.getCell('C3').note = 'with a value';
      // row 3 has a <c> for C3 only, row 9 has no <row> at all
      worksheet.getCell('B3').note = 'next to a value';
      worksheet.getCell('E9').note = {texts: [{text: 'alone', font: {bold: true}}]};

      const worksheet2 = (await roundTrip(workbook)).getWorksheet('Notes');
      expect(worksheet2.getCell('C3').note).to.equal('with a value');
      expect(worksheet2.getCell('B3').note).to.equal('next to a value');
      expect(worksheet2.getCell('B3').value).to.equal(null);
      expect(worksheet2.getCell('E9').note.texts).to.deep.equal([
        {text: 'alone', font: {bold: true}},
      ]);
      expect(worksheet2.getCell('E9').value).to.equal(null);
    });

    it('keeps the note when the workbook is written again', async () => {
      const workbook = new ExcelJS.Workbook();
      workbook.addWorksheet('Notes').getCell('B3').note = 'kept';

      const workbook3 = await roundTrip(await roundTrip(workbook));
      expect(workbook3.getWorksheet('Notes').getCell('B3').note).to.equal('kept');
    });
  });
});
