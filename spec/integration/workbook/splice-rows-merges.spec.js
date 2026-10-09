const ExcelJS = verquire('exceljs');

const LIST = {type: 'list', allowBlank: true, formulae: ['"a,b"']};

function makeWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');
  for (let row = 1; row <= 10; row++) {
    worksheet.getRow(row).values = [`r${row}`, 'x'];
  }
  worksheet.mergeCells('A5:B6');
  worksheet.getCell('C5').dataValidation = LIST;
  return {workbook, worksheet};
}

async function roundTrip(workbook) {
  const workbook2 = new ExcelJS.Workbook();
  await workbook2.xlsx.load(await workbook.xlsx.writeBuffer());
  return workbook2.getWorksheet('Sheet1');
}

describe('Worksheet', () => {
  describe('spliceRows with merged cells and data validations', () => {
    it('moves them down when rows are inserted above', async () => {
      const {workbook, worksheet} = makeWorkbook();
      worksheet.spliceRows(3, 0, ['new1'], ['new2']);

      const worksheet2 = await roundTrip(workbook);
      expect(worksheet2.model.merges).to.deep.equal(['A7:B8']);
      expect(worksheet2.getCell('A7').value).to.equal('r5');
      expect(worksheet2.getCell('B8').isMerged).to.be.true();
      expect(worksheet2.getCell('A5').isMerged).to.be.false();
      expect(worksheet2.getCell('C7').dataValidation).to.deep.equal(LIST);
      expect(worksheet2.getCell('C5').dataValidation).to.be.undefined();
    });

    it('moves them up when rows are removed above', async () => {
      const {workbook, worksheet} = makeWorkbook();
      worksheet.spliceRows(2, 2);

      const worksheet2 = await roundTrip(workbook);
      expect(worksheet2.model.merges).to.deep.equal(['A3:B4']);
      expect(worksheet2.getCell('A3').value).to.equal('r5');
      expect(worksheet2.getCell('C3').dataValidation).to.deep.equal(LIST);
    });

    it('grows a merge when rows are inserted inside it', async () => {
      const {workbook, worksheet} = makeWorkbook();
      worksheet.spliceRows(6, 0, ['new']);

      const worksheet2 = await roundTrip(workbook);
      expect(worksheet2.model.merges).to.deep.equal(['A5:B7']);
    });

    it('shrinks or drops merges and validations whose rows are removed', async () => {
      const {workbook, worksheet} = makeWorkbook();
      worksheet.mergeCells('A8:B9');
      worksheet.spliceRows(5, 1);
      worksheet.spliceRows(7, 2);

      const worksheet2 = await roundTrip(workbook);
      expect(worksheet2.model.merges).to.deep.equal(['A5:B5']);
      expect(worksheet2.getCell('C5').dataValidation).to.be.undefined();
    });
  });
});
