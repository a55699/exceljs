const ExcelJS = verquire('exceljs');

// B2:B4 is the shared formula A2*2 filled down, C2 an array formula
function makeWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');
  worksheet.getCell('A1').value = 'h';
  [1, 2, 3].forEach((value, i) => {
    worksheet.getCell(`A${i + 2}`).value = value;
  });
  worksheet.fillFormula('B2:B4', 'A2*2', [2, 4, 6]);
  worksheet.getCell('C2').value = {
    formula: 'SUM(A2:A4)',
    shareType: 'array',
    ref: 'C2',
    result: 6,
  };
  return {workbook, worksheet};
}

async function roundTrip(workbook) {
  const workbook2 = new ExcelJS.Workbook();
  await workbook2.xlsx.load(await workbook.xlsx.writeBuffer());
  return workbook2.getWorksheet('Sheet1');
}

function formulas(worksheet, addresses) {
  return addresses.map(address => worksheet.getCell(address).formula);
}

describe('Worksheet', () => {
  describe('splicing rows and columns with shared and array formulas', () => {
    it('keeps a shared formula when a row is inserted above it', async () => {
      const {workbook, worksheet} = makeWorkbook();
      worksheet.insertRow(1, ['new']);

      const worksheet2 = await roundTrip(workbook);
      // the formulas move with their cells, as other formulas of ExcelJS do
      expect(formulas(worksheet2, ['B3', 'B4', 'B5'])).to.deep.equal(['A2*2', 'A3*2', 'A4*2']);
      expect(worksheet2.getCell('C3').value).to.deep.equal({
        formula: 'SUM(A2:A4)',
        shareType: 'array',
        ref: 'C3',
        result: 6,
      });
    });

    it('keeps a shared formula when its first row is deleted', async () => {
      const {workbook, worksheet} = makeWorkbook();
      worksheet.spliceRows(2, 1);

      const worksheet2 = await roundTrip(workbook);
      expect(formulas(worksheet2, ['B2', 'B3'])).to.deep.equal(['A3*2', 'A4*2']);
      expect(worksheet2.getCell('B4').value).to.be.null();
    });

    it('keeps a shared formula and moves an array formula with columns', async () => {
      const {workbook, worksheet} = makeWorkbook();
      worksheet.spliceColumns(1, 0, []);

      const worksheet2 = await roundTrip(workbook);
      expect(formulas(worksheet2, ['C2', 'C3', 'C4'])).to.deep.equal(['A2*2', 'A3*2', 'A4*2']);
      expect(worksheet2.getCell('D2').value.ref).to.equal('D2');
    });

    it('keeps shared formulas above the splice', async () => {
      const {workbook, worksheet} = makeWorkbook();
      worksheet.spliceRows(6, 0, ['below']);

      const worksheet2 = await roundTrip(workbook);
      expect(worksheet2.getCell('B2').value.shareType).to.equal('shared');
      expect(worksheet2.getCell('B3').value.sharedFormula).to.equal('B2');
    });
  });
});
