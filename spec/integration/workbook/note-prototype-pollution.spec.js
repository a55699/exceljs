const ExcelJS = verquire('exceljs');

describe('Workbook', () => {
  describe('Notes with untrusted settings', () => {
    afterEach(() => {
      delete Object.prototype.polluted; // eslint-disable-line no-extend-native
    });

    it('does not pollute Object.prototype when writing a JSON-derived note', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');
      const json = '{"texts": [{"text": "hello"}], "__proto__": {"polluted": "yes"}}';
      worksheet.getCell('A1').value = 'value';
      worksheet.getCell('A1').note = JSON.parse(json);

      const buffer = await workbook.xlsx.writeBuffer();
      expect({}.polluted).to.be.undefined();

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      expect(workbook2.getWorksheet('Sheet1').getCell('A1').note).to.equal('hello');
    });
  });
});
