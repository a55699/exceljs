const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// Made with XlsxWriter. Sheets: Chart (a chartsheet, not kept by ExcelJS), A
// and B. Names: Rate on A (A!$A$1) and Rate on B (B!$A$1), the formula Total,
// the constant VAT, the range Plain, the print area of B and the hidden
// filter range of B.
const NAMES_FILE = './spec/integration/data/defined-names-kept.xlsx';

async function writtenNames(workbook) {
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  const xml = await zip.file('xl/workbook.xml').async('string');
  return [...xml.matchAll(/<definedName [^>]*>[^<]*<\/definedName>/g)]
    .map(match => match[0])
    .sort();
}

describe('Defined names of a loaded file', () => {
  it('keeps formulas, constants and the names of one worksheet', async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(NAMES_FILE);

    // the chartsheet is dropped, so A and B are the first and second sheets
    expect(await writtenNames(workbook)).to.deep.equal([
      '<definedName name="Plain">A!$A$3:$B$4</definedName>',
      '<definedName name="Rate" localSheetId="0">A!$A$1</definedName>',
      '<definedName name="Rate" localSheetId="1">B!$A$1</definedName>',
      '<definedName name="Total">SUM(A!$A$4:$B$4)</definedName>',
      '<definedName name="VAT">0.05</definedName>',
      // the print area is written by ExcelJS from the page setup
      '<definedName name="_xlnm.Print_Area" localSheetId="1">&apos;B&apos;!$A1:$C10</definedName>',
      '<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">B!$A$3:$B$4</definedName>',
    ]);
    expect(workbook.getWorksheet('B').pageSetup.printArea).to.equal('A1:C10');
  });

  it('writes the name of a worksheet at its new position', async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(NAMES_FILE);
    workbook.removeWorksheet(workbook.getWorksheet('A').id);

    const names = await writtenNames(workbook);
    expect(names).to.include('<definedName name="Rate" localSheetId="0">B!$A$1</definedName>');
    expect(names.filter(name => name.includes('"Rate"'))).to.have.length(1);
  });
});
