const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// Saved by Excel. Data: table Sales A1:E5 with a totals row, a calculated
// column Total ([@Qty]*[@[Unit Price]]) and a column Double filled with the
// shared formula B2*2; formulas in G refer to the table. Other Sheet: two
// formulas that refer to Sales, and table Codes C4:D5 without a header row and
// without filter buttons.
const TABLES_FILE = './spec/integration/data/tables-excel.xlsx';

async function loadFile() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(TABLES_FILE);
  return workbook;
}

async function writeAndLoad(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const tableNames = Object.keys(zip.files).filter(name =>
    /^xl\/tables\/table\d+\.xml$/.test(name)
  );
  const tablesXml = {};
  const xmls = await Promise.all(tableNames.map(name => zip.file(name).async('string')));
  xmls.forEach(xml => {
    tablesXml[xml.match(/ name="([^"]*)"/)[1]] = xml;
  });
  const reloaded = new ExcelJS.Workbook();
  await reloaded.xlsx.load(buffer);
  return {tablesXml, workbook: reloaded};
}

function formulaOf(worksheet, address) {
  return worksheet.getCell(address).value.formula;
}

describe('Tables of a loaded file', () => {
  it('reads the rows of the tables', async () => {
    const workbook = await loadFile();
    const sales = workbook.getWorksheet('Data').getTable('Sales');
    expect(sales.ref).to.equal('A1');
    expect(sales.headerRow).to.equal(true);
    expect(sales.totalsRow).to.equal(true);
    expect(sales.model.rows.map(row => row[0])).to.deep.equal(['Jan', 'Feb', 'Mar']);
    // the shared formula of each row is kept as the formula of its cell
    expect(sales.model.rows.map(row => row[4].formula)).to.deep.equal(['B2*2', 'B3*2', 'B4*2']);

    const codes = workbook.getWorksheet('Other Sheet').getTable('Codes');
    expect(codes.headerRow).to.equal(false);
    expect(codes.model.rows).to.deep.equal([
      ['x', 10],
      ['y', 20],
    ]);
  });

  it('writes the tables as they were read', async () => {
    const {tablesXml, workbook} = await writeAndLoad(await loadFile());

    const sales = tablesXml.Sales;
    expect(sales).to.include('ref="A1:E5"');
    expect(sales).to.include('totalsRowCount="1"');
    expect(sales).to.include('<autoFilter ref="A1:E4">');
    expect(sales).to.include(
      '<calculatedColumnFormula>Sales[[#This Row],[Qty]]*Sales[[#This Row],[Unit Price]]</calculatedColumnFormula>'
    );
    expect(sales).to.include('name="Qty" totalsRowFunction="sum"');

    const codes = tablesXml.Codes;
    expect(codes).to.include('ref="C4:D5"');
    expect(codes).to.include('headerRowCount="0"');
    expect(codes).not.to.include('<autoFilter');

    const data = workbook.getWorksheet('Data');
    expect(data.getCell('A1').value).to.equal('Month');
    expect(data.getCell('A5').value).to.equal('Total');
    expect(formulaOf(data, 'B5')).to.equal('SUBTOTAL(109,Sales[Qty])');
  });

  it('adds a row to a loaded table', async () => {
    const loaded = await loadFile();
    loaded.getWorksheet('Data').getTable('Sales').addRow(['Apr', 4, 8]);
    const {tablesXml, workbook} = await writeAndLoad(loaded);

    expect(tablesXml.Sales).to.include('ref="A1:E6"');
    expect(tablesXml.Sales).to.include('<autoFilter ref="A1:E5">');
    const data = workbook.getWorksheet('Data');
    expect(data.getCell('A5').value).to.equal('Apr');
    // the calculated column and the column with a filled formula get their formula
    expect(formulaOf(data, 'D5')).to.equal(
      'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Unit Price]]'
    );
    expect(formulaOf(data, 'E5')).to.equal('B5*2');
    expect(data.getCell('A6').value).to.equal('Total');
    expect(formulaOf(data, 'B6')).to.equal('SUBTOTAL(109,Sales[Qty])');
  });

  it('removes a row of a loaded table', async () => {
    const loaded = await loadFile();
    loaded.getWorksheet('Data').getTable('Sales').removeRows(0);
    const {tablesXml, workbook} = await writeAndLoad(loaded);

    expect(tablesXml.Sales).to.include('ref="A1:E4"');
    const data = workbook.getWorksheet('Data');
    expect(data.getCell('A2').value).to.equal('Feb');
    expect(formulaOf(data, 'E2')).to.equal('B2*2');
    expect(formulaOf(data, 'E3')).to.equal('B3*2');
    expect(data.getCell('A4').value).to.equal('Total');
    expect(data.getCell('A5').value).to.be.null();
  });

  it('renames a loaded table and the references to it', async () => {
    const loaded = await loadFile();
    const sales = loaded.getWorksheet('Data').getTable('Sales');
    sales.name = 'Orders';
    sales.theme = 'TableStyleLight9';
    const {tablesXml, workbook} = await writeAndLoad(loaded);

    expect(tablesXml.Orders).to.include('displayName="Orders"');
    expect(tablesXml.Orders).to.include('<tableStyleInfo name="TableStyleLight9"');
    expect(tablesXml.Orders).to.include(
      'Orders[[#This Row],[Qty]]*Orders[[#This Row],[Unit Price]]'
    );
    expect(formulaOf(workbook.getWorksheet('Data'), 'G7')).to.equal('SUM(Orders[Qty])');
    expect(formulaOf(workbook.getWorksheet('Other Sheet'), 'A2')).to.equal(
      'Orders[[#Totals],[Qty]]'
    );
  });

  it('removes a loaded table like Convert to Range', async () => {
    const loaded = await loadFile();
    loaded.getWorksheet('Data').removeTable('Sales');
    const {tablesXml, workbook} = await writeAndLoad(loaded);

    expect(Object.keys(tablesXml)).to.deep.equal(['Codes']);
    // what Excel's Convert to Range gives for these formulas
    const data = workbook.getWorksheet('Data');
    expect(formulaOf(data, 'D2')).to.equal('Data!$B2*Data!$C2');
    expect(formulaOf(data, 'B5')).to.equal('SUBTOTAL(109,Data!$B$2:$B$4)');
    expect(formulaOf(data, 'G2')).to.equal('Data!$B2+1');
    expect(formulaOf(data, 'G9')).to.equal('COUNTA(Data!$A$1:$E$5)');
    expect(formulaOf(data, 'G10')).to.equal('Data!$A$1');
    expect(formulaOf(data, 'G12')).to.equal('ROWS(Data!$A$2:$E$4)');
    expect(formulaOf(workbook.getWorksheet('Other Sheet'), 'A2')).to.equal('Data!$B$5');
  });
});
