const colCache = verquire('utils/col-cache');
const Excel = verquire('exceljs');

const spliceArray = (a, index, count, ...rest) => {
  const clone = [...a];
  clone.splice(index, count, ...rest);
  return clone;
};

const values = [
  ['Date', 'Id', 'Word'],
  [new Date('2019-08-01'), 1, 'Bird'],
  [new Date('2019-08-02'), 2, 'is'],
  [new Date('2019-08-03'), 3, 'the'],
  [new Date('2019-08-04'), 4, 'Word'],
  ['Totals', {formula: 'SUBTOTAL(104,TestTable[Id])', result: 4}, null],
];

function addTable(ref, ws) {
  return ws.addTable({
    name: 'TestTable',
    ref,
    headerRow: true,
    totalsRow: true,
    style: {
      theme: 'TableStyleDark3',
      showRowStripes: true,
    },
    columns: [
      {name: 'Date', totalsRowLabel: 'Totals', filterButton: true},
      {
        name: 'Id',
        totalsRowFunction: 'max',
        filterButton: true,
        totalsRowResult: 4,
      },
      {
        name: 'Word',
        filterButton: false,
        style: {font: {bold: true, name: 'Comic Sans MS'}},
      },
    ],
    rows: [
      [new Date('2019-08-01'), 1, 'Bird'],
      [new Date('2019-08-02'), 2, 'is'],
      [new Date('2019-08-03'), 3, 'the'],
      [new Date('2019-08-04'), 4, 'Word'],
    ],
  });
}

function checkTable(ref, ws, testValues) {
  const a = colCache.decodeAddress(ref);

  for (let i = -1; i <= testValues.length + 1; i++) {
    const vRow = testValues[i];
    const nRow = i + a.row;
    const row = nRow >= 1 && ws.getRow(nRow);
    if (!row) continue;
    for (let j = -1; j <= testValues[0].length + 1; j++) {
      const value = (vRow && vRow[j]) || null;
      const nCol = j + a.col;
      const cellValue = nCol >= 1 && row.getCell(nCol).value;
      if (!cellValue) continue;

      if (value instanceof Date) {
        expect(cellValue).to.equalDate(value);
      } else if (value === null) {
        expect(cellValue).to.be.null();
      } else if (typeof value === 'object') {
        expect(cellValue).to.deep.equal(value);
      } else {
        expect(cellValue).to.equal(value);
      }
    }
  }
}

describe('Worksheet', () => {
  describe('Table', () => {
    it('creates a table', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      addTable('A1', ws);

      checkTable('A1', ws, values);
    });

    it('removes header', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.headerRow = false;
      table.commit();

      const newValues = spliceArray(values, 0, 1);
      checkTable('A1', ws, newValues);
    });

    it('removes totals', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.totalsRow = false;
      table.commit();

      const newValues = spliceArray(values, 5, 1);
      checkTable('A1', ws, newValues);
    });

    it('moves the table', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.ref = 'C2';
      table.commit();

      checkTable('C2', ws, values);
    });

    it('removes a row', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.removeRows(1);
      table.commit();

      const newValues = spliceArray(values, 2, 1);
      checkTable('A1', ws, newValues);
    });

    it('adds a row', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.addRow([new Date('2019-08-05'), 5, 'Bird']);
      table.commit();

      const newValues = spliceArray(values, 5, 0, [
        new Date('2019-08-05'),
        5,
        'Bird',
      ]);
      checkTable('A1', ws, newValues);
    });

    it('removes a column', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.removeColumns(1);
      table.commit();

      const newValues = values.map(rVals => spliceArray(rVals, 1, 1));
      checkTable('A1', ws, newValues);
    });

    it('adds a column', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.addColumn(
        {
          name: 'Letter',
          totalsRowFunction: 'custom',
          totalsRowFormula: 'ROW()',
          totalsRowResult: 6,
          filterButton: true,
        },
        ['a', 'b', 'c', 'd'],
        2
      );
      table.commit();

      const colValues = [
        'Letter',
        'a',
        'b',
        'c',
        'd',
        {formula: 'ROW()', result: 6},
      ];
      const newValues = values.map((rVals, i) =>
        spliceArray(rVals, 2, 0, colValues[i])
      );
      checkTable('A1', ws, newValues);
    });

    it('renames a column', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      const column = table.getColumn(1);
      column.name = 'Code';
      table.commit();

      const newValues = [...values];
      newValues.splice(0, 1, ['Date', 'Code', 'Word']);
      newValues.splice(5, 1, [
        'Totals',
        {formula: 'SUBTOTAL(104,TestTable[Code])', result: 4},
        null,
      ]);

      checkTable('A1', ws, newValues);
    });

    function addSales(ws) {
      return ws.addTable({
        name: 'Sales',
        ref: 'A1',
        totalsRow: true,
        columns: [
          {name: 'Month'},
          {name: 'Qty', totalsRowFunction: 'sum'},
          {name: 'Total', calculatedColumnFormula: 'Sales[[#This Row],[Qty]]*10'},
        ],
        rows: [
          ['Jan', 1, {formula: 'Sales[[#This Row],[Qty]]*10'}],
          ['Feb', 2, {formula: 'Sales[[#This Row],[Qty]]*10'}],
        ],
      });
    }

    it('rejects table names Excel cannot open', () => {
      const ws = new Excel.Workbook().addWorksheet('blort');
      ['T1', 'Sales Table', 'R1C1', '1st'].forEach(name => {
        expect(() => ws.addTable({name, ref: 'A1', columns: [{name: 'a'}], rows: [[1]]})).to.throw(
          /is not valid/
        );
      });
      expect(() =>
        ws.addTable({name: '銷售', ref: 'A1', columns: [{name: 'a'}], rows: [[1]]})
      ).not.to.throw();
    });

    it('rejects a table name used by another table of the workbook', () => {
      const wb = new Excel.Workbook();
      addSales(wb.addWorksheet('one'));
      const ws = wb.addWorksheet('two');
      expect(() => addSales(ws)).to.throw(/already used/);
      const other = ws.addTable({name: 'Other', ref: 'A1', columns: [{name: 'a'}], rows: [[1]]});
      expect(() => {
        other.name = 'SALES';
      }).to.throw(/already used/);
    });

    it('renames a table and the references to it', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addSales(ws);
      ws.getCell('F1').value = {formula: 'SUM(Sales[Qty])+Other[Qty]'};
      wb.addWorksheet('other').getCell('A1').value = {formula: 'Sales[[#Totals],[Qty]]'};

      table.name = 'Orders';

      expect(table.displayName).to.equal('Orders');
      expect(ws.getTable('Orders')).to.equal(table);
      expect(ws.getTable('Sales')).to.be.undefined();
      expect(ws.getCell('F1').value.formula).to.equal('SUM(Orders[Qty])+Other[Qty]');
      expect(wb.getWorksheet('other').getCell('A1').value.formula).to.equal(
        'Orders[[#Totals],[Qty]]'
      );
      expect(ws.getCell('C2').value.formula).to.equal('Orders[[#This Row],[Qty]]*10');
      expect(table.getColumn(2).column.calculatedColumnFormula).to.equal(
        'Orders[[#This Row],[Qty]]*10'
      );

      table.addRow(['Mar', 3]);
      expect(ws.getCell('C4').value.formula).to.equal('Orders[[#This Row],[Qty]]*10');
      expect(ws.getCell('B5').value.formula).to.equal('SUBTOTAL(109,Orders[Qty])');
    });

    it('renames a column and the references to it', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addSales(ws);
      ws.getCell('F1').value = {formula: 'SUM(Sales[[#Data],[Qty]])'};
      ws.getCell('C3').value = {formula: '[@Qty]*10'};

      table.getColumn(1).name = 'Quantity';
      table.commit();

      expect(ws.getCell('B1').value).to.equal('Quantity');
      expect(ws.getCell('F1').value.formula).to.equal('SUM(Sales[[#Data],[Quantity]])');
      expect(ws.getCell('C2').value.formula).to.equal('Sales[[#This Row],[Quantity]]*10');
      expect(ws.getCell('B4').value.formula).to.equal('SUBTOTAL(109,Sales[Quantity])');
    });

    it('sets the theme', () => {
      const ws = new Excel.Workbook().addWorksheet('blort');
      const table = addSales(ws);
      table.theme = 'TableStyleLight9';
      expect(table.theme).to.equal('TableStyleLight9');
      expect(table.model.style.theme).to.equal('TableStyleLight9');
    });

    it('fills the calculated column of a new row and moves formulas with the rows', () => {
      const ws = new Excel.Workbook().addWorksheet('blort');
      const table = ws.addTable({
        name: 'Sales',
        ref: 'A1',
        columns: [{name: 'Qty'}, {name: 'Double', calculatedColumnFormula: 'A2*2'}],
        rows: [
          [1, {formula: 'A2*2'}],
          [2, {formula: 'A3*2'}],
          [3, {formula: 'A4*2'}],
        ],
      });

      table.addRow([4]);
      expect(ws.getCell('B5').value).to.deep.equal({formula: 'A5*2'});

      table.removeRows(0);
      expect(ws.getCell('B2').value.formula).to.equal('A2*2');
      expect(ws.getCell('B3').value.formula).to.equal('A3*2');
      expect(ws.getCell('B4').value.formula).to.equal('A4*2');
      expect(ws.getCell('A5').value).to.be.null();
      expect(table.model.tableRef).to.equal('A1:B4');
    });

    it('shows a total in the first column', () => {
      const ws = new Excel.Workbook().addWorksheet('blort');
      ws.addTable({
        name: 'Sales',
        ref: 'A1',
        totalsRow: true,
        columns: [
          {name: 'Qty', totalsRowFunction: 'sum'},
          {name: 'Name', totalsRowLabel: 'Sum'},
        ],
        rows: [[1, 'a']],
      });
      expect(ws.getCell('A3').value.formula).to.equal('SUBTOTAL(109,Sales[Qty])');
      expect(ws.getCell('B3').value).to.equal('Sum');
    });

    it('turns references into cell references when the table is removed', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      addSales(ws);
      ws.getCell('F1').value = {formula: 'SUM(Sales[Qty])'};
      wb.addWorksheet('other sheet').getCell('A1').value = {formula: 'Sales[[#Totals],[Qty]]'};

      ws.removeTable('Sales');

      expect(ws.getTables()).to.have.length(0);
      expect(ws.getCell('C2').value.formula).to.equal('blort!$B2*10');
      expect(ws.getCell('C3').value.formula).to.equal('blort!$B3*10');
      expect(ws.getCell('B4').value.formula).to.equal('SUBTOTAL(109,blort!$B$2:$B$3)');
      expect(ws.getCell('F1').value.formula).to.equal('SUM(blort!$B$2:$B$3)');
      expect(wb.getWorksheet('other sheet').getCell('A1').value.formula).to.equal('blort!$B$4');
    });
  });
});
