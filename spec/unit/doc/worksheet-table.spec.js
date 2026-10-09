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

    it('keeps one empty row in a table without rows', () => {
      const ws = new Excel.Workbook().addWorksheet('blort');
      const empty = ws.addTable({name: 'Empty', ref: 'H1', columns: [{name: 'a'}], rows: []});
      expect(empty.model.tableRef).to.equal('H1:H2');

      const table = addSales(ws);
      table.removeRows(0, 2);
      expect(table.model.tableRef).to.equal('A1:C3');
      expect(ws.getCell('A2').value).to.be.null();
      expect(ws.getCell('A3').value).to.equal('Total');
    });

    it('checks the columns and the place of a table', () => {
      const ws = new Excel.Workbook().addWorksheet('blort');
      const table = addSales(ws);
      expect(() =>
        ws.addTable({name: 'Dup', ref: 'H1', columns: [{name: 'a'}, {name: 'A'}], rows: [[1, 2]]})
      ).to.throw(/used twice/);
      expect(() =>
        ws.addTable({name: 'Over', ref: 'B2', columns: [{name: 'a'}], rows: [[1]]})
      ).to.throw(/overlaps table Sales/);
      expect(() => {
        table.getColumn(1).name = 'month';
      }).to.throw(/used twice/);

      const numbers = ws.addTable({name: 'Years', ref: 'H1', columns: [{name: 2024}], rows: [[1]]});
      expect(numbers.getColumn(0).name).to.equal('2024');
      expect(ws.getCell('H1').value).to.equal('2024');
    });

    it('escapes column names in the totals row', () => {
      const ws = new Excel.Workbook().addWorksheet('blort');
      ws.addTable({
        name: 'Prices',
        ref: 'A1',
        totalsRow: true,
        columns: [{name: 'Price [USD]', totalsRowFunction: 'sum'}],
        rows: [[1]],
      });
      expect(ws.getCell('A3').value.formula).to.equal("SUBTOTAL(109,Prices[[Price '[USD']]])");
    });

    it('renames a column when its header cell is changed', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      addSales(ws);
      ws.getCell('F1').value = {formula: 'SUM(Sales[Qty])'};

      ws.getCell('B1').value = 'Quantity';
      ws.getCell('A1').value = null;
      const {tables} = ws.model;

      expect(tables[0].columns.map(column => column.name)).to.deep.equal([
        'Column1',
        'Quantity',
        'Total',
      ]);
      expect(ws.getCell('A1').value).to.equal('Column1');
      expect(ws.getCell('F1').value.formula).to.equal('SUM(Sales[Quantity])');
      expect(ws.getCell('C2').value.formula).to.equal('Sales[[#This Row],[Quantity]]*10');
    });

    it('turns references to a removed column into #REF!', () => {
      const ws = new Excel.Workbook().addWorksheet('blort');
      const table = addSales(ws);
      ws.getCell('F1').value = {formula: 'SUM(Sales[Qty])+COUNTA(Sales[Month])'};
      table.removeColumns(1);
      table.commit();
      expect(ws.getCell('F1').value.formula).to.equal('SUM(#REF!)+COUNTA(Sales[Month])');
    });

    describe('when worksheet rows and columns are spliced', () => {
      // Sales on B2:D6: header, three data rows, totals row; the expected
      // ranges are what Excel gives for the same changes
      function addPrices(ws) {
        ws.getCell('H1').value = {formula: 'SUM(Sales[Qty])'};
        return ws.addTable({
          name: 'Sales',
          ref: 'B2',
          totalsRow: true,
          columns: [{name: 'Month'}, {name: 'Qty', totalsRowFunction: 'sum'}, {name: 'Price'}],
          rows: [
            ['Jan', 1, 5],
            ['Feb', 2, 6],
            ['Mar', 3, 7],
          ],
        });
      }
      const cases = [
        ['inserts a row at the header row', ws => ws.spliceRows(2, 0, []), 'B3:D7'],
        ['inserts a row in the data', ws => ws.insertRow(4, ['Mid', 9, 9]), 'B2:D7'],
        ['inserts a row at the totals row', ws => ws.spliceRows(6, 0, []), 'B2:D7'],
        ['inserts a row below the table', ws => ws.spliceRows(7, 0, []), 'B2:D6'],
        ['deletes the totals row', ws => ws.spliceRows(6, 1), 'B2:D5'],
        ['deletes all the data rows', ws => ws.spliceRows(3, 3), 'B2:D4'],
        ['inserts a column before the table', ws => ws.spliceColumns(2, 0, []), 'C2:E6'],
        ['inserts a column in the table', ws => ws.spliceColumns(3, 0, []), 'B2:E6'],
        ['inserts a column after the table', ws => ws.spliceColumns(5, 0, []), 'B2:D6'],
        ['deletes the first column', ws => ws.spliceColumns(2, 1), 'B2:C6'],
      ];
      cases.forEach(([title, splice, ref]) => {
        it(title, () => {
          const ws = new Excel.Workbook().addWorksheet('blort');
          addPrices(ws);
          splice(ws);
          expect(ws.getTable('Sales').model.tableRef).to.equal(ref);
        });
      });

      it('keeps the rows and names of the table', () => {
        const ws = new Excel.Workbook().addWorksheet('blort');
        const table = addPrices(ws);
        ws.spliceRows(6, 1);
        expect(table.totalsRow).to.equal(false);
        ws.spliceColumns(3, 0, []);
        expect(table.model.columns.map(column => column.name)).to.deep.equal([
          'Month',
          'Column1',
          'Qty',
          'Price',
        ]);
        expect(ws.getCell('C2').value).to.equal('Column1');
        ws.insertRow(3, [null, 'Dec', null, 0, 4]);
        expect(table.model.rows.map(row => row[0])).to.deep.equal(['Dec', 'Jan', 'Feb', 'Mar']);
      });

      it('removes a table whose rows or columns are all deleted', () => {
        const ws = new Excel.Workbook().addWorksheet('blort');
        addPrices(ws);
        ws.spliceRows(2, 5);
        expect(ws.getTables()).to.have.length(0);
        expect(ws.getCell('H1').value.formula).to.equal('SUM(#REF!)');
      });

      it('does not delete the header row alone', () => {
        const ws = new Excel.Workbook().addWorksheet('blort');
        addPrices(ws);
        expect(() => ws.spliceRows(2, 1)).to.throw(/header row of table Sales/);
        expect(() => ws.spliceRows(1, 3)).to.throw(/header row of table Sales/);
        expect(ws.getTable('Sales').model.tableRef).to.equal('B2:D6');
      });
    });
  });
});
