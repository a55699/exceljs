const {mapTableReferences, moveFormula, parseSpecifier, renderSpecifier, toCellReference} =
  verquire('utils/table-ref');

// Sales on sheet Data: A1:D5, header row, three data rows, totals row
const geometry = {
  sheetName: 'Data',
  top: 1,
  left: 1,
  headerRow: true,
  totalsRow: true,
  height: 3,
  columns: ['Month', 'Qty', 'Unit Price', 'Total'],
};

function convert(formula, row, inTable = false) {
  return mapTableReferences(formula, 'Sales', inTable, spec =>
    toCellReference(spec, geometry, row)
  );
}

describe('table-ref', () => {
  describe('toCellReference', () => {
    // the formulas as Excel writes them, and what Excel's Convert to Range
    // made of them
    const excelCases = [
      [2, 'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Unit Price]]', 'Data!$B2*Data!$C2'],
      [5, 'SUBTOTAL(109,Sales[Qty])', 'SUBTOTAL(109,Data!$B$2:$B$4)'],
      [8, 'Sales[[#Totals],[Qty]]', 'Data!$B$5'],
      [9, 'COUNTA(Sales[#All])', 'COUNTA(Data!$A$1:$D$5)'],
      [10, 'Sales[[#Headers],[Month]]', 'Data!$A$1'],
      [11, 'SUM(Sales[[Qty]:[Unit Price]])', 'SUM(Data!$B$2:$C$4)'],
      [12, 'ROWS(Sales[])', 'ROWS(Data!$A$2:$D$4)'],
      [13, 'SUM(Sales[[#Data],[Qty]])', 'SUM(Data!$B$2:$B$4)'],
      [14, 'COUNTA(Sales[#Headers])', 'COUNTA(Data!$A$1:$D$1)'],
      [15, '"Sales[Qty]"&SUM(Sales[Qty])', '"Sales[Qty]"&SUM(Data!$B$2:$B$4)'],
      [16, 'SUM(Sales[[#All],[Qty]])', 'SUM(Data!$B$1:$B$5)'],
      [17, 'SUM(Sales[[#Headers],[#Data],[Qty]])', 'SUM(Data!$B$1:$B$4)'],
    ];
    excelCases.forEach(([row, formula, expected]) => {
      it(`converts ${formula} as Excel does`, () => {
        expect(convert(formula, row)).to.equal(expected);
      });
    });

    it('converts references without the table name in a cell of the table', () => {
      expect(convert('[@Qty]*[@[Unit Price]]', 3, true)).to.equal('Data!$B3*Data!$C3');
      expect(convert('[[#This Row],Qty]*10', 3, true)).to.equal('Data!$B3*10');
    });

    it('keeps references without the table name outside the table', () => {
      expect(convert('[1]Sheet1!A1+[@Qty]', 3)).to.equal('[1]Sheet1!A1+[@Qty]');
    });

    it('keeps references to other tables and is not case sensitive', () => {
      expect(convert('sales[@Qty]+Other[Qty]', 2)).to.equal('Data!$B2+Other[Qty]');
    });

    it('gives #REF! for a missing column or part', () => {
      expect(convert('Sales[Price]', 2)).to.equal('#REF!');
      expect(
        mapTableReferences('Sales[#Totals]', 'Sales', false, spec =>
          toCellReference(spec, {...geometry, totalsRow: false}, 2)
        )
      ).to.equal('#REF!');
    });

    it('gives #REF! for the current row of a formula without a row', () => {
      expect(convert('Sales[@Qty]')).to.equal('#REF!');
    });

    it('quotes the sheet name when needed', () => {
      expect(
        mapTableReferences('SUM(Sales[Qty])', 'Sales', false, spec =>
          toCellReference(spec, {...geometry, sheetName: "Bob's Data"}, 2)
        )
      ).to.equal("SUM('Bob''s Data'!$B$2:$B$4)");
    });
  });

  describe('parseSpecifier and renderSpecifier', () => {
    const cases = [
      ['Qty', {items: [], columns: ['Qty', 'Qty']}, 'Qty'],
      ['[#This Row],[Qty]', {items: ['#this row'], columns: ['Qty', 'Qty']}, '[#This Row],[Qty]'],
      [
        '@[Unit Price]',
        {items: ['#this row'], columns: ['Unit Price', 'Unit Price']},
        '[#This Row],[Unit Price]',
      ],
      ['@Qty', {items: ['#this row'], columns: ['Qty', 'Qty']}, '[#This Row],[Qty]'],
      ['[Qty]:[Total]', {items: [], columns: ['Qty', 'Total']}, '[Qty]:[Total]'],
      ['#All', {items: ['#all'], columns: null}, '[#All]'],
      ['', {items: [], columns: null}, ''],
      ["[a'#b]", {items: [], columns: ['a#b', 'a#b']}, "[a'#b]"],
    ];
    cases.forEach(([text, spec, rendered]) => {
      it(`reads and writes ${text || 'an empty specifier'}`, () => {
        expect(parseSpecifier(text)).to.deep.equal(spec);
        expect(renderSpecifier(spec)).to.equal(rendered);
      });
    });
  });

  describe('moveFormula', () => {
    it('moves relative references and keeps structured references and texts', () => {
      expect(moveFormula('B3*2+$C$1+Sales[[#This Row],[Q1]]&"A1"', 'E3', 'E2')).to.equal(
        'B2*2+$C$1+Sales[[#This Row],[Q1]]&"A1"'
      );
      expect(moveFormula("'Q1 Sheet'!A1+A1", 'B2', 'C4')).to.equal("'Q1 Sheet'!B3+B3");
    });
  });
});
