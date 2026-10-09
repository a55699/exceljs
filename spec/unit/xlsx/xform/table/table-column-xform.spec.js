const testXformHelper = require('../test-xform-helper');

const TableColumnXform = verquire('xlsx/xform/table/table-column-xform');

const expectations = [
  {
    title: 'label',
    create() {
      return new TableColumnXform();
    },
    preparedModel: {id: 1, name: 'Foo', totalsRowLabel: 'Bar'},
    xml: '<tableColumn id="1" name="Foo" totalsRowLabel="Bar" />',
    parsedModel: {name: 'Foo', totalsRowLabel: 'Bar'},
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'function',
    create() {
      return new TableColumnXform();
    },
    preparedModel: {id: 1, name: 'Foo', totalsRowFunction: 'Baz'},
    xml: '<tableColumn id="1" name="Foo" totalsRowFunction="Baz" />',
    parsedModel: {name: 'Foo', totalsRowFunction: 'Baz'},
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'calculated column',
    create() {
      return new TableColumnXform();
    },
    preparedModel: {
      id: 2,
      name: 'Total',
      calculatedColumnFormula: 'Sales[[#This Row],[Qty]]*2',
    },
    xml: '<tableColumn id="2" name="Total"><calculatedColumnFormula>Sales[[#This Row],[Qty]]*2</calculatedColumnFormula></tableColumn>',
    parsedModel: {name: 'Total', calculatedColumnFormula: 'Sales[[#This Row],[Qty]]*2'},
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'custom total',
    create() {
      return new TableColumnXform();
    },
    preparedModel: {
      id: 1,
      name: 'Qty',
      totalsRowFunction: 'custom',
      totalsRowFormula: 'SUM(Sales[Qty])*2',
    },
    xml: '<tableColumn id="1" name="Qty" totalsRowFunction="custom"><totalsRowFormula>SUM(Sales[Qty])*2</totalsRowFormula></tableColumn>',
    parsedModel: {name: 'Qty', totalsRowFunction: 'custom', totalsRowFormula: 'SUM(Sales[Qty])*2'},
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'formula of a total that is not custom is not written',
    create() {
      return new TableColumnXform();
    },
    preparedModel: {
      id: 1,
      name: 'Qty',
      totalsRowFunction: 'sum',
      totalsRowFormula: 'SUBTOTAL(109,Sales[Qty])',
    },
    xml: '<tableColumn id="1" name="Qty" totalsRowFunction="sum" />',
    tests: ['render', 'renderIn'],
  },
];

describe('TableColumnXform', () => {
  testXformHelper(expectations);
});
