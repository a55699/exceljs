const testXformHelper = require('../test-xform-helper');

const DefinedNameXform = verquire('xlsx/xform/book/defined-name-xform');

const expectations = [
  {
    title: 'Defined Names',
    create() {
      return new DefinedNameXform();
    },
    preparedModel: {name: 'foo', ranges: ['bar!$A$1:$C$1']},
    xml: '<definedName name="foo">bar!$A$1:$C$1</definedName>',
    parsedModel: {name: 'foo', ranges: ['bar!$A$1:$C$1']},
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Print Area',
    create() {
      return new DefinedNameXform();
    },
    preparedModel: {
      name: '_xlnm.Print_Area',
      localSheetId: 0,
      ranges: ['bar!$A$1:$C$10'],
    },
    xml:
      '<definedName name="_xlnm.Print_Area" localSheetId="0">bar!$A$1:$C$10</definedName>',
    parsedModel: {
      name: '_xlnm.Print_Area',
      localSheetId: 0,
      ranges: ['bar!$A$1:$C$10'],
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'String with something that looks like a range',
    create() {
      return new DefinedNameXform();
    },
    preparedModel: {name: 'foo', ranges: [], formula: '"OFFSET($A$10;0;0;0;1)"'},
    xml: '<definedName name="foo">"OFFSET($A$10;0;0;0;1)"</definedName>',
    // a constant is kept as text
    parsedModel: {name: 'foo', ranges: [], formula: '"OFFSET($A$10;0;0;0;1)"'},
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Formula',
    create() {
      return new DefinedNameXform();
    },
    preparedModel: {name: 'TotalQty', ranges: [], formula: 'SUM(Sales[Qty])'},
    xml: '<definedName name="TotalQty">SUM(Sales[Qty])</definedName>',
    parsedModel: {name: 'TotalQty', ranges: [], formula: 'SUM(Sales[Qty])'},
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Hidden name of one worksheet',
    create() {
      return new DefinedNameXform();
    },
    preparedModel: {
      name: '_xlnm._FilterDatabase',
      localSheetId: 1,
      attributes: {hidden: '1'},
      ranges: ['Data!$A$1:$D$5'],
    },
    xml: '<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Data!$A$1:$D$5</definedName>',
    parsedModel: {
      name: '_xlnm._FilterDatabase',
      localSheetId: 1,
      attributes: {hidden: '1'},
      ranges: ['Data!$A$1:$D$5'],
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('DefinedNameXform', () => {
  testXformHelper(expectations);
});
