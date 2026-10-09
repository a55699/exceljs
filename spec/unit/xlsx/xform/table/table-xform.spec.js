const fs = require('fs');

const testXformHelper = require('../test-xform-helper');

const TableXform = verquire('xlsx/xform/table/table-xform');

const expectations = [
  {
    title: 'showing filter',
    create() {
      return new TableXform();
    },
    initialModel: null,
    preparedModel: require('./data/table.1.1'),
    xml: fs.readFileSync(`${__dirname}/data/table.1.2.xml`).toString(),
    parsedModel: require('./data/table.1.3'),
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'as Excel writes it',
    create() {
      return new TableXform();
    },
    // no headerRowCount (1), a filter button hidden on the second column only
    xml:
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B3" totalsRowShown="0">' +
      '<autoFilter ref="A1:B3"><filterColumn colId="1" hiddenButton="1"/></autoFilter>' +
      '<tableColumns count="2"><tableColumn id="1" name="Month"/><tableColumn id="2" name="Qty"/></tableColumns>' +
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>' +
      '</table>',
    parsedModel: {
      name: 'Sales',
      displayName: 'Sales',
      tableRef: 'A1:B3',
      totalsRow: false,
      headerRow: true,
      autoFilterRef: 'A1:B3',
      columns: [{name: 'Month'}, {name: 'Qty', filterButton: false}],
      style: {
        theme: 'TableStyleMedium2',
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    },
    tests: ['parse'],
  },
  {
    title: 'no header row',
    create() {
      return new TableXform();
    },
    preparedModel: {
      id: 2,
      name: 'Codes',
      tableRef: 'C4:D5',
      totalsRow: false,
      headerRow: false,
      columns: [
        {id: 1, name: 'Column1'},
        {id: 2, name: 'Column2'},
      ],
      style: {theme: 'TableStyleMedium2'},
    },
    // a table without a header row has no autoFilter
    xml:
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"' +
      ' id="2" name="Codes" displayName="Codes" ref="C4:D5" totalsRowShown="1" headerRowCount="0">' +
      '<tableColumns count="2"><tableColumn id="1" name="Column1"/><tableColumn id="2" name="Column2"/></tableColumns>' +
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/>' +
      '</table>',
    tests: ['render'],
  },
];

describe('TableXform', () => {
  testXformHelper(expectations);
});
