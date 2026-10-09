const testXformHelper = require('../test-xform-helper');

const FilterColumnXform = verquire('xlsx/xform/table/filter-column-xform');

const expectations = [
  {
    title: 'showing filter',
    create() {
      return new FilterColumnXform();
    },
    initialModel: {filterButton: true},
    preparedModel: {colId: '0', filterButton: true},
    xml: '<filterColumn colId="0" hiddenButton="0" />',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['prepare', 'render', 'renderIn', 'parse'],
    options: {index: 0},
  },
  {
    title: 'hidden filter',
    create() {
      return new FilterColumnXform();
    },
    initialModel: {filterButton: false},
    preparedModel: {colId: '1', filterButton: false},
    xml: '<filterColumn colId="1" hiddenButton="1" />',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['prepare', 'render', 'renderIn', 'parse'],
    options: {index: 1},
  },
  {
    title: 'with custom filter',
    create() {
      return new FilterColumnXform();
    },
    initialModel: {filterButton: false, customFilters: [{val: '*brandywine*'}]},
    preparedModel: {
      colId: '0',
      filterButton: false,
      customFilters: [{val: '*brandywine*'}],
    },
    xml:
      '<filterColumn colId="0" hiddenButton="1"><customFilters><customFilter val="*brandywine*"/></customFilters></filterColumn>',
    parsedModel: {
      colId: '0',
      filterButton: false,
      customFilters: [{val: '*brandywine*'}],
      filterXml: '<customFilters><customFilter val="*brandywine*"/></customFilters>',
    },
    tests: ['prepare', 'render', 'renderIn', 'parse'],
    options: {index: 0},
  },
  {
    title: 'button shown when hiddenButton is not written',
    create() {
      return new FilterColumnXform();
    },
    xml: '<filterColumn colId="2" />',
    parsedModel: {colId: '2', filterButton: true},
    tests: ['parse'],
  },
  {
    title: 'keeps a filter it does not read',
    create() {
      return new FilterColumnXform();
    },
    xml: '<filterColumn colId="0" hiddenButton="0"><top10 top="1" val="3" filterVal="7"/></filterColumn>',
    parsedModel: {
      colId: '0',
      filterButton: true,
      filterXml: '<top10 top="1" val="3" filterVal="7"/>',
    },
    preparedModel: {
      colId: '0',
      filterButton: true,
      filterXml: '<top10 top="1" val="3" filterVal="7"/>',
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('FilterColumnXform', () => {
  testXformHelper(expectations);
});
