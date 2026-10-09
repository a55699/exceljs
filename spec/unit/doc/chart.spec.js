const Workbook = verquire('doc/workbook');

function salesWorkbook() {
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet('Sales Data');
  worksheet.addRows([
    ['Month', 'East', 'West'],
    ['Jan', 10, 30],
    ['Feb', 40, 60],
    ['Mar', 50, 70],
  ]);
  return {workbook, worksheet};
}

const range = 'E2:L17';

describe('Chart', () => {
  describe('options', () => {
    it('builds series from a data range, one per column', () => {
      const {worksheet} = salesWorkbook();
      const chart = worksheet.addChart({
        type: 'bar',
        data: {sheet: worksheet, ref: 'A1:C4'},
        range,
      });

      const {series} = chart.model;
      expect(series.map(s => s.name.formula)).to.deep.equal([
        "'Sales Data'!$B$1",
        "'Sales Data'!$C$1",
      ]);
      expect(series.map(s => s.categories.formula)).to.deep.equal([
        "'Sales Data'!$A$2:$A$4",
        "'Sales Data'!$A$2:$A$4",
      ]);
      expect(series.map(s => s.values.formula)).to.deep.equal([
        "'Sales Data'!$B$2:$B$4",
        "'Sales Data'!$C$2:$C$4",
      ]);
      expect(worksheet.getCharts()).to.deep.equal([chart]);
    });

    it('builds series from a data range, one per row', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Rows');
      worksheet.addRows([
        ['', 'Jan', 'Feb'],
        ['East', 10, 40],
        ['West', 30, 60],
      ]);
      const chart = worksheet.addChart({
        type: 'line',
        data: {sheet: 'Rows', ref: 'A1:C3', seriesIn: 'rows'},
        range,
      });
      const {series} = chart.model;
      expect(series.map(s => s.name.cache)).to.deep.equal([['East'], ['West']]);
      expect(series[0].categories.formula).to.equal("'Rows'!$B$1:$C$1");
      expect(series[1].values.formula).to.equal("'Rows'!$B$3:$C$3");
    });

    it('uses the first column of the data of a scatter chart as X values', () => {
      const {worksheet} = salesWorkbook();
      const chart = worksheet.addChart({
        type: 'scatter',
        data: {sheet: worksheet, ref: 'B1:C4'},
        range,
      });
      const [series] = chart.model.series;
      expect(series.xValues.formula).to.equal("'Sales Data'!$B$2:$B$4");
      expect(series.yValues.formula).to.equal("'Sales Data'!$C$2:$C$4");
    });

    it('accepts formula strings and quotes sheet names', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet("Bob's data");
      worksheet.addRows([
        [1, 2],
        [3, 4],
      ]);
      const chart = worksheet.addChart({
        type: 'line',
        series: [
          {name: 'One', values: "'Bob''s data'!$A$1:$A$2"},
          {values: {sheet: worksheet, ref: 'B1:B2'}},
        ],
        range,
      });
      const {series} = chart.model;
      expect(series[0].name).to.deep.equal({text: 'One'});
      expect(series[0].values.formula).to.equal("'Bob''s data'!$A$1:$A$2");
      expect(series[1].values.formula).to.equal("'Bob''s data'!$B$1:$B$2");
      expect(series[1].name).to.equal(undefined);
    });

    it('sets the defaults of each type', () => {
      const {worksheet} = salesWorkbook();
      const data = {sheet: worksheet, ref: 'A1:C4'};
      const bar = worksheet.addChart({type: 'bar', data, range}).model;
      expect([bar.direction, bar.grouping]).to.deep.equal(['col', 'clustered']);
      const line = worksheet.addChart({type: 'line', data, range}).model;
      expect([line.grouping, line.markers]).to.deep.equal(['standard', true]);
      expect(worksheet.addChart({type: 'doughnut', data, range}).model.holeSize).to.equal(50);
      expect(worksheet.addChart({type: 'scatter', data, range}).model.lines).to.equal(false);
      expect(bar.legend).to.deep.equal({position: 'r'});
      expect(worksheet.addChart({type: 'pie', data, legend: false, range}).model.legend).to.equal(
        false
      );
    });

    it('throws for invalid options', () => {
      const {worksheet} = salesWorkbook();
      const data = {sheet: worksheet, ref: 'A1:C4'};
      const values = {sheet: worksheet, ref: 'B2:B4'};
      const cases = [
        [undefined, 'options are required'],
        [{type: 'radar', data, range}, 'type must be one of'],
        [{type: 'bar', range}, 'series (a non-empty array) or data is required'],
        [{type: 'bar', series: [], range}, 'series (a non-empty array) or data is required'],
        [{type: 'bar', data, series: [{values}], range}, 'use series or data, not both'],
        [{type: 'bar', series: [{name: 'x'}], range}, 'series[0] needs values'],
        [{type: 'scatter', series: [{values}], range}, 'needs xValues and yValues'],
        [
          {type: 'bar', series: [{values: {sheet: 'Nope', ref: 'B2:B4'}}], range},
          'does not exist: Nope',
        ],
        [
          {type: 'bar', series: [{values: {sheet: worksheet, ref: 'B2:C4'}}], range},
          'single row or a single column',
        ],
        [{type: 'bar', series: [{values: {sheet: worksheet, ref: 'B2:'}}], range}, 'invalid range'],
        [{type: 'bar', series: [{values: 'B2:B4'}], range}, "must be like 'Sheet 1'!$A$1:$A$4"],
        [
          {type: 'bar', series: [{values, categories: {sheet: worksheet, ref: 'A2:A3'}}], range},
          'categories and values must have the same number of cells',
        ],
        [
          {type: 'bar', series: [{name: {sheet: worksheet, ref: 'B1:C1'}, values}], range},
          'name must be a single cell',
        ],
        [{type: 'bar', direction: 'up', data, range}, "direction must be 'col' or 'bar'"],
        [{type: 'line', grouping: 'clustered', data, range}, 'grouping of a line chart'],
        [{type: 'doughnut', holeSize: 95, data, range}, 'holeSize must be a whole number'],
        [{type: 'bar', legend: {position: 'x'}, data, range}, 'legend.position must be one of'],
        [{type: 'bar', axes: {y: {min: '0'}}, data, range}, 'axes.y.min must be a number'],
        [{type: 'bar', data}, "range must be like 'E2:L17'"],
        [{type: 'bar', data, range: 'E2'}, 'range must be a cell range'],
        [{type: 'bar', data: {sheet: worksheet, ref: 'A1:C1'}, range}, 'data has no values'],
      ];
      cases.forEach(([options, message]) => {
        const addChart = () => worksheet.addChart(options);
        expect(addChart, message).to.throw(message);
      });
      expect(worksheet.getCharts()).to.deep.equal([]);
    });
  });

  describe('cached values', () => {
    it('caches numbers, text and the cell number format', () => {
      const {worksheet} = salesWorkbook();
      worksheet.getCell('B3').numFmt = '0.0';
      worksheet.getCell('C2').value = {formula: 'B2*3', result: 30};
      worksheet.getCell('C3').value = null;
      const chart = worksheet.addChart({
        type: 'bar',
        data: {sheet: worksheet, ref: 'A1:C4'},
        range,
      });

      const [east, west] = chart.model.series;
      expect(east.name).to.deep.equal({formula: "'Sales Data'!$B$1", type: 'str', cache: ['East']});
      expect(east.categories).to.deep.equal({
        formula: "'Sales Data'!$A$2:$A$4",
        type: 'str',
        cache: ['Jan', 'Feb', 'Mar'],
      });
      expect(east.values).to.deep.equal({
        formula: "'Sales Data'!$B$2:$B$4",
        type: 'num',
        formatCode: '0.0',
        cache: [10, 40, 50],
      });
      expect(west.values.cache).to.deep.equal([30, null, 70]);
    });

    it('takes the values when the model is read, not when the chart is added', () => {
      const {worksheet} = salesWorkbook();
      const chart = worksheet.addChart({
        type: 'bar',
        data: {sheet: worksheet, ref: 'A1:B4'},
        range,
      });
      worksheet.getCell('B2').value = 15;
      expect(chart.model.series[0].values.cache).to.deep.equal([15, 40, 50]);
    });

    it('writes number and date categories as numbers', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Dates');
      worksheet.addRows([
        [new Date(Date.UTC(2026, 0, 1)), 1],
        [new Date(Date.UTC(2026, 0, 2)), 2],
      ]);
      worksheet.getCell('A1').numFmt = 'yyyy-mm-dd';
      const chart = worksheet.addChart({
        type: 'line',
        series: [
          {categories: {sheet: worksheet, ref: 'A1:A2'}, values: {sheet: worksheet, ref: 'B1:B2'}},
        ],
        range,
      });
      expect(chart.model.series[0].categories).to.deep.equal({
        formula: "'Dates'!$A$1:$A$2",
        type: 'num',
        formatCode: 'yyyy-mm-dd',
        cache: [46023, 46024],
      });
    });
  });
});
