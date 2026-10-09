const BaseXform = require('../base-xform');
const ChartSeriesXform = require('./chart-series-xform');
const ChartAxisXform = require('./chart-axis-xform');

// ids of the two axes of a chart; they only need to be unique in the chart
const X_AXIS_ID = 10;
const Y_AXIS_ID = 100;

const CHART_TAGS = {
  bar: 'c:barChart',
  line: 'c:lineChart',
  area: 'c:areaChart',
  pie: 'c:pieChart',
  doughnut: 'c:doughnutChart',
  scatter: 'c:scatterChart',
};

// <c:plotArea>: the chart type element with its series, and the axes.
// Write only.
class ChartPlotAreaXform extends BaseXform {
  constructor() {
    super();
    this.series = new ChartSeriesXform();
    this.axis = new ChartAxisXform();
  }

  get tag() {
    return 'c:plotArea';
  }

  render(xmlStream, model) {
    const {type} = model;
    xmlStream.openNode(this.tag);
    xmlStream.leafNode('c:layout');

    xmlStream.openNode(CHART_TAGS[type]);
    if (type === 'bar') {
      xmlStream.leafNode('c:barDir', {val: model.direction});
    }
    if (model.grouping) {
      xmlStream.leafNode('c:grouping', {val: model.grouping});
    }
    if (type === 'scatter') {
      xmlStream.leafNode('c:scatterStyle', {val: 'lineMarker'});
    }
    const pie = type === 'pie' || type === 'doughnut';
    xmlStream.leafNode('c:varyColors', {val: pie ? 1 : 0});
    model.series.forEach((series, index) => {
      this.series.render(xmlStream, series, {index, chart: model});
    });
    if (type === 'bar') {
      xmlStream.leafNode('c:gapWidth', {val: 150});
      if (model.grouping !== 'clustered') {
        xmlStream.leafNode('c:overlap', {val: 100});
      }
    }
    if (type === 'line') {
      xmlStream.leafNode('c:marker', {val: 1});
    }
    if (pie) {
      xmlStream.leafNode('c:firstSliceAng', {val: 0});
    }
    if (type === 'doughnut') {
      xmlStream.leafNode('c:holeSize', {val: model.holeSize});
    }
    if (!pie) {
      xmlStream.leafNode('c:axId', {val: X_AXIS_ID});
      xmlStream.leafNode('c:axId', {val: Y_AXIS_ID});
    }
    xmlStream.closeNode();

    if (!pie) {
      this.renderAxes(xmlStream, model);
    }
    xmlStream.closeNode();
  }

  renderAxes(xmlStream, model) {
    const {type} = model;
    // horizontal bars have the categories on the left
    const horizontal = type === 'bar' && model.direction === 'bar';
    const crossBetween = type === 'bar' || type === 'line' ? 'between' : 'midCat';
    this.axis.render(xmlStream, {
      kind: type === 'scatter' ? 'val' : 'cat',
      id: X_AXIS_ID,
      crossId: Y_AXIS_ID,
      position: horizontal ? 'l' : 'b',
      crossBetween,
      gridlines: false,
      axis: model.axes.x,
    });
    this.axis.render(xmlStream, {
      kind: 'val',
      id: Y_AXIS_ID,
      crossId: X_AXIS_ID,
      position: horizontal ? 'b' : 'l',
      crossBetween,
      gridlines: true,
      axis: model.axes.y,
    });
  }
}

module.exports = ChartPlotAreaXform;
