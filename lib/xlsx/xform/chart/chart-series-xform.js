const BaseXform = require('../base-xform');
const ChartDataRefXform = require('./chart-data-ref-xform');

// <c:ser> of a chart, with its children in the order of the chart type's
// schema. Write only.
class ChartSeriesXform extends BaseXform {
  constructor() {
    super();
    this.dataRef = new ChartDataRefXform();
  }

  get tag() {
    return 'c:ser';
  }

  // model: the series; options: {index, chart}
  render(xmlStream, model, options) {
    const {index, chart} = options;
    xmlStream.openNode(this.tag);
    xmlStream.leafNode('c:idx', {val: index});
    xmlStream.leafNode('c:order', {val: index});
    this.renderName(xmlStream, model.name);

    switch (chart.type) {
      case 'bar':
        xmlStream.leafNode('c:invertIfNegative', {val: 0});
        this.renderCategories(xmlStream, model);
        break;
      case 'line':
        if (!chart.markers) {
          xmlStream.openNode('c:marker');
          xmlStream.leafNode('c:symbol', {val: 'none'});
          xmlStream.closeNode();
        }
        this.renderCategories(xmlStream, model);
        xmlStream.leafNode('c:smooth', {val: 0});
        break;
      case 'scatter':
        if (!chart.lines) {
          // markers only: no line between the points
          xmlStream.openNode('c:spPr');
          xmlStream.openNode('a:ln', {w: 19050});
          xmlStream.leafNode('a:noFill');
          xmlStream.closeNode();
          xmlStream.closeNode();
        }
        xmlStream.openNode('c:xVal');
        this.dataRef.render(xmlStream, model.xValues);
        xmlStream.closeNode();
        xmlStream.openNode('c:yVal');
        this.dataRef.render(xmlStream, model.yValues);
        xmlStream.closeNode();
        xmlStream.leafNode('c:smooth', {val: 0});
        break;
      default:
        // area, pie, doughnut
        this.renderCategories(xmlStream, model);
        break;
    }
    xmlStream.closeNode();
  }

  renderName(xmlStream, name) {
    if (!name) {
      return;
    }
    xmlStream.openNode('c:tx');
    if (name.text !== undefined) {
      xmlStream.leafNode('c:v', undefined, name.text);
    } else {
      this.dataRef.render(xmlStream, name);
    }
    xmlStream.closeNode();
  }

  renderCategories(xmlStream, model) {
    if (model.categories) {
      xmlStream.openNode('c:cat');
      this.dataRef.render(xmlStream, model.categories);
      xmlStream.closeNode();
    }
    xmlStream.openNode('c:val');
    this.dataRef.render(xmlStream, model.values);
    xmlStream.closeNode();
  }
}

module.exports = ChartSeriesXform;
