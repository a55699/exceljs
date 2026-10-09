const XmlStream = require('../../../utils/xml-stream');
const BaseXform = require('../base-xform');
const ChartTitleXform = require('./chart-title-xform');
const ChartPlotAreaXform = require('./chart-plot-area-xform');

// xl/charts/chartN.xml for a chart added with Worksheet#addChart. Excel's own
// extras (c14:style, extLst, printSettings) are not needed and not written.
// Write only: charts in loaded files are kept as read.
class ChartSpaceXform extends BaseXform {
  constructor() {
    super();
    this.title = new ChartTitleXform();
    this.plotArea = new ChartPlotAreaXform();
  }

  get tag() {
    return 'c:chartSpace';
  }

  render(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, ChartSpaceXform.CHART_SPACE_ATTRIBUTES);
    // without it Excel draws rounded corners
    xmlStream.leafNode('c:roundedCorners', {val: 0});

    xmlStream.openNode('c:chart');
    this.title.render(xmlStream, model.title);
    // with no title, Excel would show the series name of a one-series chart
    xmlStream.leafNode('c:autoTitleDeleted', {val: model.title ? 0 : 1});
    this.plotArea.render(xmlStream, model);
    if (model.legend) {
      xmlStream.openNode('c:legend');
      xmlStream.leafNode('c:legendPos', {val: model.legend.position});
      xmlStream.leafNode('c:overlay', {val: 0});
      xmlStream.closeNode();
    }
    xmlStream.leafNode('c:plotVisOnly', {val: 1});
    xmlStream.leafNode('c:dispBlanksAs', {val: 'gap'});
    xmlStream.closeNode();

    xmlStream.closeNode();
  }
}

ChartSpaceXform.CHART_SPACE_ATTRIBUTES = {
  'xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
  'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

module.exports = ChartSpaceXform;
