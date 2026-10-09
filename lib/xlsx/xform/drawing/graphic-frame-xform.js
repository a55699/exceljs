const BaseXform = require('../base-xform');

// <xdr:graphicFrame> holding a chart added with Worksheet#addChart.
// model: {rId, index, name}. Write only: loaded charts are kept as read.
class GraphicFrameXform extends BaseXform {
  get tag() {
    return 'xdr:graphicFrame';
  }

  prepare(model, options) {
    model.index = options.index + 1;
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {macro: ''});
    xmlStream.openNode('xdr:nvGraphicFramePr');
    xmlStream.leafNode('xdr:cNvPr', {id: model.index, name: model.name});
    xmlStream.leafNode('xdr:cNvGraphicFramePr');
    xmlStream.closeNode();
    xmlStream.openNode('xdr:xfrm');
    xmlStream.leafNode('a:off', {x: 0, y: 0});
    xmlStream.leafNode('a:ext', {cx: 0, cy: 0});
    xmlStream.closeNode();
    xmlStream.openNode('a:graphic');
    xmlStream.openNode('a:graphicData', {
      uri: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    });
    xmlStream.leafNode('c:chart', {
      'xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'r:id': model.rId,
    });
    xmlStream.closeNode();
    xmlStream.closeNode();
    xmlStream.closeNode();
  }
}

module.exports = GraphicFrameXform;
