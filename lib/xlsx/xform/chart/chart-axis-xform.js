const BaseXform = require('../base-xform');
const ChartTitleXform = require('./chart-title-xform');

// <c:catAx> or <c:valAx>. Write only.
// model: {kind: 'cat' | 'val', id, crossId, position, crossBetween, axis},
// where axis holds the user options {title, min, max, numFmt, gridlines}.
class ChartAxisXform extends BaseXform {
  constructor() {
    super();
    this.title = new ChartTitleXform();
  }

  render(xmlStream, model) {
    const {kind, id, crossId, position, crossBetween, gridlines, axis} = model;
    xmlStream.openNode(kind === 'cat' ? 'c:catAx' : 'c:valAx');
    xmlStream.leafNode('c:axId', {val: id});
    xmlStream.openNode('c:scaling');
    xmlStream.leafNode('c:orientation', {val: 'minMax'});
    if (axis.max !== undefined) {
      xmlStream.leafNode('c:max', {val: axis.max});
    }
    if (axis.min !== undefined) {
      xmlStream.leafNode('c:min', {val: axis.min});
    }
    xmlStream.closeNode();
    // without delete="0" Excel hides the axis labels
    xmlStream.leafNode('c:delete', {val: 0});
    xmlStream.leafNode('c:axPos', {val: position});
    if (axis.gridlines === undefined ? gridlines : axis.gridlines) {
      xmlStream.leafNode('c:majorGridlines');
    }
    this.title.render(xmlStream, axis.title);
    xmlStream.leafNode('c:numFmt', {
      formatCode: axis.numFmt || 'General',
      sourceLinked: axis.numFmt ? 0 : 1,
    });
    xmlStream.leafNode('c:majorTickMark', {val: 'out'});
    xmlStream.leafNode('c:minorTickMark', {val: 'none'});
    xmlStream.leafNode('c:tickLblPos', {val: 'nextTo'});
    xmlStream.leafNode('c:crossAx', {val: crossId});
    xmlStream.leafNode('c:crosses', {val: 'autoZero'});
    if (kind === 'cat') {
      xmlStream.leafNode('c:auto', {val: 1});
      xmlStream.leafNode('c:lblAlgn', {val: 'ctr'});
      xmlStream.leafNode('c:lblOffset', {val: 100});
      xmlStream.leafNode('c:noMultiLvlLbl', {val: 0});
    } else {
      xmlStream.leafNode('c:crossBetween', {val: crossBetween});
    }
    xmlStream.closeNode();
  }
}

module.exports = ChartAxisXform;
