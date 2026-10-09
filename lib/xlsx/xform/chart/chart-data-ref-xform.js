const BaseXform = require('../base-xform');

// A reference to cells with the values cached from them, as Excel writes it:
// <c:numRef> (with a format code) or <c:strRef>. Points without a value are
// left out of the cache. Write only.
class ChartDataRefXform extends BaseXform {
  render(xmlStream, model) {
    const num = model.type === 'num';
    xmlStream.openNode(num ? 'c:numRef' : 'c:strRef');
    xmlStream.leafNode('c:f', undefined, model.formula);
    xmlStream.openNode(num ? 'c:numCache' : 'c:strCache');
    if (num) {
      xmlStream.leafNode('c:formatCode', undefined, model.formatCode || 'General');
    }
    xmlStream.leafNode('c:ptCount', {val: model.cache.length});
    model.cache.forEach((value, idx) => {
      if (value !== null && value !== undefined) {
        xmlStream.openNode('c:pt', {idx});
        xmlStream.leafNode('c:v', undefined, value);
        xmlStream.closeNode();
      }
    });
    xmlStream.closeNode();
    xmlStream.closeNode();
  }
}

module.exports = ChartDataRefXform;
