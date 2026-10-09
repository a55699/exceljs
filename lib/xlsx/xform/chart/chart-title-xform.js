const BaseXform = require('../base-xform');

// <c:title> of a chart or an axis: the text as rich text, one paragraph per
// line. Write only.
class ChartTitleXform extends BaseXform {
  get tag() {
    return 'c:title';
  }

  render(xmlStream, text) {
    if (text === undefined || text === null || text === '') {
      return;
    }
    xmlStream.openNode(this.tag);
    xmlStream.openNode('c:tx');
    xmlStream.openNode('c:rich');
    xmlStream.leafNode('a:bodyPr');
    xmlStream.leafNode('a:lstStyle');
    String(text)
      .split('\n')
      .forEach(line => {
        xmlStream.openNode('a:p');
        xmlStream.openNode('a:pPr');
        xmlStream.leafNode('a:defRPr');
        xmlStream.closeNode();
        xmlStream.openNode('a:r');
        xmlStream.leafNode('a:t', undefined, line);
        xmlStream.closeNode();
        xmlStream.closeNode();
      });
    xmlStream.closeNode(); // c:rich
    xmlStream.closeNode(); // c:tx
    xmlStream.leafNode('c:overlay', {val: 0});
    xmlStream.closeNode();
  }
}

module.exports = ChartTitleXform;
