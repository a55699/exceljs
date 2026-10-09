const colCache = require('../../../utils/col-cache');
const XmlStream = require('../../../utils/xml-stream');

const BaseXform = require('../base-xform');
const TwoCellAnchorXform = require('./two-cell-anchor-xform');
const OneCellAnchorXform = require('./one-cell-anchor-xform');
const RawAnchorXform = require('./raw-anchor-xform');

function getAnchorType(model) {
  const range = typeof model.range === 'string' ? colCache.decode(model.range) : model.range;

  return range.br ? 'xdr:twoCellAnchor' : 'xdr:oneCellAnchor';
}

class DrawingXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'xdr:twoCellAnchor': new TwoCellAnchorXform(),
      'xdr:oneCellAnchor': new OneCellAnchorXform(),
    };
    // every element of the drawing is also recorded, to keep the ones that
    // are not pictures (charts, shapes, groups...) as read
    this.rawAnchor = new RawAnchorXform();
  }

  prepare(model) {
    // shape ids are unique in a drawing: number new pictures and charts after
    // the ids of the anchors kept as read
    const keptIds = model.anchors
      .filter(item => item.xml)
      .map(item =>
        [...item.xml.matchAll(/<xdr:cNvPr\s[^>]*?\bid="(\d+)"/g)].map(match => +match[1])
      );
    const firstId = Math.max(0, ...[].concat(...keptIds));
    model.anchors.forEach((item, index) => {
      if (item.xml) {
        return;
      }
      item.anchorType = getAnchorType(item);
      const anchor = this.map[item.anchorType];
      anchor.prepare(item, {index: firstId + index});
    });
  }

  get tag() {
    return 'xdr:wsDr';
  }

  render(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    // kept anchors may use namespaces declared on the drawing they came from
    xmlStream.openNode(this.tag, {...model.namespaces, ...DrawingXform.DRAWING_ATTRIBUTES});

    model.anchors.forEach(item => {
      if (item.xml) {
        xmlStream.writeXml(item.xml);
        return;
      }
      const anchor = this.map[item.anchorType];
      anchor.render(xmlStream, item);
    });

    xmlStream.closeNode();
  }

  parseOpen(node) {
    // a drawing part that is not a worksheet drawing (e.g. c:userShapes, the
    // shapes drawn over a chart) has no anchors to read
    if (this.otherDepth || (!this.model && node.name !== this.tag)) {
      this.otherDepth = (this.otherDepth || 0) + 1;
      this.model = this.model || {anchors: []};
      return true;
    }
    if (this.inAnchor) {
      if (this.parser) {
        this.parser.parseOpen(node);
      }
      this.rawAnchor.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          anchors: [],
        };
        this.namespaces = {};
        Object.keys(node.attributes).forEach(name => {
          if (name.startsWith('xmlns:')) {
            this.namespaces[name] = node.attributes[name];
          }
        });
        break;
      default:
        this.inAnchor = true;
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        this.rawAnchor.parseOpen(node);
        break;
    }
    return true;
  }

  parseText(text) {
    if (this.inAnchor) {
      if (this.parser) {
        this.parser.parseText(text);
      }
      this.rawAnchor.parseText(text);
    }
  }

  parseClose(name) {
    if (this.otherDepth) {
      this.otherDepth--;
      return this.otherDepth > 0;
    }
    if (this.inAnchor) {
      if (this.parser && !this.parser.parseClose(name)) {
        this.anchor = this.parser.model;
        this.parser = undefined;
      }
      if (!this.rawAnchor.parseClose(name)) {
        // pictures are modelled; anything else is kept as read
        if (this.anchor && this.anchor.picture) {
          this.model.anchors.push(this.anchor);
        } else {
          this.model.anchors.push(this.rawAnchor.model);
          this.model.namespaces = this.namespaces;
        }
        this.anchor = undefined;
        this.inAnchor = false;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(model, options) {
    model.anchors.forEach(anchor => {
      if (anchor.xml) {
        return;
      }
      if (anchor.br) {
        this.map['xdr:twoCellAnchor'].reconcile(anchor, options);
      } else {
        this.map['xdr:oneCellAnchor'].reconcile(anchor, options);
      }
    });
  }
}

DrawingXform.DRAWING_ATTRIBUTES = {
  'xmlns:xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
};

module.exports = DrawingXform;
