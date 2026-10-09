const BaseXform = require('../base-xform');
const utils = require('../../../utils/utils');

// Relationship attributes in DrawingML (r:id on a chart, r:embed and r:link
// on a picture, r:dm etc. on a diagram)
const REL_ATTRIBUTE = /^r:(id|embed|link|pict|dm|lo|qs|cs)$/;

function renderAttributes(attributes) {
  return Object.keys(attributes)
    .map(name => ` ${name}="${utils.xmlEncode(attributes[name])}"`)
    .join('');
}

// Records one element of a drawing (an anchor ExcelJS does not model, such as
// a chart, a shape or a group) as XML, so it can be written back as read, and
// collects the relationship ids it uses.
class RawAnchorXform extends BaseXform {
  parseOpen(node) {
    if (!this.names) {
      this.model = {xml: '', rIds: []};
      this.parts = [];
      this.names = [];
      this.leaf = false;
    }
    Object.keys(node.attributes).forEach(name => {
      if (REL_ATTRIBUTE.test(name)) {
        this.model.rIds.push(node.attributes[name]);
      }
    });
    this.names.push(node.name);
    this.parts.push(`<${node.name}${renderAttributes(node.attributes)}>`);
    this.leaf = true;
    return true;
  }

  parseText(text) {
    this.parts.push(utils.xmlEncode(text));
    this.leaf = false;
  }

  parseClose() {
    const name = this.names.pop();
    if (this.leaf) {
      const last = this.parts.length - 1;
      this.parts[last] = `${this.parts[last].slice(0, -1)}/>`;
    } else {
      this.parts.push(`</${name}>`);
    }
    this.leaf = false;
    if (this.names.length) {
      return true;
    }
    this.model.xml = this.parts.join('');
    this.names = undefined;
    return false;
  }
}

// Give the relationship attributes of a recorded anchor new ids
RawAnchorXform.renameRIds = (xml, ids) =>
  xml.replace(/ (r:(?:id|embed|link|pict|dm|lo|qs|cs))="([^"]*)"/g, (match, name, id) =>
    ids[id] ? ` ${name}="${ids[id]}"` : match
  );

module.exports = RawAnchorXform;
