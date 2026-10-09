const BaseXform = require('../base-xform');
const utils = require('../../../utils/utils');

function renderAttributes(attributes) {
  return Object.keys(attributes)
    .map(name => ` ${name}="${utils.xmlEncode(attributes[name])}"`)
    .join('');
}

// Records the XML inside a VML <v:shape>, so a shape ExcelJS does not model
// (a button, drop-down or other form control) can be written back unchanged.
// The SAX parser strips the 'x:' prefix from Excel namespace elements; it is
// put back on <ClientData> and its children, which are all in that namespace.
class VmlShapeXmlXform extends BaseXform {
  get tag() {
    return 'v:shape';
  }

  parseOpen(node) {
    if (node.name === this.tag && !this.names) {
      this.model = {attributes: {...node.attributes}, xml: ''};
      this.parts = [];
      this.names = [];
      this.inClientData = false;
      this.leaf = false;
      return true;
    }
    if (node.name === 'ClientData') {
      this.inClientData = true;
    }
    const name = this.inClientData ? `x:${node.name}` : node.name;
    this.names.push(name);
    this.parts.push(`<${name}${renderAttributes(node.attributes)}>`);
    this.leaf = true;
    return true;
  }

  parseText(text) {
    this.parts.push(utils.xmlEncode(text));
    this.leaf = false;
  }

  parseClose(name) {
    if (!this.names.length) {
      this.model.xml = this.parts.join('');
      this.names = undefined;
      return false;
    }
    const xmlName = this.names.pop();
    if (this.leaf) {
      const last = this.parts.length - 1;
      this.parts[last] = `${this.parts[last].slice(0, -1)}/>`;
    } else {
      this.parts.push(`</${xmlName}>`);
    }
    if (name === 'ClientData') {
      this.inClientData = false;
    }
    this.leaf = false;
    return true;
  }
}

module.exports = VmlShapeXmlXform;
