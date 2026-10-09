const BaseXform = require('../base-xform');

// formulas kept as child elements of a table column
const FORMULAS = ['calculatedColumnFormula', 'totalsRowFormula'];

class TableColumnXform extends BaseXform {
  get tag() {
    return 'tableColumn';
  }

  prepare(model, options) {
    model.id = options.index + 1;
  }

  render(xmlStream, model) {
    const attributes = {
      id: model.id.toString(),
      name: model.name,
      totalsRowLabel: model.totalsRowLabel,
      totalsRowFunction: model.totalsRowFunction,
      dxfId: model.dxfId,
    };
    const {calculatedColumnFormula} = model;
    // the formula of a custom total; for the other functions totalsRowFormula
    // is only the formula of the totals row cell
    const totalsRowFormula =
      model.totalsRowFunction === 'custom' ? model.totalsRowFormula : undefined;
    if (!calculatedColumnFormula && !totalsRowFormula) {
      xmlStream.leafNode(this.tag, attributes);
      return true;
    }
    xmlStream.openNode(this.tag, attributes);
    if (calculatedColumnFormula) {
      xmlStream.leafNode('calculatedColumnFormula', undefined, calculatedColumnFormula);
    }
    if (totalsRowFormula) {
      xmlStream.leafNode('totalsRowFormula', undefined, totalsRowFormula);
    }
    xmlStream.closeNode();
    return true;
  }

  parseOpen(node) {
    if (node.name === this.tag) {
      const {attributes} = node;
      this.model = {
        name: attributes.name,
        totalsRowLabel: attributes.totalsRowLabel,
        totalsRowFunction: attributes.totalsRowFunction,
        dxfId: attributes.dxfId,
      };
      this.inColumn = true;
      return true;
    }
    if (!this.inColumn) {
      return false;
    }
    if (FORMULAS.includes(node.name)) {
      this.formula = {name: node.name, text: []};
    }
    return true;
  }

  parseText(text) {
    if (this.formula) {
      this.formula.text.push(text);
    }
  }

  parseClose(name) {
    if (name === this.tag) {
      this.inColumn = false;
      return false;
    }
    if (this.formula && name === this.formula.name) {
      this.model[name] = this.formula.text.join('');
      this.formula = undefined;
    }
    return true;
  }
}

module.exports = TableColumnXform;
