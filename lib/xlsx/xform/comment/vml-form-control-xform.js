const BaseXform = require('../base-xform');

/** EMU (English Metric Units) per pixel at 96 DPI; VML anchors are in pixels */
const EMU_PER_PIXEL = 9525;

const CHECKED_STATES = {1: 'Checked', 2: 'Mixed'};

// Reads the form control parts of a VML <v:shape>: the <x:ClientData> object
// type and settings, and the label text in <v:textbox>. The model matches the
// FormCheckbox model so a loaded checkbox is written back the same way.
// Note: the SAX parser strips the 'x:' prefix from Excel namespace elements.
class VmlFormControlXform extends BaseXform {
  get tag() {
    return 'v:shape';
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {objectType: undefined};
        this.lines = [''];
        this.inTextbox = false;
        this.inClientData = false;
        this.clientData = {};
        this.text = undefined;
        break;
      case 'v:textbox':
        this.inTextbox = true;
        break;
      case 'br':
        if (this.inTextbox) {
          this.lines.push('');
        }
        break;
      case 'ClientData':
        this.inClientData = true;
        this.model.objectType = node.attributes.ObjectType;
        break;
      default:
        if (this.inClientData) {
          this.clientData[node.name] = '';
          this.text = node.name;
        }
        break;
    }
    return true;
  }

  parseText(text) {
    if (this.inTextbox) {
      this.lines[this.lines.length - 1] += text;
    } else if (this.text) {
      this.clientData[this.text] += text;
    }
  }

  parseClose(name) {
    switch (name) {
      case this.tag:
        // Excel wraps long labels over several lines with indentation;
        // only <br/> is a line break in the label
        this.model.text = this.lines
          .map(line => line.replace(/\s*\n\s*/g, ' '))
          .join('\n')
          .trim();
        return false;
      case 'v:textbox':
        this.inTextbox = false;
        return true;
      case 'ClientData':
        this.inClientData = false;
        this.normalizeModel();
        return true;
      default:
        this.text = undefined;
        return true;
    }
  }

  normalizeModel() {
    const {Anchor, FmlaLink, Checked, NoThreeD, PrintObject} = this.clientData;
    if (Anchor) {
      const anchor = Anchor.split(',').map(value => parseInt(value, 10));
      const [col, colOff, row, rowOff, endCol, endColOff, endRow, endRowOff] = anchor;
      this.model.tl = {col, colOff: colOff * EMU_PER_PIXEL, row, rowOff: rowOff * EMU_PER_PIXEL};
      this.model.br = {
        col: endCol,
        colOff: endColOff * EMU_PER_PIXEL,
        row: endRow,
        rowOff: endRowOff * EMU_PER_PIXEL,
      };
    }
    if (FmlaLink && FmlaLink.trim()) {
      this.model.link = FmlaLink.trim();
    }
    this.model.checked = CHECKED_STATES[parseInt(Checked, 10)] || 'Unchecked';
    this.model.noThreeD = NoThreeD !== undefined;
    // form controls are printed unless <x:PrintObject>False</x:PrintObject>
    this.model.print = PrintObject === undefined || PrintObject.trim() !== 'False';
  }
}

module.exports = VmlFormControlXform;
