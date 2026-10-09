const BaseXform = require('../base-xform');
const colCache = require('../../../utils/col-cache');

class DefinedNamesXform extends BaseXform {
  render(xmlStream, model) {
    // <definedNames>
    //   <definedName name="name">name.ranges.join(',')</definedName>
    //   <definedName name="_xlnm.Print_Area" localSheetId="0">name.ranges.join(',')</definedName>
    // </definedNames>
    xmlStream.openNode('definedName', {
      name: model.name,
      localSheetId: model.localSheetId,
      ...model.attributes,
    });
    xmlStream.writeText(model.formula !== undefined ? model.formula : model.ranges.join(','));
    xmlStream.closeNode();
  }

  parseOpen(node) {
    switch (node.name) {
      case 'definedName': {
        const {name, localSheetId, ...attributes} = node.attributes;
        this._parsedName = name;
        this._parsedLocalSheetId = localSheetId;
        this._parsedAttributes = attributes;
        this._parsedText = [];
        return true;
      }
      default:
        return false;
    }
  }

  parseText(text) {
    this._parsedText.push(text);
  }

  parseClose() {
    const text = this._parsedText.join('');
    this.model = {
      name: this._parsedName,
      ranges: extractRanges(text),
    };
    if (this._parsedLocalSheetId !== undefined) {
      this.model.localSheetId = parseInt(this._parsedLocalSheetId, 10);
    }
    // a formula or a constant, which is not only cell ranges, is kept as text
    if (!isRangeList(text, this.model.ranges)) {
      this.model.formula = text;
      this.model.ranges = [];
    }
    // hidden, comment...
    if (Object.keys(this._parsedAttributes).length) {
      this.model.attributes = this._parsedAttributes;
    }
    return false;
  }
}

// Whether the text of a name is only the ranges read from it, like
// 'My Sheet'!$A$1:$B$2,Sheet2!$C$3; not a formula like SUM(A!$A$1:$B$1) or a
// constant like 0.05
function isRangeList(text, ranges) {
  if (!ranges.length || ranges.join(',') !== text) {
    return false;
  }
  return ranges.every(range =>
    /^('(?:[^']|'')+'|[^'!(),\s"]+)!\$?[A-Z]*\$?\d*(:\$?[A-Z]*\$?\d*)?$/i.test(range)
  );
}

function isValidRange(range) {
  try {
    colCache.decodeEx(range);
    return true;
  } catch (err) {
    return false;
  }
}

function extractRanges(parsedText) {
  const ranges = [];
  let quotesOpened = false;
  let last = '';
  parsedText.split(',').forEach(item => {
    if (!item) {
      return;
    }
    const quotes = (item.match(/'/g) || []).length;

    if (!quotes) {
      if (quotesOpened) {
        last += `${item},`;
      } else if (isValidRange(item)) {
        ranges.push(item);
      }
      return;
    }
    const quotesEven = quotes % 2 === 0;

    if (!quotesOpened && quotesEven && isValidRange(item)) {
      ranges.push(item);
    } else if (quotesOpened && !quotesEven) {
      quotesOpened = false;
      if (isValidRange(last + item)) {
        ranges.push(last + item);
      }
      last = '';
    } else {
      quotesOpened = true;
      last += `${item},`;
    }
  });
  return ranges;
}

module.exports = DefinedNamesXform;
