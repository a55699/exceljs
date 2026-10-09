const colCache = require('./col-cache');
const {slideFormula} = require('./shared-formula');

// Structured references name a part of a table in a formula, like
// Sales[Qty], Sales[[#This Row],[Qty]], Sales[#All] or, in a cell of the
// table itself, [@Qty]. Excel writes them in files as Sales[[#This Row],[Qty]].
// These helpers find the references to one table in a formula so they can be
// renamed or turned into cell references.

const NAME_CHAR = /[\p{L}\p{N}_.\\]/u;
const NAME_START = /[\p{L}_\\]/u;

const ITEMS = ['#all', '#data', '#headers', '#totals', '#this row'];

// Read a bracketed part starting at formula[start] === '['; returns the text
// between the outer brackets and the index after the closing bracket. Inside
// the brackets ' escapes the next character.
function readBracket(formula, start) {
  let depth = 0;
  for (let i = start; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === "'") {
      i++;
    } else if (ch === '[') {
      depth++;
    } else if (ch === ']') {
      depth--;
      if (depth === 0) {
        return {inner: formula.slice(start + 1, i), end: i + 1};
      }
    }
  }
  return null;
}

function unescapeName(text) {
  return text.replace(/'(.)/g, '$1');
}

function escapeName(name) {
  return name.replace(/['[\]#]/g, "'$&");
}

// Parse what is between the brackets of a structured reference into
// {items, columns, thisRow}: items are the special items in lower case,
// columns is [first, last] or null for all columns
function parseSpecifier(inner) {
  let text = inner;
  let thisRow = false;
  if (text[0] === '@') {
    thisRow = true;
    text = text.slice(1);
  }
  if (text.trim() === '') {
    return {items: thisRow ? ['#this row'] : [], columns: null};
  }
  if (text[0] !== '[') {
    const name = unescapeName(text);
    if (name[0] === '#') {
      const item = name.toLowerCase();
      return thisRow || !ITEMS.includes(item) ? null : {items: [item], columns: null};
    }
    return {items: thisRow ? ['#this row'] : [], columns: [name, name]};
  }

  const items = thisRow ? ['#this row'] : [];
  let columns = null;
  let i = 0;
  let separator = ',';
  while (i < text.length) {
    while (text[i] === ' ') i++;
    let part;
    if (text[i] === '[') {
      part = readBracket(text, i);
    } else {
      // a column name without brackets, as in [[#This Row],Qty]
      let j = i;
      while (j < text.length && text[j] !== ',' && text[j] !== ':') {
        j += text[j] === "'" ? 2 : 1;
      }
      part = {inner: text.slice(i, j), end: j};
    }
    if (!part || !part.inner) {
      return null;
    }
    const name = unescapeName(part.inner);
    if (name[0] === '#') {
      const item = name.toLowerCase();
      if (!ITEMS.includes(item)) {
        return null;
      }
      items.push(item);
    } else if (separator === ':' && columns) {
      columns[1] = name;
    } else {
      columns = [name, name];
    }
    i = part.end;
    while (text[i] === ' ') i++;
    if (i < text.length) {
      separator = text[i];
      if (separator !== ',' && separator !== ':') {
        return null;
      }
      i++;
    }
  }
  return {items, columns};
}

const ITEMS_DISPLAY = {
  '#all': '#All',
  '#data': '#Data',
  '#headers': '#Headers',
  '#totals': '#Totals',
  '#this row': '#This Row',
};

// Write a specifier back, without the outer brackets, in the form Excel uses
// in files: Qty or [#This Row],[Qty]
function renderSpecifier({items, columns}) {
  if (
    !items.length &&
    columns &&
    columns[0] === columns[1] &&
    /^[^\s'[\]#@,:]+$/.test(columns[0])
  ) {
    return columns[0];
  }
  const parts = items.map(item => `[${ITEMS_DISPLAY[item]}]`);
  if (columns) {
    const [first, last] = columns;
    parts.push(
      first === last ? `[${escapeName(first)}]` : `[${escapeName(first)}]:[${escapeName(last)}]`
    );
  }
  return parts.join(',');
}

function quoteSheetName(name) {
  const plain =
    /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) &&
    !/^[A-Za-z]{1,3}[0-9]+$/.test(name) &&
    !/^[RrCc][0-9]*$/.test(name) &&
    !/^[Rr][0-9]*[Cc][0-9]*$/.test(name);
  return plain ? name : `'${name.replace(/'/g, "''")}'`;
}

// Turn a parsed specifier into a cell reference, as Excel's Convert to Range
// does: Data!$B$2:$B$4, or Data!$B2 for the row of the formula.
// geometry is {sheetName, top, left, headerRow, totalsRow, height, columns}
function toCellReference(spec, geometry, formulaRow) {
  const {sheetName, top, left, headerRow, totalsRow, height, columns} = geometry;

  let firstCol = left;
  let lastCol = left + columns.length - 1;
  if (spec.columns) {
    const index = name => columns.findIndex(column => column.toLowerCase() === name.toLowerCase());
    const a = index(spec.columns[0]);
    const b = index(spec.columns[1]);
    if (a < 0 || b < 0) {
      return '#REF!';
    }
    firstCol = left + Math.min(a, b);
    lastCol = left + Math.max(a, b);
  }

  const sheet = `${quoteSheetName(sheetName)}!`;
  const colRange = (rowText1, rowText2) => {
    const tl = `$${colCache.n2l(firstCol)}${rowText1}`;
    const br = `$${colCache.n2l(lastCol)}${rowText2}`;
    return tl === br ? `${sheet}${tl}` : `${sheet}${tl}:${br}`;
  };

  if (spec.items.includes('#this row')) {
    return colRange(formulaRow, formulaRow);
  }

  const dataTop = top + (headerRow ? 1 : 0);
  const dataBottom = dataTop + height - 1;
  const totals = dataBottom + 1;
  const rows = [];
  const items = spec.items.length ? spec.items : ['#data'];
  for (const item of items) {
    switch (item) {
      case '#all':
        rows.push(top, totalsRow ? totals : dataBottom);
        break;
      case '#data':
        if (height < 1) return '#REF!';
        rows.push(dataTop, dataBottom);
        break;
      case '#headers':
        if (!headerRow) return '#REF!';
        rows.push(top);
        break;
      case '#totals':
        if (!totalsRow) return '#REF!';
        rows.push(totals);
        break;
      default:
        return '#REF!';
    }
  }
  return colRange(`$${Math.min(...rows)}`, `$${Math.max(...rows)}`);
}

// Call replace(spec, qualified, inner) for each structured reference to the
// table in the formula, inner being the text between its outer brackets; it
// returns the text to put instead, or undefined to keep the reference. References without a table name ([@Qty]) are only looked at when
// inTable is true, that is when the formula is in a cell of the table.
function mapTableReferences(formula, tableName, inTable, replace) {
  const lowerName = tableName.toLowerCase();
  let out = '';
  let i = 0;
  const copyQuoted = quote => {
    let j = i + 1;
    while (j < formula.length) {
      if (formula[j] === quote) {
        if (formula[j + 1] !== quote) break;
        j++;
      }
      j++;
    }
    out += formula.slice(i, j + 1);
    i = j + 1;
  };
  const handleBracket = (start, qualified) => {
    const bracket = readBracket(formula, start);
    if (!bracket) {
      out += formula.slice(i);
      i = formula.length;
      return;
    }
    const spec = parseSpecifier(bracket.inner);
    const text = spec && replace(spec, qualified, bracket.inner);
    out += text === undefined || text === null ? formula.slice(i, bracket.end) : text;
    i = bracket.end;
  };

  while (i < formula.length) {
    const ch = formula[i];
    const prev = i > 0 ? formula[i - 1] : '';
    if (ch === '"' || ch === "'") {
      copyQuoted(ch);
    } else if (NAME_START.test(ch) && !NAME_CHAR.test(prev)) {
      let j = i;
      while (j < formula.length && NAME_CHAR.test(formula[j])) j++;
      const name = formula.slice(i, j);
      if (formula[j] === '[' && prev !== '!' && name.toLowerCase() === lowerName) {
        handleBracket(j, true);
      } else if (formula[j] === '[') {
        // another table: keep the whole reference
        const bracket = readBracket(formula, j);
        const end = bracket ? bracket.end : formula.length;
        out += formula.slice(i, end);
        i = end;
      } else {
        out += name;
        i = j;
      }
    } else if (ch === '[' && inTable) {
      handleBracket(i, false);
    } else if (ch === '[') {
      // an external workbook index like [1]Sheet1!A1
      const bracket = readBracket(formula, i);
      const end = bracket ? bracket.end : formula.length;
      out += formula.slice(i, end);
      i = end;
    } else {
      out += ch;
      i++;
    }
  }
  return out;
}

// Move a formula from one cell to another like a filled formula: the relative
// cell references move with it. Structured references, texts and quoted sheet
// names are kept as they are.
function moveFormula(formula, fromCell, toCell) {
  if (fromCell === toCell) {
    return formula;
  }
  let out = '';
  let plain = '';
  const flush = () => {
    out += plain ? slideFormula(plain, fromCell, toCell) : '';
    plain = '';
  };
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i];
    let end;
    if (ch === '"' || ch === "'") {
      end = i + 1;
      while (end < formula.length) {
        if (formula[end] === ch) {
          if (formula[end + 1] !== ch) break;
          end++;
        }
        end++;
      }
      end++;
    } else if (ch === '[') {
      const bracket = readBracket(formula, i);
      end = bracket ? bracket.end : formula.length;
    }
    if (end === undefined) {
      plain += ch;
      i++;
    } else {
      flush();
      out += formula.slice(i, end);
      i = end;
    }
  }
  flush();
  return out;
}

module.exports = {
  moveFormula,
  mapTableReferences,
  parseSpecifier,
  renderSpecifier,
  toCellReference,
  quoteSheetName,
};
