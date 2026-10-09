/* eslint-disable max-classes-per-file */
const colCache = require('../utils/col-cache');
const utils = require('../utils/utils');
const Enums = require('./enums');
const {
  mapTableReferences,
  moveFormula,
  renderSpecifier,
  toCellReference,
} = require('../utils/table-ref');

// Excel opens a file only if a table name starts with a letter, _ or \, has
// only letters, digits, _, . and \, is not like a cell reference (A1, R1C1)
// and is used by no other table of the workbook
function checkTableName(name, worksheet, except) {
  const valid =
    typeof name === 'string' &&
    name.length <= 255 &&
    /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(name) &&
    !/^[A-Za-z]{1,3}[0-9]+$/.test(name) &&
    !/^[RrCc]$/.test(name) &&
    !/^[Rr][0-9]*[Cc][0-9]*$/.test(name);
  if (!valid) {
    throw new Error(
      `Table name '${name}' is not valid: it must start with a letter, _ or \\, ` +
        'have no spaces and not be like a cell reference'
    );
  }
  const worksheets = worksheet && worksheet.workbook ? worksheet.workbook.worksheets : [];
  worksheets.forEach(ws => {
    Object.values(ws.tables).forEach(table => {
      if (table !== except && table.name.toLowerCase() === name.toLowerCase()) {
        throw new Error(`Table name '${name}' is already used`);
      }
    });
  });
}

// The value to keep for a cell of a loaded table; shared formulas become plain
// formulas so the rows can be moved
function loadedValue(cell) {
  if (!cell) {
    return null;
  }
  const {value} = cell;
  if (cell.type !== Enums.ValueType.Formula) {
    return value;
  }
  if (value.sharedFormula) {
    // the formula of a shared formula comes from its first cell
    const master = cell.worksheet.findCell(value.sharedFormula);
    if (!master || !master.value || !master.value.formula) {
      return value;
    }
    return {formula: cell.formula, result: cell.result};
  }
  if (value.shareType === 'shared') {
    return {formula: value.formula, result: value.result};
  }
  return value;
}

class Column {
  // wrapper around column model, allowing access and manipulation
  constructor(table, column, index) {
    this.table = table;
    this.column = column;
    this.index = index;
  }

  _set(name, value) {
    this.table.cacheState();
    this.column[name] = value;
  }

  /* eslint-disable lines-between-class-members */
  get name() {
    return this.column.name;
  }
  set name(value) {
    const oldName = this.column.name;
    const name = String(value);
    this.table._checkColumnName(name, this.column);
    this._set('name', name);
    this.table._renameColumnReferences(oldName, name);
  }

  get filterButton() {
    return this.column.filterButton;
  }
  set filterButton(value) {
    this.column.filterButton = value;
  }

  get style() {
    return this.column.style;
  }
  set style(value) {
    this.column.style = value;
  }

  get totalsRowLabel() {
    return this.column.totalsRowLabel;
  }
  set totalsRowLabel(value) {
    this._set('totalsRowLabel', value);
  }

  get totalsRowFunction() {
    return this.column.totalsRowFunction;
  }
  set totalsRowFunction(value) {
    this._set('totalsRowFunction', value);
  }

  get totalsRowResult() {
    return this.column.totalsRowResult;
  }
  set totalsRowResult(value) {
    this._set('totalsRowResult', value);
  }

  get totalsRowFormula() {
    return this.column.totalsRowFormula;
  }
  set totalsRowFormula(value) {
    this._set('totalsRowFormula', value);
  }
  /* eslint-enable lines-between-class-members */
}

class Table {
  constructor(worksheet, table) {
    this.worksheet = worksheet;
    // the header cell of each column and the text last written to (or read
    // from) it, to see when a header cell was changed in the sheet
    this._headers = new WeakMap();
    if (table) {
      this.table = table;
      // check things are ok first
      this.validate();

      this.store();
    }
  }

  getFormula(column) {
    // get the correct formula to apply to the totals row
    switch (column.totalsRowFunction) {
      case 'none':
        return null;
      case 'average':
        return `SUBTOTAL(101,${this._columnReference(column)})`;
      case 'countNums':
        return `SUBTOTAL(102,${this._columnReference(column)})`;
      case 'count':
        return `SUBTOTAL(103,${this._columnReference(column)})`;
      case 'max':
        return `SUBTOTAL(104,${this._columnReference(column)})`;
      case 'min':
        return `SUBTOTAL(105,${this._columnReference(column)})`;
      case 'stdDev':
        return `SUBTOTAL(106,${this._columnReference(column)})`;
      case 'var':
        return `SUBTOTAL(107,${this._columnReference(column)})`;
      case 'sum':
        return `SUBTOTAL(109,${this._columnReference(column)})`;
      case 'custom':
        return column.totalsRowFormula;
      default:
        throw new Error(`Invalid Totals Row Function: ${column.totalsRowFunction}`);
    }
  }

  // A reference to the data of a column, like Sales[Qty] or Sales[[Price '[USD']]]
  _columnReference(column) {
    return `${this.table.name}[${renderSpecifier({items: [], columns: [column.name, column.name]})}]`;
  }

  get width() {
    // width of the table
    return this.table.columns.length;
  }

  get height() {
    // height of the table data; a table without rows has one empty row, as in
    // Excel, which cannot open a table without data rows
    return Math.max(this.table.rows.length, 1);
  }

  get filterHeight() {
    // height of the table data plus optional header row
    return this.height + (this.table.headerRow ? 1 : 0);
  }

  get tableHeight() {
    // full height of the table on the sheet
    return this.filterHeight + (this.table.totalsRow ? 1 : 0);
  }

  validate() {
    const {table} = this;
    // set defaults and check is valid
    const assign = (o, name, dflt) => {
      if (o[name] === undefined) {
        o[name] = dflt;
      }
    };
    assign(table, 'headerRow', true);
    assign(table, 'totalsRow', false);

    assign(table, 'style', {});
    assign(table.style, 'theme', 'TableStyleMedium2');
    assign(table.style, 'showFirstColumn', false);
    assign(table.style, 'showLastColumn', false);
    assign(table.style, 'showRowStripes', false);
    assign(table.style, 'showColumnStripes', false);

    const assert = (test, message) => {
      if (!test) {
        throw new Error(message);
      }
    };
    assert(table.ref, 'Table must have ref');
    assert(table.columns, 'Table must have column definitions');
    assert(table.rows, 'Table must have row definitions');

    table.tl = colCache.decodeAddress(table.ref);
    const {row, col} = table.tl;
    assert(row > 0, 'Table must be on valid row');
    assert(col > 0, 'Table must be on valid col');

    const {width, filterHeight, tableHeight} = this;

    // autoFilterRef covers the header row and the data rows, as Excel writes
    // it; null means the table has no filter buttons
    if (table.autoFilterRef !== null) {
      table.autoFilterRef = table.headerRow
        ? colCache.encode(row, col, row + filterHeight - 1, col + width - 1)
        : undefined;
    }

    // tableRef is a range that includes optional headers and totals
    table.tableRef = colCache.encode(row, col, row + tableHeight - 1, col + width - 1);
    this._checkOverlap();

    const names = new Set();
    table.columns.forEach((column, i) => {
      // a header is always text in Excel
      if (typeof column.name === 'number' || typeof column.name === 'boolean') {
        column.name = String(column.name);
      }
      assert(column.name, `Column ${i} must have a name`);
      assert(
        !names.has(column.name.toLowerCase()),
        `Column name '${column.name}' is used twice in table ${table.name}`
      );
      names.add(column.name.toLowerCase());
      // the first column shows Total unless it has a function
      if (i === 0 && column.totalsRowFunction === undefined) {
        assign(column, 'totalsRowLabel', 'Total');
      }
      assign(column, 'totalsRowFunction', 'none');
      column.totalsRowFormula = this.getFormula(column);
    });
  }

  // Excel cannot open two tables that share cells
  _checkOverlap() {
    const {worksheet, table} = this;
    if (!worksheet || !worksheet.tables) {
      return;
    }
    const a = colCache.decode(table.tableRef);
    Object.values(worksheet.tables).forEach(other => {
      if (other === this || !other.table || !other.table.tableRef) {
        return;
      }
      const b = colCache.decode(other.table.tableRef);
      const top = b.top || b.row;
      const left = b.left || b.col;
      const bottom = b.bottom || b.row;
      const right = b.right || b.col;
      if (a.top <= bottom && top <= a.bottom && a.left <= right && left <= a.right) {
        throw new Error(`Table ${table.name} overlaps table ${other.table.name}`);
      }
    });
  }

  _checkColumnName(name, except) {
    if (!name) {
      throw new Error(`A column of table ${this.table.name} must have a name`);
    }
    this.table.columns.forEach(column => {
      if (column !== except && String(column.name).toLowerCase() === name.toLowerCase()) {
        throw new Error(`Column name '${name}' is used twice in table ${this.table.name}`);
      }
    });
  }

  // A column name not used in the table, as Excel makes them: Column1, Qty2
  _uniqueColumnName(base, except) {
    const used = new Set(
      this.table.columns
        .filter(column => column !== except)
        .map(column => String(column.name).toLowerCase())
    );
    if (base && !used.has(base.toLowerCase())) {
      return base;
    }
    const stem = base || 'Column';
    let n = base ? 2 : 1;
    while (used.has(`${stem}${n}`.toLowerCase())) n++;
    return `${stem}${n}`;
  }

  // A header cell changed in the sheet renames its column, as in Excel
  _syncHeaders() {
    const {table, worksheet} = this;
    if (!table.headerRow || !table.tl) {
      return;
    }
    table.columns.forEach(column => {
      const header = this._headers.get(column);
      if (!header) {
        return;
      }
      const cell = worksheet.findCell(header.row, header.col);
      const text = cell ? cell.text : '';
      if (text === header.text) {
        return;
      }
      const oldName = column.name;
      const name = this._uniqueColumnName(text, column);
      column.name = name;
      this._renameColumnReferences(oldName, name);
      worksheet.getCell(header.row, header.col).value = name;
      this._headers.set(column, {...header, text: name});
    });
  }

  store() {
    // where the table needs to store table data, headers, footers in
    // the sheet...
    const assignStyle = (cell, style) => {
      if (style) {
        Object.keys(style).forEach(key => {
          cell.style[key] = style[key];
        });
      }
    };

    const {worksheet, table} = this;
    const {row, col} = table.tl;
    let count = 0;
    if (table.headerRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const {style, name} = column;
        const cell = r.getCell(col + j);
        cell.value = name;
        this._headers.set(column, {row: r.number, col: col + j, text: name});
        assignStyle(cell, style);
      });
    }
    const firstDataRow = row + count;
    const moved = this._cache || {};
    table.rows.forEach(data => {
      const rowNumber = row + count++;
      const r = worksheet.getRow(rowNumber);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        const address = colCache.encodeAddress(rowNumber, col + j);
        let value = data[j];
        if (value === undefined) {
          // a calculated column gets its formula in a new row, as in Excel
          const formula = column.calculatedColumnFormula;
          value = formula
            ? {
                formula: moveFormula(
                  formula,
                  colCache.encodeAddress(firstDataRow, col + j),
                  address
                ),
              }
            : null;
        } else if (
          value &&
          value.formula &&
          moved.rows &&
          moved.rows.has(data) &&
          moved.columns.has(column)
        ) {
          // a formula of a row or column that has moved moves with it
          const from = colCache.encodeAddress(moved.rows.get(data), moved.columns.get(column));
          value = {...value, formula: moveFormula(value.formula, from, address)};
          data[j] = value;
        }
        cell.value = value;

        assignStyle(cell, column.style);
      });
    });
    if (!table.rows.length) {
      // the empty row of a table without rows
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        r.getCell(col + j).value = null;
      });
    }

    if (table.totalsRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        if (this.getFormula(column)) {
          cell.value = {
            formula: column.totalsRowFormula,
            result: column.totalsRowResult,
          };
        } else {
          cell.value = column.totalsRowLabel === undefined ? null : column.totalsRowLabel;
        }

        assignStyle(cell, column.style);
      });
    }
  }

  load(worksheet) {
    // read a table of a loaded worksheet: the file keeps only the table's
    // range, so the data rows are read from the cells, which are not changed
    const {table} = this;
    const range = colCache.decode(table.tableRef);
    const top = range.top || range.row;
    const left = range.left || range.col;
    const bottom = range.bottom || range.row;

    table.ref = colCache.encodeAddress(top, left);
    table.rows = [];
    const first = top + (table.headerRow ? 1 : 0);
    const last = bottom - (table.totalsRow ? 1 : 0);
    for (let r = first; r <= last; r++) {
      table.rows.push(
        table.columns.map((column, j) => loadedValue(worksheet.findCell(r, left + j)))
      );
    }
    if (table.rows.length === 1 && table.rows[0].every(value => value === null)) {
      // the empty row of a table without rows
      table.rows = [];
    }
    table.columns.forEach((column, j) => {
      if (table.headerRow) {
        this._headers.set(column, {row: top, col: left + j, text: column.name});
      }
      // no Total label is added to a loaded table
      if (column.totalsRowFunction === undefined) {
        column.totalsRowFunction = 'none';
      }
      const cell = table.totalsRow && worksheet.findCell(bottom, left + j);
      if (cell && cell.type === Enums.ValueType.Formula) {
        column.totalsRowResult = cell.result;
      }
    });

    if (table.autoFilterRef === undefined) {
      // the table has no filter buttons
      table.autoFilterRef = null;
    } else {
      table.columns.forEach(column => {
        if (column.filterButton === undefined) {
          column.filterButton = true;
        }
      });
    }

    this.validate();
  }

  get model() {
    return this.table;
  }

  set model(value) {
    this.table = value;
  }

  // ================================================================
  // TODO: Mutating methods
  cacheState() {
    if (!this._cache) {
      // where each row and column is, to move their formulas with them
      const {row, col} = this.table.tl;
      const first = row + (this.table.headerRow ? 1 : 0);
      this._cache = {
        ref: this.ref,
        width: this.width,
        tableHeight: this.tableHeight,
        rows: new Map(this.table.rows.map((data, i) => [data, first + i])),
        columns: new Map(this.table.columns.map((column, j) => [column, col + j])),
      };
    }
  }

  commit() {
    this._syncHeaders();
    // changes may have been made that might have on-sheet effects
    if (!this._cache) {
      return;
    }

    // check things are ok first
    this.validate();

    const ref = colCache.decodeAddress(this._cache.ref);
    if (this.ref !== this._cache.ref) {
      // wipe out whole table footprint at previous location
      for (let i = 0; i < this._cache.tableHeight; i++) {
        const row = this.worksheet.getRow(ref.row + i);
        for (let j = 0; j < this._cache.width; j++) {
          const cell = row.getCell(ref.col + j);
          cell.value = null;
        }
      }
    } else {
      // clear out below table if it has shrunk
      for (let i = this.tableHeight; i < this._cache.tableHeight; i++) {
        const row = this.worksheet.getRow(ref.row + i);
        for (let j = 0; j < this._cache.width; j++) {
          const cell = row.getCell(ref.col + j);
          cell.value = null;
        }
      }

      // clear out to right of table if it has lost columns
      for (let i = 0; i < this.tableHeight; i++) {
        const row = this.worksheet.getRow(ref.row + i);
        for (let j = this.width; j < this._cache.width; j++) {
          const cell = row.getCell(ref.col + j);
          cell.value = null;
        }
      }
    }

    this.store();
    this._cache = undefined;
  }

  addRow(values, rowNumber) {
    // Add a row of data, either insert at rowNumber or append
    this.cacheState();

    if (rowNumber === undefined) {
      this.table.rows.push(values);
    } else {
      this.table.rows.splice(rowNumber, 0, values);
    }

    this.commit();
  }

  removeRows(rowIndex, count = 1) {
    // Remove a rows of data
    this.cacheState();
    this.table.rows.splice(rowIndex, count);

    this.commit();
  }

  // The size and place of the table, for turning references into cells
  get _geometry() {
    const {table} = this;
    return {
      sheetName: this.worksheet.name,
      top: table.tl.row,
      left: table.tl.col,
      headerRow: table.headerRow,
      totalsRow: table.totalsRow,
      height: this.height,
      columns: table.columns.map(column => column.name),
    };
  }

  _contains(worksheet, row, col) {
    const {top, left, headerRow, totalsRow, height, columns} = this._geometry;
    const bottom = top + height + (headerRow ? 1 : 0) + (totalsRow ? 1 : 0) - 1;
    return (
      worksheet === this.worksheet &&
      row >= top &&
      row <= bottom &&
      col >= left &&
      col < left + columns.length
    );
  }

  // Call update(formula, worksheet, row, col) for each formula of the workbook,
  // in the cells and in the rows and columns of its tables; update returns the
  // new formula or undefined to keep it
  _updateFormulas(update) {
    const {workbook} = this.worksheet;
    const worksheets = workbook ? workbook.worksheets : [this.worksheet];
    worksheets.forEach(worksheet => {
      worksheet.eachRow(row => {
        row.eachCell(cell => {
          if (cell.type !== Enums.ValueType.Formula) {
            return;
          }
          const {value} = cell;
          // a shared formula follows the formula of its first cell
          if (!value.formula) {
            return;
          }
          const formula = update(value.formula, worksheet, cell.row, cell.col);
          if (formula !== undefined && formula !== value.formula) {
            cell.value = {...value, formula};
          }
        });
      });
      Object.values(worksheet.tables).forEach(table => table._updateModelFormulas(update));
    });
  }

  _updateModelFormulas(update) {
    const {table, worksheet} = this;
    const {row, col} = table.tl;
    const first = row + (table.headerRow ? 1 : 0);
    table.rows.forEach((data, i) => {
      data.forEach((value, j) => {
        if (value && value.formula) {
          const formula = update(value.formula, worksheet, first + i, col + j);
          if (formula !== undefined) {
            data[j] = {...value, formula};
          }
        }
      });
    });
    table.columns.forEach((column, j) => {
      ['calculatedColumnFormula', 'totalsRowFormula'].forEach(key => {
        if (column[key]) {
          const formula = update(column[key], worksheet, first, col + j);
          if (formula !== undefined) {
            column[key] = formula;
          }
        }
      });
    });
  }

  _renameReferences(oldName, newName) {
    this._updateFormulas(formula =>
      mapTableReferences(
        formula,
        oldName,
        false,
        (spec, qualified, inner) => `${newName}[${inner}]`
      )
    );
  }

  _renameColumnReferences(oldName, newName) {
    const lower = oldName.toLowerCase();
    const rename = name => (name.toLowerCase() === lower ? newName : name);
    this._updateFormulas((formula, worksheet, row, col) =>
      mapTableReferences(
        formula,
        this.table.name,
        this._contains(worksheet, row, col),
        (spec, qualified) => {
          if (!spec.columns || !spec.columns.some(name => name.toLowerCase() === lower)) {
            return undefined;
          }
          const inner = renderSpecifier({...spec, columns: spec.columns.map(rename)});
          return qualified ? `${this.table.name}[${inner}]` : `[${inner}]`;
        }
      )
    );
  }

  // Turn the references to this table into cell references, as Excel does
  // when a table is converted to a range
  _convertReferencesToRange() {
    const geometry = this._geometry;
    this._updateFormulas((formula, worksheet, row, col) =>
      mapTableReferences(formula, this.table.name, this._contains(worksheet, row, col), spec =>
        toCellReference(spec, geometry, row)
      )
    );
    // a pivot table made from the table gets its header and data rows
    const {top, left, headerRow, height, columns, sheetName} = geometry;
    const ref = colCache.encode(
      top,
      left,
      top + (headerRow ? 1 : 0) + height - 1,
      left + columns.length - 1
    );
    this._updatePivotSources(
      () => `<worksheetSource ref="${ref}" sheet="${utils.xmlEncode(sheetName)}"/>`
    );
  }

  _removeColumnReferences(names) {
    const lower = names.map(name => String(name).toLowerCase());
    this._updateFormulas((formula, worksheet, row, col) =>
      mapTableReferences(formula, this.table.name, this._contains(worksheet, row, col), spec =>
        spec.columns && spec.columns.some(name => lower.includes(name.toLowerCase()))
          ? '#REF!'
          : undefined
      )
    );
  }

  // When the rows or columns of a table are deleted from the worksheet, the
  // references to it become #REF!
  _removeReferences() {
    this._updateFormulas(formula =>
      mapTableReferences(formula, this.table.name, false, () => '#REF!')
    );
  }

  // Change the source of the pivot caches read from the file that use this
  // table, written as <worksheetSource name="Sales"/>
  _updatePivotSources(replace) {
    const {workbook} = this.worksheet;
    const preserved = workbook && workbook.preservedPivotTables;
    const definitions = preserved && preserved.pivotCacheDefinitions;
    if (!definitions) {
      return;
    }
    const lower = this.table.name.toLowerCase();
    Object.keys(definitions).forEach(key => {
      definitions[key] = definitions[key].replace(/<worksheetSource\b[^>]*\/>/g, element => {
        const match = element.match(/\bname="([^"]*)"/);
        return match && utils.xmlDecode(match[1]).toLowerCase() === lower
          ? replace(element)
          : element;
      });
    });
  }

  get _rowRange() {
    const {row, col} = this.table.tl;
    return {
      top: row,
      left: col,
      bottom: row + this.tableHeight - 1,
      right: col + this.width - 1,
    };
  }

  // Before rows of the worksheet are spliced: Excel does not delete the
  // header row of a table without the rest of it
  _checkSpliceRows(start, count) {
    this.commit();
    const {top, bottom} = this._rowRange;
    const end = start + count - 1;
    if (this.table.headerRow && count > 0 && start <= top && end >= top && end < bottom) {
      throw new Error(
        `Cannot delete the header row of table ${this.table.name} without the whole table; ` +
          'set headerRow to false or use removeTable'
      );
    }
  }

  // After rows of the worksheet were spliced (count rows removed at start and
  // nInserts rows inserted there), move and size the table as Excel does: rows
  // inserted in the data or at the totals row become data rows, a deleted
  // totals row turns the totals off. Returns 'removed' when all the rows of the
  // table were deleted, or the row where an empty data row must be inserted
  // when all its data rows were deleted.
  _spliceRows(start, count, nInserts) {
    const {table, worksheet} = this;
    const {top, left, bottom, right} = this._rowRange;
    const header = table.headerRow ? 1 : 0;
    const dataFirst = top + header;
    const dataLast = dataFirst + this.height - 1;
    const end = start + count - 1;
    const removed = r => count > 0 && r >= start && r <= end;
    const moved = r => (count > 0 && r > end ? r - count : r);
    if (count > 0 && start <= top && end >= bottom) {
      return 'removed';
    }

    let nData = 0;
    let firstKept;
    for (let r = dataFirst; r <= dataLast; r++) {
      if (!removed(r)) {
        nData++;
        if (firstKept === undefined) firstKept = r;
      }
    }
    const totals = table.totalsRow && !removed(bottom);
    let newTop;
    if (header) {
      newTop = moved(top);
    } else {
      newTop = moved(firstKept === undefined ? bottom : firstKept);
    }
    let inserted;
    if (nInserts) {
      const newBottom = newTop + header + nData + (totals ? 1 : 0) - 1;
      if (start <= newTop) {
        newTop += nInserts;
      } else if (start <= newBottom) {
        nData += nInserts;
        inserted = start;
      }
    }
    if (newTop === top && totals === table.totalsRow && nData === this.height && !inserted) {
      return undefined;
    }

    // a calculated column formula is written for the first data row
    const newDataFirst = newTop + header;
    table.columns.forEach((column, j) => {
      if (column.calculatedColumnFormula && newDataFirst !== dataFirst) {
        column.calculatedColumnFormula = moveFormula(
          column.calculatedColumnFormula,
          colCache.encodeAddress(dataFirst, left + j),
          colCache.encodeAddress(newDataFirst, left + j)
        );
      }
    });

    table.totalsRow = totals;
    const newBottom = newTop + header + nData + (totals ? 1 : 0) - 1;
    table.tableRef = colCache.encode(newTop, left, newBottom, right);
    if (!nData) {
      return newDataFirst;
    }
    if (inserted) {
      this._fillCalculatedColumns(inserted, nInserts, newDataFirst);
    }
    this.load(worksheet);
    return undefined;
  }

  // An empty data row was inserted for a table whose data rows were all
  // deleted
  _addEmptyRow(rowNumber) {
    const {table} = this;
    const range = colCache.decode(table.tableRef);
    const top = range.top || range.row;
    const left = range.left || range.col;
    table.tableRef = colCache.encode(
      top,
      left,
      (range.bottom || range.row) + 1,
      range.right || range.col
    );
    this._fillCalculatedColumns(rowNumber, 1, rowNumber);
    this.load(this.worksheet);
  }

  // Rows inserted in a table get the formula of its calculated columns
  _fillCalculatedColumns(first, count, dataFirst) {
    const {worksheet, table} = this;
    const range = colCache.decode(table.tableRef);
    const left = range.left || range.col;
    table.columns.forEach((column, j) => {
      const formula = column.calculatedColumnFormula;
      if (!formula) {
        return;
      }
      for (let r = first; r < first + count; r++) {
        const cell = worksheet.getCell(r, left + j);
        if (cell.value === null) {
          cell.value = {
            formula: moveFormula(
              formula,
              colCache.encodeAddress(dataFirst, left + j),
              colCache.encodeAddress(r, left + j)
            ),
          };
        }
      }
    });
  }

  // After columns of the worksheet were spliced, as _spliceRows: columns
  // inserted inside the table become table columns named Column1..., deleted
  // columns are removed from it. Returns 'removed' when all its columns were
  // deleted.
  _spliceColumns(start, count, nInserts) {
    const {table, worksheet} = this;
    const {top, left, bottom, right} = this._rowRange;
    const end = start + count - 1;
    if (count > 0 && start <= left && end >= right) {
      return 'removed';
    }
    const removedColumns = table.columns.filter(
      (column, j) => count > 0 && left + j >= start && left + j <= end
    );
    if (removedColumns.length) {
      this._removeColumnReferences(removedColumns.map(column => column.name));
      table.columns = table.columns.filter(column => !removedColumns.includes(column));
    }
    let newLeft = left;
    if (count > 0 && left > end) {
      newLeft = left - count;
    } else if (count > 0 && left >= start) {
      newLeft = start;
    }
    let insertedColumns = 0;
    if (nInserts) {
      if (start <= newLeft) {
        newLeft += nInserts;
      } else if (start <= newLeft + table.columns.length - 1) {
        for (let i = 0; i < nInserts; i++) {
          const col = start + i;
          const cell = table.headerRow && worksheet.findCell(top, col);
          const name = this._uniqueColumnName(cell ? cell.text : '');
          table.columns.splice(col - newLeft, 0, {
            name,
            totalsRowFunction: 'none',
            filterButton: true,
          });
          if (table.headerRow) {
            worksheet.getCell(top, col).value = name;
          }
          insertedColumns++;
        }
      }
    }
    if (newLeft === left && !removedColumns.length && !insertedColumns) {
      return undefined;
    }
    table.tableRef = colCache.encode(top, newLeft, bottom, newLeft + table.columns.length - 1);
    this.load(worksheet);
    return undefined;
  }

  getColumn(colIndex) {
    const column = this.table.columns[colIndex];
    return new Column(this, column, colIndex);
  }

  addColumn(column, values, colIndex) {
    // Add a new column, including column defn and values
    // Inserts at colNumber or adds to the right
    this.cacheState();

    if (colIndex === undefined) {
      this.table.columns.push(column);
      this.table.rows.forEach((row, i) => {
        row.push(values[i]);
      });
    } else {
      this.table.columns.splice(colIndex, 0, column);
      this.table.rows.forEach((row, i) => {
        row.splice(colIndex, 0, values[i]);
      });
    }
  }

  removeColumns(colIndex, count = 1) {
    // Remove a column with data
    this.cacheState();

    // as in Excel, the references to a removed column become #REF!
    this._removeColumnReferences(
      this.table.columns.slice(colIndex, colIndex + count).map(column => column.name)
    );
    this.table.columns.splice(colIndex, count);
    this.table.rows.forEach(row => {
      row.splice(colIndex, count);
    });
  }

  _assign(target, prop, value) {
    this.cacheState();
    target[prop] = value;
  }

  /* eslint-disable lines-between-class-members */
  get ref() {
    return this.table.ref;
  }
  set ref(value) {
    this._assign(this.table, 'ref', value);
  }

  get name() {
    return this.table.name;
  }
  set name(value) {
    // renaming a table renames the references to it, as in Excel
    const oldName = this.table.name;
    if (value === oldName) {
      return;
    }
    const {worksheet} = this;
    checkTableName(value, worksheet, this);
    this._renameReferences(oldName, value);
    this._updatePivotSources(element =>
      element.replace(/\bname="[^"]*"/, `name="${utils.xmlEncode(value)}"`)
    );
    this.table.name = value;
    this.table.displayName = value;
    if (worksheet && worksheet.tables[oldName] === this) {
      delete worksheet.tables[oldName];
      worksheet.tables[value] = this;
    }
  }

  get displayName() {
    return this.table.displayName || this.table.name;
  }
  set displayName(value) {
    // Excel keeps the name and the display name the same
    this.name = value;
  }

  get headerRow() {
    return this.table.headerRow;
  }
  set headerRow(value) {
    this._assign(this.table, 'headerRow', value);
  }

  get totalsRow() {
    return this.table.totalsRow;
  }
  set totalsRow(value) {
    this._assign(this.table, 'totalsRow', value);
  }

  get theme() {
    return this.table.style.theme;
  }
  set theme(value) {
    this.table.style.theme = value;
  }

  get showFirstColumn() {
    return this.table.style.showFirstColumn;
  }
  set showFirstColumn(value) {
    this.table.style.showFirstColumn = value;
  }

  get showLastColumn() {
    return this.table.style.showLastColumn;
  }
  set showLastColumn(value) {
    this.table.style.showLastColumn = value;
  }

  get showRowStripes() {
    return this.table.style.showRowStripes;
  }
  set showRowStripes(value) {
    this.table.style.showRowStripes = value;
  }

  get showColumnStripes() {
    return this.table.style.showColumnStripes;
  }
  set showColumnStripes(value) {
    this.table.style.showColumnStripes = value;
  }

  get autoFilterRef() {
    return this.table.autoFilterRef;
  }
  set autoFilterRef(value) {
    this._assign(this.table, 'autoFilterRef', value);
  }
  /* eslint-enable lines-between-class-members */
}

Table.checkName = checkTableName;

module.exports = Table;
