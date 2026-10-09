/* eslint-disable max-classes-per-file */
const colCache = require('../utils/col-cache');
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
  if (
    cell.type === Enums.ValueType.Formula &&
    (value.sharedFormula || value.shareType === 'shared')
  ) {
    return {formula: cell.formula, result: cell.result};
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
    this._set('name', value);
    this.table._renameColumnReferences(oldName, value);
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
        return `SUBTOTAL(101,${this.table.name}[${column.name}])`;
      case 'countNums':
        return `SUBTOTAL(102,${this.table.name}[${column.name}])`;
      case 'count':
        return `SUBTOTAL(103,${this.table.name}[${column.name}])`;
      case 'max':
        return `SUBTOTAL(104,${this.table.name}[${column.name}])`;
      case 'min':
        return `SUBTOTAL(105,${this.table.name}[${column.name}])`;
      case 'stdDev':
        return `SUBTOTAL(106,${this.table.name}[${column.name}])`;
      case 'var':
        return `SUBTOTAL(107,${this.table.name}[${column.name}])`;
      case 'sum':
        return `SUBTOTAL(109,${this.table.name}[${column.name}])`;
      case 'custom':
        return column.totalsRowFormula;
      default:
        throw new Error(`Invalid Totals Row Function: ${column.totalsRowFunction}`);
    }
  }

  get width() {
    // width of the table
    return this.table.columns.length;
  }

  get height() {
    // height of the table data
    return this.table.rows.length;
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

    table.columns.forEach((column, i) => {
      assert(column.name, `Column ${i} must have a name`);
      // the first column shows Total unless it has a function
      if (i === 0 && column.totalsRowFunction === undefined) {
        assign(column, 'totalsRowLabel', 'Total');
      }
      assign(column, 'totalsRowFunction', 'none');
      column.totalsRowFormula = this.getFormula(column);
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
    table.columns.forEach((column, j) => {
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
