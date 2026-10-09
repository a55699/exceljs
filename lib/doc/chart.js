const colCache = require('../utils/col-cache');
const utils = require('../utils/utils');
const {normalizePivotValue} = require('../utils/pivot-value');
const Anchor = require('./anchor');

const TYPES = ['bar', 'line', 'area', 'pie', 'doughnut', 'scatter'];

const GROUPINGS = {
  bar: ['clustered', 'stacked', 'percentStacked'],
  line: ['standard', 'stacked', 'percentStacked'],
  area: ['standard', 'stacked', 'percentStacked'],
};

const LEGEND_POSITIONS = ['r', 'l', 't', 'b', 'tr'];

// 'Sales Data'!$B$2:$B$5 or Sheet1!B2
const FORMULA = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/;

function fail(message) {
  throw new Error(`Chart: ${message}`);
}

// A cell or range reference ({sheet, ref} or a formula string) as
// {sheetName, top, left, bottom, right}
function parseRef(workbook, ref, what) {
  let sheetName;
  let address;
  if (typeof ref === 'string') {
    const match = ref.match(FORMULA);
    if (!match) {
      fail(`${what} "${ref}" must be like 'Sheet 1'!$A$1:$A$4`);
    }
    sheetName = match[1] !== undefined ? match[1].replace(/'(')/g, '$1') : match[2];
    address = match[3];
  } else if (ref && typeof ref === 'object' && ref.ref) {
    sheetName = ref.sheet && typeof ref.sheet === 'object' ? ref.sheet.name : ref.sheet;
    address = ref.ref;
  } else {
    fail(`${what} must be {sheet, ref} or a formula string`);
  }
  if (!sheetName || !workbook.getWorksheet(sheetName)) {
    fail(`${what} refers to a worksheet that does not exist: ${sheetName}`);
  }
  let range;
  try {
    range = colCache.decode(address.replace(/\$/g, ''));
  } catch (error) {
    fail(`${what} has an invalid range: ${address}`);
  }
  // a single cell decodes as {row, col}
  const box =
    'top' in range ? range : {top: range.row, left: range.col, bottom: range.row, right: range.col};
  const {top, left, bottom, right} = box;
  if (!(top > 0 && left > 0)) {
    fail(`${what} has an invalid range: ${address}`);
  }
  return {sheetName, top, left, bottom, right};
}

function refLength(ref) {
  return (ref.bottom - ref.top + 1) * (ref.right - ref.left + 1);
}

function checkLine(ref, what) {
  if (ref.top !== ref.bottom && ref.left !== ref.right) {
    fail(`${what} must be a single row or a single column`);
  }
}

// 'Sales Data'!$B$2:$B$5; the sheet name is always quoted, as Excel accepts
function formulaOf(ref) {
  const sheet = `'${ref.sheetName.replace(/'/g, "''")}'`;
  const tl = `$${colCache.n2l(ref.left)}$${ref.top}`;
  if (ref.top === ref.bottom && ref.left === ref.right) {
    return `${sheet}!${tl}`;
  }
  return `${sheet}!${tl}:$${colCache.n2l(ref.right)}$${ref.bottom}`;
}

function cellsOf(workbook, ref) {
  const worksheet = workbook.getWorksheet(ref.sheetName);
  if (!worksheet) {
    fail(`the worksheet ${ref.sheetName} used by a chart no longer exists`);
  }
  const cells = [];
  for (let row = ref.top; row <= ref.bottom; row++) {
    for (let col = ref.left; col <= ref.right; col++) {
      cells.push(worksheet.getCell(row, col));
    }
  }
  return cells;
}

// Formula and cached values of a range, as Excel stores them next to the
// reference: numbers (dates as serial numbers) for values, and text or numbers
// for categories
function dataOf(workbook, ref, kind) {
  const cells = cellsOf(workbook, ref);
  const values = cells.map(cell => normalizePivotValue(cell.value));
  const date1904 = workbook.properties && workbook.properties.date1904;
  const toNumber = value => {
    if (typeof value === 'number') {
      return value;
    }
    if (value instanceof Date) {
      return utils.dateToExcel(value, date1904);
    }
    return null;
  };
  const numeric = values.every(
    value => value === null || typeof value === 'number' || value instanceof Date
  );
  if (kind === 'num' || (kind === 'cat' && numeric && values.some(value => value !== null))) {
    const first = cells.find(cell => cell.numFmt);
    return {
      formula: formulaOf(ref),
      type: 'num',
      formatCode: (first && first.numFmt) || 'General',
      cache: values.map(toNumber),
    };
  }
  return {
    formula: formulaOf(ref),
    type: 'str',
    cache: values.map(value => {
      if (value === null) {
        return null;
      }
      return value instanceof Date ? value.toISOString() : String(value);
    }),
  };
}

// Series made from a range: one per column (or row), with the names in the
// first row (column) and the categories in the first column (row)
function seriesFromData(workbook, data) {
  if (!data || typeof data !== 'object') {
    fail('data must be {sheet, ref}');
  }
  const {seriesIn = 'columns', titlesFromData = true, categoriesFromData = true} = data;
  if (seriesIn !== 'columns' && seriesIn !== 'rows') {
    fail(`data.seriesIn must be 'columns' or 'rows', not ${seriesIn}`);
  }
  const ref = parseRef(workbook, {sheet: data.sheet, ref: data.ref}, 'data');
  const byColumn = seriesIn === 'columns';
  // positions along and across the series
  const firstLine = byColumn ? ref.left : ref.top;
  const lastLine = byColumn ? ref.right : ref.bottom;
  const firstPoint = (byColumn ? ref.top : ref.left) + (titlesFromData ? 1 : 0);
  const lastPoint = byColumn ? ref.bottom : ref.right;
  const titleAt = byColumn ? ref.top : ref.left;
  const line = (at, from, to) =>
    byColumn
      ? {sheetName: ref.sheetName, top: from, bottom: to, left: at, right: at}
      : {sheetName: ref.sheetName, left: from, right: to, top: at, bottom: at};

  if (firstPoint > lastPoint) {
    fail('data has no values below (or after) the titles');
  }
  const categories = categoriesFromData ? line(firstLine, firstPoint, lastPoint) : undefined;
  const series = [];
  for (let at = categoriesFromData ? firstLine + 1 : firstLine; at <= lastLine; at++) {
    series.push({
      name: titlesFromData ? line(at, titleAt, titleAt) : undefined,
      categories,
      values: line(at, firstPoint, lastPoint),
    });
  }
  if (!series.length) {
    fail('data has no series');
  }
  return series;
}

function parseSeries(workbook, type, series, index) {
  const what = `series[${index}]`;
  if (!series || typeof series !== 'object') {
    fail(`${what} must be an object`);
  }
  let {name} = series;
  if (name !== undefined && typeof name !== 'string') {
    name = parseRef(workbook, name, `${what}.name`);
    if (refLength(name) !== 1) {
      fail(`${what}.name must be a single cell`);
    }
  }
  const ref = key => {
    const parsed = parseRef(workbook, series[key], `${what}.${key}`);
    checkLine(parsed, `${what}.${key}`);
    return parsed;
  };
  if (type === 'scatter') {
    if (!series.xValues || !series.yValues) {
      fail(`${what} of a scatter chart needs xValues and yValues`);
    }
    const xValues = ref('xValues');
    const yValues = ref('yValues');
    if (refLength(xValues) !== refLength(yValues)) {
      fail(`${what}.xValues and yValues must have the same number of cells`);
    }
    return {name, xValues, yValues};
  }
  if (!series.values) {
    fail(`${what} needs values`);
  }
  const values = ref('values');
  const categories = series.categories ? ref('categories') : undefined;
  if (categories && refLength(categories) !== refLength(values)) {
    fail(`${what}.categories and values must have the same number of cells`);
  }
  return {name, categories, values};
}

function parseAxis(axis, what) {
  if (axis === undefined) {
    return {};
  }
  if (!axis || typeof axis !== 'object') {
    fail(`${what} must be an object`);
  }
  ['min', 'max'].forEach(key => {
    if (axis[key] !== undefined && typeof axis[key] !== 'number') {
      fail(`${what}.${key} must be a number`);
    }
  });
  const {title, min, max, numFmt, gridlines} = axis;
  return {title, min, max, numFmt, gridlines};
}

// The position of the chart: a range like 'E2:L17', {tl, br} or {tl, ext},
// as for images
function parseRange(worksheet, range) {
  if (typeof range === 'string') {
    let decoded;
    try {
      decoded = colCache.decode(range);
    } catch (error) {
      fail(`range is invalid: ${range}`);
    }
    if (!('top' in decoded) || !(decoded.top > 0 && decoded.left > 0)) {
      fail(`range must be a cell range like 'E2:L17', not ${range}`);
    }
    return {
      tl: new Anchor(worksheet, {col: decoded.left, row: decoded.top}, -1),
      br: new Anchor(worksheet, {col: decoded.right, row: decoded.bottom}, 0),
      editAs: 'twoCell',
    };
  }
  if (!range || typeof range !== 'object' || !range.tl || (!range.br && !range.ext)) {
    fail("range must be like 'E2:L17', {tl, br} or {tl, ext}");
  }
  return {
    tl: new Anchor(worksheet, range.tl, 0),
    br: range.br && new Anchor(worksheet, range.br, 0),
    ext: range.ext,
    editAs: range.editAs || (range.br ? 'twoCell' : 'oneCell'),
  };
}

class Chart {
  constructor(worksheet, options) {
    this.worksheet = worksheet;
    const {workbook} = worksheet;
    if (!options || typeof options !== 'object') {
      fail('options are required');
    }
    const {type} = options;
    if (!TYPES.includes(type)) {
      fail(`type must be one of ${TYPES.join(', ')}, not ${type}`);
    }
    this.type = type;

    if (options.series && options.data) {
      fail('use series or data, not both');
    }
    let series;
    if (options.data) {
      series = seriesFromData(workbook, options.data);
      if (type === 'scatter') {
        // the first column (or row) holds the X values
        series = series.map(({name, categories, values}) => {
          if (!categories) {
            fail('data of a scatter chart needs the X values in its first column (or row)');
          }
          return {name, xValues: categories, yValues: values};
        });
      }
    } else if (Array.isArray(options.series) && options.series.length) {
      series = options.series.map((item, index) => parseSeries(workbook, type, item, index));
    } else {
      fail('series (a non-empty array) or data is required');
    }
    this.series = series;

    if (type === 'bar') {
      this.direction = options.direction || 'col';
      if (this.direction !== 'col' && this.direction !== 'bar') {
        fail(`direction must be 'col' or 'bar', not ${options.direction}`);
      }
    }
    if (GROUPINGS[type]) {
      this.grouping = options.grouping || GROUPINGS[type][0];
      if (!GROUPINGS[type].includes(this.grouping)) {
        fail(`grouping of a ${type} chart must be one of ${GROUPINGS[type].join(', ')}`);
      }
    }
    if (type === 'line') {
      this.markers = options.markers !== false;
    }
    if (type === 'doughnut') {
      this.holeSize = options.holeSize === undefined ? 50 : options.holeSize;
      if (!Number.isInteger(this.holeSize) || this.holeSize < 10 || this.holeSize > 90) {
        fail('holeSize must be a whole number from 10 to 90');
      }
    }
    if (type === 'scatter') {
      this.lines = Boolean(options.lines);
    }

    this.title = options.title;
    this.name = options.name;
    if (options.legend === false) {
      this.legend = false;
    } else {
      const position = (options.legend && options.legend.position) || 'r';
      if (!LEGEND_POSITIONS.includes(position)) {
        fail(`legend.position must be one of ${LEGEND_POSITIONS.join(', ')}`);
      }
      this.legend = {position};
    }
    const axes = options.axes || {};
    this.axes = {x: parseAxis(axes.x, 'axes.x'), y: parseAxis(axes.y, 'axes.y')};
    this.range = parseRange(worksheet, options.range);
  }

  // The chart as written: references with the values cached from the cells
  // as they are now
  get model() {
    const {workbook} = this.worksheet;
    const name = value => {
      if (value === undefined || typeof value === 'string') {
        return value === undefined ? undefined : {text: value};
      }
      return dataOf(workbook, value, 'str');
    };
    return {
      type: this.type,
      direction: this.direction,
      grouping: this.grouping,
      markers: this.markers,
      holeSize: this.holeSize,
      lines: this.lines,
      title: this.title,
      name: this.name,
      legend: this.legend,
      axes: this.axes,
      series: this.series.map(series =>
        this.type === 'scatter'
          ? {
              name: name(series.name),
              xValues: dataOf(workbook, series.xValues, 'num'),
              yValues: dataOf(workbook, series.yValues, 'num'),
            }
          : {
              name: name(series.name),
              categories: series.categories && dataOf(workbook, series.categories, 'cat'),
              values: dataOf(workbook, series.values, 'num'),
            }
      ),
      range: {
        tl: this.range.tl.model,
        br: this.range.br && this.range.br.model,
        ext: this.range.ext,
        editAs: this.range.editAs,
      },
    };
  }
}

module.exports = Chart;
