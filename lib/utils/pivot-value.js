// Helpers to turn cell values into the plain values a pivot cache stores.
//
// Cell values can be objects (formula, rich text, hyperlink, error) and a
// formula cell returns a new object on every read, so values must be
// normalized before they are deduplicated into sharedItems and before records
// look up their sharedItems index.

function normalizePivotValue(value) {
  if (value === null || value === undefined) {
    return null;
  }
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }
  if (typeof value !== 'object') {
    return value;
  }
  if (value.formula !== undefined || value.sharedFormula !== undefined) {
    return normalizePivotValue(value.result);
  }
  if (Array.isArray(value.richText)) {
    return value.richText.map(run => run.text).join('');
  }
  if (value.hyperlink !== undefined) {
    return normalizePivotValue(value.text);
  }
  if (value.error !== undefined) {
    return String(value.error);
  }
  return String(value);
}

// Key used to match equal values. Strings match case-insensitively,
// as Excel treats pivot items that differ only in case as the same item.
function pivotValueKey(value) {
  if (value instanceof Date) {
    return `d:${value.getTime()}`;
  }
  if (typeof value === 'string') {
    return `s:${value.toLowerCase()}`;
  }
  return `${typeof value}:${value}`;
}

const TYPE_ORDER = {number: 0, object: 1, string: 2, boolean: 3};

// Sort numbers numerically, dates by time, then strings, then booleans.
function comparePivotValues(a, b) {
  const typeDiff = TYPE_ORDER[typeof a] - TYPE_ORDER[typeof b];
  if (typeDiff !== 0) {
    return typeDiff;
  }
  if (typeof a === 'string') {
    if (a === b) return 0;
    return a < b ? -1 : 1;
  }
  return Number(a) - Number(b);
}

module.exports = {normalizePivotValue, pivotValueKey, comparePivotValues};
