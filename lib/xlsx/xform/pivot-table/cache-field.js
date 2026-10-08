class CacheField {
  constructor({name, sharedItems}) {
    // string type
    //
    // {
    //   'name': 'A',
    //   'sharedItems': ['a1', 'a2', 'a3']
    // }
    //
    // or
    //
    // integer type
    //
    // {
    //   'name': 'D',
    //   'sharedItems': null
    // }
    this.name = name;
    this.sharedItems = sharedItems;
  }

  // Helper function to escape XML special characters
  escapeXml(unsafe) {
    if (unsafe == null) return '';
    return String(unsafe)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  render() {
    // PivotCache Field: http://www.datypic.com/sc/ooxml/e-ssml_cacheField-1.html
    // Shared Items: http://www.datypic.com/sc/ooxml/e-ssml_sharedItems-1.html

    // integer types
    if (this.sharedItems === null) {
      // TK(2023-07-18): left out attributes... minValue="5" maxValue="45"
      return `<cacheField name="${this.escapeXml(this.name)}" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" />
    </cacheField>`;
    }

    const items = this.sharedItems;

    // numeric types - use <n> tags
    if (items.every(isNumericItem)) {
      const numbers = items.map(item => Number(item));
      const minValue = Math.min(...numbers);
      const maxValue = Math.max(...numbers);
      const containsInteger = numbers.every(n => Number.isInteger(n));

      return `<cacheField name="${this.escapeXml(this.name)}" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="${containsInteger ? '1' : '0'}" minValue="${minValue}" maxValue="${maxValue}" count="${items.length}">
        ${numbers.map(n => `<n v="${n}" />`).join('')}
      </sharedItems>
    </cacheField>`;
    }

    // date types - use <d> tags
    if (items.every(item => item instanceof Date)) {
      const times = items.map(item => item.getTime());
      const minDate = formatDate(new Date(Math.min(...times)));
      const maxDate = formatDate(new Date(Math.max(...times)));

      return `<cacheField name="${this.escapeXml(this.name)}" numFmtId="14">
      <sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0"
        minDate="${minDate}" maxDate="${maxDate}" count="${items.length}">
        ${items.map(item => `<d v="${formatDate(item)}" />`).join('')}
      </sharedItems>
    </cacheField>`;
    }

    // string or mixed types - one tag per item type, with the same
    // sharedItems attributes Excel writes for mixed columns
    return `<cacheField name="${this.escapeXml(this.name)}" numFmtId="0">
      <sharedItems ${mixedTypeAttributes(items)} count="${items.length}">
        ${items.map(item => this.renderItem(item)).join('')}
      </sharedItems>
    </cacheField>`;
  }

  renderItem(item) {
    if (typeof item === 'number') {
      return `<n v="${item}" />`;
    }
    if (typeof item === 'boolean') {
      return `<b v="${item ? 1 : 0}" />`;
    }
    if (item instanceof Date) {
      return `<d v="${formatDate(item)}" />`;
    }
    return `<s v="${this.escapeXml(item)}" />`;
  }
}

// Decimal numbers only: rejects strings like "0x10" or "Infinity" that
// Number() accepts but are not valid xsd:double values.
const NUMERIC_STRING = /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/;

function isNumericItem(item) {
  if (typeof item === 'number') {
    return Number.isFinite(item);
  }
  return typeof item === 'string' && NUMERIC_STRING.test(item.trim());
}

function mixedTypeAttributes(items) {
  const numbers = items.filter(item => typeof item === 'number');
  const dates = items.filter(item => item instanceof Date);
  const hasText = items.some(item => typeof item === 'string' || typeof item === 'boolean');
  const typeCount = [hasText, numbers.length > 0, dates.length > 0].filter(Boolean).length;

  const attributes = [];
  if (!hasText) {
    attributes.push('containsSemiMixedTypes="0"', 'containsString="0"');
  }
  if (dates.length) {
    attributes.push('containsDate="1"');
  }
  if (typeCount > 1) {
    attributes.push('containsMixedTypes="1"');
  }
  // Like Excel, a column with dates describes its range with minDate/maxDate
  // only, even when it also contains numbers
  if (numbers.length && !dates.length) {
    attributes.push('containsNumber="1"');
    if (numbers.every(n => Number.isInteger(n))) {
      attributes.push('containsInteger="1"');
    }
    attributes.push(`minValue="${Math.min(...numbers)}"`, `maxValue="${Math.max(...numbers)}"`);
  }
  if (dates.length) {
    const times = dates.map(date => date.getTime());
    attributes.push(
      `minDate="${formatDate(new Date(Math.min(...times)))}"`,
      `maxDate="${formatDate(new Date(Math.max(...times)))}"`
    );
  }
  return attributes.join(' ');
}

function formatDate(date) {
  // xsd:dateTime without timezone, as Excel writes it
  return date.toISOString().slice(0, 19);
}

module.exports = CacheField;
