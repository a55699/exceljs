const BaseXform = require('../base-xform');
const ListXform = require('../list-xform');

const CustomFilterXform = require('./custom-filter-xform');
const FilterXform = require('./filter-xform');
const RawXmlXform = require('../drawing/raw-anchor-xform');

class FilterColumnXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      customFilters: new ListXform({
        tag: 'customFilters',
        count: false,
        empty: true,
        childXform: new CustomFilterXform(),
      }),
      filters: new ListXform({
        tag: 'filters',
        count: false,
        empty: true,
        childXform: new FilterXform(),
      }),
    };
  }

  get tag() {
    return 'filterColumn';
  }

  prepare(model, options) {
    model.colId = options.index.toString();
  }

  render(xmlStream, model) {
    const attributes = {
      colId: model.colId,
      hiddenButton: model.filterButton ? '0' : '1',
    };
    if (model.filterXml) {
      // the filter of a loaded table, kept as read
      xmlStream.openNode(this.tag, attributes);
      xmlStream.writeXml(model.filterXml);
      xmlStream.closeNode();
      return true;
    }
    if (model.customFilters) {
      xmlStream.openNode(this.tag, attributes);

      this.map.customFilters.render(xmlStream, model.customFilters);

      xmlStream.closeNode();
      return true;
    }
    xmlStream.leafNode(this.tag, attributes);
    return true;
  }

  parseOpen(node) {
    if (node.name === this.tag && !this.recording) {
      const {attributes} = node;
      this.model = {
        colId: attributes.colId,
        // hiddenButton is 0 when it is not written
        filterButton: attributes.hiddenButton !== '1',
      };
      this.recording = [];
      this.parser = undefined;
      return true;
    }
    // every filter (values, custom, top 10, dynamic, colour, icon) is kept as
    // XML; custom filters are also read into the model
    if (!this.recorder) {
      this.recorder = new RawXmlXform();
    }
    this.recorder.parseOpen(node);
    if (this.parser) {
      this.parser.parseOpen(node);
    } else if (this.map[node.name]) {
      this.parser = this.map[node.name];
      this.parser.parseOpen(node);
      if (node.name === 'customFilters') {
        this.hasCustomFilters = true;
      }
    }
    return true;
  }

  parseText(text) {
    if (this.recorder) {
      this.recorder.parseText(text);
    }
  }

  parseClose(name) {
    if (this.recorder) {
      if (!this.recorder.parseClose(name)) {
        this.recording.push(this.recorder.model.xml);
        this.recorder = undefined;
      }
      if (this.parser && !this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    // the end of filterColumn
    if (this.hasCustomFilters) {
      this.model.customFilters = this.map.customFilters.model;
    }
    if (this.recording.length) {
      this.model.filterXml = this.recording.join('');
    }
    this.recording = undefined;
    this.hasCustomFilters = false;
    return false;
  }
}

module.exports = FilterColumnXform;
