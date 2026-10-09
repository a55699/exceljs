const XmlStream = require('../../../utils/xml-stream');

const BaseXform = require('../base-xform');

// used for rendering the [Content_Types].xml file
// not used for parsing
class ContentTypesXform extends BaseXform {
  render(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);

    xmlStream.openNode('Types', ContentTypesXform.PROPERTY_ATTRIBUTES);

    const mediaHash = {};
    (model.media || []).forEach(medium => {
      if (medium.type === 'image') {
        const imageType = medium.extension;
        if (!mediaHash[imageType]) {
          mediaHash[imageType] = true;
          xmlStream.leafNode('Default', {Extension: imageType, ContentType: `image/${imageType}`});
        }
      }
    });

    xmlStream.leafNode('Default', {
      Extension: 'rels',
      ContentType: 'application/vnd.openxmlformats-package.relationships+xml',
    });
    xmlStream.leafNode('Default', {Extension: 'xml', ContentType: 'application/xml'});

    xmlStream.leafNode('Override', {
      PartName: '/xl/workbook.xml',
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    });

    model.worksheets.forEach(worksheet => {
      const name = `/xl/worksheets/sheet${worksheet.id}.xml`;
      xmlStream.leafNode('Override', {
        PartName: name,
        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      });
    });

    // Content types for programmatic pivot tables
    (model.pivotTables || []).forEach(pivotTable => {
      const {tableNumber, cacheNumber} = pivotTable;
      xmlStream.leafNode('Override', {
        PartName: `/xl/pivotCache/pivotCacheDefinition${cacheNumber}.xml`,
        ContentType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
      });
      xmlStream.leafNode('Override', {
        PartName: `/xl/pivotCache/pivotCacheRecords${cacheNumber}.xml`,
        ContentType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
      });
      xmlStream.leafNode('Override', {
        PartName: `/xl/pivotTables/pivotTable${tableNumber}.xml`,
        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
      });
    });

    // Content types for preserved pivot parts (from round-trip): one entry
    // per part that exists, since several pivot tables can share one cache
    if (model.preservedPivotTables) {
      const {pivotTables, pivotCacheDefinitions, pivotCacheRecords} = model.preservedPivotTables;
      Object.keys(pivotCacheDefinitions || {}).forEach(name => {
        xmlStream.leafNode('Override', {
          PartName: `/xl/pivotCache/${name}.xml`,
          ContentType:
            'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
        });
      });
      Object.keys(pivotCacheRecords || {}).forEach(name => {
        xmlStream.leafNode('Override', {
          PartName: `/xl/pivotCache/${name}.xml`,
          ContentType:
            'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
        });
      });
      Object.keys(pivotTables || {}).forEach(name => {
        xmlStream.leafNode('Override', {
          PartName: `/xl/pivotTables/${name}.xml`,
          ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
        });
      });
    }

    // Content types for preserved charts (from round-trip)
    if (model.preservedChartsXml) {
      Object.keys(model.preservedChartsXml).forEach(name => {
        xmlStream.leafNode('Override', {
          PartName: `/xl/charts/${name}.xml`,
          ContentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
        });
      });
    }

    // Content types for preserved chart styles
    if (model.preservedChartStylesXml) {
      Object.keys(model.preservedChartStylesXml).forEach(name => {
        xmlStream.leafNode('Override', {
          PartName: `/xl/charts/${name}.xml`,
          ContentType: 'application/vnd.ms-office.chartstyle+xml',
        });
      });
    }

    // Content types for preserved chart colors
    if (model.preservedChartColorsXml) {
      Object.keys(model.preservedChartColorsXml).forEach(name => {
        xmlStream.leafNode('Override', {
          PartName: `/xl/charts/${name}.xml`,
          ContentType: 'application/vnd.ms-office.chartcolorstyle+xml',
        });
      });
    }

    // Content types for drawings kept as read (not used by a worksheet): the
    // shapes drawn over a chart (<c:userShapes>) have their own content type
    if (model.preservedDrawingsXml) {
      Object.keys(model.preservedDrawingsXml).forEach(name => {
        const chartShapes = /^(?:<\?[^>]*>)?\s*<(?:\w+:)?userShapes[\s>]/.test(
          model.preservedDrawingsXml[name]
        );
        xmlStream.leafNode('Override', {
          PartName: `/xl/drawings/${name}.xml`,
          ContentType: chartShapes
            ? 'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml'
            : 'application/vnd.openxmlformats-officedocument.drawing+xml',
        });
      });
    }

    xmlStream.leafNode('Override', {
      PartName: '/xl/theme/theme1.xml',
      ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
    });
    xmlStream.leafNode('Override', {
      PartName: '/xl/styles.xml',
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    });

    const hasSharedStrings = model.sharedStrings && model.sharedStrings.count;
    if (hasSharedStrings) {
      xmlStream.leafNode('Override', {
        PartName: '/xl/sharedStrings.xml',
        ContentType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
      });
    }

    if (model.tables) {
      model.tables.forEach(table => {
        xmlStream.leafNode('Override', {
          PartName: `/xl/tables/${table.target}`,
          ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
        });
      });
    }

    if (model.drawings) {
      model.drawings.forEach(drawing => {
        xmlStream.leafNode('Override', {
          PartName: `/xl/drawings/${drawing.name}.xml`,
          ContentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
        });
      });
    }

    if (model.commentRefs || model.formControlRefs) {
      xmlStream.leafNode('Default', {
        Extension: 'vml',
        ContentType: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
      });

      if (model.commentRefs) {
        model.commentRefs.forEach(({commentName}) => {
          xmlStream.leafNode('Override', {
            PartName: `/xl/${commentName}.xml`,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
          });
        });
      }
    }

    if (model.formControlRefs) {
      model.formControlRefs.forEach(ref => {
        xmlStream.leafNode('Override', {
          PartName: `/xl/ctrlProps/ctrlProp${ref.id}.xml`,
          ContentType: 'application/vnd.ms-excel.controlproperties+xml',
        });
      });
    }

    xmlStream.leafNode('Override', {
      PartName: '/docProps/core.xml',
      ContentType: 'application/vnd.openxmlformats-package.core-properties+xml',
    });
    xmlStream.leafNode('Override', {
      PartName: '/docProps/app.xml',
      ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    });

    xmlStream.closeNode();
  }

  parseOpen() {
    return false;
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

ContentTypesXform.PROPERTY_ATTRIBUTES = {
  xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types',
};

module.exports = ContentTypesXform;
