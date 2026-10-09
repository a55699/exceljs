const path = require('path');
const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// Saved by Excel: charts on two sheets (chart1, chart2), see drawing-kept-anchors.spec.js
const MIXED_FILE = './spec/integration/data/drawings-mixed.xlsx';
const IMAGE_FILE = path.join(__dirname, '../data/image.png');

function salesWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sales Data');
  worksheet.addRows([
    ['Month', 'East', 'West'],
    ['Jan', 10, 30],
    ['Feb', 40, 60],
    ['Mar', 50, 70],
    ['Apr', 20, 10],
  ]);
  return {workbook, worksheet};
}

async function writeZip(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  return {buffer, zip: await JSZip.loadAsync(buffer)};
}

function readFile(zip, name) {
  return zip.file(name).async('string');
}

function relsOf(xml) {
  return [...xml.matchAll(/<Relationship ([^>]*)\/>/g)].map(match => {
    const attr = name => (match[1].match(new RegExp(`${name}="([^"]*)"`)) || [])[1];
    return {Id: attr('Id'), Type: attr('Type'), Target: attr('Target')};
  });
}

// the drawing of a worksheet: its XML, the chart parts it uses (by r:id) and
// the shape ids of its anchors
async function drawingOf(zip, sheetFile) {
  const sheetRels = relsOf(await readFile(zip, `xl/worksheets/_rels/${sheetFile}.rels`));
  const drawingRel = sheetRels.find(rel => rel.Type.endsWith('/drawing'));
  const name = path.posix.basename(drawingRel.Target, '.xml');
  const xml = await readFile(zip, `xl/drawings/${name}.xml`);
  const rels = relsOf(await readFile(zip, `xl/drawings/_rels/${name}.xml.rels`));
  const charts = [...xml.matchAll(/<c:chart [^>]*r:id="([^"]+)"/g)].map(([, rId]) => {
    const rel = rels.find(r => r.Id === rId);
    return rel && rel.Target;
  });
  const ids = [...xml.matchAll(/<xdr:cNvPr\s[^>]*?\bid="(\d+)"/g)].map(match => match[1]);
  return {name, xml, charts, ids};
}

// the chart-related overrides in [Content_Types].xml
async function chartContentTypes(zip) {
  const xml = await readFile(zip, '[Content_Types].xml');
  return [...xml.matchAll(/PartName="\/xl\/charts\/([^"]+)" ContentType="([^"]+)"/g)].map(
    ([, part, type]) => `${part} ${type.split('.').pop()}`
  );
}

describe('Workbook', () => {
  describe('Charts', () => {
    it('writes a chart part, its drawing anchor and content type', async () => {
      const {workbook, worksheet} = salesWorkbook();
      worksheet.addChart({
        type: 'bar',
        title: 'Sales by month',
        data: {sheet: worksheet, ref: 'A1:C5'},
        axes: {x: {title: 'Month'}, y: {title: 'Units', min: 0, max: 100, numFmt: '#,##0'}},
        range: 'E2:L17',
      });

      const {zip} = await writeZip(workbook);
      const drawing = await drawingOf(zip, 'sheet1.xml');
      expect(drawing.charts).to.deep.equal(['../charts/chart1.xml']);
      expect(drawing.xml).to.include('<xdr:twoCellAnchor editAs="twoCell">');
      expect(drawing.xml).to.include('<xdr:cNvPr id="1" name="Chart 1"/>');
      expect(await chartContentTypes(zip)).to.deep.equal(['chart1.xml chart+xml']);

      const chart = await readFile(zip, 'xl/charts/chart1.xml');
      expect(chart).to.include('<c:barDir val="col"/><c:grouping val="clustered"/>');
      expect(chart).to.include('<a:t>Sales by month</a:t>');
      expect(chart).to.include('<c:autoTitleDeleted val="0"/>');
      // both axes shown, with their titles and the value axis scale
      expect(chart.match(/<c:delete val="0"\/>/g)).to.have.length(2);
      expect(chart).to.include('<a:t>Month</a:t>');
      expect(chart).to.include('<c:max val="100"/><c:min val="0"/>');
      expect(chart).to.include('<c:numFmt formatCode="#,##0" sourceLinked="0"/>');
      // text categories as strRef, values as numRef, with cached values
      expect(chart).to.include(
        '<c:cat><c:strRef><c:f>&apos;Sales Data&apos;!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>Jan</c:v></c:pt>'
      );
      expect(chart).to.include(
        '<c:val><c:numRef><c:f>&apos;Sales Data&apos;!$C$2:$C$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>30</c:v></c:pt>'
      );
    });

    it('writes each chart type', async () => {
      const {workbook, worksheet} = salesWorkbook();
      const data = {sheet: worksheet, ref: 'A1:C5'};
      worksheet.addChart({
        type: 'bar',
        direction: 'bar',
        grouping: 'stacked',
        data,
        range: 'E2:L17',
      });
      worksheet.addChart({type: 'line', markers: false, data, range: 'E18:L33'});
      worksheet.addChart({type: 'area', grouping: 'percentStacked', data, range: 'E34:L49'});
      worksheet.addChart({type: 'pie', data: {sheet: worksheet, ref: 'A1:B5'}, range: 'M2:T17'});
      worksheet.addChart({
        type: 'doughnut',
        holeSize: 60,
        data,
        legend: {position: 'b'},
        range: 'M18:T33',
      });
      worksheet.addChart({
        type: 'scatter',
        data: {sheet: worksheet, ref: 'B1:C5'},
        legend: false,
        range: 'M34:T49',
      });

      const {zip} = await writeZip(workbook);
      const chart = n => readFile(zip, `xl/charts/chart${n}.xml`);
      const bar = await chart(1);
      expect(bar).to.include('<c:barDir val="bar"/><c:grouping val="stacked"/>');
      expect(bar).to.include('<c:overlap val="100"/>');
      expect(bar).to.include(
        '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>'
      );
      const line = await chart(2);
      expect(line).to.include('<c:lineChart><c:grouping val="standard"/>');
      expect(line.match(/<c:symbol val="none"\/>/g)).to.have.length(2);
      expect(await chart(3)).to.include('<c:areaChart><c:grouping val="percentStacked"/>');
      const pie = await chart(4);
      expect(pie).to.include('<c:pieChart><c:varyColors val="1"/>');
      expect(pie).not.to.include('<c:catAx>');
      const doughnut = await chart(5);
      expect(doughnut).to.include('<c:holeSize val="60"/>');
      expect(doughnut).to.include('<c:legendPos val="b"/>');
      const scatter = await chart(6);
      expect(scatter).to.include('<c:scatterStyle val="lineMarker"/>');
      expect(scatter).to.include('<c:xVal><c:numRef><c:f>&apos;Sales Data&apos;!$B$2:$B$5</c:f>');
      expect(scatter).to.include('<a:ln w="19050"><a:noFill/></a:ln>');
      expect(scatter.match(/<c:valAx>/g)).to.have.length(2);
      expect(scatter).not.to.include('<c:legend>');

      const drawing = await drawingOf(zip, 'sheet1.xml');
      expect(drawing.charts).to.have.length(6);
      expect(new Set(drawing.ids).size).to.equal(6);
    });

    it('draws charts and images in one drawing', async () => {
      const {workbook, worksheet} = salesWorkbook();
      const imageId = workbook.addImage({filename: IMAGE_FILE, extension: 'png'});
      worksheet.addImage(imageId, 'A8:C14');
      worksheet.addChart({
        type: 'line',
        data: {sheet: worksheet, ref: 'A1:C5'},
        range: {tl: {col: 4, row: 1}, ext: {width: 480, height: 288}},
      });
      const other = workbook.addWorksheet('Other');
      other.addChart({type: 'bar', data: {sheet: worksheet, ref: 'A1:C5'}, range: 'B2:I17'});

      const {zip} = await writeZip(workbook);
      const sales = await drawingOf(zip, 'sheet1.xml');
      expect(sales.xml).to.include('<xdr:pic>');
      expect(sales.xml).to.include('<xdr:oneCellAnchor editAs="oneCell">');
      expect(sales.xml).to.include('<xdr:ext cx="4572000" cy="2743200"/>');
      expect(sales.charts).to.deep.equal(['../charts/chart1.xml']);
      expect(new Set(sales.ids).size).to.equal(2);
      const otherDrawing = await drawingOf(zip, 'sheet2.xml');
      expect(otherDrawing.charts).to.deep.equal(['../charts/chart2.xml']);
      expect(otherDrawing.name).not.to.equal(sales.name);
    });

    it('numbers new charts after the charts of a loaded workbook', async () => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(MIXED_FILE);
      const data = workbook.getWorksheet('Data');
      data.addChart({type: 'pie', data: {sheet: data, ref: 'A1:B5'}, range: 'H2:M14'});

      const {zip} = await writeZip(workbook);
      const drawing = await drawingOf(zip, 'sheet1.xml');
      expect(drawing.charts).to.deep.equal(['../charts/chart1.xml', '../charts/chart3.xml']);
      // named after the kept "Chart 1"
      expect(drawing.xml).to.include('name="Chart 2"/>');
      // the new chart's shape id comes after the ids of the kept anchors
      expect(new Set(drawing.ids).size).to.equal(drawing.ids.length);
      expect(Math.max(...drawing.ids)).to.equal(+drawing.ids[drawing.ids.length - 1]);
      expect(await chartContentTypes(zip)).to.include.members([
        'chart1.xml chart+xml',
        'chart2.xml chart+xml',
        'chart3.xml chart+xml',
      ]);
      expect(await readFile(zip, 'xl/charts/chart3.xml')).to.include('<c:pieChart>');
    });

    it('keeps the chart when the workbook is loaded and written again', async () => {
      const {workbook, worksheet} = salesWorkbook();
      worksheet.addChart({
        type: 'bar',
        title: 'Kept',
        data: {sheet: worksheet, ref: 'A1:C5'},
        range: 'E2:L17',
      });
      const {buffer} = await writeZip(workbook);

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      // loaded charts are kept as read, not returned by getCharts
      expect(workbook2.getWorksheet('Sales Data').getCharts()).to.deep.equal([]);
      const {zip} = await writeZip(workbook2);
      expect((await drawingOf(zip, 'sheet1.xml')).charts).to.deep.equal(['../charts/chart1.xml']);
      expect(await readFile(zip, 'xl/charts/chart1.xml')).to.include('<a:t>Kept</a:t>');
    });
  });
});
