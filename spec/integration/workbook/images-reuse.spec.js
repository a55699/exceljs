const path = require('path');
const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const IMAGE_FILE = path.join(__dirname, '../data/image.png');
const IMAGE2_FILE = path.join(__dirname, '../data/image1.jpg');

// the picture relationships of a drawing: [r:embed of each picture, Target]
async function pictures(zip, drawing) {
  const xml = await zip.file(`xl/drawings/${drawing}.xml`).async('string');
  const rels = await zip.file(`xl/drawings/_rels/${drawing}.xml.rels`).async('string');
  return [...xml.matchAll(/r:embed="([^"]+)"/g)].map(([, rId]) => {
    const rel = rels.match(new RegExp(`<Relationship Id="${rId}"[^>]*Target="([^"]+)"`));
    return [rId, rel && rel[1]];
  });
}

describe('Workbook', () => {
  describe('Images used more than once', () => {
    it('adds the same image to two worksheets', async () => {
      const workbook = new ExcelJS.Workbook();
      const imageId = workbook.addImage({filename: IMAGE_FILE, extension: 'png'});
      workbook.addWorksheet('First').addImage(imageId, 'B2:D6');
      workbook.addWorksheet('Second').addImage(imageId, 'C3:E7');

      const buffer = await workbook.xlsx.writeBuffer();
      const zip = await JSZip.loadAsync(buffer);
      expect(await pictures(zip, 'drawing1')).to.deep.equal([['rId1', '../media/image1.png']]);
      expect(await pictures(zip, 'drawing2')).to.deep.equal([['rId1', '../media/image1.png']]);

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      expect(workbook2.getWorksheet('First').getImages()).to.have.length(1);
      expect(workbook2.getWorksheet('Second').getImages()).to.have.length(1);
    });

    it('uses one relationship per image on a worksheet', async () => {
      const workbook = new ExcelJS.Workbook();
      const imageId = workbook.addImage({filename: IMAGE_FILE, extension: 'png'});
      const imageId2 = workbook.addImage({filename: IMAGE2_FILE, extension: 'jpeg'});
      const worksheet = workbook.addWorksheet('Images');
      worksheet.addImage(imageId, 'B2:D6');
      worksheet.addImage(imageId2, 'F2:H6');
      worksheet.addImage(imageId, 'J2:L6');

      const buffer = await workbook.xlsx.writeBuffer();
      const zip = await JSZip.loadAsync(buffer);
      expect(await pictures(zip, 'drawing1')).to.deep.equal([
        ['rId1', '../media/image1.png'],
        ['rId2', '../media/image2.jpeg'],
        ['rId1', '../media/image1.png'],
      ]);

      const workbook2 = new ExcelJS.Workbook();
      await workbook2.xlsx.load(buffer);
      // media ids follow the order of the parts in the package, which is not
      // fixed; the first and last pictures share an image
      const [first, second, third] = workbook2.getWorksheet('Images').getImages();
      expect(third.imageId).to.equal(first.imageId);
      expect(second.imageId).not.to.equal(first.imageId);
    });
  });
});
