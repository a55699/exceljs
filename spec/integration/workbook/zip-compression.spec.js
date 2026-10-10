const fs = require('fs');

const ExcelJS = verquire('exceljs');

const PNG_FILE = './spec/integration/data/image.png';
const JPEG_FILE = './spec/integration/data/image1.jpg';

// The compression method of each entry, from the central directory of a zip:
// 0 is STORE and 8 is DEFLATE
function compressionMethods(buffer) {
  const end = buffer.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  const count = buffer.readUInt16LE(end + 10);
  let offset = buffer.readUInt32LE(end + 16);
  const methods = {};
  for (let i = 0; i < count; i++) {
    const nameLength = buffer.readUInt16LE(offset + 28);
    const name = buffer.toString('utf8', offset + 46, offset + 46 + nameLength);
    methods[name.replace(/^\//, '')] = buffer.readUInt16LE(offset + 10);
    offset += 46 + nameLength + buffer.readUInt16LE(offset + 30) + buffer.readUInt16LE(offset + 32);
  }
  return methods;
}

async function mediaOf(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook.model.media.map(medium => medium.buffer);
}

describe('Zip compression', () => {
  const png = fs.readFileSync(PNG_FILE);
  const jpeg = fs.readFileSync(JPEG_FILE);

  describe('Workbook', () => {
    function workbookWithImages() {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Images');
      worksheet.getCell('A1').value = 'images';
      const images = [
        workbook.addImage({buffer: png, extension: 'png'}),
        workbook.addImage({
          base64: `data:image/jpeg;base64,${jpeg.toString('base64')}`,
          extension: 'jpeg',
        }),
      ];
      worksheet.addImage(images[0], 'B2:D8');
      worksheet.addImage(images[1], 'B12:D18');
      return workbook;
    }

    it('stores images that are compressed already', async () => {
      const buffer = await workbookWithImages().xlsx.writeBuffer();
      const methods = compressionMethods(buffer);
      expect(methods['xl/media/image1.png']).to.equal(0);
      expect(methods['xl/media/image2.jpeg']).to.equal(0);
      expect(methods['xl/worksheets/sheet1.xml']).to.equal(8);
      expect(await mediaOf(buffer)).to.deep.equal([png, jpeg]);
    });

    it('stores every part with a compression level of 0', async () => {
      const buffer = await workbookWithImages().xlsx.writeBuffer({
        zip: {compression: 'DEFLATE', compressionOptions: {level: 0}},
      });
      expect(new Set(Object.values(compressionMethods(buffer)))).to.deep.equal(new Set([0]));

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      expect(workbook.getWorksheet('Images').getCell('A1').value).to.equal('images');
    });
  });

  describe('WorkbookWriter', () => {
    async function writeImage(image, options) {
      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const worksheet = workbook.addWorksheet('Images');
      worksheet.addBackgroundImage(workbook.addImage(image));
      worksheet.addRow(['images']).commit();
      worksheet.commit();
      await workbook.commit();
      return workbook.stream.read();
    }

    it('writes an image given in base64', async () => {
      const buffer = await writeImage({
        base64: `data:image/png;base64,${png.toString('base64')}`,
        extension: 'png',
      });
      expect(await mediaOf(buffer)).to.deep.equal([png]);
    });

    it('stores images that are compressed already', async () => {
      const buffer = await writeImage({buffer: jpeg, extension: 'jpeg'}, {zip: {zlib: {level: 9}}});
      const methods = compressionMethods(buffer);
      expect(methods['xl/media/image0.jpeg']).to.equal(0);
      expect(methods['xl/worksheets/sheet1.xml']).to.equal(8);
      expect(await mediaOf(buffer)).to.deep.equal([jpeg]);
    });

    it('stores an image given by its file name', async () => {
      const buffer = await writeImage({filename: PNG_FILE, extension: 'png'});
      expect(compressionMethods(buffer)['xl/media/image0.png']).to.equal(0);
      expect(await mediaOf(buffer)).to.deep.equal([png]);
    });
  });
});
