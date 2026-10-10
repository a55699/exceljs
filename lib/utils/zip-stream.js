const events = require('events');
const JSZip = require('jszip');

const StreamBuf = require('./stream-buf');
const {stringToBuffer} = require('./browser-buffer-encode');

// Images in these formats are compressed already, so deflating them again
// costs time and saves nothing: they are stored as they are
function isCompressedMedia(name) {
  return /\.(png|jpe?g|gif)$/i.test(name);
}

// =============================================================================
// The ZipWriter class
// Packs streamed data into an output zip stream
class ZipWriter extends events.EventEmitter {
  constructor(options) {
    super();
    this.options = Object.assign(
      {
        type: 'nodebuffer',
        compression: 'DEFLATE',
      },
      options
    );
    // JSZip reads a level of 0 as its default level, not as no compression
    const {compressionOptions} = this.options;
    if (compressionOptions && compressionOptions.level === 0) {
      this.options.compression = 'STORE';
    }

    this.zip = new JSZip();
    this.stream = new StreamBuf();
  }

  append(data, options) {
    const fileOptions = {};
    if (options.store) {
      fileOptions.compression = 'STORE';
    }
    if (options.hasOwnProperty('base64') && options.base64) {
      fileOptions.base64 = true;
    } else if (process.browser && typeof data === 'string') {
      // https://www.npmjs.com/package/process
      // use TextEncoder in browser
      data = stringToBuffer(data);
    }
    this.zip.file(options.name, data, fileOptions);
  }

  async finalize() {
    const content = await this.zip.generateAsync(this.options);
    this.stream.end(content);
    this.emit('finish');
  }

  // ==========================================================================
  // Stream.Readable interface
  read(size) {
    return this.stream.read(size);
  }

  setEncoding(encoding) {
    return this.stream.setEncoding(encoding);
  }

  pause() {
    return this.stream.pause();
  }

  resume() {
    return this.stream.resume();
  }

  isPaused() {
    return this.stream.isPaused();
  }

  pipe(destination, options) {
    return this.stream.pipe(destination, options);
  }

  unpipe(destination) {
    return this.stream.unpipe(destination);
  }

  unshift(chunk) {
    return this.stream.unshift(chunk);
  }

  wrap(stream) {
    return this.stream.wrap(stream);
  }
}

// =============================================================================

module.exports = {
  ZipWriter,
  isCompressedMedia,
};
