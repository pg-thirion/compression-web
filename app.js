// Compression Web — application entry point.
// Browser-only static web app that compresses PDF and DOCX files locally.
// See docs/superpowers/specs/2026-08-26-compression-web-design.md for the full design.

export const ALLOWED_LEVELS = new Set(['low', 'medium', 'high']);
export const FILE_MAX_BYTES = 100 * 1024 * 1024;
export const FILE_EXTENSIONS = new Set(['.pdf', '.docx']);

/**
 * Parse a compression level string. Strictly case-sensitive.
 * @param {string} value
 * @returns {'low'|'medium'|'high'}
 * @throws {RangeError} if value is not one of the allowed levels
 */
export function parseLevel(value) {
  if (typeof value !== 'string' || !ALLOWED_LEVELS.has(value)) {
    throw new RangeError(`Invalid compression level: ${String(value)}`);
  }
  return value;
}

/**
 * Validate a file for inclusion in the compression queue.
 * @param {{name: string, size: number, type?: string}} file
 * @returns {string|null} null if valid, otherwise a human-readable error message
 */
export function validateFile(file) {
  if (!file || typeof file.name !== 'string' || file.name === '') {
    return 'File is missing a name.';
  }

  const ext = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();
  if (!FILE_EXTENSIONS.has(ext)) {
    return `Unsupported file extension. Allowed: ${[...FILE_EXTENSIONS].join(', ')}.`;
  }

  if (typeof file.size !== 'number' || file.size < 0) {
    return 'File is missing a size.';
  }

  if (file.size > FILE_MAX_BYTES) {
    return `File exceeds maximum size of ${FILE_MAX_BYTES} bytes.`;
  }

  return null;
}

export async function compressPdf(bytes, deps = { PDFDocument: globalThis.PDFLib?.PDFDocument }) {
  const { PDFDocument } = deps;
  if (!PDFDocument) throw new Error('pdf-lib not loaded');
  try {
    const doc = await PDFDocument.load(bytes, { updateMetadata: false });
    doc.setTitle('');
    doc.setAuthor('');
    doc.setSubject('');
    doc.setKeywords([]);
    doc.setCreator('');
    doc.setProducer('');
    const out = await doc.save({ useObjectStreams: true });
    if (out.length >= bytes.length) return bytes;
    const verify = await PDFDocument.load(out, { updateMetadata: false });
    if (verify.getPageCount() !== doc.getPageCount()) return bytes;
    return out;
  } catch (err) {
    console.warn('compressPdf: re-load verification failed; returning original bytes', err);
    return bytes;
  }
}

export function minimalCoreXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/"></cp:coreProperties>'
  );
}

export function minimalAppXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
    '</Properties>'
  );
}

export function stripTrackedChanges(xml) {
  return xml
    .replace(/<w:del(?:\s[^>]*)?\/>|<w:del(?:\s[^>]*)?>[\s\S]*?<\/w:del>/g, '')
    .replace(/<w:ins(?:\s[^>]*)?\/>|<w:ins(?:\s[^>]*)?>/g, '')
    .replace(/<\/w:ins>/g, '');
}

export function stripCommentMarkers(xml) {
  return xml.replace(
    /<w:commentRangeStart\b[^>]*\/?>|<\/w:commentRangeStart>|<w:commentRangeEnd\b[^>]*\/?>|<\/w:commentRangeEnd>|<w:commentReference\b[^>]*\/?>|<\/w:commentReference>/g,
    ''
  );
}

/**
 * Re-encode an image at the supplied JPEG quality via Canvas.
 * @param {Uint8Array} bytes
 * @param {string} mimeType
 * @param {number} quality
 * @param {{createImageBitmap?: Function, OffscreenCanvas?: Function}} [deps]
 * @returns {Promise<Uint8Array>}
 */
export async function reencodeMediaImage(bytes, mimeType, quality, deps = {}) {
  const {
    createImageBitmap = globalThis.createImageBitmap,
    OffscreenCanvas = globalThis.OffscreenCanvas,
  } = deps;

  if (!createImageBitmap) throw new Error('createImageBitmap not available');
  if (!OffscreenCanvas) throw new Error('OffscreenCanvas not available');

  const blob = new Blob([bytes], { type: mimeType });
  const bitmap = await createImageBitmap(blob);
  try {
    const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
    const ctx = canvas.getContext('2d');
    ctx.drawImage(bitmap, 0, 0);
    const outBlob = await canvas.convertToBlob({ type: 'image/jpeg', quality });
    return new Uint8Array(await outBlob.arrayBuffer());
  } finally {
    bitmap.close();
  }
}

/**
 * Compress a DOCX file structurally using JSZip.
 * - Strips docProps/core.xml + docProps/app.xml to minimal stubs (when stripMetadata)
 * - Strips tracked changes / comment markers from word/document.xml (per options)
 * - Re-encodes JPEG/PNG images in word/media/ at the level's quality
 * - Re-zips with DEFLATE level 9
 * - Validates by reopening and checking word/document.xml exists
 * - Falls back to original bytes on any failure
 * @param {Uint8Array} bytes
 * @param {'low'|'medium'|'high'} [level]
 * @param {{stripMetadata?: boolean, stripTrackedChanges?: boolean, stripComments?: boolean}} [options]
 * @param {{JSZip?: any, reencodeMediaImage?: Function}} [deps]
 * @returns {Promise<Uint8Array>}
 */
export async function compressDocx(bytes, level = 'medium', options = {}, deps = {}) {
  const {
    JSZip: JSZipCtor = globalThis.JSZip,
    reencodeMediaImage: reencode = reencodeMediaImage,
  } = deps;

  if (!JSZipCtor) throw new Error('JSZip not loaded');

  const {
    stripMetadata = true,
    stripTrackedChanges: shouldStripTrackedChanges = false,
    stripComments: shouldStripComments = false,
  } = options;

  const quality = level === 'low' ? 0.5 : level === 'high' ? 0.85 : 0.7;

  try {
    const zip = await JSZipCtor.loadAsync(bytes);

    if (stripMetadata) {
      if (zip.file('docProps/core.xml')) {
        zip.file('docProps/core.xml', minimalCoreXml());
      }
      if (zip.file('docProps/app.xml')) {
        zip.file('docProps/app.xml', minimalAppXml());
      }
    }

    const docFile = zip.file('word/document.xml');
    if (docFile) {
      let docXml = await docFile.async('string');
      if (shouldStripTrackedChanges) docXml = stripTrackedChanges(docXml);
      if (shouldStripComments) docXml = stripCommentMarkers(docXml);
      zip.file('word/document.xml', docXml);
    }

    const mediaNames = Object.keys(zip.files).filter(
      (name) => /^word\/media\//.test(name) && !zip.files[name].dir
    );
    for (const name of mediaNames) {
      const lower = name.toLowerCase();
      const ext = lower.slice(lower.lastIndexOf('.'));
      if (ext !== '.jpg' && ext !== '.jpeg' && ext !== '.png') continue;
      const isJpeg = ext === '.jpg' || ext === '.jpeg';
      if (level === 'high' && isJpeg) continue;

      const mimeType = isJpeg ? 'image/jpeg' : 'image/png';
      const origBytes = await zip.file(name).async('uint8array');
      const reencoded = await reencode(origBytes, mimeType, quality);
      if (reencoded.length < origBytes.length) {
        zip.file(name, reencoded);
      }
    }

    const out = await zip.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 9 },
    });

    const verify = await JSZipCtor.loadAsync(out);
    if (!verify.file('word/document.xml')) {
      console.warn('compressDocx: validation failed — word/document.xml missing');
      return bytes;
    }

    if (out.length >= bytes.length) return bytes;
    return out;
  } catch (err) {
    console.warn('compressDocx: failed, returning original bytes', err);
    return bytes;
  }
}
