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
    .replace(/<w:del(?:\s[^>]*)?>[\s\S]*?<\/w:del>/g, '')
    .replace(/<w:ins(?:\s[^>]*)?>/g, '')
    .replace(/<\/w:ins>/g, '');
}

export function stripCommentMarkers(xml) {
  return xml.replace(
    /<w:commentRangeStart\b[^>]*\/?>|<\/w:commentRangeStart>|<w:commentRangeEnd\b[^>]*\/?>|<\/w:commentRangeEnd>|<w:commentReference\b[^>]*\/?>|<\/w:commentReference>/g,
    ''
  );
}
