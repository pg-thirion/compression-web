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

/**
 * Split an iterable of files into accepted (valid) and rejected (with errors).
 * Errors are pushed as { name, message } objects.
 * @param {Iterable<{name: string}>} files
 * @param {Array<{name: string, message: string}>} errors
 * @param {{validateFile?: Function}} [deps]
 * @returns {Array}
 */
export function filterFiles(files, errors, deps = {}) {
  const accepted = [];
  for (const file of files) {
    const err = deps.validateFile ? deps.validateFile(file) : validateFile(file);
    if (err) errors.push({ name: file.name, message: err });
    else accepted.push(file);
  }
  return accepted;
}

function defaultCreateObjectURL(blob) {
  if (typeof URL === 'undefined' || typeof URL.createObjectURL !== 'function') {
    throw new Error('URL.createObjectURL is not available');
  }
  return URL.createObjectURL(blob);
}

function defaultCreateAnchorElement() {
  if (typeof document === 'undefined') {
    throw new Error('document is not available');
  }
  return document.createElement('a');
}

function defaultScheduleRevoke(url) {
  if (typeof URL === 'undefined' || typeof URL.revokeObjectURL !== 'function') return;
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

/**
 * Build a single ZIP from compressed results and trigger a browser download.
 * @param {Array<{name: string, bytes: Uint8Array}>} results
 * @param {{
 *   JSZip?: Function,
 *   createObjectURL?: Function,
 *   createAnchorElement?: Function,
 *   scheduleRevoke?: Function,
 *   filename?: string,
 * }} [deps]
 * @returns {Promise<string>} the object URL created for the download
 */
export async function downloadResults(results, deps = {}) {
  const {
    JSZip: JSZipCtor = globalThis.JSZip,
    createObjectURL = defaultCreateObjectURL,
    createAnchorElement = defaultCreateAnchorElement,
    scheduleRevoke = defaultScheduleRevoke,
    filename = `compressed-${Date.now()}.zip`,
  } = deps;

  if (typeof JSZipCtor !== 'function') {
    throw new Error('JSZip not loaded');
  }

  const zip = new JSZipCtor();
  for (const { name, bytes } of results) {
    zip.file(name, bytes);
  }
  const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
  const url = createObjectURL(blob);
  const link = createAnchorElement();
  link.href = url;
  link.download = filename;
  link.click();
  scheduleRevoke(url);
  return url;
}

/**
 * Show or hide the DOCX options fieldset based on whether the file list
 * contains a .docx file.
 * @param {Iterable<{name: string}>|null|undefined} fileList
 * @param {{hidden: boolean}} fieldsetEl
 * @param {{hasDocx?: Function}} [deps]
 */
export function updateDocxOptionsVisibility(fileList, fieldsetEl, deps = {}) {
  const { hasDocx } = deps;
  const has = typeof hasDocx === 'function' ? hasDocx(fileList) : defaultHasDocx(fileList);
  fieldsetEl.hidden = !has;
}

function defaultHasDocx(fileList) {
  if (!fileList) return false;
  for (const f of fileList) {
    if (f && f.name && typeof f.name === 'string' && f.name.toLowerCase().endsWith('.docx')) {
      return true;
    }
  }
  return false;
}

/**
 * Read the DOCX options fieldset and return a plain options object.
 * @param {{querySelector: Function}} fieldset
 * @returns {{stripMetadata: boolean, stripTrackedChanges: boolean, stripComments: boolean}}
 */
export function readDocxOptions(fieldset) {
  return {
    stripMetadata: !!fieldset.querySelector('#strip-metadata')?.checked,
    stripTrackedChanges: !!fieldset.querySelector('#strip-tracked-changes')?.checked,
    stripComments: !!fieldset.querySelector('#strip-comments')?.checked,
  };
}

function defaultAppendError(errorsList, name, message) {
  if (typeof document === 'undefined') return;
  const li = document.createElement('li');
  li.textContent = `${name}: ${message}`;
  errorsList.append(li);
}

function defaultSetStatus(el, text) {
  if (!el) return;
  el.textContent = text;
}

function defaultClearChildren(el) {
  if (typeof document === 'undefined' || !el) return;
  while (el.firstChild) el.removeChild(el.firstChild);
}

/**
 * Orchestrate the compress-and-download flow.
 * Reads files from the input, validates them, routes each to compressPdf or
 * compressDocx by extension, and triggers downloadResults for the compressed set.
 * @param {{
 *   input: {files: Iterable},
 *   level: {value: string},
 *   docxOptions: {querySelector: Function},
 *   status: object,
 *   errorsList: object,
 *   convertButton?: {disabled: boolean},
 * }} elements
 * @param {{
 *   filterFiles?: Function,
 *   compressPdf?: Function,
 *   compressDocx?: Function,
 *   downloadResults?: Function,
 *   arrayBuffer?: Function,
 *   appendError?: Function,
 *   setStatus?: Function,
 *   clearChildren?: Function,
 * }} [deps]
 * @returns {Promise<void>}
 */
export async function onCompressClick(elements, deps = {}) {
  const {
    filterFiles: filterFn = filterFiles,
    compressPdf: compressPdfFn = compressPdf,
    compressDocx: compressDocxFn = compressDocx,
    downloadResults: downloadFn = downloadResults,
    arrayBuffer = (file) => file.arrayBuffer(),
    appendError = defaultAppendError,
    setStatus = defaultSetStatus,
    clearChildren = defaultClearChildren,
  } = deps;

  const { input, level, docxOptions, status, errorsList, convertButton } = elements;

  const errors = [];
  const files = Array.from(input.files || []);
  const accepted = filterFn(files, errors);

  clearChildren(errorsList);
  for (const { name, message } of errors) {
    appendError(errorsList, name, message);
  }

  if (accepted.length === 0) {
    setStatus(status, 'No valid files to compress.');
    return;
  }

  if (convertButton) convertButton.disabled = true;
  setStatus(
    status,
    `Compressing ${accepted.length} file${accepted.length === 1 ? '' : 's'}...`
  );

  try {
    const results = [];
    const docxOpts = readDocxOptions(docxOptions);
    const compressionLevel = level?.value ?? 'medium';

    for (const file of accepted) {
      try {
        const buffer = await arrayBuffer(file);
        const bytes = new Uint8Array(buffer);
        const lower = file.name.toLowerCase();
        const dotIdx = lower.lastIndexOf('.');
        const ext = dotIdx >= 0 ? lower.slice(dotIdx) : '';
        let outBytes;
        let outName;

        if (!FILE_EXTENSIONS.has(ext)) {
          appendError(errorsList, file.name, `Unsupported extension: ${ext}`);
          continue;
        }
        if (ext === '.pdf') {
          outBytes = await compressPdfFn(bytes);
          outName = file.name.replace(/\.pdf$/i, '') + '.compressed.pdf';
        } else {
          outBytes = await compressDocxFn(bytes, compressionLevel, docxOpts);
          outName = file.name.replace(/\.docx$/i, '') + '.compressed.docx';
        }
        results.push({ name: outName, bytes: outBytes });
      } catch (err) {
        appendError(errorsList, file.name, err.message);
      }
    }

    if (results.length > 0) {
      try {
        await downloadFn(results);
        setStatus(
          status,
          `Compressed ${results.length} file${results.length === 1 ? '' : 's'}.`
        );
      } catch (err) {
        appendError(errorsList, 'Download', err.message);
        setStatus(status, 'Compression finished but download failed.');
      }
    } else {
      setStatus(status, 'No files were compressed.');
    }
  } finally {
    if (convertButton) convertButton.disabled = false;
  }
}

// DOM attach — runs only in a browser environment.
function attachUi() {
  const fileInput = document.getElementById('files');
  const levelSelect = document.getElementById('level');
  const docxOptions = document.getElementById('docx-options');
  const status = document.getElementById('status');
  const errorsList = document.getElementById('errors');
  const convertButton = document.getElementById('convert');

  if (!fileInput || !levelSelect || !docxOptions || !status || !errorsList || !convertButton) {
    return;
  }

  const state = { files: [], errors: [] };

  function refresh() {
    defaultClearChildren(errorsList);
    for (const { name, message } of state.errors) {
      defaultAppendError(errorsList, name, message);
    }
    if (state.files.length === 0) {
      defaultSetStatus(status, 'No files selected.');
      convertButton.disabled = true;
    } else {
      defaultSetStatus(
        status,
        `${state.files.length} file${state.files.length === 1 ? '' : 's'} ready to compress.`
      );
      convertButton.disabled = false;
    }
  }

  fileInput.addEventListener('change', () => {
    const errors = [];
    state.files = filterFiles(Array.from(fileInput.files || []), errors);
    state.errors = errors;
    updateDocxOptionsVisibility(fileInput.files, docxOptions);
    refresh();
  });

  convertButton.addEventListener('click', () => {
    onCompressClick(
      { input: fileInput, level: levelSelect, docxOptions, status, errorsList, convertButton },
      {}
    );
  });

  refresh();
}

if (typeof document !== 'undefined') attachUi();
