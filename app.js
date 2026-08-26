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
 * Re-encode an image at the supplied quality via Canvas.
 * Output format matches input format so PNG transparency is preserved.
 * If maxDimension is set and the image exceeds it on its longest side,
 * the image is downscaled before re-encoding.
 * @param {Uint8Array} bytes
 * @param {string} mimeType
 * @param {number} quality
 * @param {{createImageBitmap?: Function, OffscreenCanvas?: Function, maxDimension?: number}} [deps]
 * @returns {Promise<Uint8Array>}
 */
export async function reencodeMediaImage(bytes, mimeType, quality, deps = {}) {
  const {
    createImageBitmap = globalThis.createImageBitmap,
    OffscreenCanvas = globalThis.OffscreenCanvas,
    maxDimension,
  } = deps;

  if (!createImageBitmap) throw new Error('createImageBitmap not available');
  if (!OffscreenCanvas) throw new Error('OffscreenCanvas not available');

  const blob = new Blob([bytes], { type: mimeType });
  const bitmap = await createImageBitmap(blob);
  try {
    let canvasWidth = bitmap.width;
    let canvasHeight = bitmap.height;
    let drawWidth = canvasWidth;
    let drawHeight = canvasHeight;
    if (maxDimension && maxDimension > 0) {
      const longest = Math.max(canvasWidth, canvasHeight);
      if (longest > maxDimension) {
        const scale = maxDimension / longest;
        drawWidth = Math.max(1, Math.round(canvasWidth * scale));
        drawHeight = Math.max(1, Math.round(canvasHeight * scale));
        canvasWidth = drawWidth;
        canvasHeight = drawHeight;
      }
    }
    const canvas = new OffscreenCanvas(canvasWidth, canvasHeight);
    const ctx = canvas.getContext('2d');
    ctx.drawImage(bitmap, 0, 0, drawWidth, drawHeight);
    const outBlob = await canvas.convertToBlob({ type: mimeType, quality });
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
    stripTrackedChanges: shouldStripTrackedChanges = level === 'high' || level === 'max',
    stripComments: shouldStripComments = level === 'high',
  } = options;

  const LEVEL_IMAGE = {
    low:    { quality: 0.61, maxDim: 1600 },
    medium: { quality: 0.53, maxDim: 1200 },
    high:   { quality: 0.46, maxDim: 1024 },
    max:    { quality: 0.40, maxDim: 900 },
  };
  const ls = LEVEL_IMAGE[level];

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
      const isJpeg = ext === '.jpg' || ext === '.jpeg';
      const isPng = ext === '.png';
      const isBmp = ext === '.bmp';
      const isTiff = ext === '.tif' || ext === '.tiff';

      if (!isJpeg && !isPng && !isBmp && !isTiff) continue;

      let mimeType;
      let outputExt;
      let useMaxDimension;
      let levelQuality;
      if (!ls) continue;
      useMaxDimension = ls.maxDim;
      levelQuality = ls.quality;
      if (isBmp || isTiff) {
        mimeType = 'image/jpeg';
        outputExt = '.jpg';
      } else if (isJpeg) {
        mimeType = 'image/jpeg';
        outputExt = ext;
      } else {
        mimeType = 'image/png';
        outputExt = ext;
      }

      const origBytes = await zip.file(name).async('uint8array');
      const reencoded = await reencode(origBytes, mimeType, levelQuality, {
        maxDimension: useMaxDimension,
      });
      if (reencoded.length < origBytes.length) {
        if (outputExt !== ext) {
          const newName = name.slice(0, -ext.length) + outputExt;
          zip.file(newName, reencoded);
          zip.remove(name);
          try {
            const relsPath = 'word/_rels/document.xml.rels';
            const relsEntry = zip.file(relsPath);
            if (relsEntry) {
              const relsXml = await relsEntry.async('string');
              const oldTarget = name.slice('word/'.length);
              const newTarget = newName.slice('word/'.length);
              const escapedOld = oldTarget.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              const escapedNew = newTarget.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              const relsPattern = new RegExp(
                `(<Relationship[^>]*Target=")${escapedOld}([^"]*"[^>]*\\/>)`,
                'g'
              );
              const updatedRels = relsXml.replace(relsPattern, `$1${escapedNew}$2`);
              if (updatedRels !== relsXml) {
                zip.file(relsPath, updatedRels);
              }
            }
          } catch (err) {
            // Rel update failed; new file is orphaned but document still parses.
          }
        } else {
          zip.file(name, reencoded);
        }
      }
    }

    if (level === 'high') {
      const xmlParts = Object.keys(zip.files).filter(
        (name) => !zip.files[name].dir && /\.xml$|\.rels$/i.test(name)
      );
      for (const name of xmlParts) {
        const entry = zip.file(name);
        if (!entry) continue;
        try {
          const original = await entry.async('string');
          const minified = original
            .replace(/>\s+</g, '><')
            .replace(/\s+</g, '<')
            .replace(/>\s+/g, '>')
            .trim();
          if (minified.length < original.length) {
            zip.file(name, minified);
          }
        } catch (err) {
          // Non-text content (e.g. binary disguised as XML) — leave alone.
        }
      }

      const fontNames = Object.keys(zip.files).filter(
        (name) => !zip.files[name].dir && /^word\/fonts\//i.test(name)
      );
      for (const name of fontNames) {
        zip.remove(name);
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
 * Download compressed results. With a single file, download it directly
 * (no ZIP wrapper). With multiple files, bundle them in a ZIP.
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

  let blob;
  let downloadName;

  if (results.length === 1) {
    blob = new Blob([results[0].bytes]);
    downloadName = results[0].name;
  } else {
    if (typeof JSZipCtor !== 'function') {
      throw new Error('JSZip not loaded');
    }
    const zip = new JSZipCtor();
    for (const { name, bytes } of results) {
      zip.file(name, bytes);
    }
    blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
    downloadName = filename;
  }

  const url = createObjectURL(blob);
  const link = createAnchorElement();
  link.href = url;
  link.download = downloadName;
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
 * Format a byte count as a human-readable string (B / KB / MB).
 * @param {number} n
 * @returns {string}
 */
export function formatBytes(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * Render the selected file list into a `<ul>` element. Each file becomes an
 * `<li>` showing name, formatted size, and lowercased extension (or 'unknown').
 * @param {Iterable<{name: string, size: number}>} files
 * @param {{replaceChildren?: Function, createElement?: Function, appendChild?: Function}|null} fileListEl
 * @param {{
 *   formatBytes?: (n: number) => string,
 *   createElement?: (tag: string) => any,
 *   appendChild?: (parent: any, child: any) => void,
 * }} [deps]
 */
export function renderFileList(files, fileListEl, deps = {}) {
  const {
    formatBytes: formatFn = formatBytes,
    createElement = (tag) => fileListEl && fileListEl.createElement
      ? fileListEl.createElement(tag)
      : (typeof document !== 'undefined' ? document.createElement(tag) : null),
    appendChild = (parent, child) => parent.appendChild(child),
  } = deps;

  if (!fileListEl) return;

  if (typeof fileListEl.replaceChildren === 'function') {
    fileListEl.replaceChildren();
  }

  for (const file of files || []) {
    const li = createElement('li');
    const size = formatFn(file.size);
    const ext = (file.name.toLowerCase().match(/\.([^.]+)$/) || [, 'unknown'])[1];
    if (li && 'textContent' in li) {
      li.textContent = `${file.name} — ${size} — ${ext}`;
    }
    appendChild(fileListEl, li);
  }
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
  const fileListEl = document.getElementById('file-list');
  const dropZone = document.getElementById('drop-zone');

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
    renderFileList(state.files, fileListEl);
  }

  fileInput.addEventListener('change', () => {
    const errors = [];
    state.files = filterFiles(Array.from(fileInput.files || []), errors);
    state.errors = errors;
    updateDocxOptionsVisibility(fileInput.files, docxOptions);
    refresh();
  });

  function acceptDroppedFiles(fileList) {
    const dt = new DataTransfer();
    for (const f of Array.from(fileList || [])) dt.items.add(f);
    fileInput.files = dt.files;
    const errors = [];
    state.files = filterFiles(Array.from(fileInput.files || []), errors);
    state.errors = errors;
    updateDocxOptionsVisibility(fileInput.files, docxOptions);
    refresh();
  }

  if (dropZone) {
    ['dragenter', 'dragover'].forEach((evt) => {
      dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('is-dragover');
      });
    });
    ['dragleave', 'dragend'].forEach((evt) => {
      dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('is-dragover');
      });
    });
    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('is-dragover');
      acceptDroppedFiles(e.dataTransfer ? e.dataTransfer.files : null);
    });
    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        fileInput.click();
      }
    });
  }

  convertButton.addEventListener('click', () => {
    onCompressClick(
      { input: fileInput, level: levelSelect, docxOptions, status, errorsList, convertButton },
      {}
    );
  });

  refresh();
}

if (typeof document !== 'undefined') attachUi();
