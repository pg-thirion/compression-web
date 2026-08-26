// app.js — ES module loaded by index.html

const ALLOWED_LEVELS = new Set(['low', 'medium', 'high']);
const FILE_MAX_BYTES = 100 * 1024 * 1024;
const FILE_EXTENSIONS = new Set(['.pdf', '.docx']);
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

// Flipped semantics: Low = lightest compression (best fidelity), High = strongest compression (smallest files)
const LEVEL_QUALITY = { low: 0.85, medium: 0.7, high: 0.5 };
const LEVEL_MAX_DIM = { low: Infinity, medium: 2000, high: 1500 };

function getExtension(name) {
    const lower = String(name || '').toLowerCase();
    const dotIdx = lower.lastIndexOf('.');
    return dotIdx >= 0 ? lower.slice(dotIdx) : '';
}

// ===== parseLevel =====

export function parseLevel(value) {
    if (typeof value !== 'string') {
        throw new RangeError('Level must be a string');
    }
    if (!ALLOWED_LEVELS.has(value)) {
        throw new RangeError('Level must be one of low, medium, or high');
    }
    return value;
}

// ===== validateFile =====

export function validateFile(file) {
    if (!file || typeof file.name !== 'string') {
        return 'Invalid file.';
    }
    if (Number.isNaN(file.size)) {
        return 'File size is invalid.';
    }
    if (file.size < 0) {
        return 'File size must not be negative.';
    }
    const ext = getExtension(file.name);
    if (!FILE_EXTENSIONS.has(ext)) {
        return 'Unsupported file type. Use PDF or DOCX.';
    }
    if (file.size > FILE_MAX_BYTES) {
        return 'File exceeds the 100 MB size limit.';
    }
    if (file.size < 0) {
        return 'File size must not be negative.';
    }
    return null;
}

// ===== filterFiles =====

export function filterFiles(files, errors) {
    if (!Array.isArray(errors)) {
        throw new TypeError('errors must be an array');
    }
    const accepted = [];
    for (const file of files || []) {
        const message = validateFile(file);
        if (message === null) {
            accepted.push(file);
        } else {
            errors.push(message);
        }
    }
    return accepted;
}

// ===== minimalCoreXml / minimalAppXml =====

export function minimalCoreXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        + 'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        + 'xmlns:dcterms="http://purl.org/dc/terms/" '
        + 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        + 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
        + '</cp:coreProperties>\n';
}

export function minimalAppXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        + 'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\n'
        + '</Properties>\n';
}

// ===== stripTrackedChanges =====

export function stripTrackedChanges(xml) {
    if (typeof xml !== 'string') return xml;
    // Remove <w:del>...</w:del> entirely (deleted text is gone)
    let out = xml.replace(/<w:del(?:\s[^>]*)?>[\s\S]*?<\/w:del>/g, '');
    // Unwrap <w:ins>...</w:ins> keeping content
    out = out.replace(/<w:ins(?:\s[^>]*)?>([\s\S]*?)<\/w:ins>/g, '$1');
    return out;
}

// ===== stripCommentMarkers =====

export function stripCommentMarkers(xml) {
    if (typeof xml !== 'string') return xml;
    return xml
        .replace(/<w:commentRangeStart(?:\s[^>]*)?\/?>/g, '')
        .replace(/<w:commentRangeEnd(?:\s[^>]*)?\/?>/g, '')
        .replace(/<w:commentReference(?:\s[^>]*)?\/?>/g, '');
}

// ===== default deps =====

function defaultCompressPdfDeps() {
    if (typeof globalThis === 'undefined' || !globalThis.PDFLib || !globalThis.PDFLib.PDFDocument) {
        throw new Error('pdf-lib is not loaded; include lib/pdf-lib.min.js before app.js');
    }
    if (typeof globalThis !== 'undefined' && globalThis.pdfjsLib && globalThis.pdfjsLib.GlobalWorkerOptions) {
        globalThis.pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';
    }
    return {
        PDFDocument: globalThis.PDFLib.PDFDocument,
        pdfjsLib: globalThis.pdfjsLib,
        createCanvas: defaultCreateCanvas,
    };
}

function defaultCompressDocxDeps() {
    if (typeof globalThis === 'undefined' || typeof globalThis.JSZip !== 'function') {
        throw new Error('JSZip is not loaded; include lib/jszip.min.js before app.js');
    }
    return {
        JSZip: globalThis.JSZip,
        reencodeMediaImage,
        createImageBitmap: globalThis.createImageBitmap,
        createCanvas: defaultCreateCanvas,
    };
}

function defaultCreateCanvas(width, height) {
    if (typeof document === 'undefined') {
        throw new Error('createCanvas is not available in this environment');
    }
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    return canvas;
}

// ===== compressPdf =====

export async function compressPdf(bytes, level, deps = defaultCompressPdfDeps()) {
    // Phase 1: structural compression (metadata wipe + object streams)
    let pdf;
    let structuralBytes = bytes;
    try {
        pdf = await deps.PDFDocument.load(bytes, { updateMetadata: false });
        pdf.setTitle('');
        pdf.setAuthor('');
        pdf.setSubject('');
        pdf.setKeywords([]);
        pdf.setCreator('');
        pdf.setProducer('');
        structuralBytes = await pdf.save({ useObjectStreams: true });
    } catch {
        structuralBytes = bytes;
    }
    let bestBytes = structuralBytes.byteLength < bytes.byteLength ? structuralBytes : bytes;

    // Phase 2: raster re-encode at Medium/High (only if pdf.js is available)
    if (level !== 'low' && deps.pdfjsLib && bestBytes.byteLength > 0) {
        const quality = LEVEL_QUALITY[level];
        try {
            const rasterBytes = await rasterReencodePdf(bestBytes, quality, deps);
            if (rasterBytes.byteLength < bestBytes.byteLength) {
                bestBytes = rasterBytes;
            }
        } catch {
            // fall back to structural
        }
    }

    // Validate by re-loading and checking page count
    try {
        const expectedPages = pdf
            ? pdf.getPageCount()
            : (await deps.PDFDocument.load(bytes, { updateMetadata: false })).getPageCount();
        const reloaded = await deps.PDFDocument.load(bestBytes, { updateMetadata: false });
        if (reloaded.getPageCount() !== expectedPages) {
            bestBytes = bytes;
        }
    } catch {
        bestBytes = bytes;
    }

    return new Blob([bestBytes], { type: 'application/pdf' });
}

async function rasterReencodePdf(inputBytes, quality, deps) {
    const pdfJsDoc = await deps.pdfjsLib.getDocument({ data: inputBytes.slice() }).promise;
    try {
        const newPdf = await deps.PDFDocument.create();
        for (let i = 1; i <= pdfJsDoc.numPages; i++) {
            const page = await pdfJsDoc.getPage(i);
            const viewport = page.getViewport({ scale: 1.0 });
            const canvas = deps.createCanvas(viewport.width, viewport.height);
            const ctx = canvas.getContext('2d');
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, viewport.width, viewport.height);
            await page.render({ canvasContext: ctx, viewport }).promise;
            const blob = await new Promise((resolve, reject) => {
                canvas.toBlob(b => (b ? resolve(b) : reject(new Error('toBlob failed'))), 'image/jpeg', quality);
            });
            const jpegBytes = new Uint8Array(await blob.arrayBuffer());
            const embedded = await newPdf.embedJpg(jpegBytes);
            const newPage = newPdf.addPage([viewport.width, viewport.height]);
            newPage.drawImage(embedded, { x: 0, y: 0, width: viewport.width, height: viewport.height });
        }
        return await newPdf.save({ useObjectStreams: true });
    } finally {
        await pdfJsDoc.destroy();
    }
}

// ===== reencodeMediaImage =====

export async function reencodeMediaImage(bytes, sourceMime, targetMime, quality, deps) {
    if (!deps) {
        deps = {
            createImageBitmap: globalThis.createImageBitmap,
            createCanvas: defaultCreateCanvas,
            maxDimension: Infinity,
        };
    }
    const sourceBlob = new Blob([bytes], { type: sourceMime });
    const bitmap = await deps.createImageBitmap(sourceBlob);
    try {
        let w = bitmap.width;
        let h = bitmap.height;
        if (deps.maxDimension && Number.isFinite(deps.maxDimension)) {
            const longest = Math.max(w, h);
            if (longest > deps.maxDimension) {
                const scale = deps.maxDimension / longest;
                w = Math.max(1, Math.round(w * scale));
                h = Math.max(1, Math.round(h * scale));
            }
        }
        const canvas = deps.createCanvas(w, h);
        const ctx = canvas.getContext('2d');
        if (targetMime === 'image/jpeg') {
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, w, h);
        }
        ctx.drawImage(bitmap, 0, 0, w, h);
        const outBlob = await new Promise((resolve, reject) => {
            canvas.toBlob(
                (b) => (b ? resolve(b) : reject(new Error('toBlob failed'))),
                targetMime,
                targetMime === 'image/png' ? undefined : quality
            );
        });
        return new Uint8Array(await outBlob.arrayBuffer());
    } finally {
        if (bitmap && typeof bitmap.close === 'function') {
            bitmap.close();
        }
    }
}

// ===== compressDocx =====

export async function compressDocx(bytes, level, options, deps) {
    if (!deps) {
        deps = defaultCompressDocxDeps();
    }
    options = options || {};
    const originalBytes = bytes;
    try {
        const zip = new deps.JSZip();
        await zip.loadAsync(bytes);
        // Replace core.xml + app.xml with stubs
        if (options.stripMetadata !== false) {
            zip.file('docProps/core.xml', minimalCoreXml());
            zip.file('docProps/app.xml', minimalAppXml());
        }
        // Apply XML edits to document.xml
        if (zip.file('word/document.xml') !== null) {
            let docXml = await zip.fileAsync('word/document.xml');
            let updated = docXml;
            if (options.stripTrackedChanges) {
                updated = stripTrackedChanges(updated);
            }
            if (options.stripComments) {
                updated = stripCommentMarkers(updated);
            }
            if (updated !== docXml) {
                zip.file('word/document.xml', updated);
            }
        }
        // Re-encode media images
        const quality = LEVEL_QUALITY[level];
        const maxDim = LEVEL_MAX_DIM[level];
        const convertPng = level !== 'low';

        const mediaPaths = [];
        for (const path of Object.keys(zip.files)) {
            if (!path.startsWith('word/media/')) continue;
            const lower = path.toLowerCase();
            const ext = lower.slice(lower.lastIndexOf('.') + 1);
            if (ext !== 'png' && ext !== 'jpg' && ext !== 'jpeg') continue;
            mediaPaths.push(path);
        }

        const reencodeFn = deps.reencodeMediaImage || reencodeMediaImage;
        for (const path of mediaPaths) {
            const ext = path.toLowerCase().slice(path.toLowerCase().lastIndexOf('.') + 1);
            // At low level, PNGs are not re-encoded (would be a no-op PNG→PNG)
            if (ext === 'png' && level === 'low') continue;
            const data = await zip.fileAsync(path);
            const sourceMime = ext === 'png' ? 'image/png' : 'image/jpeg';
            const targetMime = ext === 'png' ? 'image/jpeg' : 'image/jpeg';
            try {
                const reencoded = await reencodeFn(data, sourceMime, targetMime, quality, {
                    createImageBitmap: deps.createImageBitmap,
                    createCanvas: deps.createCanvas,
                    maxDimension: maxDim,
                });
                if (reencoded && reencoded.byteLength < data.byteLength) {
                    zip.file(path, reencoded);
                }
            } catch (e) {
                // Skip this image on failure, keep original
            }
        }
        const out = await zip.generateAsync({
            type: 'uint8array',
            compression: 'DEFLATE',
            compressionOptions: { level: 9 },
        });
        // Validate by reopening
        try {
            const reopen = new deps.JSZip();
            await reopen.loadAsync(out);
            const docEntry = reopen.file('word/document.xml');
            if (!docEntry) {
                return new Blob([originalBytes], { type: DOCX_MIME });
            }
        } catch (e) {
            return new Blob([originalBytes], { type: DOCX_MIME });
        }
        return new Blob([out], { type: DOCX_MIME });
    } catch (err) {
        return new Blob([originalBytes], { type: DOCX_MIME });
    }
}

// ===== downloadResults =====

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

export async function downloadResults(results, deps = {}) {
    const JSZip = deps.JSZip ?? globalThis.JSZip;
    const createObjectURL = deps.createObjectURL ?? defaultCreateObjectURL;
    const createAnchorElement = deps.createAnchorElement ?? defaultCreateAnchorElement;
    const scheduleRevoke = deps.scheduleRevoke ?? defaultScheduleRevoke;
    if (typeof JSZip !== 'function') {
        throw new Error('JSZip is not available');
    }
    const zip = new JSZip();
    const usedNames = new Set();
    for (const result of results) {
        const originalName = result.file.name;
        const dotIdx = originalName.lastIndexOf('.');
        const base = dotIdx > 0 ? originalName.slice(0, dotIdx) : originalName;
        const ext = dotIdx > 0 ? originalName.slice(dotIdx) : '';
        let outName = `${base}.compressed${ext}`;
        let counter = 2;
        while (usedNames.has(outName.toLowerCase())) {
            outName = `${base}-${counter}.compressed${ext}`;
            counter += 1;
        }
        usedNames.add(outName.toLowerCase());
        zip.file(outName, result.blob);
    }
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const url = createObjectURL(zipBlob);
    const link = createAnchorElement();
    link.href = url;
    const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    link.download = `compressed-${ts}.zip`;
    link.click();
    scheduleRevoke(url);
    return url;
}

// ===== Session history & metrics =====

const session = {
    history: [],
    totalSavedBytes: 0,
    filesProcessed: 0,
};

export function getSession() {
    return session;
}

export function resetSession() {
    session.history.length = 0;
    session.totalSavedBytes = 0;
    session.filesProcessed = 0;
}

export function recordCompression(fileName, originalBytes, compressedBytes) {
    const savedBytes = Math.max(0, originalBytes - compressedBytes);
    const percentSaved = originalBytes > 0 ? (savedBytes / originalBytes) * 100 : 0;
    session.history.unshift({ fileName, originalBytes, compressedBytes, savedBytes, percentSaved, timestamp: Date.now() });
    if (session.history.length > 20) {
        session.history.length = 20;
    }
    session.totalSavedBytes += savedBytes;
    session.filesProcessed += 1;
    renderSession();
}

export function renderSession() {
    if (typeof document === 'undefined') return;
    const summary = document.getElementById('summary');
    const history = document.getElementById('history');
    const list = document.getElementById('history-list');
    if (!summary || !history || !list) return;
    if (session.filesProcessed > 0) {
        summary.hidden = false;
        history.hidden = false;
        const totalEl = document.getElementById('totalSavedMB');
        const filesEl = document.getElementById('filesProcessed');
        if (totalEl) totalEl.textContent = (session.totalSavedBytes / (1024 * 1024)).toFixed(2) + ' MB';
        if (filesEl) filesEl.textContent = String(session.filesProcessed);
        while (list.firstChild) list.removeChild(list.firstChild);
        for (const entry of session.history) {
            const li = document.createElement('li');
            const nameSpan = document.createElement('span');
            nameSpan.className = 'hist-name';
            nameSpan.textContent = entry.fileName;
            const savedSpan = document.createElement('span');
            savedSpan.className = 'hist-saved';
            const saved = (entry.savedBytes / 1024).toFixed(1);
            const pct = entry.percentSaved.toFixed(1);
            savedSpan.textContent = '−' + saved + ' KB (−' + pct + '%)';
            li.appendChild(nameSpan);
            li.appendChild(document.createTextNode(' '));
            li.appendChild(savedSpan);
            list.appendChild(li);
        }
    } else {
        summary.hidden = true;
        history.hidden = true;
    }
}

// ===== UI attach =====

function getElement(id) {
    if (typeof document === 'undefined') return null;
    return document.getElementById(id);
}

function setText(el, text) {
    if (!el) return;
    while (el.firstChild) el.removeChild(el.firstChild);
    el.append(document.createTextNode(text));
}

function renderErrors(el, errors) {
    if (!el) return;
    while (el.firstChild) el.removeChild(el.firstChild);
    for (const msg of errors) {
        const li = document.createElement('li');
        li.append(document.createTextNode(msg));
        el.append(li);
    }
}

function hasDocx(files) {
    for (const f of files) {
        if (getExtension(f.name) === '.docx') return true;
    }
    return false;
}

async function onConvertClick(state, levelSelect, docxOptions, downloadFn, compressPdfFn, compressDocxFn, recordFn) {
    const errors = [];
    const accepted = filterFiles(state.files, errors);
    if (!accepted.length) {
        return { errors, results: [] };
    }
    let level;
    try {
        level = parseLevel(levelSelect.value);
    } catch (err) {
        errors.push(err.message);
        return { errors, results: [] };
    }
    const results = [];
    for (const file of accepted) {
        try {
            const bytes = new Uint8Array(await file.arrayBuffer());
            const ext = getExtension(file.name);
            let blob;
            if (ext === '.pdf') {
                blob = await compressPdfFn(bytes, level);
            } else if (ext === '.docx') {
                blob = await compressDocxFn(bytes, level, docxOptions);
            } else {
                continue;
            }
            if (recordFn) {
                recordFn(file.name, file.size, blob.size);
            }
            results.push({ file, blob });
        } catch (err) {
            errors.push(`${file.name}: ${err.message}`);
        }
    }
    return { errors, results };
}

function attachUi() {
    const dropZone = getElement('drop-zone');
    const fileInput = getElement('files');
    const levelSelect = getElement('level');
    const convertButton = getElement('convert');
    const statusEl = getElement('status');
    const errorsEl = getElement('errors');
    const docxOptionsFieldset = getElement('docx-options');
    const stripTrackedChangesCb = getElement('strip-tracked-changes');
    const stripCommentsCb = getElement('strip-comments');
    const stripMetadataCb = getElement('strip-metadata');

    if (!fileInput || !convertButton) return;

    const state = { files: [], errors: [], processing: false };

    function updateDocxOptionsVisibility() {
        if (!docxOptionsFieldset) return;
        docxOptionsFieldset.hidden = !hasDocx(state.files);
    }

    function updateConvertButton() {
        convertButton.disabled = state.processing || state.files.length === 0;
    }

    function refreshStatus() {
        const n = state.files.length;
        setText(statusEl, n === 0 ? 'No files selected.' : `${n} file${n === 1 ? '' : 's'} ready to compress.`);
    }

    function acceptFiles(fileList) {
        const incoming = Array.from(fileList || []);
        const errors = [];
        const accepted = filterFiles(incoming, errors);
        state.files = accepted;
        state.errors = errors;
        renderErrors(errorsEl, errors);
        refreshStatus();
        updateDocxOptionsVisibility();
        updateConvertButton();
    }

    function preventDefaults(e) { e.preventDefault(); e.stopPropagation(); }

    function handleDrop(e) {
        if (state.processing) return;
        preventDefaults(e);
        if (e.dataTransfer && e.dataTransfer.files) acceptFiles(e.dataTransfer.files);
    }

    function openFilePicker() {
        if (state.processing) return;
        fileInput.click();
    }

    if (dropZone) {
        dropZone.addEventListener('dragenter', preventDefaults);
        dropZone.addEventListener('dragover', preventDefaults);
        dropZone.addEventListener('drop', handleDrop);
        dropZone.addEventListener('click', openFilePicker);
        dropZone.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openFilePicker(); }
        });
    }

    fileInput.addEventListener('change', () => acceptFiles(fileInput.files));

    convertButton.addEventListener('click', async () => {
        const docxOptions = {
            stripMetadata: stripMetadataCb ? stripMetadataCb.checked : true,
            stripTrackedChanges: stripTrackedChangesCb ? stripTrackedChangesCb.checked : false,
            stripComments: stripCommentsCb ? stripCommentsCb.checked : false,
        };
        state.processing = true;
        updateConvertButton();
        setText(statusEl, `Compressing ${state.files.length} file${state.files.length === 1 ? '' : 's'}...`);
        const { errors, results } = await onConvertClick(
            state,
            levelSelect,
            docxOptions,
            downloadResults,
            compressPdf,
            compressDocx,
            recordCompression
        );
        state.errors.push(...errors);
        renderErrors(errorsEl, state.errors);
        if (results.length > 0) {
            try {
                await downloadResults(results);
                setText(statusEl, `Compressed ${results.length} file${results.length === 1 ? '' : 's'}.`);
            } catch (err) {
                state.errors.push(`Download failed: ${err.message}`);
                renderErrors(errorsEl, state.errors);
                setText(statusEl, 'Compression finished but download failed.');
            }
        } else {
            setText(statusEl, 'No files were compressed.');
        }
        state.processing = false;
        updateConvertButton();
    });

    // Initialize
    updateDocxOptionsVisibility();
    refreshStatus();
    updateConvertButton();
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', attachUi);
    } else {
        attachUi();
    }
}