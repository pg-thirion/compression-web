import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
    parseLevel,
    validateFile,
    filterFiles,
    compressPdf,
    compressDocx,
    reencodeMediaImage,
    stripTrackedChanges,
    stripCommentMarkers,
    minimalCoreXml,
    minimalAppXml,
    downloadResults,
    recordCompression,
    renderSession,
    getSession,
    resetSession,
} from '../app.js';

// ===== parseLevel =====

test('parseLevel accepts "low"', () => {
    assert.equal(parseLevel('low'), 'low');
});

test('parseLevel accepts "medium"', () => {
    assert.equal(parseLevel('medium'), 'medium');
});

test('parseLevel accepts "high"', () => {
    assert.equal(parseLevel('high'), 'high');
});

test('parseLevel rejects empty string', () => {
    assert.throws(() => parseLevel(''), RangeError);
});

test('parseLevel rejects "extreme"', () => {
    assert.throws(() => parseLevel('extreme'), RangeError);
});

test('parseLevel rejects "LOW" (case-sensitive)', () => {
    assert.throws(() => parseLevel('LOW'), RangeError);
});

test('parseLevel rejects " low " with whitespace', () => {
    assert.throws(() => parseLevel(' low '), RangeError);
});

test('parseLevel rejects null', () => {
    assert.throws(() => parseLevel(null), RangeError);
});

// ===== validateFile =====

test('validateFile accepts a normal PDF', () => {
    assert.equal(validateFile({ name: 'doc.pdf', size: 1024 }), null);
});

test('validateFile accepts a normal DOCX', () => {
    assert.equal(validateFile({ name: 'report.docx', size: 1024 }), null);
});

test('validateFile rejects a .gif', () => {
    assert.match(validateFile({ name: 'cat.gif', size: 1024 }), /unsupported/i);
});

test('validateFile rejects an oversize PDF', () => {
    assert.match(validateFile({ name: 'big.pdf', size: 200 * 1024 * 1024 }), /100.*MB/i);
});

test('validateFile rejects an oversize DOCX', () => {
    assert.match(validateFile({ name: 'big.docx', size: 200 * 1024 * 1024 }), /100.*MB/i);
});

test('validateFile rejects negative size', () => {
    assert.match(validateFile({ name: 'bad.pdf', size: -1 }), /negative/i);
});

test('validateFile rejects NaN size', () => {
    assert.match(validateFile({ name: 'bad.pdf', size: NaN }), /invalid/i);
});

test('validateFile rejects missing file', () => {
    assert.match(validateFile(null), /invalid/i);
});

// ===== compressPdf =====

test('compressPdf: output smaller than input + validation passes', async () => {
    const deps = makePdfDeps({ pages: 3, outputSize: 100 });
    const bytes = new Uint8Array(200);
    const blob = await compressPdf(bytes, 'low', deps);
    assert.equal(blob.size, 100);
    assert.equal(deps.loadCalls.length, 2);
});

test('compressPdf: metadata cleared on output document', async () => {
    const deps = makePdfDeps({ pages: 1, outputSize: 80 });
    await compressPdf(new Uint8Array(200), 'low', deps);
    const w = deps.lastWritten;
    assert.equal(w.title, '');
    assert.equal(w.author, '');
    assert.equal(w.subject, '');
    assert.deepEqual(w.keywords, []);
    assert.equal(w.creator, '');
    assert.equal(w.producer, '');
});

test('compressPdf: uses useObjectStreams:true', async () => {
    const deps = makePdfDeps({ pages: 1, outputSize: 80 });
    await compressPdf(new Uint8Array(200), 'low', deps);
    assert.equal(deps.lastSaveOptions.useObjectStreams, true);
});

test('compressPdf: keeps output when smaller', async () => {
    const deps = makePdfDeps({ pages: 1, outputSize: 80 });
    const blob = await compressPdf(new Uint8Array(200), 'low', deps);
    assert.equal(blob.size, 80);
});

test('compressPdf: falls back to original when output is larger', async () => {
    const deps = makePdfDeps({ pages: 1, outputSize: 500 });
    const blob = await compressPdf(new Uint8Array(200), 'low', deps);
    assert.equal(blob.size, 200);
});

test('compressPdf: falls back to original when validation fails', async () => {
    const deps = makePdfDeps({ pages: 1, outputSize: 80, validationPageCount: 99 });
    const blob = await compressPdf(new Uint8Array(200), 'low', deps);
    assert.equal(blob.size, 200);
});

test('compressPdf: falls back to original on throw', async () => {
    const deps = makePdfDeps({ pages: 1, outputSize: 80, throwOnLoad: true });
    const blob = await compressPdf(new Uint8Array(200), 'low', deps);
    assert.equal(blob.size, 200);
});

test('compressPdf: load called with updateMetadata:false', async () => {
    const deps = makePdfDeps({ pages: 1, outputSize: 80 });
    await compressPdf(new Uint8Array(200), 'low', deps);
    assert.equal(deps.loadCalls[0].updateMetadata, false);
});

test('compressPdf: does not call raster path at low level', async () => {
    const deps = makePdfDeps({ pages: 1, outputSize: 80 });
    await compressPdf(new Uint8Array(200), 'low', deps);
    assert.equal(deps.rasterCalls, 0);
});

test('compressPdf: skips raster path when pdfjsLib is not provided', async () => {
    const deps = makePdfDeps({ pages: 1, outputSize: 80, withPdfJs: false });
    await compressPdf(new Uint8Array(200), 'medium', deps);
    assert.equal(deps.rasterCalls, 0);
});

test('compressPdf: uses raster path at medium/high when pdfjsLib is provided', async () => {
    const deps = makePdfDeps({ pages: 2, outputSize: 100, rasterSize: 50, withPdfJs: true });
    const blob = await compressPdf(new Uint8Array(200), 'medium', deps);
    assert.equal(deps.rasterCalls, 1);
    assert.equal(blob.size, 50);
});

test('compressPdf: keeps structural output when raster is larger', async () => {
    const deps = makePdfDeps({ pages: 2, outputSize: 80, rasterSize: 500, withPdfJs: true });
    const blob = await compressPdf(new Uint8Array(200), 'high', deps);
    assert.equal(deps.rasterCalls, 1);
    assert.equal(blob.size, 80);
});

// ===== minimalCoreXml / minimalAppXml =====

test('minimalCoreXml returns valid-looking empty stub', () => {
    const xml = minimalCoreXml();
    assert.match(xml, /<\?xml/);
    assert.match(xml, /<cp:coreProperties/);
    assert.match(xml, /<\/cp:coreProperties>/);
});

test('minimalAppXml returns valid-looking empty stub', () => {
    const xml = minimalAppXml();
    assert.match(xml, /<\?xml/);
    assert.match(xml, /<Properties/);
    assert.match(xml, /<\/Properties>/);
});

// ===== stripTrackedChanges =====

test('stripTrackedChanges removes <w:del>...</w:del>', () => {
    const xml = '<w:body><w:del>removed</w:del><w:p>keep</w:p></w:body>';
    const out = stripTrackedChanges(xml);
    assert.equal(out.includes('<w:del>'), false);
    assert.match(out, /<w:p>keep<\/w:p>/);
});

test('stripTrackedChanges unwraps <w:ins>...</w:ins>', () => {
    const xml = '<w:body><w:ins>inserted</w:ins><w:p>keep</w:p></w:body>';
    const out = stripTrackedChanges(xml);
    assert.equal(out.includes('<w:ins>'), false);
    assert.match(out, /inserted/);
    assert.match(out, /<w:p>keep<\/w:p>/);
});

// ===== stripCommentMarkers =====

test('stripCommentMarkers removes commentRangeStart/End/Reference self-closing', () => {
    const xml = '<w:body><w:commentRangeStart/><w:p>hello</w:p><w:commentRangeEnd/><w:commentReference/></w:body>';
    const out = stripCommentMarkers(xml);
    assert.equal(out.includes('commentRangeStart'), false);
    assert.equal(out.includes('commentRangeEnd'), false);
    assert.equal(out.includes('commentReference'), false);
    assert.match(out, /<w:p>hello<\/w:p>/);
});

// ===== reencodeMediaImage =====

test('reencodeMediaImage: bitmap to canvas to toBlob to Uint8Array', async () => {
    const out = new Uint8Array([7, 8, 9]);
    const deps = makeImageDeps({ outputBytes: out, width: 10, height: 10 });
    const result = await reencodeMediaImage(new Uint8Array([1, 2, 3]), 'image/jpeg', 'image/jpeg', 0.7, deps);
    assert.deepEqual([...result], [7, 8, 9]);
    assert.equal(deps.toBlobCalledWith.mime, 'image/jpeg');
    assert.equal(deps.toBlobCalledWith.quality, 0.7);
    assert.equal(deps.closed(), true);
});

test('reencodeMediaImage: closes bitmap on toBlob error', async () => {
    const deps = makeImageDeps({ toBlobShouldFail: true, width: 10, height: 10 });
    await assert.rejects(() => reencodeMediaImage(new Uint8Array([1]), 'image/jpeg', 'image/jpeg', 0.5, deps));
    assert.equal(deps.closed(), true);
});

test('reencodeMediaImage: PNG target passes undefined quality to toBlob', async () => {
    const out = new Uint8Array([7, 8, 9]);
    const deps = makeImageDeps({ outputBytes: out, width: 10, height: 10 });
    await reencodeMediaImage(new Uint8Array([1, 2, 3]), 'image/png', 'image/png', 0.7, deps);
    assert.equal(deps.toBlobCalledWith.mime, 'image/png');
    assert.equal(deps.toBlobCalledWith.quality, undefined);
});

test('reencodeMediaImage: downscales when maxDimension is finite and exceeded', async () => {
    const deps = makeImageDeps({ outputBytes: new Uint8Array([1, 2]), width: 4000, height: 2000, maxDimension: 2000 });
    await reencodeMediaImage(new Uint8Array([1]), 'image/jpeg', 'image/jpeg', 0.7, deps);
    assert.equal(deps.createdCanvasWidth, 2000);
    assert.equal(deps.createdCanvasHeight, 1000);
});

test('reencodeMediaImage: does not upscale when image is smaller than maxDimension', async () => {
    const deps = makeImageDeps({ outputBytes: new Uint8Array([1]), width: 800, height: 600 });
    await reencodeMediaImage(new Uint8Array([1]), 'image/jpeg', 'image/jpeg', 0.7, deps);
    assert.equal(deps.createdCanvasWidth, 800);
    assert.equal(deps.createdCanvasHeight, 600);
});

test('reencodeMediaImage: no scaling when maxDimension is Infinity', async () => {
    const deps = makeImageDeps({ outputBytes: new Uint8Array([1]), width: 5000, height: 5000 });
    await reencodeMediaImage(new Uint8Array([1]), 'image/jpeg', 'image/jpeg', 0.7, deps);
    assert.equal(deps.createdCanvasWidth, 5000);
    assert.equal(deps.createdCanvasHeight, 5000);
});

// ===== compressDocx =====

test('compressDocx: replaces core.xml + app.xml with stubs', async () => {
    const deps = makeDocxDeps();
    await compressDocx(new Uint8Array(200), 'medium', { stripMetadata: true }, deps);
    const files = deps.zippedFiles();
    assert.match(files['docProps/core.xml'], /<cp:coreProperties/);
    assert.match(files['docProps/app.xml'], /<Properties/);
});

test('compressDocx: applies stripTrackedChanges when option set', async () => {
    const deps = makeDocxDeps({
        documentXml: '<w:body><w:del>gone</w:del><w:ins>kept</w:ins><w:p>text</w:p></w:body>',
    });
    await compressDocx(new Uint8Array(200), 'medium', { stripTrackedChanges: true }, deps);
    const out = deps.zippedFiles()['word/document.xml'];
    assert.equal(out.includes('<w:del>'), false);
    assert.match(out, /kept/);
});

test('compressDocx: applies stripCommentMarkers when option set', async () => {
    const deps = makeDocxDeps({
        documentXml: '<w:body><w:commentRangeStart/><w:p>text</w:p><w:commentRangeEnd/></w:body>',
    });
    await compressDocx(new Uint8Array(200), 'medium', { stripComments: true }, deps);
    const out = deps.zippedFiles()['word/document.xml'];
    assert.equal(out.includes('commentRangeStart'), false);
});

test('compressDocx: generates with DEFLATE level 9', async () => {
    const deps = makeDocxDeps();
    await compressDocx(new Uint8Array(200), 'medium', {}, deps);
    assert.equal(deps.generateOptions().type, 'uint8array');
    assert.equal(deps.generateOptions().compression, 'DEFLATE');
    assert.equal(deps.generateOptions().compressionOptions.level, 9);
});

test('compressDocx: re-encode media images when smaller', async () => {
    const deps = makeDocxDeps({
        media: { 'word/media/photo.jpg': new Uint8Array(500) },
        reencodedBytes: new Uint8Array(100),
    });
    await compressDocx(new Uint8Array(200), 'medium', {}, deps);
    assert.equal(deps.zippedFiles()['word/media/photo.jpg'].length, 100);
});

test('compressDocx: skips re-encode when output is larger', async () => {
    const deps = makeDocxDeps({
        media: { 'word/media/photo.jpg': new Uint8Array(100) },
        reencodedBytes: new Uint8Array(500),
    });
    await compressDocx(new Uint8Array(200), 'medium', {}, deps);
    assert.equal(deps.zippedFiles()['word/media/photo.jpg'].length, 100);
});

test('compressDocx: skips PNG media at low level', async () => {
    const deps = makeDocxDeps({
        media: { 'word/media/photo.png': new Uint8Array(100) },
    });
    await compressDocx(new Uint8Array(200), 'low', {}, deps);
    assert.equal(deps.reencodeCalls(), 0);
});

test('compressDocx: converts PNG to JPEG at medium level', async () => {
    const deps = makeDocxDeps({
        media: { 'word/media/photo.png': new Uint8Array(500) },
        reencodedBytes: new Uint8Array(100),
    });
    await compressDocx(new Uint8Array(200), 'medium', {}, deps);
    assert.equal(deps.reencodeCalls(), 1);
    assert.equal(deps.lastSourceMime(), 'image/png');
    assert.equal(deps.lastTargetMime(), 'image/jpeg');
    assert.equal(deps.zippedFiles()['word/media/photo.png'].length, 100);
});

test('compressDocx: converts PNG to JPEG at high level', async () => {
    const deps = makeDocxDeps({
        media: { 'word/media/photo.png': new Uint8Array(500) },
        reencodedBytes: new Uint8Array(100),
    });
    await compressDocx(new Uint8Array(200), 'high', {}, deps);
    assert.equal(deps.reencodeCalls(), 1);
    assert.equal(deps.lastTargetMime(), 'image/jpeg');
});

test('compressDocx: quality mapping low/medium/high (flipped)', async () => {
    const depsLow = makeDocxDeps({
        media: { 'word/media/a.jpg': new Uint8Array(500) },
        reencodedBytes: new Uint8Array(100),
    });
    await compressDocx(new Uint8Array(200), 'low', {}, depsLow);
    assert.equal(depsLow.lastQuality(), 0.85);

    const depsHigh = makeDocxDeps({
        media: { 'word/media/a.jpg': new Uint8Array(500) },
        reencodedBytes: new Uint8Array(100),
    });
    await compressDocx(new Uint8Array(200), 'high', {}, depsHigh);
    assert.equal(depsHigh.lastQuality(), 0.5);
});

test('compressDocx: passes maxDimension to reencodeMediaImage deps', async () => {
    const deps = makeDocxDeps({
        media: { 'word/media/a.jpg': new Uint8Array(500) },
        reencodedBytes: new Uint8Array(100),
    });
    await compressDocx(new Uint8Array(200), 'medium', {}, deps);
    assert.equal(deps.lastMaxDimension(), 2000);
});

test('compressDocx: falls back to original on any failure', async () => {
    const deps = makeDocxDeps({ throwOnLoad: true });
    const blob = await compressDocx(new Uint8Array(200), 'medium', {}, deps);
    assert.equal(blob.size, 200);
});

test('compressDocx: falls back when validation fails (reopen missing document.xml)', async () => {
    const deps = makeDocxDeps({ validationMissingDocument: true });
    const blob = await compressDocx(new Uint8Array(200), 'medium', {}, deps);
    assert.equal(blob.size, 200);
});

// ===== filterFiles =====

test('filterFiles returns valid files + pushes errors', () => {
    const files = [
        { name: 'good.pdf', size: 1000 },
        { name: 'bad.gif', size: 1000 },
        { name: 'report.docx', size: 1000 },
        { name: 'huge.pdf', size: 200 * 1024 * 1024 },
    ];
    const errors = [];
    const accepted = filterFiles(files, errors);
    assert.equal(accepted.length, 2);
    assert.equal(accepted[0].name, 'good.pdf');
    assert.equal(accepted[1].name, 'report.docx');
    assert.equal(errors.length, 2);
});

// ===== downloadResults =====

test('downloadResults: uses provided anchor + clicks + revokes', async () => {
    let clicked = false;
    let revokedUrl = null;
    let assignedHref = null;
    let assignedDownload = null;
    const anchor = { click() { clicked = true; } };
    Object.defineProperty(anchor, 'href', { get() { return assignedHref; }, set(v) { assignedHref = v; } });
    Object.defineProperty(anchor, 'download', { get() { return assignedDownload; }, set(v) { assignedDownload = v; } });
    const deps = makeJsZipDeps();
    deps.createAnchorElement = () => anchor;
    deps.createObjectURL = () => 'blob:abc';
    deps.scheduleRevoke = (url) => { revokedUrl = url; };
    const url = await downloadResults(
        [{ file: { name: 'out.pdf' }, blob: new Uint8Array([1, 2, 3]) }],
        deps
    );
    assert.equal(clicked, true);
    assert.equal(url, 'blob:abc');
    assert.equal(revokedUrl, 'blob:abc');
});

test('downloadResults: zip contains the provided file with .compressed.pdf suffix', async () => {
    const deps = makeJsZipDeps();
    deps.createObjectURL = () => 'blob:abc';
    deps.createAnchorElement = () => ({ click() {}, set href(v) {}, set download(v) {} });
    deps.scheduleRevoke = () => {};
    await downloadResults(
        [{ file: { name: 'doc.pdf' }, blob: new Uint8Array([1]) }],
        deps
    );
    const names = Object.keys(deps.zippedNames);
    assert.ok(names.includes('doc.compressed.pdf'), 'expected doc.compressed.pdf in ' + JSON.stringify(names));
});

// ===== Session history & metrics =====

test('recordCompression: updates session state', () => {
    resetSession();
    recordCompression('test.pdf', 1000, 600);
    const s = getSession();
    assert.equal(s.filesProcessed, 1);
    assert.equal(s.totalSavedBytes, 400);
    assert.equal(s.history.length, 1);
    assert.equal(s.history[0].fileName, 'test.pdf');
    assert.equal(s.history[0].savedBytes, 400);
    assert.equal(s.history[0].percentSaved, 40);
});

test('recordCompression: caps history at 20 entries', () => {
    resetSession();
    for (let i = 0; i < 25; i++) {
        recordCompression(`file${i}.pdf`, 1000, 500);
    }
    const s = getSession();
    assert.equal(s.history.length, 20);
    assert.equal(s.filesProcessed, 25);
    assert.equal(s.history[0].fileName, 'file24.pdf');
    assert.equal(s.history[19].fileName, 'file5.pdf');
});

test('recordCompression: handles no savings gracefully', () => {
    resetSession();
    recordCompression('big.pdf', 1000, 1000);
    const s = getSession();
    assert.equal(s.totalSavedBytes, 0);
    assert.equal(s.history[0].percentSaved, 0);
});

test('recordCompression: handles zero originalBytes', () => {
    resetSession();
    recordCompression('empty.pdf', 0, 0);
    const s = getSession();
    assert.equal(s.totalSavedBytes, 0);
    assert.equal(s.history[0].percentSaved, 0);
});

test('renderSession: updates DOM elements with totals', () => {
    resetSession();
    const mockDoc = makeMockDocument();
    const origDoc = globalThis.document;
    globalThis.document = mockDoc;
    try {
        recordCompression('a.pdf', 2 * 1024 * 1024, 1 * 1024 * 1024);
        recordCompression('b.pdf', 1024 * 1024, 512 * 1024);
        const totalEl = mockDoc.getElementById('totalSavedMB');
        const filesEl = mockDoc.getElementById('filesProcessed');
        const list = mockDoc.getElementById('history-list');
        const summary = mockDoc.getElementById('summary');
        const history = mockDoc.getElementById('history');
        assert.equal(summary.hidden, false);
        assert.equal(history.hidden, false);
        assert.match(totalEl.textContent, /MB$/);
        assert.equal(filesEl.textContent, '2');
        assert.equal(list.children.length, 2);
        assert.equal(list.children[0].querySelector('.hist-name').textContent, 'b.pdf');
    } finally {
        globalThis.document = origDoc;
        resetSession();
    }
});

test('renderSession: hides summary/history when nothing processed', () => {
    resetSession();
    const mockDoc = makeMockDocument();
    const origDoc = globalThis.document;
    globalThis.document = mockDoc;
    try {
        renderSession();
        const summary = mockDoc.getElementById('summary');
        const history = mockDoc.getElementById('history');
        assert.equal(summary.hidden, true);
        assert.equal(history.hidden, true);
    } finally {
        globalThis.document = origDoc;
    }
});

// ===== Test helpers =====

function makePdfDeps(opts) {
    opts = opts || {};
    const pages = opts.pages != null ? opts.pages : 1;
    const outputSize = opts.outputSize != null ? opts.outputSize : 50;
    const validationPageCount = opts.validationPageCount;
    const throwOnLoad = opts.throwOnLoad || false;
    const withPdfJs = opts.withPdfJs || false;
    const rasterSize = opts.rasterSize;
    const loadCalls = [];
    const written = { title: null, author: null, subject: null, keywords: null, creator: null, producer: null };
    let lastSaveOptions = null;
    let rasterCalls = 0;
    const docOriginal = {
        setTitle: (v) => { written.title = v; },
        setAuthor: (v) => { written.author = v; },
        setSubject: (v) => { written.subject = v; },
        setKeywords: (v) => { written.keywords = v; },
        setCreator: (v) => { written.creator = v; },
        setProducer: (v) => { written.producer = v; },
        getPageCount: () => pages,
        save: async (opts2) => {
            lastSaveOptions = opts2;
            return new Uint8Array(outputSize);
        },
    };
    const docValidated = {
        getPageCount: () => validationPageCount != null ? validationPageCount : pages,
    };
    let callCount = 0;
    const pdfjsLib = withPdfJs ? {
        getDocument: () => ({
            promise: Promise.resolve({
                numPages: pages,
                getPage: async () => ({
                    getViewport: () => ({ width: 100, height: 100 }),
                    render: () => ({ promise: Promise.resolve() }),
                }),
                destroy: async () => {},
            }),
        }),
    } : undefined;
    return {
        loadCalls,
        lastWritten: written,
        get lastSaveOptions() { return lastSaveOptions; },
        get rasterCalls() { return rasterCalls; },
        pdfjsLib,
        createCanvas: (w, h) => {
            const ctx = { drawImage() {}, fillStyle: '', fillRect() {} };
            return { width: w, height: h, getContext: () => ctx, toBlob: (cb) => cb({ arrayBuffer: async () => new Uint8Array([1]).buffer }) };
        },
        PDFDocument: {
            create: async () => {
                rasterCalls++;
                return {
                    addPage: () => ({ drawImage: () => {} }),
                    embedJpg: async () => ({}),
                    save: async () => new Uint8Array(rasterSize != null ? rasterSize : outputSize),
                };
            },
            load: async (bytes, loadOpts) => {
                loadCalls.push(Object.assign({ bytes: bytes.byteLength }, loadOpts));
                if (throwOnLoad && callCount === 0) {
                    callCount++;
                    throw new Error('boom');
                }
                callCount++;
                return callCount === 1 ? docOriginal : docValidated;
            },
        },
    };
}

function makeImageDeps(opts) {
    opts = opts || {};
    const outputBytes = opts.outputBytes || new Uint8Array([0, 1, 2]);
    const toBlobShouldFail = opts.toBlobShouldFail || false;
    const width = opts.width != null ? opts.width : 10;
    const height = opts.height != null ? opts.height : 10;
    const maxDimension = opts.maxDimension;
    const state = { closed: false, toBlobCalledWith: null };
    let createdCanvasWidth = null;
    let createdCanvasHeight = null;
    const deps = {
        state,
        closed: () => state.closed,
        get toBlobCalledWith() { return state.toBlobCalledWith; },
        get createdCanvasWidth() { return createdCanvasWidth; },
        get createdCanvasHeight() { return createdCanvasHeight; },
        createImageBitmap: async () => ({
            width: width,
            height: height,
            close() { state.closed = true; },
        }),
        createCanvas: (w, h) => {
            createdCanvasWidth = w;
            createdCanvasHeight = h;
            const ctx = { drawImage() {}, fillStyle: '', fillRect() {} };
            return {
                width: w,
                height: h,
                getContext: () => ctx,
                toBlob: (cb, mime, quality) => {
                    state.toBlobCalledWith = { mime: mime, quality: quality };
                    if (toBlobShouldFail) {
                        cb(null);
                    } else {
                        cb({ arrayBuffer: async () => outputBytes.buffer });
                    }
                },
            };
        },
    };
    if (maxDimension !== undefined) {
        deps.maxDimension = maxDimension;
    }
    return deps;
}

function makeJsZipDeps() {
    const zippedNames = {};
    const zip = {
        file(name, data) { zippedNames[name] = data; },
        generateAsync: async (opts) => {
            zip._opts = opts;
            return new Uint8Array([1, 2, 3, 4]);
        },
    };
    return {
        zippedNames,
        get zip() { return zip; },
        JSZip: function () { return zip; },
    };
}

function makeDocxDeps(opts) {
    opts = opts || {};
    const documentXml = opts.documentXml || '<w:body><w:p>hello</w:p></w:body>';
    const media = opts.media || {};
    const reencodedBytes = opts.reencodedBytes || new Uint8Array([0]);
    const throwOnLoad = opts.throwOnLoad || false;
    const validationMissingDocument = opts.validationMissingDocument || false;

    const state = {
        fileMap: Object.assign({
            'word/document.xml': documentXml,
            'docProps/core.xml': '<?xml version="1.0"?><cp:coreProperties xmlns:cp="x"><dc:creator>Original</dc:creator></cp:coreProperties>',
            'docProps/app.xml': '<?xml version="1.0"?><Properties xmlns="x"><Application>Original</Application></Properties>',
        }, media),
        generateOptions: null,
        reencodeCalls: 0,
        lastQuality: null,
        lastMaxDimension: null,
        lastSourceMime: null,
        lastTargetMime: null,
        jsZipCalls: 0,
    };

    function buildZip(opt) {
        opt = opt || {};
        const mode = opt.mode || 'compress';
        const isValidate = mode === 'validate';
        const baseMap = (isValidate && validationMissingDocument)
            ? Object.assign({}, state.fileMap, { 'word/document.xml': undefined })
            : state.fileMap;
        let loadCount = 0;
        return {
            files: baseMap,
            file(name, data) {
                if (data === undefined) {
                    return baseMap[name] !== undefined ? { name } : null;
                }
                baseMap[name] = data;
            },
            async fileAsync(name) { return baseMap[name]; },
            forEach(cb) {
                for (const n of Object.keys(baseMap)) {
                    if (baseMap[n] === undefined) continue;
                    cb({ name: n, fn: async () => baseMap[n] });
                }
            },
            generateAsync: async (gopts) => {
                if (!isValidate) state.generateOptions = gopts;
                return new Uint8Array([9, 9, 9]);
            },
            async loadAsync(bytes) {
                loadCount++;
                if (throwOnLoad && loadCount === 1) throw new Error('load fail');
                return this;
            },
        };
    }

    const deps = {
        state: state,
        zippedFiles: () => state.fileMap,
        generateOptions: () => state.generateOptions,
        reencodeCalls: () => state.reencodeCalls,
        lastQuality: () => state.lastQuality,
        lastMaxDimension: () => state.lastMaxDimension,
        lastSourceMime: () => state.lastSourceMime,
        lastTargetMime: () => state.lastTargetMime,
        JSZip: function () {
            state.jsZipCalls++;
            const mode = state.jsZipCalls === 1 ? 'compress' : 'validate';
            return buildZip({ mode: mode });
        },
        reencodeMediaImage: async (bytes, sourceMime, targetMime, quality, reencodeDeps) => {
            state.reencodeCalls++;
            state.lastQuality = quality;
            state.lastSourceMime = sourceMime;
            state.lastTargetMime = targetMime;
            if (reencodeDeps && reencodeDeps.maxDimension !== undefined) {
                state.lastMaxDimension = reencodeDeps.maxDimension;
            }
            return reencodedBytes;
        },
    };
    return deps;
}

function makeMockDocument() {
    const elements = {
        'summary': makeMockElement('summary'),
        'history': makeMockElement('history'),
        'history-list': makeMockElement('history-list'),
        'totalSavedMB': makeMockElement('totalSavedMB'),
        'filesProcessed': makeMockElement('filesProcessed'),
    };
    return {
        getElementById(id) { return elements[id] || null; },
        createElement(tag) {
            return makeMockElement(tag);
        },
        createTextNode(text) {
            return { nodeType: 3, textContent: text };
        },
    };
}

function makeMockElement(tag) {
    const children = [];
    const el = {
        tagName: tag,
        hidden: false,
        textContent: '',
        className: '',
        children,
        firstChild: null,
        get firstChild() { return children[0] || null; },
        appendChild(child) {
            children.push(child);
            return child;
        },
        removeChild(child) {
            const idx = children.indexOf(child);
            if (idx >= 0) children.splice(idx, 1);
            return child;
        },
        querySelector(sel) {
            return findInChildren(this, sel);
        },
    };
    return el;
}

function findInChildren(root, sel) {
    const cls = sel.startsWith('.') ? sel.slice(1) : null;
    function walk(node) {
        if (!node || !node.children) return null;
        for (const c of node.children) {
            if (cls && c.className === cls) return c;
            const found = walk(c);
            if (found) return found;
        }
        return null;
    }
    return walk(root);
}