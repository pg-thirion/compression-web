import { createRequire } from 'node:module';
import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  parseLevel,
  validateFile,
  ALLOWED_LEVELS,
  FILE_MAX_BYTES,
  FILE_EXTENSIONS,
  compressPdf,
  minimalCoreXml,
  minimalAppXml,
  stripTrackedChanges,
  stripCommentMarkers,
  reencodeMediaImage,
  compressDocx,
} from '../app.js';

const require = createRequire(import.meta.url);
const { PDFDocument } = require('../lib/pdf-lib.min.js');
const JSZip = require('../lib/jszip.min.js');
globalThis.JSZip = JSZip;

async function buildDocx({ includeMedia = false, tracked = false, comments = false, coreMeta = 'real', mediaBytes = null } = {}) {
  const z = new JSZip();
  z.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  const coreXml = coreMeta === 'real'
    ? '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">' + 'A'.repeat(2000) + '</dc:title><dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">' + 'B'.repeat(500) + '</dc:creator><dc:subject xmlns:dc="http://purl.org/dc/elements/1.1/">' + 'C'.repeat(800) + '</dc:subject></cp:coreProperties>'
    : '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>';
  z.file('docProps/core.xml', coreXml);
  z.file('docProps/app.xml', '<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>x</Application></Properties>');
  let body = '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>';
  if (tracked) body += '<w:ins w:id="1">tracked accept</w:ins><w:del w:id="2">tracked reject</w:del>';
  if (comments) body += '<w:commentRangeStart/><w:t>commented</w:t><w:commentRangeEnd/><w:commentReference w:id="1"/>';
  z.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`);
  if (includeMedia) {
    z.file('word/media/image1.jpg', mediaBytes || new Uint8Array([0xff, 0xd8, 0xff, 0xe0]));
  }
  return await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

async function buildPdf({ pages = 1, metadata = {} } = {}) {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pages; i++) doc.addPage();
  if (metadata.title) doc.setTitle(metadata.title);
  if (metadata.author) doc.setAuthor(metadata.author);
  if (metadata.subject) doc.setSubject(metadata.subject);
  if (metadata.keywords) doc.setKeywords(metadata.keywords);
  if (metadata.creator) doc.setCreator(metadata.creator);
  if (metadata.producer) doc.setProducer(metadata.producer);
  return await doc.save();
}

describe('constants', () => {
  it('ALLOWED_LEVELS is a Set with low, medium, high', () => {
    assert.ok(ALLOWED_LEVELS instanceof Set);
    assert.deepEqual([...ALLOWED_LEVELS].sort(), ['high', 'low', 'medium']);
  });

  it('FILE_MAX_BYTES is 100 * 1024 * 1024', () => {
    assert.equal(FILE_MAX_BYTES, 100 * 1024 * 1024);
  });

  it('FILE_EXTENSIONS is a Set with .pdf and .docx', () => {
    assert.ok(FILE_EXTENSIONS instanceof Set);
    assert.deepEqual([...FILE_EXTENSIONS].sort(), ['.docx', '.pdf']);
  });
});

describe('parseLevel', () => {
  it('returns "low" for "low"', () => {
    assert.equal(parseLevel('low'), 'low');
  });

  it('returns "medium" for "medium"', () => {
    assert.equal(parseLevel('medium'), 'medium');
  });

  it('returns "high" for "high"', () => {
    assert.equal(parseLevel('high'), 'high');
  });

  it('throws RangeError for "LOW" (uppercase)', () => {
    assert.throws(() => parseLevel('LOW'), RangeError);
  });

  it('throws RangeError for "extreme"', () => {
    assert.throws(() => parseLevel('extreme'), RangeError);
  });

  it('throws RangeError for empty string', () => {
    assert.throws(() => parseLevel(''), RangeError);
  });

  it('throws RangeError for null', () => {
    assert.throws(() => parseLevel(null), RangeError);
  });

  it('throws RangeError for undefined', () => {
    assert.throws(() => parseLevel(undefined), RangeError);
  });

  it('throws RangeError for a number', () => {
    assert.throws(() => parseLevel(1), RangeError);
  });

  it('throws RangeError for an object', () => {
    assert.throws(() => parseLevel({ level: 'low' }), RangeError);
  });
});

describe('validateFile', () => {
  it('returns null for a valid PDF', () => {
    const file = { name: 'report.pdf', size: 1024, type: 'application/pdf' };
    assert.equal(validateFile(file), null);
  });

  it('returns null for a valid DOCX', () => {
    const file = {
      name: 'report.docx',
      size: 1024,
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    };
    assert.equal(validateFile(file), null);
  });

  it('returns error string for wrong extension', () => {
    const file = { name: 'photo.png', size: 1024, type: 'image/png' };
    const result = validateFile(file);
    assert.equal(typeof result, 'string');
    assert.match(result, /extension/i);
  });

  it('returns error string for file exceeding FILE_MAX_BYTES', () => {
    const file = {
      name: 'huge.pdf',
      size: FILE_MAX_BYTES + 1,
      type: 'application/pdf',
    };
    const result = validateFile(file);
    assert.equal(typeof result, 'string');
    assert.match(result, /size|too large|exceed/i);
  });

  it('returns error string for missing name', () => {
    const file = { name: '', size: 1024, type: 'application/pdf' };
    const result = validateFile(file);
    assert.equal(typeof result, 'string');
  });

  it('returns error string for missing size', () => {
    const file = { name: 'report.pdf', type: 'application/pdf' };
    const result = validateFile(file);
    assert.equal(typeof result, 'string');
  });

  it('accepts a file at exactly FILE_MAX_BYTES', () => {
    const file = { name: 'edge.pdf', size: FILE_MAX_BYTES, type: 'application/pdf' };
    assert.equal(validateFile(file), null);
  });

  it('returns error for a file with no extension', () => {
    const file = { name: 'noext', size: 1024, type: '' };
    const result = validateFile(file);
    assert.equal(typeof result, 'string');
  });

  it('is case-insensitive on extension matching', () => {
    const file = { name: 'Report.PDF', size: 1024, type: 'application/pdf' };
    assert.equal(validateFile(file), null);
  });
});

describe('compressPdf', () => {
  it('returns Uint8Array that re-loads with same page count (1 page)', async () => {
    const input = await buildPdf({ pages: 1 });
    const result = await compressPdf(input, { PDFDocument });
    assert.ok(result instanceof Uint8Array);
    const reloaded = await PDFDocument.load(result, { updateMetadata: false });
    assert.equal(reloaded.getPageCount(), 1);
  });

  it('returns Uint8Array that re-loads with same page count (3 pages)', async () => {
    const input = await buildPdf({ pages: 3 });
    const result = await compressPdf(input, { PDFDocument });
    const reloaded = await PDFDocument.load(result, { updateMetadata: false });
    assert.equal(reloaded.getPageCount(), 3);
  });

  it('wipes metadata', async () => {
    const input = await buildPdf({
      metadata: {
        title: 'My Title',
        author: 'Jane Doe',
        subject: 'Test Subject',
        keywords: ['pdf', 'test'],
        creator: 'Notepad',
        producer: 'MyProducer',
      },
    });
    const result = await compressPdf(input, { PDFDocument });
    const reloaded = await PDFDocument.load(result, { updateMetadata: false });
    assert.equal(reloaded.getTitle() ?? '', '');
    assert.equal(reloaded.getAuthor() ?? '', '');
    assert.equal(reloaded.getSubject() ?? '', '');
    const keywords = reloaded.getKeywords();
    assert.ok(
      keywords == null || (Array.isArray(keywords) && keywords.length === 0) || keywords === '',
      `expected empty keywords, got ${JSON.stringify(keywords)}`
    );
    assert.equal(reloaded.getCreator() ?? '', '');
    assert.equal(reloaded.getProducer() ?? '', '');
  });

  it('returns output strictly smaller than input when input has metadata', async () => {
    const input = await buildPdf({
      metadata: {
        title: 'A long enough title to occupy some bytes in the output',
        author: 'Author Name',
        subject: 'A longer subject line that also takes bytes',
        keywords: ['keyword1', 'keyword2', 'keyword3'],
        creator: 'Some Creator Application',
        producer: 'Some Producer Application',
      },
    });
    const result = await compressPdf(input, { PDFDocument });
    assert.ok(result.length < input.length, `expected ${result.length} < ${input.length}`);
  });

  it('never returns a larger output', async () => {
    const input = await buildPdf({ pages: 1 });
    const result = await compressPdf(input, { PDFDocument });
    assert.ok(result.length <= input.length, `expected ${result.length} <= ${input.length}`);
  });

  it('falls back to original bytes when input is malformed', async () => {
    const garbage = new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    const result = await compressPdf(garbage, { PDFDocument });
    assert.equal(result, garbage);
  });

  it('accepts custom deps', async () => {
    const sentinel = new Uint8Array([42, 42, 42]);
    let loadCallCount = 0;
    const stubDoc = {
      setTitle: () => {},
      setAuthor: () => {},
      setSubject: () => {},
      setKeywords: () => {},
      setCreator: () => {},
      setProducer: () => {},
      save: async () => sentinel,
      getPageCount: () => 1,
    };
    const stubPDFDocument = {
      load: async () => {
        loadCallCount++;
        return stubDoc;
      },
    };
    const input = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    const result = await compressPdf(input, { PDFDocument: stubPDFDocument });
    assert.ok(loadCallCount > 0, 'stub PDFDocument.load was never called — deps were not used');
    assert.equal(result, sentinel, 'result was not the stub sentinel — deps were not consulted');
  });
});

describe('minimalCoreXml', () => {
  it('returns a non-empty string', () => {
    const xml = minimalCoreXml();
    assert.equal(typeof xml, 'string');
    assert.ok(xml.length > 0);
  });

  it('starts with an XML prolog', () => {
    assert.ok(minimalCoreXml().startsWith('<?xml'));
  });

  it('declares the cp core-properties root element', () => {
    assert.match(minimalCoreXml(), /<cp:coreProperties[\s>]/);
  });

  it('declares the cp namespace', () => {
    assert.match(
      minimalCoreXml(),
      /xmlns:cp="http:\/\/schemas\.openxmlformats\.org\/package\/2006\/metadata\/core-properties"/
    );
  });
});

describe('minimalAppXml', () => {
  it('returns a non-empty string', () => {
    const xml = minimalAppXml();
    assert.equal(typeof xml, 'string');
    assert.ok(xml.length > 0);
  });

  it('starts with an XML prolog', () => {
    assert.ok(minimalAppXml().startsWith('<?xml'));
  });

  it('declares a Properties root element with extended-properties namespace', () => {
    const xml = minimalAppXml();
    assert.match(xml, /<Properties[\s>]/);
    assert.match(
      xml,
      /xmlns="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/extended-properties"/
    );
  });
});

describe('stripTrackedChanges', () => {
  it('drops <w:del>...</w:del> blocks including their contents', () => {
    const input = '<w:p><w:r><w:t>before </w:t></w:r><w:del>removed text</w:del><w:r><w:t>after</w:t></w:r></w:p>';
    const output = stripTrackedChanges(input);
    assert.equal(output.includes('<w:del>'), false);
    assert.equal(output.includes('removed text'), false);
    assert.match(output, /before/);
    assert.match(output, /after/);
  });

  it('drops <w:del> blocks that carry attributes', () => {
    const input = '<w:del w:id="1" w:author="x">stuff</w:del>';
    const output = stripTrackedChanges(input);
    assert.equal(output, '');
  });

  it('unwraps <w:ins>...</w:ins> keeping the inner content', () => {
    const input = '<w:p><w:r><w:t>before </w:t></w:r><w:ins>inserted</w:ins><w:r><w:t>after</w:t></w:r></w:p>';
    const output = stripTrackedChanges(input);
    assert.equal(output.includes('<w:ins>'), false);
    assert.equal(output.includes('</w:ins>'), false);
    assert.match(output, /inserted/);
  });

  it('unwraps <w:ins> that carry attributes', () => {
    const input = '<w:ins w:id="2" w:author="x">kept</w:ins>';
    const output = stripTrackedChanges(input);
    assert.equal(output, 'kept');
  });

  it('handles both <w:del> and <w:ins> in the same fragment', () => {
    const input = 'a<w:del>b</w:del>c<w:ins>d</w:ins>e';
    const output = stripTrackedChanges(input);
    assert.equal(output, 'acde');
  });

  it('leaves fragments without tracked changes unchanged', () => {
    const input = '<w:p><w:r><w:t>plain</w:t></w:r></w:p>';
    assert.equal(stripTrackedChanges(input), input);
  });

  it('leaves a plain string without tracked changes unchanged', () => {
    assert.equal(stripTrackedChanges('hello world'), 'hello world');
  });

  it('drops a <w:del> nested inside an <w:ins> (drop-del first, then unwrap-ins)', () => {
    const input = '<w:ins>keep<w:del>drop</w:del>more</w:ins>';
    const output = stripTrackedChanges(input);
    assert.equal(output, 'keepmore');
  });

  it('drops a self-closing <w:del/> without consuming following content', () => {
    const input = '<w:del w:id="1"/>KEEP';
    const output = stripTrackedChanges(input);
    assert.equal(output, 'KEEP');
  });

  it('handles interleaved self-closing and paired <w:del> without corrupting content', () => {
    const input = '<w:del w:id="1"/>KEEP<w:del w:id="2">GONE</w:del>TAIL';
    const output = stripTrackedChanges(input);
    assert.equal(output, 'KEEPTAIL');
  });

  it('drops a self-closing <w:ins/> (the unwrap form is symmetric with <w:del/>)', () => {
    const input = '<w:ins w:id="1"/>';
    const output = stripTrackedChanges(input);
    assert.equal(output, '');
  });

  it('still drops paired <w:del> content (regression guard for the alternation)', () => {
    const input = '<w:del w:id="2">GONE</w:del>TAIL';
    const output = stripTrackedChanges(input);
    assert.equal(output, 'TAIL');
    assert.equal(output.includes('GONE'), false);
  });
});

describe('stripCommentMarkers', () => {
  it('removes <w:commentRangeStart/> markers', () => {
    const input = '<w:p>a<w:commentRangeStart/>b</w:p>';
    const output = stripCommentMarkers(input);
    assert.equal(output.includes('commentRangeStart'), false);
    assert.match(output, /a/);
    assert.match(output, /b/);
  });

  it('removes <w:commentRangeEnd/> markers', () => {
    const input = '<w:p>a<w:commentRangeEnd/>b</w:p>';
    const output = stripCommentMarkers(input);
    assert.equal(output.includes('commentRangeEnd'), false);
  });

  it('removes <w:commentReference w:id="0"/> markers', () => {
    const input = '<w:p>a<w:commentReference w:id="0"/>b</w:p>';
    const output = stripCommentMarkers(input);
    assert.equal(output.includes('commentReference'), false);
    assert.match(output, /a/);
    assert.match(output, /b/);
  });

  it('removes all three marker types together', () => {
    const input =
      '<w:p><w:commentRangeStart w:id="0"/>text<w:commentRangeEnd w:id="0"/><w:commentReference w:id="0"/></w:p>';
    const output = stripCommentMarkers(input);
    assert.equal(output.includes('commentRangeStart'), false);
    assert.equal(output.includes('commentRangeEnd'), false);
    assert.equal(output.includes('commentReference'), false);
    assert.match(output, /text/);
  });

  it('handles paired <w:commentReference ...></w:commentReference> form', () => {
    const input = '<w:p>a<w:commentReference w:id="0"></w:commentReference>b</w:p>';
    const output = stripCommentMarkers(input);
    assert.equal(output.includes('commentReference'), false);
    assert.equal(output, '<w:p>ab</w:p>');
  });

  it('leaves fragments without comment markers unchanged', () => {
    const input = '<w:p><w:r><w:t>plain</w:t></w:r></w:p>';
    assert.equal(stripCommentMarkers(input), input);
  });

  it('leaves a plain string without comment markers unchanged', () => {
    assert.equal(stripCommentMarkers('hello world'), 'hello world');
  });
});

function createImageMocks({ convertToBlobImpl } = {}) {
  let closeCount = 0;
  let blobReceived = null;
  let convertToBlobOpts = null;
  let getContextCalled = false;
  let drawImageCalled = false;
  let canvasWidth = null;
  let canvasHeight = null;

  const fakeBitmap = {
    width: 100,
    height: 50,
    close() { closeCount++; },
  };

  const createImageBitmap = async (blob) => {
    blobReceived = blob;
    return fakeBitmap;
  };

  function MockOffscreenCanvas(width, height) {
    canvasWidth = width;
    canvasHeight = height;
  }
  MockOffscreenCanvas.prototype.getContext = function () {
    getContextCalled = true;
    return { drawImage: () => { drawImageCalled = true; } };
  };
  MockOffscreenCanvas.prototype.convertToBlob = function (opts) {
    convertToBlobOpts = opts;
    if (convertToBlobImpl) return convertToBlobImpl(opts);
    return Promise.resolve(new Blob([new Uint8Array([1, 2, 3, 4])], { type: 'image/jpeg' }));
  };

  return {
    deps: { createImageBitmap, OffscreenCanvas: MockOffscreenCanvas },
    state: () => ({
      closeCount,
      blobReceived,
      convertToBlobOpts,
      getContextCalled,
      drawImageCalled,
      canvasWidth,
      canvasHeight,
    }),
  };
}

describe('reencodeMediaImage', () => {
  it('returns a Uint8Array from convertToBlob output', async () => {
    const { deps, state } = createImageMocks();
    const result = await reencodeMediaImage(
      new Uint8Array([10, 20, 30]),
      'image/jpeg',
      0.7,
      deps
    );
    assert.ok(result instanceof Uint8Array);
    assert.equal(result.length, 4);
  });

  it('sizes the canvas to the bitmap dimensions', async () => {
    const { deps, state } = createImageMocks();
    await reencodeMediaImage(new Uint8Array([1, 2, 3]), 'image/jpeg', 0.5, deps);
    const s = state();
    assert.equal(s.canvasWidth, 100);
    assert.equal(s.canvasHeight, 50);
  });

  it('passes { type: "image/jpeg", quality } to convertToBlob', async () => {
    const { deps, state } = createImageMocks();
    await reencodeMediaImage(new Uint8Array([1, 2, 3]), 'image/jpeg', 0.42, deps);
    assert.deepEqual(state().convertToBlobOpts, { type: 'image/jpeg', quality: 0.42 });
  });

  it('wraps bytes in a Blob with the supplied mimeType', async () => {
    const { deps, state } = createImageMocks();
    await reencodeMediaImage(new Uint8Array([1, 2, 3]), 'image/png', 0.5, deps);
    const blob = state().blobReceived;
    assert.ok(blob instanceof Blob);
    assert.equal(blob.type, 'image/png');
  });

  it('closes the bitmap on success', async () => {
    const { deps, state } = createImageMocks();
    await reencodeMediaImage(new Uint8Array([1, 2, 3]), 'image/jpeg', 0.7, deps);
    assert.equal(state().closeCount, 1);
  });

  it('closes the bitmap even when convertToBlob throws', async () => {
    const { deps, state } = createImageMocks({
      convertToBlobImpl: () => Promise.reject(new Error('encode failed')),
    });
    await assert.rejects(
      reencodeMediaImage(new Uint8Array([1, 2, 3]), 'image/jpeg', 0.7, deps),
      /encode failed/
    );
    assert.equal(state().closeCount, 1);
  });
});

describe('compressDocx', () => {
  it('happy path: returns Uint8Array with word/document.xml and size ≤ original', async () => {
    const input = await buildDocx({ coreMeta: 'real' });
    const result = await compressDocx(input, 'medium', { stripMetadata: true });
    assert.ok(result instanceof Uint8Array);
    const verify = await JSZip.loadAsync(result);
    assert.ok(verify.file('word/document.xml') !== null);
    assert.ok(result.length <= input.length);
  });

  it('strips metadata by default: result is smaller and core.xml no longer contains the title', async () => {
    const input = await buildDocx({ coreMeta: 'real' });
    const result = await compressDocx(input, 'medium');
    assert.ok(result.length < input.length, `expected ${result.length} < ${input.length}`);
    const verify = await JSZip.loadAsync(result);
    const coreXml = await verify.file('docProps/core.xml').async('string');
    assert.equal(coreXml.includes('AAAAAAAAAA'), false);
  });

  it('does NOT strip metadata when stripMetadata is false', async () => {
    const input = await buildDocx({ coreMeta: 'real' });
    const result = await compressDocx(input, 'medium', { stripMetadata: false });
    const verify = await JSZip.loadAsync(result);
    const coreXml = await verify.file('docProps/core.xml').async('string');
    assert.equal(coreXml.includes('AAAAAAAAAA'), true);
  });

  it('strips tracked changes when option is on: drops <w:del>, unwraps <w:ins>', async () => {
    const input = await buildDocx({ tracked: true });
    const result = await compressDocx(input, 'medium', { stripTrackedChanges: true });
    const verify = await JSZip.loadAsync(result);
    const docXml = await verify.file('word/document.xml').async('string');
    assert.equal(docXml.includes('<w:del'), false);
    assert.equal(docXml.includes('tracked reject'), false);
    assert.equal(docXml.includes('<w:ins'), false);
    assert.equal(docXml.includes('tracked accept'), true);
  });

  it('strips comment markers when option is on', async () => {
    const input = await buildDocx({ comments: true });
    const result = await compressDocx(input, 'medium', { stripComments: true });
    const verify = await JSZip.loadAsync(result);
    const docXml = await verify.file('word/document.xml').async('string');
    assert.equal(docXml.includes('commentRangeStart'), false);
    assert.equal(docXml.includes('commentRangeEnd'), false);
    assert.equal(docXml.includes('commentReference'), false);
    assert.match(docXml, /commented/);
  });

  it('re-encodes media via injected stub: smaller bytes replace original', async () => {
    const origMedia = new Uint8Array(50).fill(7);
    const input = await buildDocx({ includeMedia: true, mediaBytes: origMedia });
    const smallerBytes = new Uint8Array([1, 2, 3]);
    const stubReencode = async () => smallerBytes;
    const result = await compressDocx(input, 'medium', {}, { reencodeMediaImage: stubReencode });
    const verify = await JSZip.loadAsync(result);
    const newMedia = await verify.file('word/media/image1.jpg').async('uint8array');
    assert.equal(newMedia.length, 3);
  });

  it('media: leaves file unchanged when re-encoded output is larger', async () => {
    const origMedia = new Uint8Array([1, 2, 3]);
    const input = await buildDocx({ includeMedia: true, mediaBytes: origMedia });
    const largerBytes = new Uint8Array(50).fill(9);
    const stubReencode = async () => largerBytes;
    const result = await compressDocx(input, 'medium', {}, { reencodeMediaImage: stubReencode });
    const verify = await JSZip.loadAsync(result);
    const newMedia = await verify.file('word/media/image1.jpg').async('uint8array');
    assert.equal(newMedia.length, 3, 'should keep original smaller bytes');
  });

  it('skips re-encoding JPEG when level is High', async () => {
    const input = await buildDocx({ includeMedia: true });
    let called = false;
    const stubReencode = async () => {
      called = true;
      return new Uint8Array([1, 2, 3]);
    };
    await compressDocx(input, 'high', {}, { reencodeMediaImage: stubReencode });
    assert.equal(called, false, 'reencodeMediaImage should not be called for JPEG at High level');
  });

  it('falls back to original bytes when JSZip.loadAsync throws', async () => {
    const input = await buildDocx();
    const stubJSZip = {
      loadAsync: async () => { throw new Error('corrupt zip'); },
    };
    const result = await compressDocx(input, 'medium', {}, { JSZip: stubJSZip });
    assert.equal(result, input);
  });

  it('falls back to original bytes when result lacks word/document.xml', async () => {
    const input = await buildDocx();
    const fakeZip = {
      files: {},
      file: (name) => (name === 'word/document.xml' ? null : null),
      generateAsync: async () => new Uint8Array([1, 2, 3]),
    };
    const stubJSZip = {
      loadAsync: async () => fakeZip,
    };
    const result = await compressDocx(input, 'medium', {}, { JSZip: stubJSZip });
    assert.equal(result, input);
  });
});
