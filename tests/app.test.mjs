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
} from '../app.js';

const require = createRequire(import.meta.url);
const { PDFDocument } = require('../lib/pdf-lib.min.js');

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
