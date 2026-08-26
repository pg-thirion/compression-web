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
    const input = await buildPdf({ pages: 2 });
    const result = await compressPdf(input, { PDFDocument });
    const reloaded = await PDFDocument.load(result, { updateMetadata: false });
    assert.equal(reloaded.getPageCount(), 2);
  });
});
