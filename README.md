# Compression Web

Compression Web compresses PDF and DOCX files entirely in your browser. Files never leave your machine — the app is a static page that uses vendored JavaScript libraries to do all the work locally, then bundles the results into a single ZIP for download.

## Accepted formats

- PDF
- DOCX

Maximum 100 MB per file. There is no hard cap on the total size in flight, but browser memory is the practical limit.

## Compression levels

| Level | Compression | Use when |
|-------|-------------|----------|
| Low | Lightest | Best fidelity, biggest files |
| Medium | Balanced | Default for most files |
| High | Strongest | Smallest files, lower fidelity |

At Low, both PDFs and DOCX receive structural compression only — metadata wipe, object-stream repacking, DOCX XML cleanup, JPEG re-encoding. The original image content is preserved.

At Medium and High, **PDFs are rasterized**: each page is rendered to a JPEG at the level's quality and re-embedded as an image. Text in the output is no longer searchable, and any vector content (form fields, links, annotations) is lost. This is a deliberate trade-off — it produces much smaller output for image-heavy PDFs (scans, slides, image-only documents).

At Medium and High, **DOCX media images are downsampled** to fit inside 2000 px (Medium) or 1500 px (High) on the longest side, and PNGs are re-encoded as JPEG. JPEG transparency is lost (replaced with white).

DOCX compression also depends on the DOCX options below.

## DOCX options

The DOCX options fieldset appears only when at least one DOCX file is in the queue:

- **Strip tracked changes** (off by default) — removes `<w:del>` blocks and unwraps `<w:ins>` blocks from `word/document.xml`.
- **Strip comments** (off by default) — removes `<w:commentRangeStart>`, `<w:commentRangeEnd>`, and `<w:commentReference>` markers.
- **Strip metadata** (on by default) — replaces `docProps/core.xml` and `docProps/app.xml` with empty stubs.

JPEG quality for re-encoding: Low 0.85, Medium 0.7, High 0.5 (matches the level semantics — Low keeps fidelity, High compresses harder).

## Session history & metrics

After each successful compression, the page shows two summary tiles ("saved this session" and "files processed") and a list of the last 20 compressions with per-file size savings. The history is purely in-memory: it clears on page refresh and is never persisted, sent over the network, or written to storage.

## Architecture

- Single static page: `index.html`, `app.js`, `lib/jszip.min.js`, `lib/pdf-lib.min.js`, `lib/pdf.min.js`, `lib/pdf.worker.min.js`.
- pdf-lib (1.17.1) for PDF structural manipulation — loaded as a UMD bundle exposing `globalThis.PDFLib`.
- pdf.js for PDF rasterization (page rendering to canvas) at Medium/High — worker source configured in `app.js`.
- JSZip for reading and writing DOCX files (they are ZIP containers).
- HTML5 Canvas + `createImageBitmap` for re-encoding DOCX media images.
- All processing happens in the browser; the static site has no backend.

## Development

Tests use the built-in Node.js test runner:

```bash
node --test tests/app.test.mjs
```

To run the page locally:

```bash
python -m http.server 8000
```

Then open `http://localhost:8000`.

## Deployment

The deployment script uploads the static site to S3 in `eu-west-3` and invalidates CloudFront:

```bash
AWS_ACCOUNT_ID=664882921031 ./deploy.sh
```

The bucket name defaults to `hbtrained-compression-web-${ACCOUNT_ID}`. Override with:

```bash
BUCKET_NAME=my-bucket AWS_ACCOUNT_ID=664882921031 ./deploy.sh
```

Public reads are granted only for the static site files (`index.html`, `app.js`, all `lib/*` bundles). CloudFront distribution ID and domain are read from `~/.square-fit-batch/state` (line 1 = distribution ID, line 2 = domain). The script does not grant public upload.

## Known limitations

- **Raster re-encoding at Medium/High destroys text searchability in PDFs.** The trade-off is intentional — rasterization produces much smaller output for image-heavy PDFs but text is no longer selectable or searchable in the result.
- **PDF vector content is lost at Medium/High.** Form fields, links, annotations, and any other non-image content are flattened into JPEG page images.
- **Validation is structural, not pixel-diff.** Output PDFs are re-loaded and page count is verified; DOCX output is reopened and `word/document.xml` is checked. Structural integrity is verified, not visual fidelity.
- **Single-threaded.** PDF rasterization and DOCX image re-encoding are single-threaded JavaScript in the browser.
- **Validation can be over-conservative.** If the optimized output is larger than the original, or the re-load fails for any reason, the original bytes are returned unchanged — so no data is ever lost.
- **DOCX PNGs at Medium/High become JPEG.** Alpha channel is lost (replaced with white background). To preserve PNG, use the Low level.
- **Session history is in-memory only.** Clearing on page refresh is by design — there is no localStorage persistence.

## Files

- `index.html` — static interface
- `app.js` — validation, conversion, ZIP download, session history
- `lib/jszip.min.js` — local JSZip runtime
- `lib/pdf-lib.min.js` — local pdf-lib runtime
- `lib/pdf.min.js` — local pdf.js runtime (used for PDF rasterization at Medium/High)
- `lib/pdf.worker.min.js` — local pdf.js worker (configured in app.js)
- `deploy.sh` — S3 + CloudFront deployment
- `tests/app.test.mjs` — unit tests