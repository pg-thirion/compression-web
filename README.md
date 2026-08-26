# Compression Web

Compression Web compresses PDF and DOCX files entirely in the browser. Files never leave the user's machine: the app is a static page that uses vendored JavaScript libraries (pdf-lib and JSZip) to do all the work locally, then bundles the results into a single ZIP for download. There is no backend.

## Accepted formats

- PDF
- DOCX

Maximum 100 MB per file. There is no hard cap on the total size in flight; browser memory is the practical limit.

## Compression levels

Four levels are offered. They differ only in how aggressively DOCX media (images) is re-encoded; PDF behavior is identical at every level.

| Level | JPEG quality | Max image dimension | Use when |
|-------|--------------|---------------------|----------|
| Low    | 61% | 1600 px | Best fidelity, lightest change |
| Medium | 53% | 1200 px | Balanced (default) |
| High   | 46% | 1024 px | Stronger savings for image-heavy DOCX |
| Max    | 40% | 900 px  | Smallest possible output |

Each step back from Max adds roughly +15% quality. The dropdown labels show the values in brackets (e.g. `Medium (Balanced, Q=53%, Cap 1200px)`) so the trade-off is visible before the user runs a job.

### PDF

Every level does the same work on PDF: metadata wipe, re-save with `useObjectStreams: true`. If the rewritten output is not smaller than the original, the original bytes are returned unchanged. No rasterization, no downsampling.

### DOCX

Every level strips metadata and re-zips with maximum DEFLATE. At every level, embedded images in `word/media/` are re-encoded: JPEG, PNG, BMP, and TIFF are all decoded to a canvas, scaled to fit within the level's `maxDim` cap on the longest edge, then re-encoded as JPEG at the level's quality. If the re-encoded output is not smaller than the original bytes, the original is kept.

Additional text-affecting operations, applied by default at the levels indicated:

| Operation | Where | Notes |
|-----------|-------|-------|
| Strip tracked changes (`<w:del>`, unwrap `<w:ins>`) | High, Max | Off at Low/Medium so text edits are preserved. |
| Strip comment markers | High only | Off at Max to keep doc legible when shared. |
| XML minify (whitespace strip on XML parts) | High only | Off at Max. |
| Font strip (drop `word/fonts/` entries) | High only | Off at Max. |

These options can also be toggled manually in the DOCX options fieldset.

## DOCX options

The DOCX options fieldset appears only when at least one DOCX file is in the queue:

- **Strip tracked changes** (off by default) — removes `<w:del>` blocks and unwraps `<w:ins>` blocks from `word/document.xml`.
- **Strip comments** (off by default) — removes comment range markers and references from `word/document.xml`.
- **Strip metadata** (on by default) — replaces `docProps/core.xml` and `docProps/app.xml` with empty stubs.

## Behavior

- Drag-and-drop or click the drop zone to queue files. Each file is listed with size and extension. Invalid files (wrong extension, oversize) are reported in the errors list and skipped.
- The Compress button runs the selected level across every accepted file, then triggers a single ZIP download (`compressed-<timestamp>.zip`) containing `*.compressed.pdf` / `*.compressed.docx` entries.
- Transparent PNGs are preserved: only opaque images are candidates for re-encoding where doing so would help. (At present all images run through the same quality/cap pipeline; if re-encoded bytes are larger, the original is kept.)
- Single file at a time is processed; multiple files run sequentially on the main thread.

## Architecture

Single static page, no backend:

- `index.html` — interface (drop zone, level selector, DOCX options, file list, status, errors, compress button, favicon).
- `app.js` — validation, compression, image re-encoding, ZIP packaging, download trigger.
- `favicon.png` — site icon (served without auth; see CloudFront auth note below).
- `lib/jszip.min.js` — vendored JSZip runtime (read/write DOCX as ZIP).
- `lib/pdf-lib.min.js` — vendored pdf-lib runtime (load/save PDF).
- `deploy.sh` — idempotent S3 + CloudFront deployment.
- `tests/app.test.mjs` — unit tests (Node.js built-in test runner).

Not included: no pdf.js, no rasterization, no downsampling of text, no worker threads. All processing runs on the main thread.

## Development

Run the test suite:

```bash
node --test tests/app.test.mjs
```

> The test file is partly behind the current code shape (download-results behavior and the JPEG-at-High assertion predate the latest changes). Several tests are expected to fail until they catch up; the core helpers (`parseLevel`, `validateFile`, `compressPdf`, `stripTrackedChanges`, `stripCommentMarkers`, `reencodeMediaImage`) all pass.

Serve the page locally:

```bash
python -m http.server 8000
```

Then open `http://localhost:8000`.

## Deployment

`deploy.sh` uploads the static site to S3 in `eu-west-3` and invalidates CloudFront if a distribution is configured:

```bash
AWS_ACCOUNT_ID=664882921031 ./deploy.sh
```

The bucket name defaults to `hbtrained-compression-web`. Override with:

```bash
BUCKET_NAME=my-bucket AWS_ACCOUNT_ID=664882921031 ./deploy.sh
```

Public reads are granted only for the static site files (`index.html`, `favicon.png`, `app.js`, `lib/jszip.min.js`, `lib/pdf-lib.min.js`). The script does not grant public upload. CloudFront distribution ID and domain are read from `~/.square-fit-batch/state` (line 1 = distribution ID, line 2 = domain) and used for cache invalidation.

### CloudFront auth + favicon bypass

The CloudFront distribution fronts the site with a Basic-Auth viewer-request function (`square-fit-basic-auth`). The function allows `GET /favicon.png` (and `/favicon.ico`) through without credentials so the browser can fetch the tab icon without prompting. All other paths require the auth header. To update the function, see `docs/` or rerun the standard `aws cloudfront update-function` + `aws cloudfront publish-function` flow against `square-fit-basic-auth` in `us-east-1`.

## Known limitations

- **Fidelity-preserving only.** No Ghostscript-equivalent exists in the browser, so scanned and image-heavy PDFs will see little or no compression. The app never re-encodes PDF page content.
- **No rasterization.** Text remains selectable, vector content (form fields, links, annotations) is preserved at every level.
- **Validation is structural.** Output PDFs are re-loaded; output DOCX ZIPs are reopened and `word/document.xml` is checked. Page count and word/document.xml presence are verified, not visual fidelity.
- **Single-threaded.** All compression runs as single-threaded JavaScript on the main thread.
- **No data loss if smaller fails.** If the optimized output is not smaller than the original, or if re-loading fails for any reason, the original bytes are returned unchanged.