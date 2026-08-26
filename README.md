# Compression Web

Compression Web compresses PDF and DOCX files entirely in the browser. Files never leave the user's machine: the app is a static page that uses vendored JavaScript libraries (pdf-lib and JSZip) to do all the work locally, then bundles the results into a single ZIP for download. There is no backend.

## Accepted formats

- PDF
- DOCX

Maximum 100 MB per file. There is no hard cap on the total size in flight; browser memory is the practical limit.

## Compression levels

All three levels preserve fidelity exactly. No rasterization, no downsampling, no text-to-image conversion occurs at any level.

| Level | What it does | Use when |
|-------|--------------|----------|
| Low | Structural compression only | Best fidelity, lightest change |
| Medium | Structural compression + DOCX media re-encoding | Default for most files |
| High | Structural compression + DOCX media re-encoding at higher quality | Smallest output for image-heavy DOCX |

For PDF, every level does the same work: metadata wipe, re-save with `useObjectStreams: true`. If the rewritten output is not smaller than the original, the original bytes are returned unchanged.

For DOCX, every level strips metadata and re-zips with maximum DEFLATE. At Medium and High, embedded images in `word/media/` are decoded and re-encoded as JPEG at the level's quality:

| Level | JPEG quality |
|-------|--------------|
| Low | 0.5 |
| Medium | 0.7 |
| High | 0.85 |

## DOCX options

The DOCX options fieldset appears only when at least one DOCX file is in the queue:

- **Strip tracked changes** (off by default) — removes `<w:del>` blocks and unwraps `<w:ins>` blocks from `word/document.xml`.
- **Strip comments** (off by default) — removes comment range markers and references from `word/document.xml`.
- **Strip metadata** (on by default) — replaces `docProps/core.xml` and `docProps/app.xml` with empty stubs.

## Architecture

Single static page, no backend:

- `index.html` — interface (drop zone, level selector, DOCX options, file list, status, errors, compress button).
- `app.js` — validation, compression, ZIP packaging, download trigger.
- `lib/jszip.min.js` — vendored JSZip runtime (read/write DOCX as ZIP).
- `lib/pdf-lib.min.js` — vendored pdf-lib runtime (load/save PDF).
- `deploy.sh` — idempotent S3 + CloudFront deployment.
- `tests/app.test.mjs` — unit tests (Node.js built-in test runner).

Not included: no pdf.js, no rasterization, no downsampling, no worker threads. All processing runs on the main thread.

## Development

Run the test suite:

```bash
node --test tests/app.test.mjs
```

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

The bucket name defaults to `hbtrained-compression-web-${AWS_ACCOUNT_ID}`. Override with:

```bash
BUCKET_NAME=my-bucket AWS_ACCOUNT_ID=664882921031 ./deploy.sh
```

Public reads are granted only for the static site files (`index.html`, `app.js`, `lib/jszip.min.js`, `lib/pdf-lib.min.js`). The script does not grant public upload. CloudFront distribution ID and domain are read from `~/.square-fit-batch/state` (line 1 = distribution ID, line 2 = domain) and used for cache invalidation.

## Known limitations

- **Fidelity-preserving only.** No Ghostscript-equivalent exists in the browser, so scanned and image-heavy PDFs will see little or no compression. The app never re-encodes PDF page content.
- **No rasterization.** Text remains selectable, vector content (form fields, links, annotations) is preserved at every level.
- **Validation is structural.** Output PDFs are re-loaded; output DOCX ZIPs are reopened and `word/document.xml` is checked. Page count and word/document.xml presence are verified, not visual fidelity.
- **Single-threaded.** All compression runs as single-threaded JavaScript on the main thread.
- **No data loss if smaller fails.** If the optimized output is not smaller than the original, or if re-loading fails for any reason, the original bytes are returned unchanged.