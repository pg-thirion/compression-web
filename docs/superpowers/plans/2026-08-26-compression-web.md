# Compression Web — Implementation Plan

**Goal:** Browser-only static web app that compresses PDF and DOCX files locally.

**Stack:** pdf-lib, JSZip, HTML5 Canvas. Node test runner.

**Files:** index.html, app.js, deploy.sh, README.md, lib/jszip.min.js (vendored), lib/pdf-lib.min.js (vendored), tests/app.test.mjs.

---

## Task 1: Bootstrap + vendor libs

- Copy `lib/jszip.min.js` from Square Fit Batch worktree
- Download pdf-lib 1.17.1 UMD to `lib/pdf-lib.min.js` via curl
- Verify both files load in node

## Task 2: Validation + level parsing (TDD)

Tests in `tests/app.test.mjs`, impl in `app.js`:

- `parseLevel(value)` → 'low'|'medium'|'high', throws RangeError otherwise
- `validateFile(file)` → null or error message
- Constants: `ALLOWED_LEVELS`, `FILE_MAX_BYTES = 100*1024*1024`, `FILE_EXTENSIONS = new Set(['.pdf', '.docx'])`

## Task 3: PDF compression (TDD)

`compressPdf(bytes, deps)`:
- Load with `updateMetadata: false`
- Wipe metadata (title, author, subject, keywords, creator, producer)
- Save with `useObjectStreams: true`
- Keep output only if smaller AND re-loads with same page count
- Otherwise return original blob
- Default deps: `{ PDFDocument: globalThis.PDFLib.PDFDocument }`

## Task 4: DOCX XML helpers (TDD)

- `minimalCoreXml()` / `minimalAppXml()` → empty stub metadata XML
- `stripTrackedChanges(xml)` → drop `<w:del>...</w:del>` entirely, unwrap `<w:ins>...</w:ins>` keeping content
- `stripCommentMarkers(xml)` → drop `<w:commentRangeStart/>`, `<w:commentRangeEnd/>`, `<w:commentReference/>`

## Task 5: DOCX media re-encoding (TDD)

`reencodeMediaImage(bytes, mimeType, quality, deps)`:
- createImageBitmap → canvas drawImage → toBlob('image/jpeg', quality) → Uint8Array
- Always close bitmap in finally
- Test mocks createImageBitmap + createCanvas

## Task 6: DOCX compression (TDD)

`compressDocx(bytes, level, options, deps)`:
- JSZip.loadAsync(bytes)
- Replace core.xml + app.xml with stubs (when options.stripMetadata !== false)
- Apply stripTrackedChanges / stripCommentMarkers to word/document.xml
- For each `word/media/*.{jpg,jpeg}`: re-encode via reencodeMediaImage; replace if smaller
- generateAsync with DEFLATE level 9
- Validate by reopening and checking word/document.xml exists
- Fallback to original blob on any failure

## Task 7: Filter + download + click handler

- `filterFiles(files, errors)` returns accepted, pushes errors
- `downloadResults(results, deps)` mirrors Square Fit: JSZip, blob URL, click, revoke
- Click handler: read files, validate, route PDF/DOCX by extension, build zip, trigger download

## Task 8: DOCX options UI gating

- index.html: add fieldset with strip-tracked-changes, strip-comments, strip-metadata checkboxes
- app.js: show fieldset only when at least one .docx in queue; update on file selection change

## Task 9: deploy.sh

Mirror Square Fit deploy.sh:
- `BUCKET=hbtrained-compression-web-1` (or similar)
- `REGION=eu-west-3`
- Idempotent bucket create + public-access-block + website config + narrow policy
- Upload 4 files with cache headers
- CloudFront invalidation via `~/.square-fit-batch/state` (which holds the distribution ID)
- Print HTTPS URL

## Task 10: README — usage, architecture, limits

## Task 11: Deploy

- Run deploy.sh
- ACM cert for compress.hbtrained.com in us-east-1
- CloudFront alias update
- Cloudflare DNS records (validation + final CNAME)
- Verify with curl + Basic Auth
