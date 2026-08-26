# Browser-Based Document Compression Web App

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Browser-only static web app that compresses PDF and DOCX files locally in the user's browser, mirroring the functionality of the existing Compression App (`compression-system` + `doc-compressor`) without any backend.

**Architecture:** Single-page HTML/CSS/JS web app. All processing happens client-side via vendored libraries (pdf-lib, JSZip). Files never leave the user's browser. Static hosting on S3 + CloudFront, identical deployment pattern to the Square Fit Batch web app.

**Tech Stack:** pdf-lib (PDF manipulation), JSZip (DOCX as ZIP), HTML5 Canvas (image re-encoding), S3 + CloudFront + ACM cert + CloudFront Function (HTTP Basic Auth).

---

## Functional design

### PDF compression (pdf-lib)

Structural-only — preserves fidelity exactly:

1. Load PDF with `PDFDocument.load(bytes, { updateMetadata: false })`
2. Wipe metadata (Author, Title, Producer, Creator, Subject, Keywords, ModDate)
3. Re-save with `useObjectStreams: true`
4. If output is smaller than original, return output; otherwise return original
5. Validate by re-loading the output; on failure, fall back to original

Typical savings:
- Text PDFs with embedded fonts: 20–50%
- Image-heavy / scanned PDFs: 5–15%
- Already-optimized PDFs: 0–5%

**No Ghostscript-equivalent in the browser** — fidelity-preserving only. Raster re-encoding can come later if needed.

### DOCX compression (JSZip + Canvas)

DOCX files are ZIP containers with XML + embedded images.

1. Load DOCX as ZIP with JSZip
2. Strip `docProps/core.xml` and `docProps/app.xml` to a minimal stub
3. Optionally strip tracked changes (`<w:ins>`, `<w:del>`) and comments from `word/document.xml` based on user-toggled options
4. For each image in `word/media/`:
   - PNG / JPG / JPEG: decode via `createImageBitmap`, re-encode via canvas `toBlob('image/jpeg', quality)` at the level's quality (Low 0.5 / Medium 0.7 / High 0.85)
   - If compression level is High and image is already JPEG, skip re-encoding
5. Re-zip with DEFLATE level 9 (`new JSZip()` with `compression: 'DEFLATE'`, level 9)
6. Validate by reopening the ZIP and confirming `word/document.xml` exists

### UI

- Drop zone (same pattern as Square Fit Batch)
- Compression level dropdown: Low / Medium / High
- DOCX options fieldset (only visible when at least one DOCX file is in the queue):
  - Strip tracked changes (default off)
  - Strip comments (default off)
  - Strip metadata (default on)
- File list with size + format + status
- Status / errors pane
- "Compress & download" button

### Output

A single ZIP named `compressed-{timestamp}.zip` containing one entry per input file with `.compressed.{pdf|docx}` extension. Originals are not modified.

---

## Deployment

- **New S3 bucket** for static hosting
- **New ACM cert** for `compress.hbtrained.com` in `us-east-1` (DNS validation)
- **Existing CloudFront distribution** (`E3AQZAKCYSUKF2`) gets a new alias `compress.hbtrained.com`
- **Existing CloudFront Function** for HTTP Basic Auth applies to all viewer requests on the distribution (covers the new alias automatically)
- **Cloudflare DNS** for `compress.hbtrained.com` → `d1a08qccil1nlo.cloudfront.net` (DNS-only, grey cloud)

### Idempotent deploy (deploy.sh)

Same pattern as Square Fit Batch: create bucket if missing, configure public access block + website, apply narrow policy, upload files with cache headers, invalidate CloudFront.

### Limits

- Per-file: 100 MB max (raised over Square Fit because PDFs and DOCX can be larger than images)
- Total in-flight: bounded by browser memory; no explicit cap

---

## Repo

- **GitHub**: `pg-thirion/compression-web` (new repo, separate from `square-fit-batch`)
- Local path: `J:\Shared drives\HB Trained\Admin\Scripts\Central Program\Complete Programs\Compression Web`

---

## Known limitations (documented in README)

- No Ghostscript-equivalent: scanned PDFs won't compress much
- No raster re-encoding: fidelity is preserved exactly
- Validation is structural (re-open), not pixel-diff
- No multi-threading: PDF compression is single-threaded JS in the browser
