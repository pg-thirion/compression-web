// Compression Web — application entry point.
// Browser-only static web app that compresses PDF and DOCX files locally.
// See docs/superpowers/specs/2026-08-26-compression-web-design.md for the full design.

export const ALLOWED_LEVELS = new Set(['low', 'medium', 'high']);
export const FILE_MAX_BYTES = 100 * 1024 * 1024;
export const FILE_EXTENSIONS = new Set(['.pdf', '.docx']);

/**
 * Parse a compression level string. Strictly case-sensitive.
 * @param {string} value
 * @returns {'low'|'medium'|'high'}
 * @throws {RangeError} if value is not one of the allowed levels
 */
export function parseLevel(value) {
  if (typeof value !== 'string' || !ALLOWED_LEVELS.has(value)) {
    throw new RangeError(`Invalid compression level: ${String(value)}`);
  }
  return value;
}

/**
 * Validate a file for inclusion in the compression queue.
 * @param {{name: string, size: number, type?: string}} file
 * @returns {string|null} null if valid, otherwise a human-readable error message
 */
export function validateFile(file) {
  if (!file || typeof file.name !== 'string' || file.name === '') {
    return 'File is missing a name.';
  }

  const ext = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();
  if (!FILE_EXTENSIONS.has(ext)) {
    return `Unsupported file extension. Allowed: ${[...FILE_EXTENSIONS].join(', ')}.`;
  }

  if (typeof file.size !== 'number' || file.size < 0) {
    return 'File is missing a size.';
  }

  if (file.size > FILE_MAX_BYTES) {
    return `File exceeds maximum size of ${FILE_MAX_BYTES} bytes.`;
  }

  return null;
}
