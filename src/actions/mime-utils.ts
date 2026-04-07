/**
 * MIME type to file extension mapping utilities.
 */

const contentTypeToExtension: Record<string, string> = {
  "image/jpeg": "jpg",
  "image/png": "png",
  "image/gif": "gif",
  "image/webp": "webp",
  "image/bmp": "bmp",
  "image/svg+xml": "svg",
  "image/tiff": "tiff",
  "image/x-icon": "ico",
  "image/avif": "avif",
};

/**
 * Derive a file extension from a Content-Type header value.
 *
 * Falls back to "jpg" when the content type is unrecognised.
 */
export function extensionFromContentType(contentType: string): string {
  const mimeType = contentType.split(";")[0].trim().toLowerCase();
  return contentTypeToExtension[mimeType] ?? "jpg";
}
