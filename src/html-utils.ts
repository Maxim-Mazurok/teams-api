/**
 * Shared HTML entity decoding utilities.
 *
 * Used by both the VTT transcript parser and the action formatters
 * to convert HTML entities back to plain text.
 */

/** Decode common HTML entities to plain text. */
export function decodeHtmlEntities(text: string): string {
  return text
    .replace(/&nbsp;/g, " ")
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#8203;/g, "") // zero-width space
    .replace(/&#(\d+);/g, (_, code: string) =>
      String.fromCharCode(Number(code)),
    );
}
