/**
 * Shared HTML utilities for entity decoding, escaping, and Teams markup.
 *
 * Used by both the VTT transcript parser, the action formatters,
 * and the quote/inline-reply builder.
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

/** Escape characters that are special in HTML to prevent injection. */
export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Build a Teams inline-quote blockquote HTML element.
 *
 * Teams represents inline quotes using a `<blockquote>` with Skype schema
 * microdata attributes. This is the format Teams Desktop/Web produces when
 * a user right-clicks a message and selects "Reply".
 *
 * @param messageId  - The ID of the quoted message (OriginalArrivalTime).
 * @param senderMri  - Full MRI of the quoted message sender (e.g. "8:orgid:{uuid}").
 * @param senderDisplayName - Display name of the quoted message sender.
 * @param previewText - Plain-text preview of the quoted message content.
 */
export function buildQuoteHtml(options: {
  messageId: string;
  senderMri: string;
  senderDisplayName: string;
  previewText: string;
}): string {
  const truncatedPreview =
    options.previewText.length > 200
      ? `${options.previewText.slice(0, 197)}...`
      : options.previewText;

  return (
    `<blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="${escapeHtml(options.messageId)}">` +
    `<strong itemprop="mri" itemid="${escapeHtml(options.senderMri)}">${escapeHtml(options.senderDisplayName)}</strong>` +
    `<p itemprop="preview">${escapeHtml(truncatedPreview)}</p>` +
    `</blockquote>`
  );
}
