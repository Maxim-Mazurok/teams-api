import { describe, it, expect } from "vitest";
import {
  decodeHtmlEntities,
  escapeHtml,
  buildQuoteHtml,
} from "../../src/html-utils.js";

describe("decodeHtmlEntities", () => {
  it("decodes &nbsp; to space", () => {
    expect(decodeHtmlEntities("hello&nbsp;world")).toBe("hello world");
  });

  it("decodes &quot; to double quote", () => {
    expect(decodeHtmlEntities("&quot;quoted&quot;")).toBe('"quoted"');
  });

  it("decodes &amp; to ampersand", () => {
    expect(decodeHtmlEntities("a &amp; b")).toBe("a & b");
  });

  it("decodes &lt; to less than", () => {
    expect(decodeHtmlEntities("a &lt; b")).toBe("a < b");
  });

  it("decodes &gt; to greater than", () => {
    expect(decodeHtmlEntities("a &gt; b")).toBe("a > b");
  });

  it("removes &#8203; (zero-width space)", () => {
    expect(decodeHtmlEntities("hello&#8203;world")).toBe("helloworld");
  });

  it("decodes numeric character references (&#<number>;)", () => {
    expect(decodeHtmlEntities("&#65;&#66;&#67;")).toBe("ABC");
  });

  it("decodes multiple entities in one string", () => {
    expect(decodeHtmlEntities("a &amp; b &lt; c &gt; d")).toBe("a & b < c > d");
  });

  it("returns string unchanged when no entities present", () => {
    expect(decodeHtmlEntities("just plain text")).toBe("just plain text");
  });

  it("does not decode &apos; (not handled)", () => {
    expect(decodeHtmlEntities("It&apos;s")).toBe("It&apos;s");
  });

  it("handles mixed content with unhandled and handled entities", () => {
    expect(decodeHtmlEntities("It&apos;s &lt;b&gt;bold&lt;/b&gt;")).toBe(
      "It&apos;s <b>bold</b>",
    );
  });

  it("handles multiple &nbsp; in a row", () => {
    expect(decodeHtmlEntities("a&nbsp;&nbsp;&nbsp;b")).toBe("a   b");
  });

  it("handles empty string", () => {
    expect(decodeHtmlEntities("")).toBe("");
  });
});

describe("escapeHtml", () => {
  it("escapes ampersands", () => {
    expect(escapeHtml("a & b")).toBe("a &amp; b");
  });

  it("escapes angle brackets", () => {
    expect(escapeHtml("<script>alert(1)</script>")).toBe(
      "&lt;script&gt;alert(1)&lt;/script&gt;",
    );
  });

  it("escapes double quotes", () => {
    expect(escapeHtml('"hello"')).toBe("&quot;hello&quot;");
  });

  it("escapes all special characters together", () => {
    expect(escapeHtml('a & b < c > d "e"')).toBe(
      "a &amp; b &lt; c &gt; d &quot;e&quot;",
    );
  });

  it("returns plain text unchanged", () => {
    expect(escapeHtml("hello world")).toBe("hello world");
  });

  it("handles empty string", () => {
    expect(escapeHtml("")).toBe("");
  });
});

describe("buildQuoteHtml", () => {
  const baseOptions = {
    messageId: "1719500000000",
    senderMri: "8:orgid:abc-def-123",
    senderDisplayName: "Jane Doe",
    previewText: "Hello, this is a test message",
  };

  it("produces a blockquote with Skype schema microdata", () => {
    const result = buildQuoteHtml(baseOptions);
    expect(result).toContain('itemtype="http://schema.skype.com/Reply"');
    expect(result).toContain('itemid="1719500000000"');
    expect(result).toContain('itemprop="mri"');
    expect(result).toContain('itemprop="preview"');
  });

  it("includes sender display name and MRI", () => {
    const result = buildQuoteHtml(baseOptions);
    expect(result).toContain("Jane Doe");
    expect(result).toContain('itemid="8:orgid:abc-def-123"');
  });

  it("includes the preview text", () => {
    const result = buildQuoteHtml(baseOptions);
    expect(result).toContain("Hello, this is a test message");
  });

  it("truncates preview text longer than 200 characters", () => {
    const longText = "x".repeat(250);
    const result = buildQuoteHtml({ ...baseOptions, previewText: longText });
    expect(result).toContain("x".repeat(197) + "...");
    expect(result).not.toContain("x".repeat(201));
  });

  it("does not truncate preview text of exactly 200 characters", () => {
    const exactText = "y".repeat(200);
    const result = buildQuoteHtml({ ...baseOptions, previewText: exactText });
    expect(result).toContain("y".repeat(200));
    expect(result).not.toContain("...");
  });

  it("escapes HTML in sender display name", () => {
    const result = buildQuoteHtml({
      ...baseOptions,
      senderDisplayName: '<script>alert("xss")</script>',
    });
    expect(result).not.toContain("<script>");
    expect(result).toContain("&lt;script&gt;");
  });

  it("escapes HTML in preview text", () => {
    const result = buildQuoteHtml({
      ...baseOptions,
      previewText: "a <b>bold</b> & important",
    });
    expect(result).toContain("a &lt;b&gt;bold&lt;/b&gt; &amp; important");
  });

  it("wraps content in blockquote tags", () => {
    const result = buildQuoteHtml(baseOptions);
    expect(result).toMatch(/^<blockquote\b/);
    expect(result).toMatch(/<\/blockquote>$/);
  });
});
