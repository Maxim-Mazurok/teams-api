import { describe, it, expect } from "vitest";
import { decodeHtmlEntities } from "../../src/html-utils.js";

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
