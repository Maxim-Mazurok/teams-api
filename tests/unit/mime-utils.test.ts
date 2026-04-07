import { describe, it, expect } from "vitest";
import { extensionFromContentType } from "../../src/actions/mime-utils.js";

describe("extensionFromContentType", () => {
  it("should return jpg for image/jpeg", () => {
    expect(extensionFromContentType("image/jpeg")).toBe("jpg");
  });

  it("should return png for image/png", () => {
    expect(extensionFromContentType("image/png")).toBe("png");
  });

  it("should return gif for image/gif", () => {
    expect(extensionFromContentType("image/gif")).toBe("gif");
  });

  it("should return webp for image/webp", () => {
    expect(extensionFromContentType("image/webp")).toBe("webp");
  });

  it("should handle content type with charset parameter", () => {
    expect(extensionFromContentType("image/png; charset=utf-8")).toBe("png");
  });

  it("should be case insensitive", () => {
    expect(extensionFromContentType("Image/PNG")).toBe("png");
  });

  it("should fall back to jpg for unknown content types", () => {
    expect(extensionFromContentType("application/octet-stream")).toBe("jpg");
  });
});
