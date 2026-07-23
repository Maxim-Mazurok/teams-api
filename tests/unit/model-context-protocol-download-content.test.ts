import { resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { describe, expect, it } from "vitest";
import type { DownloadResult } from "../../src/actions/file-actions.js";
import {
  buildDownloadContentBlocks,
  createDownloadOutputResults,
} from "../../src/model-context-protocol-download-content.js";

function createDownloadResult(
  overrides: Partial<DownloadResult> = {},
): DownloadResult {
  return {
    fileName: "report.pdf",
    fileType: "pdf",
    size: 4,
    contentType: "application/pdf",
    savedTo: "/downloads/report.pdf",
    data: Buffer.from("data"),
    ...overrides,
  };
}

describe("buildDownloadContentBlocks", () => {
  it("returns image data inline", () => {
    const imageData = Buffer.from("image data");

    expect(
      buildDownloadContentBlocks([
        createDownloadResult({
          fileName: "photo.png",
          fileType: "png",
          contentType: "image/png",
          savedTo: "/downloads/photo.png",
          data: imageData,
        }),
      ]),
    ).toEqual([
      {
        type: "image",
        data: imageData.toString("base64"),
        mimeType: "image/png",
      },
    ]);
  });

  it("recognizes image MIME types case-insensitively", () => {
    const imageData = Buffer.from("image data");

    expect(
      buildDownloadContentBlocks([
        createDownloadResult({
          fileName: "photo.png",
          fileType: "png",
          contentType: "Image/PNG",
          savedTo: "/downloads/photo.png",
          data: imageData,
        }),
      ]),
    ).toEqual([
      {
        type: "image",
        data: imageData.toString("base64"),
        mimeType: "Image/PNG",
      },
    ]);
  });

  it("returns text file content inline", () => {
    const savedTo = resolve("/downloads/meeting notes#review.txt");

    expect(
      buildDownloadContentBlocks([
        createDownloadResult({
          fileName: "notes.txt",
          fileType: "txt",
          contentType: "text/plain",
          savedTo,
          data: Buffer.from("meeting notes"),
        }),
      ]),
    ).toEqual([
      {
        type: "resource",
        resource: {
          uri: pathToFileURL(savedTo).toString(),
          mimeType: "text/plain",
          text: "meeting notes",
        },
      },
    ]);
  });

  it("recognizes text content by extension when the MIME type is generic", () => {
    expect(
      buildDownloadContentBlocks([
        createDownloadResult({
          fileName: "config.yaml",
          fileType: "yaml",
          contentType: "application/octet-stream",
          savedTo: "/downloads/config.yaml",
          data: Buffer.from("enabled: true"),
        }),
      ]),
    ).toEqual([
      {
        type: "resource",
        resource: {
          uri: "file:///downloads/config.yaml",
          mimeType: "application/octet-stream",
          text: "enabled: true",
        },
      },
    ]);
  });

  it("omits non-image binary content", () => {
    expect(buildDownloadContentBlocks([createDownloadResult()])).toEqual([]);
  });

  it("removes binary data from detailed download output", () => {
    expect(createDownloadOutputResults([createDownloadResult()])).toEqual([
      {
        fileName: "report.pdf",
        fileType: "pdf",
        size: 4,
        contentType: "application/pdf",
        savedTo: "/downloads/report.pdf",
        byteLength: 4,
      },
    ]);
  });
});
