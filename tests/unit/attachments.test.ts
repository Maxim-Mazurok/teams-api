/**
 * Unit tests for attachment parsing, building, and API functions (src/api/attachments.ts).
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import {
  parseInlineImages,
  parseFileAttachments,
  fetchAmsImage,
  uploadAmsImage,
  buildAmsImageTag,
  uploadSharePointFile,
  buildFilesPropertyJson,
} from "../../src/api/attachments.js";
import { ApiAuthError } from "../../src/api/common.js";
import type { TeamsToken } from "../../src/types.js";

const testToken: TeamsToken = {
  skypeToken: "test-skype-token",
  region: "apac",
  bearerToken: "test-bearer-token",
  amsToken: "test-ams-token",
};

const originalFetch = globalThis.fetch;

beforeEach(() => {
  globalThis.fetch = vi.fn();
});

afterEach(() => {
  globalThis.fetch = originalFetch;
});

describe("parseInlineImages", () => {
  it("returns empty array for content with no images", () => {
    expect(parseInlineImages("Hello world")).toEqual([]);
    expect(parseInlineImages("")).toEqual([]);
  });

  it("parses a single inline image with dimensions", () => {
    const content =
      'Some text <img src="https://as-prod.asyncgw.teams.microsoft.com/v1/objects/0-abc-d1-def123/views/imgo" ' +
      'itemscope="" itemtype="http://schema.skype.com/AMSImage" ' +
      'style="width:400px; height:300px"> more text';

    const images = parseInlineImages(content);
    expect(images).toHaveLength(1);
    expect(images[0].amsObjectId).toBe("0-abc-d1-def123");
    expect(images[0].width).toBe(400);
    expect(images[0].height).toBe(300);
    expect(images[0].fullSizeUrl).toBe(
      "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/0-abc-d1-def123/views/imgpsh_fullsize_anim",
    );
    expect(images[0].contentPosition).toBe(10); // position of <img in the string
  });

  it("parses multiple inline images", () => {
    const content =
      '<img src="https://as-prod.asyncgw.teams.microsoft.com/v1/objects/obj-1/views/imgo" ' +
      'itemtype="http://schema.skype.com/AMSImage">' +
      'text between' +
      '<img src="https://as-prod.asyncgw.teams.microsoft.com/v1/objects/obj-2/views/imgo" ' +
      'itemtype="http://schema.skype.com/AMSImage">';

    const images = parseInlineImages(content);
    expect(images).toHaveLength(2);
    expect(images[0].amsObjectId).toBe("obj-1");
    expect(images[1].amsObjectId).toBe("obj-2");
    expect(images[1].contentPosition).toBeGreaterThan(images[0].contentPosition);
  });

  it("ignores non-AMS images", () => {
    const content =
      '<img src="https://example.com/photo.jpg" itemtype="http://schema.skype.com/Other">';
    expect(parseInlineImages(content)).toEqual([]);
  });

  it("ignores images without src attribute", () => {
    const content =
      '<img itemtype="http://schema.skype.com/AMSImage">';
    expect(parseInlineImages(content)).toEqual([]);
  });

  it("handles images without dimensions", () => {
    const content =
      '<img src="https://as-prod.asyncgw.teams.microsoft.com/v1/objects/no-dims/views/imgo" ' +
      'itemtype="http://schema.skype.com/AMSImage">';

    const images = parseInlineImages(content);
    expect(images).toHaveLength(1);
    expect(images[0].width).toBeNull();
    expect(images[0].height).toBeNull();
  });
});

describe("parseFileAttachments", () => {
  it("returns empty array for null/undefined input", () => {
    expect(parseFileAttachments(null)).toEqual([]);
    expect(parseFileAttachments(undefined)).toEqual([]);
  });

  it("returns empty array for invalid JSON string", () => {
    expect(parseFileAttachments("not valid json")).toEqual([]);
  });

  it("returns empty array for non-array JSON string", () => {
    expect(parseFileAttachments('{"key": "value"}')).toEqual([]);
  });

  it("parses a single file attachment from JSON string", () => {
    const rawFiles = JSON.stringify([
      {
        "@type": "http://schema.skype.com/File",
        itemid: "item-123",
        fileName: "report.pdf",
        fileType: ".pdf",
        fileInfo: {
          fileUrl: "https://sharepoint.com/report.pdf",
          shareUrl: "https://sharepoint.com/share/report.pdf",
        },
      },
    ]);

    const files = parseFileAttachments(rawFiles);
    expect(files).toHaveLength(1);
    expect(files[0]).toEqual({
      itemId: "item-123",
      fileName: "report.pdf",
      fileType: ".pdf",
      fileUrl: "https://sharepoint.com/report.pdf",
      shareUrl: "https://sharepoint.com/share/report.pdf",
    });
  });

  it("parses file attachments from an array directly", () => {
    const rawFiles = [
      {
        "@type": "http://schema.skype.com/File",
        itemid: "item-1",
        fileName: "doc.docx",
        fileType: ".docx",
        fileInfo: { fileUrl: "https://sp.com/doc.docx", shareUrl: "" },
      },
      {
        "@type": "http://schema.skype.com/File",
        id: "item-2",
        fileName: "sheet.xlsx",
        fileType: ".xlsx",
        objectUrl: "https://sp.com/sheet.xlsx",
      },
    ];

    const files = parseFileAttachments(rawFiles);
    expect(files).toHaveLength(2);
    expect(files[0].itemId).toBe("item-1");
    expect(files[0].fileName).toBe("doc.docx");
    expect(files[1].itemId).toBe("item-2");
    expect(files[1].fileName).toBe("sheet.xlsx");
    expect(files[1].fileUrl).toBe("https://sp.com/sheet.xlsx");
  });

  it("filters out entries without File type", () => {
    const rawFiles = JSON.stringify([
      {
        "@type": "http://schema.skype.com/File",
        fileName: "valid.pdf",
        fileType: ".pdf",
      },
      {
        "@type": "http://schema.skype.com/Link",
        fileName: "link.html",
      },
    ]);

    const files = parseFileAttachments(rawFiles);
    expect(files).toHaveLength(1);
    expect(files[0].fileName).toBe("valid.pdf");
  });

  it("filters out entries without fileName", () => {
    const rawFiles = JSON.stringify([
      {
        "@type": "http://schema.skype.com/File",
        fileType: ".pdf",
      },
    ]);

    expect(parseFileAttachments(rawFiles)).toEqual([]);
  });

  it("uses fallback fields when primary fields are missing", () => {
    const rawFiles = JSON.stringify([
      {
        "@type": "http://schema.skype.com/File",
        id: "fallback-id",
        fileName: "fallback.txt",
        objectUrl: "https://sp.com/fallback.txt",
      },
    ]);

    const files = parseFileAttachments(rawFiles);
    expect(files).toHaveLength(1);
    expect(files[0].itemId).toBe("fallback-id");
    expect(files[0].fileUrl).toBe("https://sp.com/fallback.txt");
    expect(files[0].shareUrl).toBe("");
  });
});

describe("buildAmsImageTag", () => {
  it("builds tag without dimensions", () => {
    const tag = buildAmsImageTag("my-object-id");
    expect(tag).toBe(
      '<img src="https://as-prod.asyncgw.teams.microsoft.com/v1/objects/my-object-id/views/imgo" ' +
        'itemscope="" itemtype="http://schema.skype.com/AMSImage">',
    );
  });

  it("builds tag with dimensions", () => {
    const tag = buildAmsImageTag("my-object-id", 640, 480);
    expect(tag).toContain('style="width:640px; height:480px"');
    expect(tag).toContain('itemtype="http://schema.skype.com/AMSImage"');
  });

  it("omits style when only one dimension is provided", () => {
    const tag = buildAmsImageTag("my-object-id", 640);
    expect(tag).not.toContain("style=");
  });
});

describe("fetchAmsImage", () => {
  it("fetches an image with skype token auth", async () => {
    const imageBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47]); // PNG magic bytes
    const mockFetch = vi.fn().mockResolvedValue({
      ok: true,
      headers: new Headers({ "content-type": "image/png" }),
      arrayBuffer: () => Promise.resolve(imageBytes.buffer.slice(0, imageBytes.byteLength)),
    });
    globalThis.fetch = mockFetch;

    const result = await fetchAmsImage(testToken, "test-object-id");
    expect(result.contentType).toBe("image/png");
    expect(result.size).toBe(4);

    const calledUrl = mockFetch.mock.calls[0][0];
    expect(calledUrl).toContain("test-object-id/views/imgo");

    const calledOptions = mockFetch.mock.calls[0][1];
    expect(calledOptions.headers.Authorization).toBe(
      "skype_token test-skype-token",
    );
  });

  it("uses full-size view when specified", async () => {
    const mockFetch = vi.fn().mockResolvedValue({
      ok: true,
      headers: new Headers({ "content-type": "image/jpeg" }),
      arrayBuffer: () => Promise.resolve(new ArrayBuffer(0)),
    });
    globalThis.fetch = mockFetch;

    await fetchAmsImage(testToken, "test-object-id", "imgpsh_fullsize_anim");

    const calledUrl = mockFetch.mock.calls[0][0];
    expect(calledUrl).toContain("test-object-id/views/imgpsh_fullsize_anim");
  });

  it("throws ApiAuthError on 401", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 401,
      statusText: "Unauthorized",
    });

    await expect(fetchAmsImage(testToken, "test-id")).rejects.toThrow(
      ApiAuthError,
    );
  });

  it("throws generic error on other failures", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 500,
      statusText: "Internal Server Error",
    });

    await expect(fetchAmsImage(testToken, "test-id")).rejects.toThrow(
      "Failed to fetch AMS image: 500 Internal Server Error",
    );
  });
});

describe("uploadAmsImage", () => {
  const imageData = Buffer.from([0x89, 0x50, 0x4e, 0x47]);
  const conversationId = "19:abc123@thread.v2";

  it("throws when amsToken is missing", async () => {
    const tokenWithoutAms: TeamsToken = {
      skypeToken: "test-skype",
      region: "apac",
    };

    await expect(
      uploadAmsImage(tokenWithoutAms, imageData, "test.png", conversationId),
    ).rejects.toThrow("AMS token is required");
  });

  it("creates object and uploads content in two steps", async () => {
    const mockFetch = vi
      .fn()
      // Step 1: create object
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ id: "created-object-id" }),
      })
      // Step 2: upload content
      .mockResolvedValueOnce({
        ok: true,
      });
    globalThis.fetch = mockFetch;

    const result = await uploadAmsImage(
      testToken,
      imageData,
      "screenshot.png",
      conversationId,
    );

    expect(result.amsObjectId).toBe("created-object-id");

    // Verify create request
    expect(mockFetch).toHaveBeenCalledTimes(2);
    const createCall = mockFetch.mock.calls[0];
    expect(createCall[1].method).toBe("POST");
    expect(createCall[1].headers["Authorization"]).toBe(
      "Bearer test-ams-token",
    );
    expect(createCall[1].headers["x-ms-client-version"]).toBe(
      "1415/26022704215",
    );
    const createBody = JSON.parse(createCall[1].body as string);
    expect(createBody.type).toBe("pish/image");
    expect(createBody.permissions[conversationId]).toEqual(["read"]);
    expect(createBody.filename).toBe("screenshot.png");

    // Verify upload request
    const uploadCall = mockFetch.mock.calls[1];
    expect(uploadCall[0]).toContain("created-object-id/content/imgpsh");
    expect(uploadCall[1].method).toBe("PUT");
    expect(uploadCall[1].headers["Authorization"]).toBe(
      "Bearer test-ams-token",
    );
  });

  it("throws ApiAuthError on 401 during create", async () => {
    globalThis.fetch = vi.fn().mockResolvedValueOnce({
      ok: false,
      status: 401,
      statusText: "Unauthorized",
    });

    await expect(
      uploadAmsImage(testToken, imageData, "test.png", conversationId),
    ).rejects.toThrow(ApiAuthError);
  });

  it("throws generic error on non-401 failure during create", async () => {
    globalThis.fetch = vi.fn().mockResolvedValueOnce({
      ok: false,
      status: 400,
      statusText: "Bad Request",
      text: () => Promise.resolve("Missing required field"),
    });

    await expect(
      uploadAmsImage(testToken, imageData, "test.png", conversationId),
    ).rejects.toThrow("Failed to create AMS object: 400 Bad Request");
  });

  it("throws ApiAuthError on 401 during upload step", async () => {
    globalThis.fetch = vi
      .fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ id: "obj-id" }),
      })
      .mockResolvedValueOnce({
        ok: false,
        status: 401,
        statusText: "Unauthorized",
      });

    await expect(
      uploadAmsImage(testToken, imageData, "test.png", conversationId),
    ).rejects.toThrow(ApiAuthError);
  });
});

describe("uploadSharePointFile", () => {
  const fileData = Buffer.from("# Hello World\n");
  const testEmail = "user.name@company.com";

  const tokenWithSharePoint: TeamsToken = {
    skypeToken: "test-skype-token",
    region: "apac",
    sharePointToken: "test-sharepoint-token",
    sharePointHost: "company-my.sharepoint.com",
  };

  it("throws when sharePointToken is missing", async () => {
    const tokenWithoutSharePoint: TeamsToken = {
      skypeToken: "test-skype",
      region: "apac",
    };

    await expect(
      uploadSharePointFile(tokenWithoutSharePoint, fileData, "test.md", testEmail),
    ).rejects.toThrow("SharePoint token is required");
  });

  it("throws when sharePointHost is missing", async () => {
    const tokenWithoutHost: TeamsToken = {
      skypeToken: "test-skype",
      region: "apac",
      sharePointToken: "test-sharepoint-token",
    };

    await expect(
      uploadSharePointFile(tokenWithoutHost, fileData, "test.md", testEmail),
    ).rejects.toThrow("SharePoint host is required");
  });

  it("uploads file content via PUT and returns metadata", async () => {
    const mockFetch = vi.fn().mockResolvedValueOnce({
      ok: true,
      json: () =>
        Promise.resolve({
          id: "sp-item-id",
          name: "test.md",
          webDavUrl: "https://company-my.sharepoint.com/personal/user_name_company_com/Documents/Microsoft%20Teams%20Chat%20Files/test.md",
          webUrl: "https://company-my.sharepoint.com/personal/user_name_company_com/Documents/Microsoft%20Teams%20Chat%20Files/test.md",
          sharepointIds: {
            listItemUniqueId: "unique-item-id",
            siteId: "site-id-123",
          },
        }),
    });
    globalThis.fetch = mockFetch;

    const result = await uploadSharePointFile(
      tokenWithSharePoint,
      fileData,
      "test.md",
      testEmail,
    );

    expect(result.itemId).toBe("unique-item-id");
    expect(result.siteId).toBe("site-id-123");
    expect(result.fileName).toBe("test.md");
    expect(result.fileType).toBe("md");
    expect(result.siteBaseUrl).toBe("https://company-my.sharepoint.com");
    expect(result.personalPath).toBe("/personal/user_name_company_com");

    // Verify the PUT request
    const calledUrl = mockFetch.mock.calls[0][0] as string;
    expect(calledUrl).toContain("/personal/user_name_company_com/_api/v2.0/drive/root:");
    expect(calledUrl).toContain("Microsoft%20Teams%20Chat%20Files/test.md");
    expect(calledUrl).toContain("@name.conflictBehavior=rename");

    const calledOptions = mockFetch.mock.calls[0][1] as RequestInit;
    expect(calledOptions.method).toBe("PUT");
    expect((calledOptions.headers as Record<string, string>).Authorization).toBe(
      "Bearer test-sharepoint-token",
    );
  });

  it("throws ApiAuthError on 401", async () => {
    globalThis.fetch = vi.fn().mockResolvedValueOnce({
      ok: false,
      status: 401,
      statusText: "Unauthorized",
    });

    await expect(
      uploadSharePointFile(tokenWithSharePoint, fileData, "test.md", testEmail),
    ).rejects.toThrow(ApiAuthError);
  });

  it("throws generic error on non-401 failure", async () => {
    globalThis.fetch = vi.fn().mockResolvedValueOnce({
      ok: false,
      status: 403,
      statusText: "Forbidden",
      text: () => Promise.resolve("Access denied"),
    });

    await expect(
      uploadSharePointFile(tokenWithSharePoint, fileData, "test.md", testEmail),
    ).rejects.toThrow("Failed to upload file to SharePoint: 403 Forbidden");
  });

  it("handles file names without extension", async () => {
    globalThis.fetch = vi.fn().mockResolvedValueOnce({
      ok: true,
      json: () =>
        Promise.resolve({
          id: "sp-item-id",
          name: "Makefile",
          webDavUrl: "https://company-my.sharepoint.com/dav/Makefile",
          webUrl: "https://company-my.sharepoint.com/web/Makefile",
          sharepointIds: {
            listItemUniqueId: "item-id",
            siteId: "site-id",
          },
        }),
    });

    const result = await uploadSharePointFile(
      tokenWithSharePoint,
      fileData,
      "Makefile",
      testEmail,
    );

    expect(result.fileType).toBe("");
    expect(result.fileName).toBe("Makefile");
  });
});

describe("buildFilesPropertyJson", () => {
  it("builds valid JSON for a single file", () => {
    const json = buildFilesPropertyJson([
      {
        itemId: "item-123",
        siteId: "site-456",
        fileName: "report.pdf",
        fileType: "pdf",
        fileUrl: "https://sp.com/report.pdf",
        webDavUrl: "https://sp.com/dav/report.pdf",
        siteBaseUrl: "https://company-my.sharepoint.com",
        personalPath: "/personal/user_company_com",
      },
    ]);

    const parsed = JSON.parse(json) as Array<Record<string, unknown>>;
    expect(parsed).toHaveLength(1);
    expect(parsed[0]["@type"]).toBe("http://schema.skype.com/File");
    expect(parsed[0].version).toBe(2);
    expect(parsed[0].id).toBe("item-123");
    expect(parsed[0].itemid).toBe("item-123");
    expect(parsed[0].fileName).toBe("report.pdf");
    expect(parsed[0].fileType).toBe("pdf");
    expect(parsed[0].state).toBe("active");
    expect(parsed[0].permissionScope).toBe("users");

    const sharepointIds = parsed[0].sharepointIds as Record<string, string>;
    expect(sharepointIds.listItemUniqueId).toBe("item-123");
    expect(sharepointIds.siteId).toBe("site-456");
  });

  it("builds valid JSON for multiple files", () => {
    const json = buildFilesPropertyJson([
      {
        itemId: "item-1",
        siteId: "site-1",
        fileName: "doc.docx",
        fileType: "docx",
        fileUrl: "https://sp.com/doc.docx",
        webDavUrl: "https://sp.com/dav/doc.docx",
        siteBaseUrl: "https://sp.com",
        personalPath: "/personal/user",
      },
      {
        itemId: "item-2",
        siteId: "site-2",
        fileName: "sheet.xlsx",
        fileType: "xlsx",
        fileUrl: "https://sp.com/sheet.xlsx",
        webDavUrl: "https://sp.com/dav/sheet.xlsx",
        siteBaseUrl: "https://sp.com",
        personalPath: "/personal/user",
      },
    ]);

    const parsed = JSON.parse(json) as Array<Record<string, unknown>>;
    expect(parsed).toHaveLength(2);
    expect(parsed[0].fileName).toBe("doc.docx");
    expect(parsed[1].fileName).toBe("sheet.xlsx");
  });

  it("returns empty array JSON for no files", () => {
    const json = buildFilesPropertyJson([]);
    expect(json).toBe("[]");
  });
});
