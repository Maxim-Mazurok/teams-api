/**
 * Attachment parsing and download utilities for Teams messages.
 *
 * Handles two types of attachments:
 * - Inline images (AMS) — embedded as `<img>` tags in message HTML
 * - File attachments (SharePoint) — referenced in `properties.files` JSON
 */

import type { TeamsToken, ImageAttachment, FileAttachment } from "../types.js";
import { fetchWithRetry, ApiAuthError } from "./common.js";

const AMS_BASE = "https://as-prod.asyncgw.teams.microsoft.com/v1/objects";

/**
 * Teams client version header required by the AMS API.
 * AMS uses this to determine the "platform id" and rejects requests without it.
 */
const AMS_CLIENT_VERSION = "1415/26022704215";

/**
 * Extract inline image attachments from message HTML content.
 *
 * Parses `<img>` tags with `itemtype="http://schema.skype.com/AMSImage"` and
 * extracts the AMS object ID, URL, and dimensions from the tag attributes.
 */
export function parseInlineImages(content: string): ImageAttachment[] {
  const images: ImageAttachment[] = [];
  const imagePattern =
    /<img\s+[^>]*itemtype="http:\/\/schema\.skype\.com\/AMSImage"[^>]*>/gi;

  let match: RegExpExecArray | null;
  while ((match = imagePattern.exec(content)) !== null) {
    const tag = match[0];
    const contentPosition = match.index;

    const srcMatch = tag.match(/src="([^"]+)"/);
    if (!srcMatch) continue;

    const url = srcMatch[1];
    const amsIdMatch = url.match(/\/objects\/([^/]+)\//);
    if (!amsIdMatch) continue;

    const amsObjectId = amsIdMatch[1];

    let width: number | null = null;
    let height: number | null = null;
    const widthMatch = tag.match(/width:(\d+)px/);
    const heightMatch = tag.match(/height:(\d+)px/);
    if (widthMatch) width = Number(widthMatch[1]);
    if (heightMatch) height = Number(heightMatch[1]);

    const fullSizeUrl = `${AMS_BASE}/${amsObjectId}/views/imgpsh_fullsize_anim`;

    images.push({
      amsObjectId,
      url,
      fullSizeUrl,
      width,
      height,
      contentPosition,
    });
  }

  return images;
}

/** Raw file object from `properties.files` JSON. */
interface RawFileEntry {
  itemid?: string;
  id?: string;
  fileName?: string;
  fileType?: string;
  fileInfo?: {
    fileUrl?: string;
    shareUrl?: string;
  };
  objectUrl?: string;
  title?: string;
  "@type"?: string;
}

/**
 * Parse file attachments from the raw `properties.files` JSON string.
 *
 * File attachments are SharePoint-hosted documents, videos, and other files
 * shared through Teams. They have a separate schema from inline images.
 */
export function parseFileAttachments(rawFiles: unknown): FileAttachment[] {
  let entries: RawFileEntry[];

  if (typeof rawFiles === "string") {
    try {
      entries = JSON.parse(rawFiles) as RawFileEntry[];
    } catch {
      return [];
    }
  } else if (Array.isArray(rawFiles)) {
    entries = rawFiles as RawFileEntry[];
  } else {
    return [];
  }

  if (!Array.isArray(entries)) return [];

  return entries
    .filter(
      (entry) =>
        entry["@type"] === "http://schema.skype.com/File" && entry.fileName,
    )
    .map((entry) => ({
      itemId: entry.itemid ?? entry.id ?? "",
      fileName: entry.fileName ?? entry.title ?? "",
      fileType: entry.fileType ?? "",
      fileUrl: entry.fileInfo?.fileUrl ?? entry.objectUrl ?? "",
      shareUrl: entry.fileInfo?.shareUrl ?? "",
    }));
}

/**
 * Fetch an image from the AMS (Async Media Service).
 *
 * Returns the raw binary data and content type. Uses the skype token
 * for authentication (same pattern as transcript fetching).
 *
 * @param view - AMS view name: "imgo" (compressed), "imgpsh_fullsize_anim" (full-size)
 */
export async function fetchAmsImage(
  token: TeamsToken,
  amsObjectId: string,
  view: "imgo" | "imgpsh_fullsize_anim" = "imgo",
): Promise<{ data: Buffer; contentType: string; size: number }> {
  const url = `${AMS_BASE}/${amsObjectId}/views/${view}`;

  const response = await fetchWithRetry(url, {
    headers: {
      Authorization: `skype_token ${token.skypeToken}`,
    },
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `AMS image fetch authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    throw new Error(
      `Failed to fetch AMS image: ${response.status} ${response.statusText}`,
    );
  }

  const contentType = response.headers.get("content-type") ?? "image/jpeg";
  const arrayBuffer = await response.arrayBuffer();
  const data = Buffer.from(arrayBuffer);

  return { data, contentType, size: data.length };
}

/**
 * Upload an image to the AMS (Async Media Service).
 *
 * Creates a new AMS object with permissions for the target conversation,
 * uploads the image data, and returns the object ID that can be referenced
 * in message HTML via `<img>` tags.
 *
 * Requires the AMS/IC3 Bearer token (audience: ic3.teams.office.com), not the
 * middle-tier bearer or skype token.
 */
export async function uploadAmsImage(
  token: TeamsToken,
  imageData: Buffer,
  fileName: string,
  conversationId: string,
): Promise<{ amsObjectId: string }> {
  if (!token.amsToken) {
    throw new Error(
      "AMS token is required for image upload but was not captured during authentication.",
    );
  }

  // Step 1: Create the AMS object with inline permissions
  const createResponse = await fetchWithRetry(`${AMS_BASE}/`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token.amsToken}`,
      "Content-Type": "application/json",
      "x-ms-client-version": AMS_CLIENT_VERSION,
    },
    body: JSON.stringify({
      type: "pish/image",
      permissions: { [conversationId]: ["read"] },
      sharingMode: "Attached",
      filename: fileName,
    }),
  });

  if (!createResponse.ok) {
    if (createResponse.status === 401) {
      throw new ApiAuthError(
        `AMS object creation failed: ${createResponse.status} ${createResponse.statusText}`,
      );
    }
    const errorText = await createResponse.text();
    throw new Error(
      `Failed to create AMS object: ${createResponse.status} ${createResponse.statusText} — ${errorText}`,
    );
  }

  const createData = (await createResponse.json()) as { id: string };
  const amsObjectId = createData.id;

  // Step 2: Upload the image content to the fixed "imgpsh" content path
  const uploadResponse = await fetchWithRetry(
    `${AMS_BASE}/${amsObjectId}/content/imgpsh`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token.amsToken}`,
        "Content-Type": "application/octet-stream",
        "x-ms-client-version": AMS_CLIENT_VERSION,
      },
      body: imageData,
    },
  );

  if (!uploadResponse.ok) {
    if (uploadResponse.status === 401) {
      throw new ApiAuthError(
        `AMS image upload failed: ${uploadResponse.status} ${uploadResponse.statusText}`,
      );
    }
    const errorText = await uploadResponse.text();
    throw new Error(
      `Failed to upload AMS image: ${uploadResponse.status} ${uploadResponse.statusText} — ${errorText}`,
    );
  }

  return { amsObjectId };
}

/** Response from SharePoint file upload. */
export interface SharePointUploadResult {
  /** SharePoint item unique ID. */
  itemId: string;
  /** SharePoint site ID. */
  siteId: string;
  /** File name (may differ from input if conflict-renamed). */
  fileName: string;
  /** File extension without dot. */
  fileType: string;
  /** Direct SharePoint file URL. */
  fileUrl: string;
  /** WebDAV URL for the file. */
  webDavUrl: string;
  /** SharePoint site base URL. */
  siteBaseUrl: string;
  /** SharePoint personal site path segment (e.g. "/personal/user_domain_com/"). */
  personalPath: string;
  /** Sharing link URL (empty string if sharing scope is "none"). */
  shareUrl: string;
  /** Sharing link ID returned by the createLink API (empty string if sharing scope is "none"). */
  shareId: string;
  /** SharePoint drive item ID (used for the createLink API call). */
  driveItemId: string;
}

/**
 * Derive the SharePoint personal site URL components from a token and email.
 *
 * The SharePoint host comes from `token.sharePointHost` (captured from the
 * MSAL cache during authentication).
 * The personal path is derived from the email by replacing `.` and `@` with `_`.
 */
function deriveSharePointSiteInfo(
  token: TeamsToken,
  email: string,
): { siteBaseUrl: string; personalPath: string } {
  if (!token.sharePointHost) {
    throw new Error(
      "SharePoint host is required for file upload but was not captured during authentication. " +
        "Re-authenticate to capture the SharePoint host.",
    );
  }

  const siteBaseUrl = `https://${token.sharePointHost}`;

  // email "user.name@company.com" → "user_name_company_com"
  const personalSegment = email.replace(/[.@]/g, "_");
  const personalPath = `/personal/${personalSegment}`;

  return { siteBaseUrl, personalPath };
}

/** Response from the SharePoint createLink API. */
interface SharePointSharingLink {
  /** The sharing link URL that recipients can use to access the file. */
  shareUrl: string;
  /** The opaque sharing link ID. */
  shareId: string;
}

/**
 * Sharing options for the createLink API.
 *
 * - `{ scope: "organization" }` — anyone in the org with the link can edit.
 * - `{ scope: "users", emails: [...] }` — only the named recipients can edit.
 */
export type SharingLinkOptions =
  | { scope: "organization" }
  | { scope: "users"; emails: string[] };

/**
 * Create a sharing link for a SharePoint file.
 *
 * Without this step, files uploaded to the sender's OneDrive are private —
 * chat participants would see a "Request Access" prompt instead of the file.
 *
 * Uses the SharePoint REST API v2.0 `createLink` endpoint.
 *
 * @param options - Sharing scope: "organization" (anyone in the org) or
 *   "users" with specific recipient emails.
 */
export async function createSharePointSharingLink(
  token: TeamsToken,
  siteBaseUrl: string,
  personalPath: string,
  driveItemId: string,
  options: SharingLinkOptions = { scope: "organization" },
): Promise<SharePointSharingLink> {
  if (!token.sharePointToken) {
    throw new Error(
      "SharePoint token is required for creating sharing links but was not captured during authentication. " +
        "Re-authenticate to capture the SharePoint token.",
    );
  }

  const createLinkUrl = `${siteBaseUrl}${personalPath}/_api/v2.0/drive/items/${driveItemId}/createLink`;

  const requestBody: Record<string, unknown> = {
    type: "edit",
    scope: options.scope,
  };

  if (options.scope === "users") {
    requestBody.recipients = options.emails.map((email) => ({ email }));
  }

  const response = await fetchWithRetry(createLinkUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token.sharePointToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(requestBody),
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `SharePoint createLink authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    const errorText = await response.text();
    throw new Error(
      `Failed to create SharePoint sharing link: ${response.status} ${response.statusText} — ${errorText}`,
    );
  }

  const data = (await response.json()) as {
    shareId: string;
    link: {
      webUrl: string;
    };
  };

  return {
    shareUrl: data.link.webUrl,
    shareId: data.shareId,
  };
}

/**
 * Upload a file to SharePoint OneDrive for Business (Teams Chat Files folder).
 *
 * Files shared in Teams conversations are stored in the sender's OneDrive
 * under "Microsoft Teams Chat Files". This function uploads a file there
 * using the SharePoint REST API, matching the same flow the Teams web client uses.
 *
 * After uploading, a sharing link is created according to the provided
 * sharing options so that chat participants can access the file.
 *
 * @param email - The sender's corporate email (used to derive the personal site path)
 * @param sharingOptions - Controls who gets access. Pass `null` for no sharing link.
 */
export async function uploadSharePointFile(
  token: TeamsToken,
  fileData: Buffer,
  fileName: string,
  email: string,
  sharingOptions: SharingLinkOptions | null = { scope: "organization" },
): Promise<SharePointUploadResult> {
  if (!token.sharePointToken) {
    throw new Error(
      "SharePoint token is required for file upload but was not captured during authentication. " +
        "Re-authenticate to capture the SharePoint token.",
    );
  }

  const { siteBaseUrl, personalPath } = deriveSharePointSiteInfo(token, email);

  const encodedFileName = encodeURIComponent(fileName);
  const uploadUrl =
    `${siteBaseUrl}${personalPath}/_api/v2.0/drive/root:/Microsoft%20Teams%20Chat%20Files/${encodedFileName}:/content` +
    `?@name.conflictBehavior=rename&$select=*,sharepointIds,webDavUrl`;

  const response = await fetchWithRetry(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token.sharePointToken}`,
      "Content-Type": "application/octet-stream",
    },
    body: fileData,
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `SharePoint file upload authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    const errorText = await response.text();
    throw new Error(
      `Failed to upload file to SharePoint: ${response.status} ${response.statusText} — ${errorText}`,
    );
  }

  const data = (await response.json()) as {
    id: string;
    name: string;
    webDavUrl: string;
    webUrl: string;
    sharepointIds: {
      listItemUniqueId: string;
      siteId: string;
    };
  };

  const fileExtension = data.name.includes(".")
    ? (data.name.split(".").pop() ?? "")
    : "";

  let shareUrl = "";
  let shareId = "";

  if (sharingOptions !== null) {
    const sharingLink = await createSharePointSharingLink(
      token,
      siteBaseUrl,
      personalPath,
      data.id,
      sharingOptions,
    );
    shareUrl = sharingLink.shareUrl;
    shareId = sharingLink.shareId;
  }

  return {
    itemId: data.sharepointIds.listItemUniqueId,
    siteId: data.sharepointIds.siteId,
    fileName: data.name,
    fileType: fileExtension,
    fileUrl: data.webUrl,
    webDavUrl: data.webDavUrl,
    siteBaseUrl,
    personalPath,
    shareUrl,
    shareId,
    driveItemId: data.id,
  };
}

/**
 * Build the `properties.files` JSON string for a message with file attachments.
 *
 * This JSON is included in the message body when sending messages with
 * SharePoint-hosted file attachments.
 */
export function buildFilesPropertyJson(
  uploadResults: SharePointUploadResult[],
): string {
  const files = uploadResults.map((result) => ({
    "@type": "http://schema.skype.com/File",
    version: 2,
    id: result.itemId,
    itemid: result.itemId,
    fileName: result.fileName,
    fileType: result.fileType,
    title: result.fileName,
    type: result.fileType,
    state: "active",
    objectUrl: `${result.siteBaseUrl}${result.personalPath}/Documents/Microsoft%20Teams%20Chat%20Files/${encodeURIComponent(result.fileName)}`,
    baseUrl: `${result.siteBaseUrl}${result.personalPath}/`,
    permissionScope: "users",
    sharepointIds: {
      listItemUniqueId: result.itemId,
      siteId: result.siteId,
    },
    fileInfo: {
      itemId: null,
      fileUrl: `${result.siteBaseUrl}${result.personalPath}/Documents/Microsoft%20Teams%20Chat%20Files/${encodeURIComponent(result.fileName)}`,
      siteUrl: `${result.siteBaseUrl}${result.personalPath}/`,
      serverRelativeUrl: `${result.personalPath}/Documents/Microsoft Teams Chat Files/${result.fileName}`,
      shareUrl: result.shareUrl,
      shareId: result.shareId,
    },
    fileChicletState: {
      serviceName: "p2p",
      state: "active",
    },
  }));

  return JSON.stringify(files);
}

/**
 * Build an HTML `<img>` tag for an AMS-hosted image.
 *
 * This produces the same markup that the Teams web client generates.
 */
export function buildAmsImageTag(
  amsObjectId: string,
  width?: number,
  height?: number,
): string {
  const src = `${AMS_BASE}/${amsObjectId}/views/imgo`;
  const styleAttribute =
    width && height ? ` style="width:${width}px; height:${height}px"` : "";
  return `<img src="${src}" itemscope="" itemtype="http://schema.skype.com/AMSImage"${styleAttribute}>`;
}

/**
 * Download a file attachment from SharePoint.
 *
 * File attachments in Teams are hosted on SharePoint (OneDrive for Business).
 * This function downloads the file content using the SharePoint REST API
 * and the captured SharePoint bearer token.
 *
 * @param fileUrl - Direct SharePoint URL from `FileAttachment.fileUrl`
 */
export async function fetchSharePointFile(
  token: TeamsToken,
  fileUrl: string,
  itemId: string,
): Promise<{
  data: Buffer;
  contentType: string;
  size: number;
  fileName: string;
}> {
  if (!token.sharePointToken) {
    throw new Error(
      "SharePoint token is required for file download but was not captured during authentication. " +
        "Re-authenticate to capture the SharePoint token.",
    );
  }

  // Extract the SharePoint site base URL and the personal site path from fileUrl.
  // fileUrl looks like: https://tenant-my.sharepoint.com/personal/user_name_company_com/Documents/...
  const parsedUrl = new URL(fileUrl);
  const siteBaseUrl = `${parsedUrl.protocol}//${parsedUrl.host}`;
  const pathSegments = decodeURIComponent(parsedUrl.pathname).split("/");
  // pathSegments: ['', 'personal', 'user_name_company_com', 'Documents', ...]
  const personalPath = `/${pathSegments[1]}/${pathSegments[2]}`;

  // Use the SharePoint drive items API with the item's unique ID.
  // This avoids path-encoding issues with special characters in filenames,
  // and works for files on any user's personal OneDrive.
  const downloadUrl = `${siteBaseUrl}${personalPath}/_api/v2.0/drive/items/${itemId}/content`;

  const response = await fetchWithRetry(downloadUrl, {
    headers: {
      Authorization: `Bearer ${token.sharePointToken}`,
    },
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `SharePoint file download authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    throw new Error(
      `Failed to download SharePoint file: ${response.status} ${response.statusText}`,
    );
  }

  const contentType =
    response.headers.get("content-type") ?? "application/octet-stream";
  const arrayBuffer = await response.arrayBuffer();
  const data = Buffer.from(arrayBuffer);

  // Extract filename from the original file URL path
  const fileName = pathSegments[pathSegments.length - 1] ?? "unknown";

  return { data, contentType, size: data.length, fileName };
}
