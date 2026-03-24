/**
 * Transcript fetching and VTT parsing for Teams meetings.
 *
 * Handles extracting transcript URLs from recording messages,
 * fetching VTT content from the AMS (Async Media Service),
 * and parsing VTT into structured entries.
 */

import type {
  TeamsToken,
  TranscriptEntry,
  TranscriptResult,
} from "../types.js";
import { fetchWithRetry, ApiAuthError } from "./common.js";
import { fetchMessagesPage } from "./chat-service.js";
import { decodeHtmlEntities } from "../html-utils.js";

/**
 * Extract the AMS transcript URL from a `RichText/Media_CallRecording` message.
 *
 * The message content is XML containing `<item type="amsTranscript" uri="...">`.
 * Returns null if no transcript URL is found.
 */
export function extractTranscriptUrl(messageContent: string): string | null {
  const match = messageContent.match(
    /<item\s+[^>]*type="amsTranscript"[^>]*\buri="([^"]+)"/,
  );
  return match?.[1] ?? null;
}

/**
 * Extract the meeting title from a `RichText/Media_CallRecording` message.
 *
 * The title is in the `<OriginalName>` element inside the XML content.
 */
export function extractMeetingTitle(messageContent: string): string {
  const match = messageContent.match(/<OriginalName\b[^>]*v="([^"]*)"[^>]*\/>/);
  return match?.[1] ?? "Unknown Meeting";
}

/**
 * Check whether a `RichText/Media_CallRecording` message represents
 * a successful recording (as opposed to started/failed).
 */
export function isSuccessfulRecording(messageContent: string): boolean {
  return /<RecordingStatus\b[^>]*status="Success"/.test(messageContent);
}

/**
 * Parse VTT content into structured transcript entries.
 *
 * Handles the Teams VTT format with `<v Speaker Name>text</v>` tags
 * and HTML entities.
 */
export function parseVtt(vttContent: string): TranscriptEntry[] {
  const entries: TranscriptEntry[] = [];
  const lines = vttContent.split("\n");

  let currentStartTime = "";
  let currentEndTime = "";

  for (const line of lines) {
    // Match timestamp lines: "00:00:00.000 --> 00:00:05.000"
    const timestampMatch = line.match(
      /^(\d{2}:\d{2}:\d{2}\.\d{3})\s+-->\s+(\d{2}:\d{2}:\d{2}\.\d{3})/,
    );
    if (timestampMatch) {
      currentStartTime = timestampMatch[1];
      currentEndTime = timestampMatch[2];
      continue;
    }

    // Match speaker lines: "<v Speaker Name>text</v>"
    const speakerMatch = line.match(/^<v\s+([^>]+)>(.+)<\/v>\s*$/);
    if (speakerMatch && currentStartTime) {
      entries.push({
        speaker: decodeHtmlEntities(speakerMatch[1]),
        startTime: currentStartTime,
        endTime: currentEndTime,
        text: decodeHtmlEntities(speakerMatch[2]),
      });
      currentStartTime = "";
      currentEndTime = "";
    }
  }

  return entries;
}

/**
 * Fetch a transcript from the AMS (Async Media Service).
 *
 * Uses `Authorization: skype_token <token>` header (note: different format
 * from the Chat Service `Authentication: skypetoken=<token>` header).
 */
export async function fetchTranscriptVtt(
  token: TeamsToken,
  amsTranscriptUrl: string,
): Promise<string> {
  const response = await fetchWithRetry(amsTranscriptUrl, {
    headers: {
      Authorization: `skype_token ${token.skypeToken}`,
    },
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `Transcript fetch authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    throw new Error(
      `Failed to fetch transcript: ${response.status} ${response.statusText}`,
    );
  }

  return response.text();
}

/**
 * Fetch and parse the transcript for a conversation.
 *
 * 1. Fetches messages from the conversation
 * 2. Finds the latest `RichText/Media_CallRecording` message with status="Success"
 * 3. Extracts the AMS transcript URL from the XML content
 * 4. Fetches the VTT content from AMS
 * 5. Parses the VTT into structured entries
 */
export async function fetchTranscript(
  token: TeamsToken,
  conversationId: string,
): Promise<TranscriptResult> {
  // Fetch messages to find the recording message
  const pageSize = 200;
  const maxPages = 20;
  let backwardLink: string | undefined;
  let amsTranscriptUrl: string | null = null;
  let meetingTitle = "Unknown Meeting";

  for (let pageIndex = 0; pageIndex < maxPages; pageIndex++) {
    const page = await fetchMessagesPage(
      token,
      conversationId,
      pageSize,
      backwardLink,
    );

    for (const message of page.messages) {
      if (
        message.messageType === "RichText/Media_CallRecording" &&
        isSuccessfulRecording(message.content)
      ) {
        const transcriptUrl = extractTranscriptUrl(message.content);
        if (transcriptUrl) {
          amsTranscriptUrl = transcriptUrl;
          meetingTitle = extractMeetingTitle(message.content);
          break;
        }
      }
    }

    if (amsTranscriptUrl || !page.backwardLink) break;
    backwardLink = page.backwardLink;
  }

  if (!amsTranscriptUrl) {
    throw new Error(
      "No meeting transcript found in this conversation. " +
        "Make sure the conversation contains a recorded meeting with a transcript.",
    );
  }

  const rawVtt = await fetchTranscriptVtt(token, amsTranscriptUrl);
  const entries = parseVtt(rawVtt);

  return { meetingTitle, rawVtt, entries };
}
