// cspell:words cvid powerbar
import { randomUUID } from "node:crypto";
import type {
  TeamsToken,
  SearchMessagesOptions,
  MessageSearchPage,
} from "../types.js";
import { ApiAuthError, fetchWithRetry } from "./common.js";

const SEARCH_EXTENSION_PREFIX = "SkypeSpaces_ConversationPost_Extension_";
const SEARCH_EXTENSION_FIELDS = [
  "FromSkypeInternalId",
  "ThreadType",
  "ParentMessageId",
].map((field) => `Extension_${SEARCH_EXTENSION_PREFIX}${field}_String`);

interface SearchResponse {
  EntitySets?: Array<{
    EntityType?: string;
    Properties?: { HasParseException?: boolean };
    IsPartial?: boolean;
    ResultSets?: Array<{
      Total?: number;
      MoreResultsAvailable?: boolean;
      Results?: Array<{
        Id?: string;
        Source?: {
          InternetMessageId?: string;
          ClientThreadId?: string;
          DateTimeSent?: string;
          Subject?: string;
          Preview?: string;
          ParentFolderDisplayName?: string;
          From?: { EmailAddress?: { Name?: string; Address?: string } };
          Sender?: { EmailAddress?: { Name?: string; Address?: string } };
          Extensions?: Record<string, string | null>;
        };
      }>;
    }>;
  }>;
}

function quoteSearchValue(value: string): string {
  return `"${value.replaceAll("\\", "\\\\").replaceAll('"', '\\"')}"`;
}

function buildSearchQuery(options: SearchMessagesOptions): string {
  const query = options.query.trim();
  if (!query) throw new Error("Search query must not be empty");
  const filters = ["NOT (isClientSoftDeleted:TRUE)", `(${query})`];
  if (options.from !== undefined) {
    if (!options.from.trim())
      throw new Error("Search sender must not be empty");
    filters.push(`from:${quoteSearchValue(options.from.trim())}`);
  }
  if (
    options.messageType !== undefined &&
    !["all", "chat", "channel"].includes(options.messageType)
  ) {
    throw new Error("Search message type must be all, chat, or channel");
  }
  if (options.messageType === "chat" || options.messageType === "channel") {
    const threadTypes =
      options.messageType === "channel" ? "topic OR space" : "chat";
    filters.push(
      `Extension_${SEARCH_EXTENSION_PREFIX}ThreadType_String:(${threadTypes})`,
    );
  }
  if (options.conversationId !== undefined) {
    if (!options.conversationId.trim())
      throw new Error("Search conversation ID must not be empty");
    filters.push(
      `ClientThreadId:${quoteSearchValue(options.conversationId.trim())}`,
    );
  }
  for (const [name, value, operator] of [
    ["sentAfter", options.sentAfter, ">="],
    ["sentBefore", options.sentBefore, "<="],
  ] as const) {
    if (value === undefined) continue;
    if (
      !/^\d{4}-\d{2}-\d{2}$/.test(value) ||
      !Number.isFinite(Date.parse(value)) ||
      new Date(value).toISOString().slice(0, 10) !== value
    ) {
      throw new Error(`${name} must be a valid date in YYYY-MM-DD format`);
    }
    filters.push(`sent${operator}${value}`);
  }
  if (
    options.sentAfter &&
    options.sentBefore &&
    options.sentAfter > options.sentBefore
  ) {
    throw new Error("sentAfter must not be later than sentBefore");
  }
  return filters.join(" AND ");
}

export async function searchMessages(
  token: TeamsToken,
  options: SearchMessagesOptions,
): Promise<MessageSearchPage> {
  if (!token.substrateToken) {
    throw new ApiAuthError(
      "Substrate token is missing; re-authentication required for message search",
    );
  }
  const queryString = buildSearchQuery(options);
  const offset = options.offset ?? 0;
  const pageSize = options.pageSize ?? 25;
  if (!Number.isInteger(offset) || offset < 0) {
    throw new Error("Search offset must be a non-negative integer");
  }
  if (!Number.isInteger(pageSize) || pageSize < 1 || pageSize > 100) {
    throw new Error("Search page size must be an integer between 1 and 100");
  }
  const response = await fetchWithRetry(
    "https://substrate.office.com/searchservice/api/v2/query",
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token.substrateToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        cvid: randomUUID(),
        scenario: { Name: "powerbar" },
        EntityRequests: [
          {
            entityType: "Message",
            contentSources: ["Teams"],
            fields: SEARCH_EXTENSION_FIELDS,
            propertySet: "Optimized",
            query: {
              queryString,
              displayQueryString: options.query.trim(),
            },
            from: offset,
            size: pageSize,
          },
        ],
      }),
    },
  );
  if (response.status === 401 || response.status === 403) {
    throw new ApiAuthError(
      `Message search authentication failed: ${response.status}`,
    );
  }
  if (!response.ok) {
    throw new Error(`Message search failed: ${response.status}`);
  }
  const data = (await response.json()) as SearchResponse;
  const entitySet = data.EntitySets?.find(
    (entitySet) => entitySet.EntityType === "Message",
  );
  if (!entitySet?.ResultSets) {
    throw new Error("Message search response is missing message result sets");
  }
  if (entitySet.Properties?.HasParseException) {
    throw new Error("Message search query could not be parsed");
  }
  if (entitySet.IsPartial) {
    throw new Error(
      "Message search returned incomplete results; retry the search",
    );
  }
  const resultSets = entitySet.ResultSets;
  const messages = resultSets.flatMap((resultSet) =>
    (resultSet.Results ?? []).map((result) => {
      const source = result.Source;
      const sender = source?.From?.EmailAddress ?? source?.Sender?.EmailAddress;
      const messageId = source?.InternetMessageId;
      const conversationId = source?.ClientThreadId;
      const parentMessageId =
        source?.Extensions?.[`${SEARCH_EXTENSION_PREFIX}ParentMessageId`] ??
        undefined;
      const webUrl =
        messageId && conversationId
          ? `https://teams.microsoft.com/l/message/${encodeURIComponent(conversationId)}/${encodeURIComponent(messageId)}${parentMessageId ? `?parentMessageId=${encodeURIComponent(parentMessageId)}` : ""}`
          : undefined;
      return {
        id: result.Id,
        messageId,
        conversationId,
        parentMessageId,
        threadType:
          source?.Extensions?.[`${SEARCH_EXTENSION_PREFIX}ThreadType`] ??
          undefined,
        senderMri:
          source?.Extensions?.[
            `${SEARCH_EXTENSION_PREFIX}FromSkypeInternalId`
          ] ?? undefined,
        senderDisplayName: sender?.Name,
        senderEmail: sender?.Address,
        sentAt: source?.DateTimeSent,
        subject: source?.Subject,
        preview: source?.Preview,
        conversationName: source?.ParentFolderDisplayName,
        webUrl,
      };
    }),
  );
  const hasMore = resultSets.some(
    (resultSet) => resultSet.MoreResultsAvailable === true,
  );
  return {
    messages,
    total: resultSets.reduce(
      (total, resultSet) => total + (resultSet.Total ?? 0),
      0,
    ),
    hasMore,
    ...(hasMore ? { nextOffset: offset + pageSize } : {}),
  };
}
