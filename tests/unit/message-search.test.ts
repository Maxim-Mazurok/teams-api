// cspell:words unstub
import { afterEach, describe, expect, it, vi } from "vitest";
import { searchMessages } from "../../src/api/message-search.js";
import { ApiAuthError } from "../../src/api/common.js";
import type { SearchMessagesOptions } from "../../src/types.js";

const token = {
  skypeToken: "chat-token",
  substrateToken: "search-token",
  region: "apac",
};

afterEach(() => vi.unstubAllGlobals());

describe("searchMessages", () => {
  it("sends a Teams-only search and parses results with pagination", async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          EntitySets: [
            {
              EntityType: "Message",
              ResultSets: [
                {
                  Total: 42,
                  MoreResultsAvailable: true,
                  Results: [
                    {
                      Id: "result-1",
                      Source: {
                        InternetMessageId: "message-1",
                        ClientThreadId: "19:channel@thread.tacv2",
                        DateTimeSent: "2026-10-01T09:00:00Z",
                        Subject: "Release",
                        Preview: "Release ready",
                        From: {
                          EmailAddress: {
                            Name: "Alice Smith",
                            Address: "alice@contoso.com",
                          },
                        },
                        Links: [
                          {
                            Url: "https://contoso.com/release",
                            LinkType: "Web",
                          },
                        ],
                        Extensions: {
                          SkypeSpaces_ConversationPost_Extension_ParentMessageId:
                            "parent-1",
                          SkypeSpaces_ConversationPost_Extension_ThreadType:
                            "channel",
                        },
                      },
                    },
                  ],
                },
              ],
            },
          ],
        }),
      ),
    );
    vi.stubGlobal("fetch", fetchMock);
    const page = await searchMessages(token, {
      query: "release",
      offset: 5,
      pageSize: 10,
    });
    expect(page).toMatchObject({
      total: 42,
      hasMore: true,
      nextOffset: 15,
      messages: [
        {
          messageId: "message-1",
          senderEmail: "alice@contoso.com",
          preview: "Release ready",
          parentMessageId: "parent-1",
          threadType: "channel",
          webUrl:
            "https://teams.microsoft.com/l/message/19%3Achannel%40thread.tacv2/message-1?parentMessageId=parent-1",
        },
      ],
    });
    const [url, request] = fetchMock.mock.calls[0];
    expect(url).toBe("https://substrate.office.com/searchservice/api/v2/query");
    expect(request.headers.Authorization).toBe("Bearer search-token");
    expect(JSON.parse(request.body).EntityRequests).toEqual([
      expect.objectContaining({
        entityType: "Message",
        contentSources: ["Teams"],
        from: 5,
        size: 10,
        query: {
          queryString: "NOT (isClientSoftDeleted:TRUE) AND (release)",
          displayQueryString: "release",
        },
      }),
    ]);
  });

  it("distinguishes missing metadata from explicitly empty fields", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        new Response(
          JSON.stringify({
            EntitySets: [
              {
                EntityType: "Message",
                ResultSets: [
                  {
                    Total: 2,
                    Results: [
                      {
                        Id: "missing-metadata",
                        Source: {
                          Extensions: {
                            SkypeSpaces_ConversationPost_Extension_ParentMessageId:
                              null,
                          },
                        },
                      },
                      {
                        Id: "empty-metadata",
                        Source: {
                          Subject: "",
                          Preview: "",
                          From: { EmailAddress: { Address: "" } },
                        },
                      },
                    ],
                  },
                ],
              },
            ],
          }),
        ),
      ),
    );
    const page = await searchMessages(token, { query: "release" });
    expect(page.messages[0].senderEmail).toBeUndefined();
    expect(page.messages[0].subject).toBeUndefined();
    expect(page.messages[0].preview).toBeUndefined();
    expect(page.messages[0].messageId).toBeUndefined();
    expect(page.messages[0].parentMessageId).toBeUndefined();
    expect(page.messages[0].webUrl).toBeUndefined();
    expect(JSON.parse(JSON.stringify(page.messages[0]))).toEqual({
      id: "missing-metadata",
    });
    expect(page.messages[1]).toMatchObject({
      senderEmail: "",
      subject: "",
      preview: "",
    });
  });

  it("sums approximate result-set totals without deduplicating results", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        new Response(
          JSON.stringify({
            EntitySets: [
              {
                EntityType: "Message",
                ResultSets: [
                  { Total: 10, Results: [{ Id: "shared-result" }] },
                  {
                    Total: 20,
                    Results: [{ Id: "shared-result" }],
                    MoreResultsAvailable: true,
                  },
                  { Results: [] },
                ],
              },
            ],
          }),
        ),
      ),
    );
    const page = await searchMessages(token, { query: "release" });
    expect(page.total).toBe(30);
    expect(page.messages.map((message) => message.id)).toEqual([
      "shared-result",
      "shared-result",
    ]);
    expect(page.hasMore).toBe(true);
    expect(page.nextOffset).toBe(25);
  });

  it("rejects missing search credentials without making a request", async () => {
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);
    await expect(
      searchMessages(
        { ...token, substrateToken: undefined },
        { query: "release" },
      ),
    ).rejects.toBeInstanceOf(ApiAuthError);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it.each(["", "Bad Request"])(
    "reports a stable HTTP error regardless of status text: %j",
    async (statusText) => {
      vi.stubGlobal(
        "fetch",
        vi
          .fn()
          .mockResolvedValue(
            new Response("bad request", { status: 400, statusText }),
          ),
      );
      await expect(searchMessages(token, { query: "release" })).rejects.toThrow(
        new Error("Message search failed: 400"),
      );
    },
  );

  it("combines sender, channel, conversation, and inclusive date filters", async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          EntitySets: [
            {
              EntityType: "Message",
              ResultSets: [{ Total: 0, Results: [] }],
            },
          ],
        }),
      ),
    );
    vi.stubGlobal("fetch", fetchMock);
    const page = await searchMessages(token, {
      query: "*",
      from: "alice@contoso.com",
      messageType: "channel",
      conversationId: "19:channel@thread.tacv2",
      sentAfter: "2026-10-01",
      sentBefore: "2026-10-26",
    });
    expect(page).toEqual({ messages: [], total: 0, hasMore: false });
    expect(
      JSON.parse(fetchMock.mock.calls[0][1].body).EntityRequests[0].query
        .queryString,
    ).toBe(
      'NOT (isClientSoftDeleted:TRUE) AND (*) AND from:"alice@contoso.com" AND Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String:(topic OR space) AND ClientThreadId:"19:channel@thread.tacv2" AND sent>=2026-10-01 AND sent<=2026-10-26',
    );
  });

  it.each([
    { query: " " },
    { query: "release", offset: -1 },
    { query: "release", offset: 1.5 },
    { query: "release", pageSize: 0 },
    { query: "release", pageSize: 101 },
    { query: "release", pageSize: NaN },
    { query: "release", from: " " },
    { query: "release", conversationId: " " },
    { query: "release", sentAfter: "2026-02-30" },
    { query: "release", sentBefore: "yesterday" },
    { query: "release", sentAfter: "2026-10-26", sentBefore: "2026-10-01" },
  ])("rejects invalid options without fetching: %j", async (options) => {
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);
    await expect(searchMessages(token, options)).rejects.toThrow();
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it.each(["", "channels", "CHANNEL", " "])(
    "rejects invalid message scope without fetching: %j",
    async (messageType) => {
      const fetchMock = vi.fn();
      vi.stubGlobal("fetch", fetchMock);
      await expect(
        searchMessages(token, {
          query: "release",
          messageType: messageType as SearchMessagesOptions["messageType"],
        }),
      ).rejects.toThrow("Search message type must be all, chat, or channel");
      expect(fetchMock).not.toHaveBeenCalled();
    },
  );

  it.each([401, 403])(
    "throws an authentication error for HTTP %s",
    async (status) => {
      vi.stubGlobal(
        "fetch",
        vi.fn().mockResolvedValue(new Response("unauthorized", { status })),
      );
      await expect(
        searchMessages(token, { query: "release" }),
      ).rejects.toBeInstanceOf(ApiAuthError);
    },
  );

  it.each([
    {},
    {
      EntitySets: [
        {
          EntityType: "Message",
          Properties: { HasParseException: true },
          ResultSets: [],
        },
      ],
    },
    {
      EntitySets: [{ EntityType: "Message", IsPartial: true, ResultSets: [] }],
    },
  ])("rejects missing, invalid, or incomplete results: %j", async (body) => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(new Response(JSON.stringify(body))),
    );
    await expect(searchMessages(token, { query: "release" })).rejects.toThrow();
  });
});
