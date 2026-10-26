import { describe, expect, it, vi } from "vitest";
import type { Page } from "playwright";
import { captureTokensFromPage } from "../../src/auth/token-capture.js";

interface FetchRequestPausedEvent {
  requestId: string;
  request: {
    url: string;
    headers: Record<string, string>;
  };
}

function createPageStub(
  events: FetchRequestPausedEvent[],
  searchEvents: FetchRequestPausedEvent[] = [],
) {
  let requestPausedHandler:
    | ((event: FetchRequestPausedEvent) => Promise<void>)
    | null = null;

  const devToolsSession = {
    send: vi.fn().mockResolvedValue(undefined),
    on: vi.fn(
      (
        eventName: string,
        handler: (event: FetchRequestPausedEvent) => Promise<void>,
      ) => {
        if (eventName === "Fetch.requestPaused") {
          requestPausedHandler = handler;
        }
      },
    ),
    detach: vi.fn().mockResolvedValue(undefined),
  };

  const searchInput = {
    fill: vi.fn().mockResolvedValue(undefined),
    press: vi.fn(async () => {
      for (const event of searchEvents) {
        await requestPausedHandler?.(event);
      }
    }),
  };
  const searchRegion = {
    getByRole: vi.fn(() => ({ first: () => searchInput })),
  };
  const page = {
    context: () => ({
      newCDPSession: vi.fn().mockResolvedValue(devToolsSession),
    }),
    reload: vi.fn(async () => {
      if (!requestPausedHandler) {
        throw new Error("Fetch.requestPaused handler was not registered");
      }
      for (const event of events) {
        await requestPausedHandler(event);
      }
    }),
    evaluate: vi.fn().mockResolvedValue({}),
    getByRole: vi.fn((role: string) => {
      if (role !== "search") throw new Error("Expected search landmark");
      return searchRegion;
    }),
    url: () => "https://teams.cloud.microsoft/v2/",
  } as unknown as Page;

  return { page, devToolsSession, searchInput, searchRegion };
}

describe("captureTokensFromPage", () => {
  it("captures a missing Substrate token without relying on localized search labels", async () => {
    vi.useFakeTimers();
    try {
      const { page, searchInput, searchRegion } = createPageStub(
        [
          {
            requestId: "chat",
            request: {
              url: "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations",
              headers: { "x-skypetoken": "chat-token" },
            },
          },
        ],
        [
          {
            requestId: "search",
            request: {
              url: "https://substrate.office.com/searchservice/api/v2/query",
              headers: { authorization: "Bearer search-token" },
            },
          },
        ],
      );
      const pending = captureTokensFromPage(page, vi.fn(), 30 * 1_000);
      await vi.advanceTimersByTimeAsync(1_000);
      expect((await pending).substrateToken).toBe("search-token");
      expect(page.getByRole).toHaveBeenCalledWith("search");
      expect(searchRegion.getByRole).toHaveBeenCalledWith("combobox");
      expect(searchInput.fill).toHaveBeenCalledWith(
        '"teams-api-authentication-token-capture"',
        { timeout: 10 * 1_000 },
      );
      expect(searchInput.press).toHaveBeenCalledWith("Enter", {
        timeout: 5 * 1_000,
      });
    } finally {
      vi.useRealTimers();
    }
  });

  it("keeps chat authentication usable when search UI is unavailable", async () => {
    vi.useFakeTimers();
    try {
      const { page, searchInput } = createPageStub([
        {
          requestId: "chat",
          request: {
            url: "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations",
            headers: { "x-skypetoken": "chat-token" },
          },
        },
      ]);
      const searchError = new Error("Sensitive diagnostic details");
      searchError.name = "TimeoutError";
      searchInput.fill.mockRejectedValue(searchError);
      const log = vi.fn();
      const pending = captureTokensFromPage(page, log, 30 * 1_000);
      await vi.advanceTimersByTimeAsync(6 * 1_000);
      expect((await pending).skypeToken).toBe("chat-token");
      expect(searchInput.press).not.toHaveBeenCalled();
      expect(log).toHaveBeenCalledWith(
        "Search UI unavailable (TimeoutError); continuing with captured chat credentials",
      );
      expect(log.mock.calls.flat().join("\n")).not.toContain(
        searchError.message,
      );
    } finally {
      vi.useRealTimers();
    }
  });

  it("prefers a Chat Service Skype token over earlier Teams web requests", async () => {
    const { page, searchInput } = createPageStub([
      {
        requestId: "web-request",
        request: {
          url: "https://teams.cloud.microsoft/api/mt/amer/beta/users/fetchShortProfile",
          headers: {
            "x-skypetoken": "web-token",
            authorization: "Bearer bearer-token",
          },
        },
      },
      {
        requestId: "substrate-request",
        request: {
          url: "https://substrate.office.com/search/api/v1/suggestions?scenario=peoplepicker.newChat",
          headers: {
            authorization: "Bearer substrate-token",
          },
        },
      },
      {
        requestId: "chat-service-request",
        request: {
          url: "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations",
          headers: {
            "x-skypetoken": "skypetoken=chat-service-token",
          },
        },
      },
    ]);

    const token = await captureTokensFromPage(page, vi.fn(), 30_000);

    expect(token.skypeToken).toBe("chat-service-token");
    expect(token.region).toBe("emea");
    expect(token.bearerToken).toBe("bearer-token");
    expect(token.substrateToken).toBe("substrate-token");
    expect(searchInput.fill).not.toHaveBeenCalled();
  });

  it("falls back to the Teams web chat Bearer token when no Skype token is issued", async () => {
    vi.useFakeTimers();
    try {
      const { page } = createPageStub([
        {
          requestId: "web-chat-request",
          request: {
            url: "https://teams.cloud.microsoft/api/chatsvc/apac/v1/users/ME/conversations?view=mychats&pageSize=1",
            headers: { authorization: "Bearer chat-bearer-token" },
          },
        },
      ]);

      const pending = captureTokensFromPage(page, vi.fn(), 90_000);
      await vi.advanceTimersByTimeAsync(15_000);
      const token = await pending;

      expect(token.skypeToken).toBe("");
      expect(token.chatBearerToken).toBe("chat-bearer-token");
      expect(token.region).toBe("apac");
    } finally {
      vi.useRealTimers();
    }
  });

  it.each(["before", "after"])(
    "prefers a non-chat Skype token arriving %s the chat Bearer token",
    async (arrivalOrder) => {
      vi.useFakeTimers();
      try {
        const skypeRequest = {
          requestId: "profile-request",
          request: {
            url: "https://teams.cloud.microsoft/api/mt/apac/beta/users/fetchShortProfile",
            headers: { "x-skypetoken": "skypetoken=fallback-token" },
          },
        };
        const bearerRequest = {
          requestId: "chat-request",
          request: {
            url: "https://teams.cloud.microsoft/api/chatsvc/apac/v1/users/ME/conversations",
            headers: { authorization: "Bearer chat-bearer-token" },
          },
        };
        const requests =
          arrivalOrder === "before"
            ? [skypeRequest, bearerRequest]
            : [bearerRequest, skypeRequest];
        const { page } = createPageStub([
          ...requests,
          {
            requestId: "search-request",
            request: {
              url: "https://substrate.office.com/search/api/v1/suggestions",
              headers: { authorization: "Bearer substrate-token" },
            },
          },
        ]);
        const pending = captureTokensFromPage(page, vi.fn(), 90 * 1_000);

        await vi.advanceTimersByTimeAsync(5 * 1_000);
        const token = await pending;

        expect(token.skypeToken).toBe("fallback-token");
        expect(token.chatBearerToken).toBeUndefined();
      } finally {
        vi.useRealTimers();
      }
    },
  );

  it("prefers the Skype token when both a chat Bearer token and a Skype token are seen", async () => {
    const { page } = createPageStub([
      {
        requestId: "web-chat-request",
        request: {
          url: "https://teams.cloud.microsoft/api/chatsvc/emea/v1/users/ME/properties",
          headers: { authorization: "Bearer chat-bearer-token" },
        },
      },
      {
        requestId: "chat-service-request",
        request: {
          url: "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations",
          headers: { "x-skypetoken": "skypetoken=chat-service-token" },
        },
      },
    ]);

    const token = await captureTokensFromPage(page, vi.fn(), 30_000);

    expect(token.skypeToken).toBe("chat-service-token");
    expect(token.chatBearerToken).toBeUndefined();
  });
});
