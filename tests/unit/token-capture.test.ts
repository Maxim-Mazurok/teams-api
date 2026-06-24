import { describe, expect, it, vi } from "vitest";
import { captureTokensFromPage } from "../../src/auth/token-capture.js";

interface FetchRequestPausedEvent {
  requestId: string;
  request: {
    url: string;
    headers: Record<string, string>;
  };
}

function createPageStub(events: FetchRequestPausedEvent[]) {
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
    url: () => "https://teams.cloud.microsoft/v2/",
  };

  return { page, devToolsSession };
}

describe("captureTokensFromPage", () => {
  it("prefers a Chat Service Skype token over earlier Teams web requests", async () => {
    const { page } = createPageStub([
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
  });
});
