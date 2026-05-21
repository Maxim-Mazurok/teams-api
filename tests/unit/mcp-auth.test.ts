import { describe, it, expect, vi, beforeEach } from "vitest";
import {
  AuthenticationInProgressError,
  McpAuthManager,
  NeedsEmailError,
  type McpAuthManagerOptions,
} from "../../src/mcp-auth.js";
import type {
  AutoLoginOptions,
  InteractiveLoginOptions,
  ManualTokenOptions,
  SmartLoginOptions,
} from "../../src/types.js";

class FakeTeamsClient {
  public email: string | null = null;

  setEmail(email: string): void {
    this.email = email;
  }
}

function createPendingPromise<Value>(): {
  promise: Promise<Value>;
  resolve: (value: Value) => void;
  reject: (reason?: unknown) => void;
} {
  let resolvePromise: (value: Value) => void = () => {};
  let rejectPromise: (reason?: unknown) => void = () => {};
  const promise = new Promise<Value>((resolve, reject) => {
    resolvePromise = resolve;
    rejectPromise = reject;
  });

  return {
    promise,
    resolve: resolvePromise,
    reject: rejectPromise,
  };
}

function createFactory(client = new FakeTeamsClient()) {
  return {
    create: vi
      .fn<(options: AutoLoginOptions) => Promise<FakeTeamsClient>>()
      .mockResolvedValue(client),
    connect: vi
      .fn<(options?: SmartLoginOptions) => Promise<FakeTeamsClient>>()
      .mockResolvedValue(client),
    fromDebugSession: vi
      .fn<(options?: ManualTokenOptions) => Promise<FakeTeamsClient>>()
      .mockResolvedValue(client),
    fromInteractiveLogin: vi
      .fn<(options?: InteractiveLoginOptions) => Promise<FakeTeamsClient>>()
      .mockResolvedValue(client),
    fromToken: vi
      .fn<
        (
          skypeToken: string,
          region: string,
          bearerToken?: string,
          substrateToken?: string,
        ) => FakeTeamsClient
      >()
      .mockReturnValue(client),
  };
}

function createManager(
  options: Omit<
    McpAuthManagerOptions<FakeTeamsClient>,
    "clientFactory" | "recordAuth"
  > & {
    clientFactory?: McpAuthManagerOptions<FakeTeamsClient>["clientFactory"];
  } = {},
) {
  const recordAuth = vi.fn();
  const manager = new McpAuthManager<FakeTeamsClient>({
    ...options,
    clientFactory: options.clientFactory ?? createFactory(),
    recordAuth,
  });

  return { manager, recordAuth };
}

beforeEach(() => {
  vi.resetAllMocks();
});

describe("McpAuthManager", () => {
  it("creates a token client immediately when TEAMS_TOKEN is configured", async () => {
    const factory = createFactory();
    const { manager, recordAuth } = createManager({
      clientFactory: factory,
      environment: {
        TEAMS_BEARER_TOKEN: "bearer-token",
        TEAMS_EMAIL: "user@contoso.com",
        TEAMS_REGION: "apac",
        TEAMS_SUBSTRATE_TOKEN: "substrate-token",
        TEAMS_TOKEN: "skype-token",
      },
    });

    const client = await manager.getClient();

    expect(client.email).toBe("user@contoso.com");
    expect(factory.fromToken).toHaveBeenCalledWith(
      "skype-token",
      "apac",
      "bearer-token",
      "substrate-token",
    );
    expect(recordAuth).toHaveBeenCalledWith({
      strategy: "token",
      success: true,
    });
  });

  it("requires email before TEAMS_AUTO can authenticate", async () => {
    const factory = createFactory();
    const { manager } = createManager({
      clientFactory: factory,
      environment: { TEAMS_AUTO: "true" },
    });

    await expect(manager.getClient()).rejects.toBeInstanceOf(NeedsEmailError);
    expect(factory.create).not.toHaveBeenCalled();
  });

  it("runs eager startup auth when an email is configured", async () => {
    const factory = createFactory();
    const log = vi.fn();
    const { manager } = createManager({
      clientFactory: factory,
      environment: {
        TEAMS_EMAIL: "user@contoso.com",
      },
      log,
    });

    await manager.authenticateOnStartup();

    expect(factory.connect).toHaveBeenCalledWith(
      expect.objectContaining({
        email: "user@contoso.com",
        log,
      }),
    );
    expect(log).toHaveBeenCalledWith(
      "Starting Teams authentication during MCP server startup...",
    );
    expect(log).toHaveBeenCalledWith(
      "Microsoft Teams authentication successful.",
    );
  });

  it("skips eager startup auth without configured token or email", async () => {
    const factory = createFactory();
    const { manager } = createManager({
      clientFactory: factory,
      environment: {},
    });

    await manager.authenticateOnStartup();

    expect(factory.connect).not.toHaveBeenCalled();
    expect(factory.fromToken).not.toHaveBeenCalled();
  });

  it("starts two-phase browser auth when only the tool call supplies email", async () => {
    const client = new FakeTeamsClient();
    const pendingLogin = createPendingPromise<FakeTeamsClient>();
    const factory = createFactory(client);
    factory.fromInteractiveLogin.mockReturnValue(pendingLogin.promise);
    const log = vi.fn();
    const { manager } = createManager({
      canAttemptAutoLogin: () => false,
      clientFactory: factory,
      environment: { TEAMS_LOGIN: "true" },
      log,
    });

    await expect(manager.getClient("user@contoso.com")).rejects.toBeInstanceOf(
      AuthenticationInProgressError,
    );
    expect(factory.fromInteractiveLogin).toHaveBeenCalledWith(
      expect.objectContaining({
        email: "user@contoso.com",
        log,
      }),
    );

    pendingLogin.resolve(client);
    await expect(manager.getClient("user@contoso.com")).rejects.toBeInstanceOf(
      AuthenticationInProgressError,
    );
    await pendingLogin.promise;
    const authenticatedClient = await manager.getClient("user@contoso.com");

    expect(authenticatedClient).toBe(client);
    expect(authenticatedClient.email).toBe("user@contoso.com");
    expect(log).toHaveBeenCalledWith(
      "Starting Microsoft Teams login in the background...",
    );
  });

  it("does not split a tool-email flow when auto-login can run silently", async () => {
    const factory = createFactory();
    const { manager } = createManager({
      canAttemptAutoLogin: () => true,
      clientFactory: factory,
      environment: {},
    });

    await manager.getClient("user@contoso.com");

    expect(factory.connect).toHaveBeenCalledWith(
      expect.objectContaining({
        email: "user@contoso.com",
      }),
    );
  });
});
