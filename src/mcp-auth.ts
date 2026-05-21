import { canAttemptAutoLogin as defaultCanAttemptAutoLogin } from "./platform.js";
import { recordAuth as defaultRecordAuth } from "./telemetry.js";
import type {
  AuthLogFunction,
  AutoLoginOptions,
  InteractiveLoginOptions,
  ManualTokenOptions,
  SmartLoginOptions,
} from "./types.js";

interface AuthenticatedTeamsClient {
  setEmail(email: string): void;
}

interface TeamsClientAuthFactory<Client extends AuthenticatedTeamsClient> {
  create(options: AutoLoginOptions): Promise<Client>;
  connect(options?: SmartLoginOptions): Promise<Client>;
  fromInteractiveLogin(options?: InteractiveLoginOptions): Promise<Client>;
  fromDebugSession(options?: ManualTokenOptions): Promise<Client>;
  fromToken(
    skypeToken: string,
    region: string,
    bearerToken?: string,
    substrateToken?: string,
  ): Client;
}

type AuthRecord = Parameters<typeof defaultRecordAuth>[0];
type RecordAuthFunction = (record: AuthRecord) => void;
type CanAttemptAutoLoginFunction = () => boolean;

interface McpAuthEnvironment {
  TEAMS_AUTO?: string;
  TEAMS_BEARER_TOKEN?: string;
  TEAMS_DEBUG?: string;
  TEAMS_DEBUG_PORT?: string;
  TEAMS_EMAIL?: string;
  TEAMS_LOGIN?: string;
  TEAMS_REGION?: string;
  TEAMS_SUBSTRATE_TOKEN?: string;
  TEAMS_TOKEN?: string;
}

interface ResolvedMcpAuthEnvironment {
  auto: boolean;
  bearerToken?: string;
  debug: boolean;
  debugPort: number;
  email?: string;
  login: boolean;
  region?: string;
  substrateToken?: string;
  token?: string;
}

export interface McpAuthManagerOptions<
  Client extends AuthenticatedTeamsClient,
> {
  canAttemptAutoLogin?: CanAttemptAutoLoginFunction;
  clientFactory: TeamsClientAuthFactory<Client>;
  environment?: McpAuthEnvironment;
  log?: AuthLogFunction;
  recordAuth?: RecordAuthFunction;
}

export class NeedsEmailError extends Error {
  constructor() {
    super(
      "I need your corporate email address to log into Teams. " +
        "Please provide your email and call this tool again.",
    );
    this.name = "NeedsEmailError";
  }
}

export class AuthenticationInProgressError extends Error {
  constructor() {
    super(
      "A browser window has been opened for Microsoft Teams login. " +
        "Please complete the login, then call this tool again.",
    );
    this.name = "AuthenticationInProgressError";
  }
}

function getEnvironmentValue(
  environment: McpAuthEnvironment,
  name: keyof McpAuthEnvironment,
): string | undefined {
  const value = environment[name];
  return value && value.length > 0 ? value : undefined;
}

function resolveMcpAuthEnvironment(
  environment: McpAuthEnvironment,
): ResolvedMcpAuthEnvironment {
  const debugPortValue = getEnvironmentValue(environment, "TEAMS_DEBUG_PORT");

  return {
    auto: environment.TEAMS_AUTO === "true",
    bearerToken: getEnvironmentValue(environment, "TEAMS_BEARER_TOKEN"),
    debug: environment.TEAMS_DEBUG === "true",
    debugPort: debugPortValue ? Number(debugPortValue) : 9222,
    email: getEnvironmentValue(environment, "TEAMS_EMAIL"),
    login: environment.TEAMS_LOGIN === "true",
    region: getEnvironmentValue(environment, "TEAMS_REGION"),
    substrateToken: getEnvironmentValue(environment, "TEAMS_SUBSTRATE_TOKEN"),
    token: getEnvironmentValue(environment, "TEAMS_TOKEN"),
  };
}

function formatError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export class McpAuthManager<Client extends AuthenticatedTeamsClient> {
  private readonly canAttemptAutoLogin: CanAttemptAutoLoginFunction;
  private readonly clientFactory: TeamsClientAuthFactory<Client>;
  private readonly environment: McpAuthEnvironment;
  private readonly log: AuthLogFunction;
  private readonly recordAuth: RecordAuthFunction;
  private authenticationPromise: Promise<Client> | null = null;
  private clientInstance: Client | null = null;

  constructor(options: McpAuthManagerOptions<Client>) {
    this.canAttemptAutoLogin =
      options.canAttemptAutoLogin ?? defaultCanAttemptAutoLogin;
    this.clientFactory = options.clientFactory;
    this.environment = options.environment ?? process.env;
    this.log = options.log ?? (() => {});
    this.recordAuth = options.recordAuth ?? defaultRecordAuth;
  }

  async authenticateOnStartup(): Promise<void> {
    const authEnvironment = resolveMcpAuthEnvironment(this.environment);
    if (!authEnvironment.token && !authEnvironment.email) {
      return;
    }

    try {
      this.log("Starting Teams authentication during MCP server startup...");
      await this.getClient();
    } catch (error) {
      this.log(
        `Teams authentication during MCP server startup failed: ${formatError(
          error,
        )}`,
      );
    }
  }

  async getClient(toolEmail?: string): Promise<Client> {
    if (this.clientInstance) {
      return this.clientInstance;
    }

    const authEnvironment = resolveMcpAuthEnvironment(this.environment);
    if (this.shouldUseTwoPhaseAuthentication(toolEmail, authEnvironment)) {
      if (!this.authenticationPromise) {
        this.log("Starting Microsoft Teams login in the background...");
        void this.startAuthentication(toolEmail).catch((error: unknown) => {
          this.log(`Microsoft Teams login failed: ${formatError(error)}`);
        });
      }
      throw new AuthenticationInProgressError();
    }

    return this.startAuthentication(toolEmail);
  }

  private shouldUseTwoPhaseAuthentication(
    toolEmail: string | undefined,
    authEnvironment: ResolvedMcpAuthEnvironment,
  ): boolean {
    if (
      !toolEmail ||
      authEnvironment.email ||
      authEnvironment.token ||
      authEnvironment.debug
    ) {
      return false;
    }

    if (authEnvironment.login) {
      return true;
    }

    if (authEnvironment.auto) {
      return !this.canAttemptAutoLogin();
    }

    return !this.canAttemptAutoLogin();
  }

  private startAuthentication(toolEmail: string | undefined): Promise<Client> {
    this.authenticationPromise ??= this.authenticate(toolEmail)
      .then((client) => {
        this.clientInstance = client;
        return client;
      })
      .finally(() => {
        this.authenticationPromise = null;
      });

    return this.authenticationPromise;
  }

  private async authenticate(toolEmail: string | undefined): Promise<Client> {
    const authEnvironment = resolveMcpAuthEnvironment(this.environment);
    const email = authEnvironment.email ?? toolEmail;
    const log = this.log;

    if (authEnvironment.token) {
      if (!authEnvironment.region) {
        throw new Error("TEAMS_REGION is required when TEAMS_TOKEN is set");
      }
      const client = this.clientFactory.fromToken(
        authEnvironment.token,
        authEnvironment.region,
        authEnvironment.bearerToken,
        authEnvironment.substrateToken,
      );
      if (email) {
        client.setEmail(email);
      }
      this.recordAuth({ strategy: "token", success: true });
      this.log("Microsoft Teams authentication successful.");
      return client;
    }

    if (authEnvironment.auto) {
      if (!email) {
        throw new NeedsEmailError();
      }
      try {
        const client = await this.clientFactory.create({
          email,
          region: authEnvironment.region,
          headless: true,
          verbose: false,
          log,
        });
        this.recordAuth({ strategy: "auto", success: true });
        this.log("Microsoft Teams authentication successful.");
        return client;
      } catch (error) {
        this.recordAuth({ strategy: "auto", success: false, error });
        throw error;
      }
    }

    if (authEnvironment.login) {
      try {
        const client = await this.clientFactory.fromInteractiveLogin({
          region: authEnvironment.region,
          email,
          verbose: false,
          log,
        });
        if (email) {
          client.setEmail(email);
        }
        this.recordAuth({ strategy: "login", success: true });
        this.log("Microsoft Teams authentication successful.");
        return client;
      } catch (error) {
        this.recordAuth({ strategy: "login", success: false, error });
        throw error;
      }
    }

    if (authEnvironment.debug) {
      try {
        const client = await this.clientFactory.fromDebugSession({
          debugPort: authEnvironment.debugPort,
          region: authEnvironment.region,
        });
        this.recordAuth({ strategy: "debug", success: true });
        this.log("Microsoft Teams authentication successful.");
        return client;
      } catch (error) {
        this.recordAuth({ strategy: "debug", success: false, error });
        throw error;
      }
    }

    try {
      const client = await this.clientFactory.connect({
        email,
        region: authEnvironment.region,
        verbose: false,
        log,
      });
      if (email) {
        client.setEmail(email);
      }
      this.recordAuth({ strategy: "auto", success: true });
      this.log("Microsoft Teams authentication successful.");
      return client;
    } catch (error) {
      this.recordAuth({ strategy: "auto", success: false, error });
      throw error;
    }
  }
}
