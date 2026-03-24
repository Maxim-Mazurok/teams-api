/**
 * Common utilities shared across API modules.
 *
 * Contains error classes, retry logic, and shared helpers
 * used by chat-service, middle-tier, substrate, and transcript modules.
 */

export class ApiAuthError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ApiAuthError";
  }
}

export class ApiRateLimitError extends Error {
  public readonly retryAfterMilliseconds: number;

  constructor(message: string, retryAfterMilliseconds: number) {
    super(message);
    this.name = "ApiRateLimitError";
    this.retryAfterMilliseconds = retryAfterMilliseconds;
  }
}

const MAX_RETRY_ATTEMPTS = 5;
const INITIAL_BACKOFF_MILLISECONDS = 2_000;

function parseRetryAfter(response: Response): number {
  const retryAfterHeader = response.headers.get("Retry-After");
  if (!retryAfterHeader) {
    return INITIAL_BACKOFF_MILLISECONDS;
  }
  const seconds = Number(retryAfterHeader);
  if (!Number.isNaN(seconds) && seconds > 0) {
    return seconds * 1_000;
  }
  return INITIAL_BACKOFF_MILLISECONDS;
}

function delay(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

/**
 * Wrapper around `fetch` that automatically retries on 429 (rate limit) responses.
 *
 * Uses the `Retry-After` header when present; otherwise applies exponential backoff
 * starting at 2 seconds. Retries up to 5 times before throwing `ApiRateLimitError`.
 */
export async function fetchWithRetry(
  input: string | URL | Request,
  init?: RequestInit,
): Promise<Response> {
  for (let attempt = 0; attempt <= MAX_RETRY_ATTEMPTS; attempt++) {
    const response = await fetch(input, init);

    if (response.status !== 429) {
      return response;
    }

    if (attempt === MAX_RETRY_ATTEMPTS) {
      const errorText = await response.text();
      throw new ApiRateLimitError(
        `Rate limit exceeded after ${MAX_RETRY_ATTEMPTS + 1} attempts: ${response.status} ${errorText}`,
        parseRetryAfter(response),
      );
    }

    const retryAfterMilliseconds = parseRetryAfter(response);
    const backoffMilliseconds = retryAfterMilliseconds * Math.pow(2, attempt);
    await delay(backoffMilliseconds);
  }

  throw new Error("Unreachable: fetchWithRetry loop exited unexpectedly");
}
