/**
 * Page state diagnostics for the authentication flow.
 *
 * Provides structured descriptions of where the user is in the
 * Microsoft login flow, used for progress reporting and error messages.
 */

/** Describes the current state of the page during the login flow. */
export interface PageState {
  url: string;
  phase:
    | "login-page"
    | "mfa-challenge"
    | "redirect"
    | "teams-loading"
    | "teams-loaded"
    | "unknown";
  detail: string;
}

/**
 * Diagnose the current page state during authentication.
 * Returns a structured description of where the user is in the login flow.
 */
export async function diagnosePageState(page: {
  url: () => string;
  evaluate: (script: string) => Promise<unknown>;
}): Promise<PageState> {
  const url = page.url();

  if (
    url.includes("login.microsoftonline.com") ||
    url.includes("login.microsoft.com")
  ) {
    const pageContent = (await page
      .evaluate("document.body?.innerText?.slice(0, 500) ?? ''")
      .catch(() => "")) as string;

    if (
      pageContent.includes("Approve sign in request") ||
      pageContent.includes("Enter code") ||
      pageContent.includes("Verify your identity") ||
      pageContent.includes("approve") ||
      url.includes("/common/SAS")
    ) {
      return {
        url,
        phase: "mfa-challenge",
        detail: "Waiting for MFA/2FA approval",
      };
    }

    return {
      url,
      phase: "login-page",
      detail: "On the Microsoft login page",
    };
  }

  if (
    url.includes("teams.cloud.microsoft") ||
    url.includes("teams.microsoft.com")
  ) {
    const isLoaded = (await page
      .evaluate(
        "(document.querySelector('#app') ?? document.querySelector(\"[data-tid='app-layout']\") ?? document.querySelector('main')) !== null",
      )
      .catch(() => false)) as boolean;

    if (isLoaded) {
      return {
        url,
        phase: "teams-loaded",
        detail: "Teams application loaded",
      };
    }

    return {
      url,
      phase: "teams-loading",
      detail: "Teams page reached, application loading",
    };
  }

  if (url.includes("microsoftonline.com") || url.includes("microsoft.com")) {
    return {
      url,
      phase: "redirect",
      detail: "Redirecting through Microsoft authentication",
    };
  }

  return { url, phase: "unknown", detail: `Unexpected page: ${url}` };
}
