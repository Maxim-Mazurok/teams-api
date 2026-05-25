/**
 * Image description utilities.
 *
 * The Teams API provides image bytes; this module sends those bytes to an
 * OpenAI-compatible vision endpoint when the user has configured credentials.
 */

import { fetchWithRetry } from "./common.js";

const DEFAULT_OPENAI_COMPATIBLE_BASE_URL = "https://api.openai.com/v1";
const DEFAULT_OPENAI_COMPATIBLE_MODEL = "gpt-4o-mini";
const DEFAULT_IMAGE_DESCRIPTION_PROMPT =
  "Describe this Teams image for someone who cannot see it. Focus on visible text, UI state, diagrams, charts, and any actionable context. Be concise but complete.";

interface ImageDescriptionConfig {
  endpointUrl: string;
  apiKey: string;
  model: string;
}

export interface DescribeImageOptions {
  imageData: Buffer;
  contentType: string;
  prompt?: string;
  environment?: NodeJS.ProcessEnv;
  fetchFunction?: typeof fetch;
}

export interface ImageDescriptionResult {
  provider: "openai-compatible";
  model: string;
  description: string;
}

interface OpenAiCompatibleVisionResponse {
  choices?: Array<{
    message?: {
      content?: string | null;
    };
  }>;
}

function readImageDescriptionConfig(
  environment: NodeJS.ProcessEnv,
): ImageDescriptionConfig {
  const apiKey =
    environment.TEAMS_IMAGE_DESCRIPTION_API_KEY ?? environment.OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error(
      "Image description requires TEAMS_IMAGE_DESCRIPTION_API_KEY or OPENAI_API_KEY to be set.",
    );
  }

  const baseUrl = (
    environment.TEAMS_IMAGE_DESCRIPTION_BASE_URL ??
    DEFAULT_OPENAI_COMPATIBLE_BASE_URL
  ).replace(/\/+$/, "");
  const model =
    environment.TEAMS_IMAGE_DESCRIPTION_MODEL ??
    DEFAULT_OPENAI_COMPATIBLE_MODEL;

  return {
    endpointUrl: `${baseUrl}/chat/completions`,
    apiKey,
    model,
  };
}

export async function describeImageWithOpenAiCompatibleVision(
  options: DescribeImageOptions,
): Promise<ImageDescriptionResult> {
  const environment = options.environment ?? process.env;
  const fetchFunction = options.fetchFunction ?? fetchWithRetry;
  const config = readImageDescriptionConfig(environment);
  const prompt = options.prompt ?? DEFAULT_IMAGE_DESCRIPTION_PROMPT;

  const response = await fetchFunction(config.endpointUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${config.apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: config.model,
      messages: [
        {
          role: "user",
          content: [
            {
              type: "text",
              text: prompt,
            },
            {
              type: "image_url",
              image_url: {
                url: `data:${options.contentType};base64,${options.imageData.toString("base64")}`,
              },
            },
          ],
        },
      ],
      max_tokens: 500,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Image description request failed: ${response.status} ${response.statusText} — ${errorText}`,
    );
  }

  const responseBody =
    (await response.json()) as OpenAiCompatibleVisionResponse;
  const description = responseBody.choices?.[0]?.message?.content?.trim();
  if (!description) {
    throw new Error(
      "Image description response did not include message content.",
    );
  }

  return {
    provider: "openai-compatible",
    model: config.model,
    description,
  };
}
