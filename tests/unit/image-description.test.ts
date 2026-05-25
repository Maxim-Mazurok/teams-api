import { describe, expect, it, vi } from "vitest";
import { describeImageWithOpenAiCompatibleVision } from "../../src/api/image-description.js";

describe("describeImageWithOpenAiCompatibleVision", () => {
  it("should call an OpenAI-compatible vision endpoint", async () => {
    const fetchFunction = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      statusText: "OK",
      json: () =>
        Promise.resolve({
          choices: [{ message: { content: "A screenshot of a task board." } }],
        }),
      text: () => Promise.resolve(""),
    });
    const imageData = Buffer.from("image-bytes");

    const result = await describeImageWithOpenAiCompatibleVision({
      imageData,
      contentType: "image/png",
      prompt: "Describe the UI.",
      environment: {
        TEAMS_IMAGE_DESCRIPTION_API_KEY: "test-key",
        TEAMS_IMAGE_DESCRIPTION_BASE_URL: "https://vision.example.test/v1/",
        TEAMS_IMAGE_DESCRIPTION_MODEL: "vision-model",
      } as NodeJS.ProcessEnv,
      fetchFunction,
    });

    expect(result).toEqual({
      provider: "openai-compatible",
      model: "vision-model",
      description: "A screenshot of a task board.",
    });
    expect(fetchFunction).toHaveBeenCalledWith(
      "https://vision.example.test/v1/chat/completions",
      expect.objectContaining({
        method: "POST",
        headers: {
          Authorization: "Bearer test-key",
          "Content-Type": "application/json",
        },
      }),
    );
    const requestBody = JSON.parse(
      fetchFunction.mock.calls[0][1].body as string,
    ) as {
      messages: Array<{
        content: Array<{
          text?: string;
          image_url?: {
            url: string;
          };
        }>;
      }>;
    };
    expect(requestBody.messages[0].content[0].text).toBe("Describe the UI.");
    expect(requestBody.messages[0].content[1].image_url?.url).toBe(
      `data:image/png;base64,${imageData.toString("base64")}`,
    );
  });

  it("should throw when no API key is configured", async () => {
    await expect(
      describeImageWithOpenAiCompatibleVision({
        imageData: Buffer.from("image"),
        contentType: "image/png",
        environment: {} as NodeJS.ProcessEnv,
        fetchFunction: vi.fn(),
      }),
    ).rejects.toThrow("TEAMS_IMAGE_DESCRIPTION_API_KEY or OPENAI_API_KEY");
  });

  it("should surface provider error responses", async () => {
    const fetchFunction = vi.fn().mockResolvedValue({
      ok: false,
      status: 429,
      statusText: "Too Many Requests",
      text: () => Promise.resolve("rate limited"),
    });

    await expect(
      describeImageWithOpenAiCompatibleVision({
        imageData: Buffer.from("image"),
        contentType: "image/png",
        environment: {
          TEAMS_IMAGE_DESCRIPTION_API_KEY: "test-key",
        } as NodeJS.ProcessEnv,
        fetchFunction,
      }),
    ).rejects.toThrow(
      "Image description request failed: 429 Too Many Requests",
    );
  });
});
