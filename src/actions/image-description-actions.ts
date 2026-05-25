/**
 * Image description action definitions.
 *
 * Actions: describe-image.
 */

import type { ImageAttachment } from "../types.js";
import { describeImageWithOpenAiCompatibleVision } from "../api/image-description.js";
import { type ActionDefinition } from "./formatters.js";
import {
  conversationParameters,
  resolveConversationId,
} from "./conversation-resolution.js";

export const ImageDescriptionSourceType = {
  AmsObject: "ams-object",
  Message: "message",
} as const;

export type ImageDescriptionSourceType =
  (typeof ImageDescriptionSourceType)[keyof typeof ImageDescriptionSourceType];

export interface DescribeImageResult {
  description: string;
  provider: string;
  model: string;
  image: {
    sourceType: ImageDescriptionSourceType;
    amsObjectId: string;
    contentType: string;
    size: number;
    messageId?: string;
    imageIndex?: number;
    width?: number | null;
    height?: number | null;
  };
}

interface ImageTarget {
  sourceType: ImageDescriptionSourceType;
  amsObjectId: string;
  messageId?: string;
  imageIndex?: number;
  width?: number | null;
  height?: number | null;
}

function resolveImageIndex(value: number | undefined): number {
  const imageIndex = value ?? 0;
  if (!Number.isInteger(imageIndex) || imageIndex < 0) {
    throw new Error("imageIndex must be a non-negative integer.");
  }
  return imageIndex;
}

function resolveMessageImage(
  images: ImageAttachment[],
  imageIndex: number,
): ImageAttachment {
  const image = images[imageIndex];
  if (!image) {
    throw new Error(
      `Message image index ${imageIndex} is out of range; the message has ${images.length} image(s).`,
    );
  }
  return image;
}

async function resolveImageTarget(
  client: Parameters<ActionDefinition["execute"]>[0],
  parameters: Record<string, unknown>,
): Promise<ImageTarget> {
  const amsObjectId = parameters.amsObjectId as string | undefined;
  if (amsObjectId) {
    return {
      sourceType: ImageDescriptionSourceType.AmsObject,
      amsObjectId,
    };
  }

  const messageId = parameters.messageId as string | undefined;
  if (!messageId) {
    throw new Error(
      "Provide either amsObjectId, or messageId plus chat/to/conversationId.",
    );
  }

  const { conversationId } = await resolveConversationId(client, parameters);
  const messages = await client.getMessages(conversationId);
  const message = messages.find((candidate) => candidate.id === messageId);
  if (!message) {
    throw new Error(
      `Message ${messageId} not found in conversation ${conversationId}`,
    );
  }

  if (message.images.length === 0) {
    throw new Error(`Message ${messageId} has no inline images.`);
  }

  const imageIndex = resolveImageIndex(parameters.imageIndex as number);
  const image = resolveMessageImage(message.images, imageIndex);

  return {
    sourceType: ImageDescriptionSourceType.Message,
    amsObjectId: image.amsObjectId,
    messageId,
    imageIndex,
    width: image.width,
    height: image.height,
  };
}

export const describeImageAction: ActionDefinition = {
  name: "describe-image",
  title: "Describe Image",
  description:
    "Describe an inline Teams image using a configured OpenAI-compatible vision model. " +
    "Use amsObjectId directly, or identify a message by chat/to/conversationId plus messageId. " +
    "Set TEAMS_IMAGE_DESCRIPTION_API_KEY or OPENAI_API_KEY before using this tool.",
  parameters: [
    ...conversationParameters,
    {
      name: "messageId",
      type: "string",
      description:
        "ID of the message containing the image. Required unless amsObjectId is provided.",
      required: false,
    },
    {
      name: "imageIndex",
      type: "number",
      description:
        "Zero-based index of the inline image within the message (default: 0).",
      required: false,
      default: 0,
    },
    {
      name: "amsObjectId",
      type: "string",
      description:
        "Direct AMS object ID for an inline Teams image. When provided, conversation and message lookup are skipped.",
      required: false,
    },
    {
      name: "fullSize",
      type: "boolean",
      description:
        "Download the full-size AMS image before describing it (default: true).",
      required: false,
      default: true,
    },
    {
      name: "prompt",
      type: "string",
      description:
        "Optional custom instruction for the vision model. Defaults to a concise Teams-context description prompt.",
      required: false,
    },
  ],
  execute: async (client, parameters) => {
    const imageTarget = await resolveImageTarget(client, parameters);
    const fullSize = (parameters.fullSize as boolean | undefined) ?? true;
    const imageData = await client.downloadImage(
      imageTarget.amsObjectId,
      fullSize,
    );
    const descriptionResult = await describeImageWithOpenAiCompatibleVision({
      imageData: imageData.data,
      contentType: imageData.contentType,
      prompt: parameters.prompt as string | undefined,
    });

    return {
      description: descriptionResult.description,
      provider: descriptionResult.provider,
      model: descriptionResult.model,
      image: {
        ...imageTarget,
        contentType: imageData.contentType,
        size: imageData.size,
      },
    } satisfies DescribeImageResult;
  },
  formatConcise: (result) => {
    const descriptionResult = result as DescribeImageResult;
    const lines = [
      "## Image description",
      "",
      descriptionResult.description,
      "",
      `- Provider: ${descriptionResult.provider}`,
      `- Model: ${descriptionResult.model}`,
      `- AMS object ID: ${descriptionResult.image.amsObjectId}`,
      `- Content type: ${descriptionResult.image.contentType}`,
      `- Size: ${descriptionResult.image.size} bytes`,
    ];

    if (descriptionResult.image.messageId) {
      lines.push(`- Message ID: ${descriptionResult.image.messageId}`);
    }
    if (descriptionResult.image.imageIndex !== undefined) {
      lines.push(`- Image index: ${descriptionResult.image.imageIndex}`);
    }
    if (
      descriptionResult.image.width !== undefined &&
      descriptionResult.image.height !== undefined
    ) {
      lines.push(
        `- Dimensions: ${descriptionResult.image.width ?? "unknown"} x ${descriptionResult.image.height ?? "unknown"}`,
      );
    }

    return lines.join("\n");
  },
};
