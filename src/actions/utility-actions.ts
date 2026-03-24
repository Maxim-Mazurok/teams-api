/**
 * Utility action definitions.
 *
 * Actions: whoami, get-members, get-transcript.
 */

import type { Member, TranscriptResult } from "../types.js";
import {
  type ActionDefinition,
  toonHeader,
  formatTimestamp,
  groupBySpeaker,
} from "./formatters.js";
import {
  conversationParameters,
  resolveConversationId,
} from "./conversation-resolution.js";

export const getMembers: ActionDefinition = {
  name: "get-members",
  title: "Get Conversation Members",
  description:
    "List members of a conversation. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "At least one identifier is required. " +
    "Display names are resolved via the Teams profile API when available, with message history as fallback. " +
    "Note: 1:1 chat members may have empty display names if profile resolution is unavailable.",
  parameters: [...conversationParameters],
  execute: async (client, parameters) => {
    const { conversationId } = await resolveConversationId(client, parameters);
    return client.getMembers(conversationId);
  },
  formatResult: (result) => {
    const members = result as Member[];
    const people = members.filter((member) => member.memberType === "person");
    const bots = members.filter((member) => member.memberType === "bot");
    const lines = [`\n${people.length} people, ${bots.length} bots:\n`];
    for (const member of people) {
      const name = member.displayName || "(unknown)";
      lines.push(`  ${name} (${member.role}) — ${member.id}`);
    }
    if (bots.length > 0) {
      lines.push("");
      lines.push("  Bots/Apps:");
      for (const bot of bots) {
        const name = bot.displayName || "(unnamed bot)";
        lines.push(`  ${name} — ${bot.id}`);
      }
    }
    return lines.join("\n");
  },
  formatMarkdown: (result) => {
    const members = result as Member[];
    const people = members.filter((member) => member.memberType === "person");
    const bots = members.filter((member) => member.memberType === "bot");
    const lines = [
      `## Members (${people.length} people, ${bots.length} bots)`,
      "",
    ];
    if (people.length > 0) {
      lines.push("| Name | Role | ID |");
      lines.push("|------|------|----|");
      for (const member of people) {
        const name = member.displayName || "(unknown)";
        lines.push(`| ${name} | ${member.role} | ${member.id} |`);
      }
    }
    if (bots.length > 0) {
      lines.push("", "### Bots/Apps", "");
      lines.push("| Name | ID |");
      lines.push("|------|----|");
      for (const bot of bots) {
        const name = bot.displayName || "(unnamed bot)";
        lines.push(`| ${name} | ${bot.id} |`);
      }
    }
    return lines.join("\n");
  },
  formatToon: (result) => {
    const members = result as Member[];
    const people = members.filter((member) => member.memberType === "person");
    const bots = members.filter((member) => member.memberType === "bot");
    const lines = [
      toonHeader("👥", `${people.length} People, ${bots.length} Bots`),
    ];
    for (const member of people) {
      const name = member.displayName || "(unknown)";
      lines.push("");
      lines.push(`  👤 ${name} · ${member.role}`);
      lines.push(`     ${member.id}`);
    }
    if (bots.length > 0) {
      lines.push("");
      lines.push("  🤖 Bots/Apps:");
      for (const bot of bots) {
        const name = bot.displayName || "(unnamed bot)";
        lines.push(`     🤖 ${name} — ${bot.id}`);
      }
    }
    return lines.join("\n");
  },
};

export const whoami: ActionDefinition = {
  name: "whoami",
  title: "Current User Info",
  description:
    "Get the display name and region of the currently authenticated user.",
  parameters: [],
  execute: async (client) => {
    const displayName = await client.getCurrentUserDisplayName();
    const token = client.getToken();
    return { displayName, region: token.region };
  },
  formatResult: (result) => {
    const { displayName, region } = result as {
      displayName: string;
      region: string;
    };
    return `${displayName} (region: ${region})`;
  },
  formatMarkdown: (result) => {
    const { displayName, region } = result as {
      displayName: string;
      region: string;
    };
    return [`## ${displayName}`, "", `- **Region:** ${region}`].join("\n");
  },
  formatToon: (result) => {
    const { displayName, region } = result as {
      displayName: string;
      region: string;
    };
    return [toonHeader("🙋", displayName), `  📍 region: ${region}`].join("\n");
  },
};

export const getTranscript: ActionDefinition = {
  name: "get-transcript",
  title: "Get Meeting Transcript",
  description:
    "Get the meeting transcript from a conversation that contains a recorded meeting. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "Use --raw-vtt to get the original VTT file instead of parsed output.",
  parameters: [
    ...conversationParameters,
    {
      name: "rawVtt",
      type: "boolean",
      description:
        "Return the original VTT file content instead of parsed transcript (default: false)",
      required: false,
      default: false,
    },
  ],
  execute: async (client, parameters) => {
    const { conversationId } = await resolveConversationId(client, parameters);
    const rawVtt = (parameters.rawVtt as boolean | undefined) ?? false;

    const transcriptResult = await client.getTranscript(conversationId);

    if (rawVtt) {
      return { rawVtt: transcriptResult.rawVtt, format: "vtt" as const };
    }

    return transcriptResult;
  },
  formatResult: (result) => {
    const data = result as TranscriptResult | { rawVtt: string; format: "vtt" };

    if ("format" in data && data.format === "vtt") {
      return data.rawVtt;
    }

    const transcript = data as TranscriptResult;
    const groups = groupBySpeaker(transcript.entries);
    const lines = [
      `\nTranscript: ${transcript.meetingTitle} (${transcript.entries.length} segments)\n`,
    ];

    for (const group of groups) {
      const time = formatTimestamp(group.startTime);
      lines.push(`  [${time}] ${group.speaker}:`);
      lines.push(`    ${group.segments.join(" ")}`);
    }

    return lines.join("\n");
  },
  formatMarkdown: (result) => {
    const data = result as TranscriptResult | { rawVtt: string; format: "vtt" };

    if ("format" in data && data.format === "vtt") {
      return ["```vtt", data.rawVtt, "```"].join("\n");
    }

    const transcript = data as TranscriptResult;
    const groups = groupBySpeaker(transcript.entries);
    const lines = [
      `## Transcript: ${transcript.meetingTitle}`,
      "",
      `*${transcript.entries.length} segments*`,
      "",
    ];

    for (const group of groups) {
      const time = formatTimestamp(group.startTime);
      lines.push(`**${group.speaker}** *(${time})*`, "");
      lines.push(group.segments.join(" "), "");
    }

    return lines.join("\n");
  },
  formatToon: (result) => {
    const data = result as TranscriptResult | { rawVtt: string; format: "vtt" };

    if ("format" in data && data.format === "vtt") {
      return data.rawVtt;
    }

    const transcript = data as TranscriptResult;
    const groups = groupBySpeaker(transcript.entries);
    const lines = [
      toonHeader(
        "🎙️",
        `Transcript: ${transcript.meetingTitle} (${transcript.entries.length} segments)`,
      ),
    ];

    for (const group of groups) {
      const time = formatTimestamp(group.startTime);
      lines.push("");
      lines.push(`  🗣️  ${group.speaker} · ${time}`);
      lines.push(`      ${group.segments.join(" ")}`);
    }

    return lines.join("\n");
  },
};
