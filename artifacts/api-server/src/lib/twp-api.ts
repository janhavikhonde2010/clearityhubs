import { logger } from "./logger";

const TWP_BASE = "https://growth.thewiseparrot.club/api/v1/whatsapp";

export class TwpAuthError extends Error {
  status = 401;
  constructor(message = "Invalid API token or Phone Number ID") {
    super(message);
    this.name = "TwpAuthError";
  }
}

export interface RawSubscriber {
  subscriber_id: number;
  chat_id: string;
  first_name: string;
  last_name: string;
  label_names: string | null;
  last_message_time: string | null;
}

export interface RawMessage {
  id: number;
  sender: "user" | "bot" | "system" | "sequence" | "agent";
  agent_name: string | null;
  message_content: string | null;
  conversation_time: string;
  wa_message_id: string | null;
}

export interface ButtonClick {
  buttonName: string;
  timestamp: string;
}

export interface ProcessedSubscriber {
  name: string;
  phoneNumber: string;
  subscriberId: number | null;
  labelName: string;
  allLabelNames: string[];
  lastMessageTime: string | null;
  assignedAgent: string | null;
  assignedSequence: string | null;
  userReplyCount: number;
  twpReplyCount: number;
  isDormant: boolean;
  lastUserMessage: string | null;
  lastUserMessageTime: string | null;
  assignmentHistory: Array<{ agent: string; timestamp: string }>;
  postSequenceReplies: number;
  buttonClicks: ButtonClick[];
}

function extractTextFromContent(content: string | null): string | null {
  if (!content) return null;
  try {
    const parsed = JSON.parse(content);
    // Try to extract text body from WhatsApp message structure
    const text =
      parsed?.entry?.[0]?.changes?.[0]?.value?.messages?.[0]?.text?.body ||
      parsed?.text?.body ||
      parsed?.body?.text ||
      null;
    return text;
  } catch {
    return content;
  }
}

/**
 * Extracts button clicks from user messages.
 * Handles both flat { interactive: { button_reply: { title } } }
 * and deeply nested WhatsApp API webhook payloads.
 */
function extractButtonClicks(messages: RawMessage[]): ButtonClick[] {
  const clicks: ButtonClick[] = [];
  for (const msg of messages) {
    if (msg.sender !== "user") continue;
    if (!msg.message_content) continue;

    let parsed: Record<string, unknown>;
    try {
      parsed = JSON.parse(msg.message_content);
    } catch {
      continue;
    }

    // Try multiple paths where button_reply.title can live
    const candidates = [
      // Flat format: { interactive: { type: "button_reply", button_reply: { title } } }
      (parsed?.interactive as Record<string, unknown>)?.button_reply,
      // Nested WhatsApp webhook format
      (
        (parsed as Record<string, unknown[]>)?.entry?.[0] as Record<string, unknown[]>
      )?.changes?.[0]
        ? (
            (
              (parsed as Record<string, unknown[]>).entry[0] as Record<string, unknown[]>
            ).changes[0] as Record<string, Record<string, Record<string, unknown>[]>>
          )?.value?.messages?.[0]?.interactive?.button_reply
        : null,
      // Alternative: { type: "interactive", interactive: { button_reply } }
      (parsed?.type === "interactive"
        ? (parsed?.interactive as Record<string, unknown>)?.button_reply
        : null),
    ];

    for (const candidate of candidates) {
      if (candidate && typeof candidate === "object") {
        const title = (candidate as Record<string, unknown>).title;
        if (typeof title === "string" && title.trim()) {
          clicks.push({ buttonName: title.trim(), timestamp: msg.conversation_time });
          break;
        }
      }
    }
  }
  return clicks;
}

function parseConversationMessages(messageStr: string | object): RawMessage[] {
  let raw: Record<string, RawMessage>;
  if (typeof messageStr === "string") {
    try {
      raw = JSON.parse(messageStr);
    } catch {
      return [];
    }
  } else {
    raw = messageStr as Record<string, RawMessage>;
  }

  return Object.values(raw).sort(
    (a, b) => new Date(a.conversation_time).getTime() - new Date(b.conversation_time).getTime(),
  );
}

function extractAssignments(messages: RawMessage[]): Array<{ agent: string; timestamp: string }> {
  const history: Array<{ agent: string; timestamp: string }> = [];
  for (const msg of messages) {
    if (msg.sender === "system") {
      const content = msg.message_content ?? "";
      const match = content.match(/Conversation was assigned to (.+)/i);
      if (match) {
        history.push({ agent: match[1].trim(), timestamp: msg.conversation_time });
      }
    }
  }
  return history;
}

function extractSequence(messages: RawMessage[]): string | null {
  // Look for "Subscribed to sequence: <name>" in system messages
  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i];
    if (msg.sender === "system") {
      const content = msg.message_content ?? "";
      const match = content.match(/[Ss]ubscribed to sequence:\s*(.+)/);
      if (match) return match[1].trim();
    }
    if (msg.sender === "sequence") {
      // sequence messages indicate a sequence was running
      if (msg.agent_name) return msg.agent_name;
    }
  }
  return null;
}

function countPostSequenceReplies(messages: RawMessage[]): number {
  let lastSequenceIdx = -1;
  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i];
    if (
      msg.sender === "sequence" ||
      (msg.sender === "system" &&
        (msg.message_content ?? "").match(/[Ss]ubscribed to sequence/))
    ) {
      lastSequenceIdx = i;
    }
  }
  if (lastSequenceIdx === -1) return 0;
  let count = 0;
  for (let i = lastSequenceIdx + 1; i < messages.length; i++) {
    if (messages[i].sender === "user") count++;
  }
  return count;
}

function isDormantLead(messages: RawMessage[]): boolean {
  const userMessages = messages.filter((m) => m.sender === "user");
  if (userMessages.length === 0) return true;
  const last = userMessages[userMessages.length - 1];
  const lastTime = new Date(last.conversation_time).getTime();
  const sevenDaysAgo = Date.now() - 7 * 24 * 60 * 60 * 1000;
  return lastTime < sevenDaysAgo;
}

export async function fetchAccountInfo(
  apiToken: string,
  phoneNumberId: string,
): Promise<{ businessName: string | null }> {
  const candidates = [
    `${TWP_BASE}/account/profile?apiToken=${encodeURIComponent(apiToken)}&phone_number_id=${encodeURIComponent(phoneNumberId)}`,
    `${TWP_BASE}/account/info?apiToken=${encodeURIComponent(apiToken)}&phone_number_id=${encodeURIComponent(phoneNumberId)}`,
    `${TWP_BASE}/account?apiToken=${encodeURIComponent(apiToken)}&phone_number_id=${encodeURIComponent(phoneNumberId)}`,
    `${TWP_BASE}/profile?apiToken=${encodeURIComponent(apiToken)}&phone_number_id=${encodeURIComponent(phoneNumberId)}`,
  ];
  for (const url of candidates) {
    try {
      const res = await fetch(url, { signal: AbortSignal.timeout(5000) });
      if (!res.ok) continue;
      const data = await res.json() as Record<string, unknown>;
      if (data.status !== "1") continue;
      const msg = data.message as Record<string, unknown> | undefined;
      if (!msg) continue;
      const name =
        (msg.business_name as string | undefined) ||
        (msg.name as string | undefined) ||
        (msg.account_name as string | undefined) ||
        (msg.display_name as string | undefined);
      if (name) return { businessName: name };
    } catch {
      // try next
    }
  }
  return { businessName: null };
}

export async function fetchSubscribers(apiToken: string, phoneNumberId: string): Promise<RawSubscriber[]> {
  const url = `${TWP_BASE}/subscriber/list?apiToken=${encodeURIComponent(apiToken)}&phone_number_id=${encodeURIComponent(phoneNumberId)}&limit=10000`;
  logger.info({ url: url.split("?")[0] }, "Fetching subscribers");
  const res = await fetch(url, { signal: AbortSignal.timeout(20_000) });
  if (!res.ok) {
    logger.error({ status: res.status }, "Failed to fetch subscribers");
    if (res.status === 401 || res.status === 403) {
      throw new TwpAuthError();
    }
    throw new Error(`Subscriber API returned ${res.status}`);
  }
  const data = await res.json() as { status: string; message: RawSubscriber[] | string };
  if (data.status !== "1") {
    logger.warn({ status: data.status }, "Subscriber API returned non-1 status");
    return [];
  }
  if (Array.isArray(data.message)) return data.message;
  return [];
}

export async function fetchConversation(
  apiToken: string,
  phoneNumberId: string,
  phoneNumber: string,
): Promise<RawMessage[]> {
  const url = `${TWP_BASE}/get/conversation?apiToken=${encodeURIComponent(apiToken)}&phone_number_id=${encodeURIComponent(phoneNumberId)}&phone_number=${encodeURIComponent(phoneNumber)}&limit=500`;
  try {
    const res = await fetch(url, { signal: AbortSignal.timeout(8_000) });
    if (!res.ok) {
      logger.warn({ status: res.status, phoneNumber }, "Failed to fetch conversation");
      return [];
    }
    const data = await res.json() as { status: string; message: string | object };
    if (data.status !== "1" || !data.message) return [];
    return parseConversationMessages(data.message);
  } catch (err) {
    logger.warn({ err, phoneNumber }, "Error fetching conversation");
    return [];
  }
}

// Maximum number of subscribers to process per request.
// Keeps total processing time well under the 30s proxy timeout.
const MAX_SUBSCRIBERS = 30_000;

export async function processSubscribers(
  apiToken: string,
  phoneNumberId: string,
  subscribers: RawSubscriber[],
): Promise<ProcessedSubscriber[]> {
  const results: ProcessedSubscriber[] = [];

  // Sort by most-recent activity so the most relevant leads are always included
  const sorted = [...subscribers].sort((a, b) => {
    const ta = a.last_message_time ? new Date(a.last_message_time).getTime() : 0;
    const tb = b.last_message_time ? new Date(b.last_message_time).getTime() : 0;
    return tb - ta;
  });
  const limited = sorted.slice(0, MAX_SUBSCRIBERS);

  if (subscribers.length > MAX_SUBSCRIBERS) {
    logger.warn(
      { total: subscribers.length, processing: MAX_SUBSCRIBERS },
      "Subscriber list truncated to stay within timeout budget",
    );
  }

  // Process in batches of 30 to balance speed vs API load
  const BATCH_SIZE = 30;
  for (let i = 0; i < limited.length; i += BATCH_SIZE) {
    const batch = limited.slice(i, i + BATCH_SIZE);
    const batchResults = await Promise.all(
      batch.map(async (sub) => {
        const phoneNumber = sub.chat_id ?? "";
        const name = [sub.first_name, sub.last_name].filter(Boolean).join(" ").trim() || phoneNumber;

        let messages: RawMessage[] = [];
        if (phoneNumber) {
          messages = await fetchConversation(apiToken, phoneNumberId, phoneNumber);
        }

        const userMessages = messages.filter((m) => m.sender === "user");
        const twpMessages = messages.filter(
          (m) => m.sender === "bot" || m.sender === "sequence" || m.sender === "agent",
        );
        const assignmentHistory = extractAssignments(messages);
        const assignedAgent =
          assignmentHistory.length > 0
            ? assignmentHistory[assignmentHistory.length - 1].agent
            : null;
        const assignedSequence = extractSequence(messages);

        const lastUserMsg = userMessages[userMessages.length - 1];
        const lastUserText = lastUserMsg ? extractTextFromContent(lastUserMsg.message_content) : null;

        // Parse label_names (comma-separated string)
        const allLabelNames = sub.label_names
          ? sub.label_names.split(",").map((l) => l.trim()).filter(Boolean)
          : [];
        const labelName = allLabelNames[0] ?? "Unlabeled";

        const buttonClicks = extractButtonClicks(messages);

        return {
          name,
          phoneNumber,
          subscriberId: sub.subscriber_id ?? null,
          labelName,
          allLabelNames,
          lastMessageTime: sub.last_message_time,
          assignedAgent,
          assignedSequence,
          userReplyCount: userMessages.length,
          twpReplyCount: twpMessages.length,
          isDormant: isDormantLead(messages),
          lastUserMessage: lastUserText,
          lastUserMessageTime: lastUserMsg?.conversation_time ?? null,
          assignmentHistory,
          postSequenceReplies: countPostSequenceReplies(messages),
          buttonClicks,
        };
      }),
    );
    results.push(...batchResults);
  }

  return results;
}
