import { Router, type IRouter } from "express";
import {
  GetSubscribersQueryParams,
  GetDashboardSummaryQueryParams,
  GetAgentStatsQueryParams,
  GetSequenceStatsQueryParams,
  GetLabelStatsQueryParams,
} from "@workspace/api-zod";
import { fetchSubscribers, processSubscribers, fetchAccountInfo, type ProcessedSubscriber } from "../lib/twp-api";

const router: IRouter = Router();

// In-memory cache to avoid re-fetching on every stat endpoint call
const cache = new Map<string, { data: ProcessedSubscriber[]; timestamp: number }>();
const CACHE_TTL = 5 * 60 * 1000; // 5 minutes

// In-flight promise deduplication — prevents all 6 simultaneous dashboard
// requests from each triggering their own processSubscribers() call when the
// cache is cold.
const inflight = new Map<string, Promise<ProcessedSubscriber[]>>();

async function getProcessedSubscribers(apiToken: string, phoneNumberId: string): Promise<ProcessedSubscriber[]> {
  const key = `${apiToken}:${phoneNumberId}`;

  // Return cached result if still fresh
  const cached = cache.get(key);
  if (cached && Date.now() - cached.timestamp < CACHE_TTL) {
    return cached.data;
  }

  // If a fetch is already in progress for this key, share its promise
  const existing = inflight.get(key);
  if (existing) return existing;

  // Start a new fetch and register it so concurrent callers can join
  const promise = (async () => {
    try {
      const rawSubs = await fetchSubscribers(apiToken, phoneNumberId);
      const processed = await processSubscribers(apiToken, phoneNumberId, rawSubs);
      cache.set(key, { data: processed, timestamp: Date.now() });
      return processed;
    } finally {
      inflight.delete(key);
    }
  })();

  inflight.set(key, promise);
  return promise;
}

router.get("/subscribers", async (req, res): Promise<void> => {
  const parsed = GetSubscribersQueryParams.safeParse(req.query);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const { apiToken, phoneNumberId } = parsed.data;
  const subscribers = await getProcessedSubscribers(apiToken, phoneNumberId);
  res.json({ subscribers, total: subscribers.length });
});

router.get("/dashboard/summary", async (req, res): Promise<void> => {
  const parsed = GetDashboardSummaryQueryParams.safeParse(req.query);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const { apiToken, phoneNumberId } = parsed.data;
  const subscribers = await getProcessedSubscribers(apiToken, phoneNumberId);

  const totalLeads = subscribers.length;
  const dormantLeads = subscribers.filter((s) => s.isDormant).length;
  const activeLeads = totalLeads - dormantLeads;
  const totalUserReplies = subscribers.reduce((sum, s) => sum + s.userReplyCount, 0);
  const totalTwpReplies = subscribers.reduce((sum, s) => sum + s.twpReplyCount, 0);
  const withSequences = subscribers.filter((s) => s.assignedSequence !== null);
  const totalSequencesSent = withSequences.length;
  const leadsReactivatedAfterSequence = withSequences.filter((s) => s.postSequenceReplies > 0).length;
  const reactivationRate =
    totalSequencesSent > 0
      ? Math.round((leadsReactivatedAfterSequence / totalSequencesSent) * 100 * 100) / 100
      : 0;

  res.json({
    totalLeads,
    dormantLeads,
    activeLeads,
    totalUserReplies,
    totalTwpReplies,
    totalSequencesSent,
    leadsReactivatedAfterSequence,
    reactivationRate,
  });
});

router.get("/dashboard/agent-stats", async (req, res): Promise<void> => {
  const parsed = GetAgentStatsQueryParams.safeParse(req.query);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const { apiToken, phoneNumberId } = parsed.data;
  const subscribers = await getProcessedSubscribers(apiToken, phoneNumberId);

  const agentMap = new Map<string, { leads: ProcessedSubscriber[] }>();
  for (const sub of subscribers) {
    const agent = sub.assignedAgent ?? "Unassigned";
    if (!agentMap.has(agent)) agentMap.set(agent, { leads: [] });
    agentMap.get(agent)!.leads.push(sub);
  }

  const agents = Array.from(agentMap.entries()).map(([agentName, { leads }]) => {
    const activeLeads = leads.filter((l) => !l.isDormant).length;
    const dormantLeads = leads.filter((l) => l.isDormant).length;
    const avgUserReplies =
      leads.length > 0
        ? Math.round((leads.reduce((s, l) => s + l.userReplyCount, 0) / leads.length) * 100) / 100
        : 0;
    return {
      agentName,
      leadsAssigned: leads.length,
      activeLeads,
      dormantLeads,
      avgUserReplies,
    };
  });

  agents.sort((a, b) => b.leadsAssigned - a.leadsAssigned);
  res.json({ agents });
});

router.get("/dashboard/sequence-stats", async (req, res): Promise<void> => {
  const parsed = GetSequenceStatsQueryParams.safeParse(req.query);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const { apiToken, phoneNumberId } = parsed.data;
  const subscribers = await getProcessedSubscribers(apiToken, phoneNumberId);

  const seqMap = new Map<string, { leads: ProcessedSubscriber[] }>();
  for (const sub of subscribers) {
    if (!sub.assignedSequence) continue;
    if (!seqMap.has(sub.assignedSequence)) seqMap.set(sub.assignedSequence, { leads: [] });
    seqMap.get(sub.assignedSequence)!.leads.push(sub);
  }

  const sequences = Array.from(seqMap.entries()).map(([sequenceName, { leads }]) => {
    const repliesAfterSequence = leads.filter((l) => l.postSequenceReplies > 0).length;
    const reactivationRate =
      leads.length > 0
        ? Math.round((repliesAfterSequence / leads.length) * 100 * 100) / 100
        : 0;
    return {
      sequenceName,
      totalSent: leads.length,
      repliesAfterSequence,
      reactivationRate,
    };
  });

  sequences.sort((a, b) => b.reactivationRate - a.reactivationRate);
  res.json({ sequences });
});

router.get("/dashboard/button-stats", async (req, res): Promise<void> => {
  const parsed = GetLabelStatsQueryParams.safeParse(req.query);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const { apiToken, phoneNumberId } = parsed.data;
  const subscribers = await getProcessedSubscribers(apiToken, phoneNumberId);

  // Aggregate button clicks across all subscribers
  const buttonMap = new Map<string, {
    totalClicks: number;
    uniqueLeads: Map<string, { name: string; phoneNumber: string }>;
  }>();
  let totalClicks = 0;

  for (const sub of subscribers) {
    for (const click of sub.buttonClicks) {
      if (!buttonMap.has(click.buttonName)) {
        buttonMap.set(click.buttonName, { totalClicks: 0, uniqueLeads: new Map() });
      }
      const entry = buttonMap.get(click.buttonName)!;
      entry.totalClicks++;
      entry.uniqueLeads.set(sub.phoneNumber, { name: sub.name, phoneNumber: sub.phoneNumber });
      totalClicks++;
    }
  }

  const buttons = Array.from(buttonMap.entries())
    .map(([buttonName, { totalClicks: clicks, uniqueLeads }]) => ({
      buttonName,
      totalClicks: clicks,
      uniqueLeads: uniqueLeads.size,
      clickRate: totalClicks > 0 ? Math.round((clicks / totalClicks) * 10000) / 100 : 0,
      subscribers: Array.from(uniqueLeads.values()),
    }))
    .sort((a, b) => b.totalClicks - a.totalClicks);

  const mostClicked = buttons[0] ?? null;
  const leastClicked = buttons[buttons.length - 1] ?? null;
  const bestConversion = [...buttons].sort((a, b) => b.uniqueLeads - a.uniqueLeads)[0] ?? null;

  res.json({
    totalClicks,
    buttons,
    insights: {
      mostClicked: mostClicked?.buttonName ?? null,
      leastClicked: leastClicked?.buttonName ?? null,
      bestConversion: bestConversion?.buttonName ?? null,
    },
  });
});

router.get("/dashboard/label-stats", async (req, res): Promise<void> => {
  const parsed = GetLabelStatsQueryParams.safeParse(req.query);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const { apiToken, phoneNumberId } = parsed.data;
  const subscribers = await getProcessedSubscribers(apiToken, phoneNumberId);

  const labelMap = new Map<string, { leads: ProcessedSubscriber[] }>();
  for (const sub of subscribers) {
    const label = sub.labelName || "Unlabeled";
    if (!labelMap.has(label)) labelMap.set(label, { leads: [] });
    labelMap.get(label)!.leads.push(sub);
  }

  const labels = Array.from(labelMap.entries()).map(([labelName, { leads }]) => ({
    labelName,
    count: leads.length,
    dormantCount: leads.filter((l) => l.isDormant).length,
  }));

  labels.sort((a, b) => b.count - a.count);
  res.json({ labels });
});

router.get("/labels/list", async (req, res): Promise<void> => {
  const parsed = GetDashboardSummaryQueryParams.safeParse(req.query);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const { apiToken, phoneNumberId } = parsed.data;

  const url = new URL("https://growth.thewiseparrot.club/api/v1/whatsapp/label/list");
  url.searchParams.set("apiToken", apiToken);
  url.searchParams.set("phone_number_id", phoneNumberId);

  const twpRes = await fetch(url.toString(), { signal: AbortSignal.timeout(10_000) });
  const raw = await twpRes.json() as { status?: string; message?: unknown };

  if (!twpRes.ok || raw.status !== "1") {
    res.json({ labels: [] });
    return;
  }

  // TWP label list returns an array in message field
  const items = Array.isArray(raw.message) ? raw.message as Array<{ id?: unknown; name?: unknown; label_name?: unknown }> : [];
  const labels = items.map((item) => ({
    id: String(item.id ?? ""),
    name: String(item.name ?? item.label_name ?? ""),
  })).filter((l) => l.id && l.name);

  res.json({ labels });
});

async function createSubscriber(apiToken: string, phoneNumberId: string, phoneNumber: string, name: string): Promise<void> {
  const params = new URLSearchParams({
    apiToken,
    phoneNumberID: phoneNumberId,
    name,
    phoneNumber,
  });
  await fetch(
    "https://growth.thewiseparrot.club/api/v1/whatsapp/subscriber/create",
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: params.toString(), signal: AbortSignal.timeout(10_000) }
  );
}

router.post("/labels/bulk-assign", async (req, res): Promise<void> => {
  const { apiToken, phoneNumberId, labelId, phoneNumbers, names } = req.body as {
    apiToken?: string;
    phoneNumberId?: string;
    labelId?: string;
    phoneNumbers?: unknown;
    names?: unknown;
  };

  if (!apiToken || !phoneNumberId || !labelId?.trim() || !Array.isArray(phoneNumbers) || phoneNumbers.length === 0) {
    res.status(400).json({ error: "apiToken, phoneNumberId, labelId and phoneNumbers[] are required" });
    return;
  }

  const numbers = (phoneNumbers as unknown[]).map(String).filter(Boolean);
  const nameList = Array.isArray(names) ? (names as unknown[]).map(String) : [];
  const errors: { phone: string; reason: string }[] = [];
  let succeeded = 0;
  let created = 0;

  // Process in batches of 5 concurrent requests to avoid overwhelming the TWP API
  const BATCH = 5;
  for (let i = 0; i < numbers.length; i += BATCH) {
    const batch = numbers.slice(i, i + BATCH);
    await Promise.all(batch.map(async (phone, batchIdx) => {
      const globalIdx = i + batchIdx;
      const assignParams = new URLSearchParams({
        apiToken,
        phone_number_id: phoneNumberId,
        phone_number: phone,
        label_ids: labelId,
      });
      try {
        const r = await fetch(
          "https://growth.thewiseparrot.club/api/v1/whatsapp/subscriber/chat/assign-labels",
          { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: assignParams.toString(), signal: AbortSignal.timeout(10_000) }
        );
        const raw = await r.json() as { status?: string; message?: string };
        if (r.ok && raw.status === "1") {
          succeeded++;
        } else {
          // Subscriber not found — create them, then retry assign
          const subscriberName = (nameList[globalIdx] ?? "").trim() || phone;
          await createSubscriber(apiToken, phoneNumberId, phone, subscriberName);
          created++;

          const r2 = await fetch(
            "https://growth.thewiseparrot.club/api/v1/whatsapp/subscriber/chat/assign-labels",
            { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: assignParams.toString(), signal: AbortSignal.timeout(10_000) }
          );
          const raw2 = await r2.json() as { status?: string; message?: string };
          if (r2.ok && raw2.status === "1") {
            succeeded++;
          } else {
            errors.push({ phone, reason: raw2.message ?? `HTTP ${r2.status}` });
          }
        }
      } catch (err) {
        errors.push({ phone, reason: err instanceof Error ? err.message : "Unknown error" });
      }
    }));
  }

  res.json({ total: numbers.length, succeeded, created, failed: errors.length, errors });
});

router.post("/labels/assign-subscriber", async (req, res): Promise<void> => {
  const { apiToken, phoneNumberId, phoneNumber, labelIds, name } = req.body as {
    apiToken?: string;
    phoneNumberId?: string;
    phoneNumber?: string;
    labelIds?: string;
    name?: string;
  };

  if (!apiToken || !phoneNumberId || !phoneNumber?.trim() || !labelIds?.trim()) {
    res.status(400).json({ error: "apiToken, phoneNumberId, phoneNumber and labelIds are required" });
    return;
  }

  const phone = phoneNumber.trim();
  const params = new URLSearchParams({
    apiToken,
    phone_number_id: phoneNumberId,
    phone_number: phone,
    label_ids: labelIds.trim(),
  });

  const twpRes = await fetch(
    "https://growth.thewiseparrot.club/api/v1/whatsapp/subscriber/chat/assign-labels",
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: params.toString() }
  );

  const raw = await twpRes.json() as { status?: string; message?: string };

  if (!twpRes.ok || raw.status !== "1") {
    // Subscriber not found — create them, then retry assign
    const subscriberName = (name ?? "").trim() || phone;
    await createSubscriber(apiToken, phoneNumberId, phone, subscriberName);

    const twpRes2 = await fetch(
      "https://growth.thewiseparrot.club/api/v1/whatsapp/subscriber/chat/assign-labels",
      { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: params.toString() }
    );
    const raw2 = await twpRes2.json() as { status?: string; message?: string };

    if (!twpRes2.ok || raw2.status !== "1") {
      const msg = typeof raw2.message === "string" ? raw2.message : "Failed to assign subscriber";
      res.status(400).json({ success: false, message: msg });
      return;
    }

    res.json({ success: true, message: `New subscriber created and assigned successfully` });
    return;
  }

  res.json({ success: true, message: raw.message ?? "Subscriber assigned successfully" });
});

router.post("/labels/create", async (req, res): Promise<void> => {
  const { apiToken, phoneNumberId, labelName } = req.body as {
    apiToken?: string;
    phoneNumberId?: string;
    labelName?: string;
  };

  if (!apiToken || !phoneNumberId || !labelName?.trim()) {
    res.status(400).json({ error: "apiToken, phoneNumberId and labelName are required" });
    return;
  }

  const params = new URLSearchParams({
    apiToken,
    phone_number_id: phoneNumberId,
    label_name: labelName.trim(),
  });

  const twpRes = await fetch(
    "https://growth.thewiseparrot.club/api/v1/whatsapp/label/create",
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: params.toString() }
  );

  const raw = await twpRes.json() as { status?: string; message?: string };

  if (!twpRes.ok || raw.status !== "1") {
    const msg = typeof raw.message === "string" ? raw.message : "Failed to create label";
    res.status(400).json({ success: false, message: msg });
    return;
  }

  res.json({ success: true, message: raw.message ?? "Label created successfully" });
});

router.get("/account-info", async (req, res): Promise<void> => {
  const parsed = GetDashboardSummaryQueryParams.safeParse(req.query);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const { apiToken, phoneNumberId } = parsed.data;
  const info = await fetchAccountInfo(apiToken, phoneNumberId);
  res.json(info);
});

export default router;
