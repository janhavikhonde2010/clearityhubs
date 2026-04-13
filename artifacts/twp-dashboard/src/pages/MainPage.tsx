import { useState, useMemo, useEffect } from "react";
import { useQueryClient } from "@tanstack/react-query";
import {
  useGetDashboardSummary,
  useGetLabelStats,
  useGetAgentStats,
  useGetSequenceStats,
  useGetSubscribers,
  useGetButtonStats,
  useGetAccountInfo,
  useCreateLabel,
  useGetLabelList,
  useAssignSubscriberToLabel,
  useBulkAssignSubscribersToLabel,
  useGetTemplateList,
  useSendTemplateToLabel,
} from "@workspace/api-client-react";
import * as XLSX from "xlsx";
import { useCredentials } from "@/contexts/CredentialsContext";
import { CredentialsPrompt } from "@/components/layout/CredentialsPrompt";
import { Skeleton } from "@/components/ui/skeleton";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip as RechartsTooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend,
} from "recharts";
import {
  Users, Activity, Moon, TrendingUp, MessageSquare,
  Headset, AlertTriangle, Search, LogOut,
  MousePointerClick, Trophy, Minus, Zap, Flame,
  RefreshCw, Download, ArrowLeft, ChevronLeft, ChevronRight,
  BarChart2, Award, Building2, Tag, Plus, CheckCircle2, XCircle,
  Upload, FileSpreadsheet, X, AlertCircle, Send, Eye,
} from "lucide-react";
import { format } from "date-fns";

/* ── Brand colours ── */
const BLUE    = "#2563EB";
const BLUE_L  = "#3B82F6";
const BLUE_D  = "#1E3A8A";
const ORANGE  = "#F97316";
const ORANGE_L = "#FB923C";
const ORANGE_D = "#EA580C";
const GREEN   = "#22C55E";
const GREEN_D = "#16A34A";
const RED     = "#EF4444";
const GRAY    = "#6B7280";
const BORDER  = "#E5E7EB";
const PAGE_SIZE = 5;

/* ── CSV utility ── */
function downloadCSV(filename: string, rows: Record<string, unknown>[]) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const csv = [headers.join(","), ...rows.map((r) => headers.map((h) => JSON.stringify(r[h] ?? "")).join(","))].join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

/* ── Pagination hook ── */
function usePagination<T>(items: T[], pageSize = PAGE_SIZE) {
  const [page, setPage] = useState(0);
  const totalPages = Math.max(1, Math.ceil(items.length / pageSize));
  const from = page * pageSize;
  const to = Math.min(from + pageSize, items.length);
  const slice = items.slice(from, to);
  useEffect(() => { setPage(0); }, [items.length]);
  return {
    slice, page, totalPages,
    from: items.length === 0 ? 0 : from + 1,
    to, total: items.length,
    prevPage: () => setPage((p) => Math.max(0, p - 1)),
    nextPage: () => setPage((p) => Math.min(totalPages - 1, p + 1)),
    resetPage: () => setPage(0),
  };
}

/* ── Custom chart tooltip ── */
function ChTooltip({ active, payload, label, pct }: {
  active?: boolean; payload?: Array<{ name: string; value: number; fill?: string; color?: string }>;
  label?: string; pct?: boolean;
}) {
  if (!active || !payload?.length) return null;
  return (
    <div className="ch-card px-4 py-3 text-sm" style={{ border: `1px solid ${BORDER}`, minWidth: 140 }}>
      {label && <p className="text-xs font-semibold mb-1.5" style={{ color: GRAY }}>{label}</p>}
      {payload.map((p, i) => (
        <div key={i} className="flex items-center gap-2 py-0.5">
          <span className="w-2.5 h-2.5 rounded-sm shrink-0" style={{ background: p.fill || p.color }} />
          <span style={{ color: GRAY }}>{p.name}:</span>
          <span className="font-bold" style={{ color: "#111827" }}>
            {pct ? `${Number(p.value).toFixed(1)}%` : Number(p.value).toLocaleString()}
          </span>
        </div>
      ))}
    </div>
  );
}

/* ═══════════════ MAIN PAGE ═══════════════ */
export default function MainPage() {
  const { credentials, isConfigured, clearCredentials } = useCredentials();
  const queryClient = useQueryClient();
  const [refreshing, setRefreshing] = useState(false);

  const params = {
    apiToken: credentials?.apiToken ?? "",
    phoneNumberId: credentials?.phoneNumberId ?? "",
  };
  const queryOpts = { query: { enabled: isConfigured } };
  const { data: accountData } = useGetAccountInfo(params, queryOpts);
  const accountLabel = accountData?.businessName || `ID: ${credentials?.phoneNumberId ?? ""}`;

  if (!isConfigured) return <CredentialsPrompt />;

  const handleRefresh = async () => {
    setRefreshing(true);
    await queryClient.invalidateQueries();
    await queryClient.refetchQueries();
    setTimeout(() => setRefreshing(false), 800);
  };

  return (
    <div className="min-h-screen" style={{ background: "#F8FAFC" }}>

      {/* ── Header ── */}
      <header className="sticky top-0 z-50 bg-white border-b" style={{ borderColor: BORDER, boxShadow: "0 1px 3px rgba(0,0,0,0.04)" }}>
        <div className="max-w-screen-2xl mx-auto px-5 h-16 flex items-center gap-3">
          {/* Logo */}
          <div className="flex items-center gap-2.5 shrink-0">
            <div className="w-9 h-9 rounded-xl flex items-center justify-center shrink-0"
              style={{ background: "#EFF6FF", border: "1.5px solid #DBEAFE" }}>
              <img src="/logo.png" alt="Clarity Hub" className="w-6 h-6 object-contain" />
            </div>
            <span className="font-bold text-base tracking-tight" style={{ color: "#111827" }}>Clarity Hub</span>
          </div>

          {/* Account name badge */}
          {accountData?.businessName ? (
            <div className="ch-badge-blue hidden sm:inline-flex items-center gap-1">
              <Building2 className="w-3 h-3" />
              {accountData.businessName}
            </div>
          ) : (
            <div className="ch-badge-blue hidden sm:inline-flex">{accountLabel}</div>
          )}

          <div className="flex-1" />

          {/* Actions */}
          <a
            href="https://growth.thewiseparrot.club/dashboard"
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors"
            style={{ background: "#EFF6FF", color: BLUE, border: "1px solid #DBEAFE" }}
            onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.background = "#DBEAFE"; }}
            onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.background = "#EFF6FF"; }}
          >
            <ArrowLeft className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Dashboard</span>
          </a>

          <button onClick={handleRefresh}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors"
            style={{ background: "#F9FAFB", color: GRAY, border: `1px solid ${BORDER}` }}>
            <RefreshCw className={`w-3.5 h-3.5 ${refreshing ? "animate-spin" : ""}`} />
            <span className="hidden sm:inline">Refresh</span>
          </button>

          <button onClick={clearCredentials}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors"
            style={{ color: RED, border: "1px solid #FECACA", background: "#FEF2F2" }}
            onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.background = "#FEE2E2"; }}
            onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.background = "#FEF2F2"; }}>
            <LogOut className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Disconnect</span>
          </button>
        </div>
      </header>

      <main className="max-w-screen-2xl mx-auto px-4 md:px-6 py-8 space-y-10">
        <DashboardContent />
      </main>

      <div className="text-center py-5 text-xs" style={{ color: "#D1D5DB" }}>
        Clarity Hub · WhatsApp Lead Analytics · Powered by The Wise Parrot
      </div>
    </div>
  );
}

/* ═══════════════ LABEL MANAGEMENT CARD ═══════════════ */
function LabelManagementCard({ apiToken, phoneNumberId }: { apiToken: string; phoneNumberId: string }) {
  const [tab, setTab] = useState<"create" | "assign">("create");
  const [assignMode, setAssignMode] = useState<"single" | "bulk">("single");

  // ─── Create Label state ───
  const [labelName, setLabelName] = useState("");
  const [createFeedback, setCreateFeedback] = useState<{ ok: boolean; msg: string } | null>(null);
  const queryClient = useQueryClient();

  const { mutate: createLabel, isPending: creating } = useCreateLabel({
    mutation: {
      onSuccess: (data) => {
        setCreateFeedback({ ok: data.success, msg: data.message });
        if (data.success) {
          setLabelName("");
          void queryClient.invalidateQueries({ queryKey: ["/api/dashboard/label-stats"] });
          void queryClient.invalidateQueries({ queryKey: ["/api/labels/list"] });
        }
        setTimeout(() => setCreateFeedback(null), 5000);
      },
      onError: () => {
        setCreateFeedback({ ok: false, msg: "Failed to create label. Please check your credentials." });
        setTimeout(() => setCreateFeedback(null), 5000);
      },
    },
  });

  // ─── Shared label list ───
  const { data: labelListData, isLoading: loadingLabels } = useGetLabelList(
    { apiToken, phoneNumberId },
    { query: { enabled: tab === "assign", retry: 0 } }
  );
  const labelList = labelListData?.labels ?? [];

  // ─── Single assign state ───
  const [phone, setPhone] = useState("");
  const [subscriberName, setSubscriberName] = useState("");
  const [selectedLabelId, setSelectedLabelId] = useState("");
  const [assignFeedback, setAssignFeedback] = useState<{ ok: boolean; msg: string } | null>(null);

  const { mutate: assignSubscriber, isPending: assigning } = useAssignSubscriberToLabel({
    mutation: {
      onSuccess: (data) => {
        setAssignFeedback({ ok: data.success, msg: data.message });
        if (data.success) { setPhone(""); }
        setTimeout(() => setAssignFeedback(null), 5000);
      },
      onError: () => {
        setAssignFeedback({ ok: false, msg: "Failed to assign subscriber. Please try again." });
        setTimeout(() => setAssignFeedback(null), 5000);
      },
    },
  });

  // ─── Bulk assign (Excel) state ───
  const [bulkLabelId, setBulkLabelId] = useState("");
  const [parsedPhones, setParsedPhones] = useState<string[]>([]);
  const [parsedNames, setParsedNames] = useState<string[]>([]);
  const [parseError, setParseError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [bulkResult, setBulkResult] = useState<{ total: number; succeeded: number; created: number; failed: number; errors: { phone: string; reason: string }[] } | null>(null);

  const { mutate: bulkAssign, isPending: bulkAssigning } = useBulkAssignSubscribersToLabel({
    mutation: {
      onSuccess: (data) => { setBulkResult(data); },
      onError: () => { setBulkResult(null); setParseError("Bulk assignment failed. Please try again."); },
    },
  });

  const handleFileUpload = (file: File) => {
    setParseError(null);
    setParsedPhones([]);
    setParsedNames([]);
    setBulkResult(null);
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target!.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { header: 1 }) as unknown[][];

        // Extract phone numbers: look for a column named "phone" / "phone_number" / "mobile" or just take first column
        const header = rows[0]?.map((h) => String(h ?? "").toLowerCase().trim()) ?? [];
        const phoneColIdx = header.findIndex((h) => h.includes("phone") || h.includes("mobile") || h.includes("number"));
        const colIdx = phoneColIdx >= 0 ? phoneColIdx : 0;

        // Also extract names if a "name" column exists
        const nameColIdx = header.findIndex((h) => h === "name" || h === "full name" || h === "fullname" || h === "subscriber name");

        const dataRows = rows.slice(phoneColIdx >= 0 ? 1 : 0);

        const phones: string[] = [];
        const names: string[] = [];

        for (const row of dataRows) {
          const phone = String((row as unknown[])[colIdx] ?? "").replace(/\D/g, "").trim();
          if (phone.length >= 7) {
            phones.push(phone);
            const name = nameColIdx >= 0 ? String((row as unknown[])[nameColIdx] ?? "").trim() : "";
            names.push(name);
          }
        }

        if (phones.length === 0) {
          setParseError("No valid phone numbers found. Make sure your Excel has a column named 'phone' or 'phone_number'.");
          return;
        }
        setParsedPhones(phones);
        setParsedNames(names);
      } catch {
        setParseError("Could not read the file. Please upload a valid .xlsx or .xls file.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const inputCls = "flex-1 rounded-xl border px-4 py-2.5 text-sm outline-none";
  const inputStyle = { borderColor: BORDER, color: "#111827", background: "#F9FAFB" };
  const selectStyle = { ...inputStyle, minWidth: 180, flex: "0 0 auto" };

  const LabelSelect = ({ value, onChange }: { value: string; onChange: (v: string) => void }) => (
    <select value={value} onChange={(e) => onChange(e.target.value)}
      className="rounded-xl border px-3 py-2.5 text-sm outline-none" style={selectStyle}
      onFocus={(e) => { e.currentTarget.style.borderColor = BLUE; e.currentTarget.style.background = "#fff"; }}
      onBlur={(e)  => { e.currentTarget.style.borderColor = BORDER; e.currentTarget.style.background = "#F9FAFB"; }}>
      <option value="">{loadingLabels ? "Loading…" : labelList.length === 0 ? "No labels found" : "Select a label"}</option>
      {labelList.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
    </select>
  );

  return (
    <div className="rounded-2xl border mb-5 overflow-hidden" style={{ background: "#fff", borderColor: BORDER }}>
      {/* Main tab bar */}
      <div className="flex border-b" style={{ borderColor: BORDER }}>
        {(["create", "assign"] as const).map((t) => {
          const active = tab === t;
          return (
            <button key={t} onClick={() => setTab(t)}
              className="flex items-center gap-2 px-5 py-3.5 text-sm font-semibold transition-colors"
              style={{ color: active ? ORANGE_D : GRAY, borderBottom: active ? `2px solid ${ORANGE}` : "2px solid transparent", background: active ? "#FFFBF5" : "transparent" }}>
              {t === "create" ? <><Plus className="w-3.5 h-3.5" />Create Label</> : <><Users className="w-3.5 h-3.5" />Assign Subscriber</>}
            </button>
          );
        })}
      </div>

      <div className="p-5">
        {tab === "create" ? (
          <>
            <p className="text-xs mb-3" style={{ color: GRAY }}>Add a new pipeline stage label to your WhatsApp account.</p>
            <form onSubmit={(e) => { e.preventDefault(); if (!labelName.trim() || creating) return; setCreateFeedback(null); createLabel({ data: { apiToken, phoneNumberId, labelName: labelName.trim() } }); }}
              className="flex items-center gap-3">
              <input type="text" value={labelName} onChange={(e) => setLabelName(e.target.value)}
                placeholder="e.g. Hot Lead, Closed, Follow-up…" maxLength={80}
                className={inputCls} style={inputStyle}
                onFocus={(e) => { e.currentTarget.style.borderColor = ORANGE; e.currentTarget.style.background = "#fff"; }}
                onBlur={(e)  => { e.currentTarget.style.borderColor = BORDER; e.currentTarget.style.background = "#F9FAFB"; }} />
              <button type="submit" disabled={!labelName.trim() || creating}
                className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl font-semibold text-sm disabled:opacity-50"
                style={{ background: ORANGE, color: "#fff", minWidth: 120 }}>
                {creating ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Plus className="w-4 h-4" />}
                {creating ? "Creating…" : "Create Label"}
              </button>
            </form>
            {createFeedback && (
              <div className="mt-3 flex items-center gap-2 rounded-xl px-4 py-2.5 text-sm font-medium"
                style={{ background: createFeedback.ok ? "#F0FDF4" : "#FEF2F2", color: createFeedback.ok ? GREEN_D : RED, border: `1px solid ${createFeedback.ok ? "#BBF7D0" : "#FECACA"}` }}>
                {createFeedback.ok ? <CheckCircle2 className="w-4 h-4 shrink-0" /> : <XCircle className="w-4 h-4 shrink-0" />}
                {createFeedback.msg}
              </div>
            )}
          </>
        ) : (
          <>
            {/* Sub-mode toggle */}
            <div className="flex gap-1 mb-4 p-1 rounded-xl w-fit" style={{ background: "#F3F4F6" }}>
              {(["single", "bulk"] as const).map((m) => (
                <button key={m} onClick={() => { setAssignMode(m); setBulkResult(null); setParseError(null); setParsedPhones([]); setParsedNames([]); setFileName(null); }}
                  className="flex items-center gap-1.5 px-4 py-1.5 rounded-lg text-xs font-semibold transition-all"
                  style={{ background: assignMode === m ? "#fff" : "transparent", color: assignMode === m ? "#111827" : GRAY, boxShadow: assignMode === m ? "0 1px 3px rgba(0,0,0,.10)" : "none" }}>
                  {m === "single" ? <><Users className="w-3 h-3" />Single</> : <><FileSpreadsheet className="w-3 h-3" />Import Excel</>}
                </button>
              ))}
            </div>

            {assignMode === "single" ? (
              <>
                <div className="flex flex-col gap-2">
                  <div className="flex items-center gap-3">
                    <input type="text" value={phone} onChange={(e) => setPhone(e.target.value)}
                      placeholder="Phone number (e.g. 919876543210)" maxLength={20}
                      className={inputCls} style={inputStyle}
                      onFocus={(e) => { e.currentTarget.style.borderColor = BLUE; e.currentTarget.style.background = "#fff"; }}
                      onBlur={(e)  => { e.currentTarget.style.borderColor = BORDER; e.currentTarget.style.background = "#F9FAFB"; }} />
                    <input type="text" value={subscriberName} onChange={(e) => setSubscriberName(e.target.value)}
                      placeholder="Name (for new subscribers)" maxLength={80}
                      className={inputCls} style={inputStyle}
                      onFocus={(e) => { e.currentTarget.style.borderColor = BLUE; e.currentTarget.style.background = "#fff"; }}
                      onBlur={(e)  => { e.currentTarget.style.borderColor = BORDER; e.currentTarget.style.background = "#F9FAFB"; }} />
                  </div>
                  <div className="flex items-center gap-3">
                    <LabelSelect value={selectedLabelId} onChange={setSelectedLabelId} />
                    <button onClick={() => { if (!phone.trim() || !selectedLabelId || assigning) return; setAssignFeedback(null); assignSubscriber({ data: { apiToken, phoneNumberId, phoneNumber: phone.trim(), labelIds: selectedLabelId, name: subscriberName.trim() || undefined } }); }}
                      disabled={!phone.trim() || !selectedLabelId || assigning}
                      className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl font-semibold text-sm disabled:opacity-50"
                      style={{ background: BLUE, color: "#fff", minWidth: 110 }}>
                      {assigning ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Users className="w-4 h-4" />}
                      {assigning ? "Assigning…" : "Assign"}
                    </button>
                  </div>
                </div>
                <p className="text-xs mt-1.5" style={{ color: GRAY }}>If the subscriber doesn't exist yet, they'll be created automatically and added to the label.</p>
                {assignFeedback && (
                  <div className="mt-3 flex items-center gap-2 rounded-xl px-4 py-2.5 text-sm font-medium"
                    style={{ background: assignFeedback.ok ? "#F0FDF4" : "#FEF2F2", color: assignFeedback.ok ? GREEN_D : RED, border: `1px solid ${assignFeedback.ok ? "#BBF7D0" : "#FECACA"}` }}>
                    {assignFeedback.ok ? <CheckCircle2 className="w-4 h-4 shrink-0" /> : <XCircle className="w-4 h-4 shrink-0" />}
                    {assignFeedback.msg}
                  </div>
                )}
              </>
            ) : (
              <>
                {/* Excel Upload Area */}
                {!parsedPhones.length ? (
                  <label
                    className="flex flex-col items-center justify-center gap-3 rounded-2xl border-2 border-dashed py-8 cursor-pointer transition-colors"
                    style={{ borderColor: BORDER, background: "#FAFAFA" }}
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleFileUpload(f); }}
                  >
                    <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{ background: "#EFF6FF" }}>
                      <Upload className="w-5 h-5" style={{ color: BLUE }} />
                    </div>
                    <div className="text-center">
                      <p className="text-sm font-semibold" style={{ color: "#111827" }}>Drop your Excel file here</p>
                      <p className="text-xs mt-0.5" style={{ color: GRAY }}>or click to browse — .xlsx / .xls supported</p>
                      <p className="text-xs mt-1" style={{ color: GRAY }}>Columns: <span className="font-mono" style={{ color: BLUE }}>phone</span>/<span className="font-mono" style={{ color: BLUE }}>phone_number</span> (required) · <span className="font-mono" style={{ color: BLUE }}>name</span> (optional, for new subscribers)</p>
                    </div>
                    <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFileUpload(f); }} />
                  </label>
                ) : (
                  <div className="rounded-2xl border p-4" style={{ borderColor: BORDER, background: "#F9FAFB" }}>
                    <div className="flex items-center justify-between mb-3">
                      <div className="flex items-center gap-2">
                        <FileSpreadsheet className="w-4 h-4" style={{ color: GREEN_D }} />
                        <span className="text-sm font-semibold" style={{ color: "#111827" }}>{fileName}</span>
                        <span className="text-xs px-2 py-0.5 rounded-full font-semibold" style={{ background: "#DCFCE7", color: GREEN_D }}>
                          {parsedPhones.length} numbers found
                        </span>
                      </div>
                      <button onClick={() => { setParsedPhones([]); setParsedNames([]); setFileName(null); setBulkResult(null); setParseError(null); }}
                        className="w-6 h-6 rounded-lg flex items-center justify-center hover:bg-gray-200 transition-colors">
                        <X className="w-3.5 h-3.5" style={{ color: GRAY }} />
                      </button>
                    </div>
                    {/* Phone number preview */}
                    <div className="rounded-xl border overflow-hidden mb-3" style={{ borderColor: BORDER }}>
                      <div className="max-h-28 overflow-y-auto">
                        <table className="w-full text-xs">
                          <thead style={{ background: "#F3F4F6" }}>
                            <tr>
                              <th className="px-3 py-2 text-left font-semibold" style={{ color: GRAY }}>#</th>
                              <th className="px-3 py-2 text-left font-semibold" style={{ color: GRAY }}>Phone Number</th>
                              {parsedNames.some(n => n) && <th className="px-3 py-2 text-left font-semibold" style={{ color: GRAY }}>Name</th>}
                            </tr>
                          </thead>
                          <tbody>
                            {parsedPhones.slice(0, 5).map((p, i) => (
                              <tr key={i} style={{ borderTop: `1px solid ${BORDER}` }}>
                                <td className="px-3 py-1.5" style={{ color: GRAY }}>{i + 1}</td>
                                <td className="px-3 py-1.5 font-mono" style={{ color: "#111827" }}>{p}</td>
                                {parsedNames.some(n => n) && <td className="px-3 py-1.5" style={{ color: "#374151" }}>{parsedNames[i] || <span style={{ color: GRAY, fontStyle: "italic" }}>—</span>}</td>}
                              </tr>
                            ))}
                            {parsedPhones.length > 5 && (
                              <tr style={{ borderTop: `1px solid ${BORDER}` }}>
                                <td colSpan={parsedNames.some(n => n) ? 3 : 2} className="px-3 py-1.5 text-center" style={{ color: GRAY }}>
                                  …and {parsedPhones.length - 5} more
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                )}

                {parseError && (
                  <div className="mt-3 flex items-center gap-2 rounded-xl px-4 py-2.5 text-sm font-medium"
                    style={{ background: "#FEF2F2", color: RED, border: `1px solid #FECACA` }}>
                    <AlertCircle className="w-4 h-4 shrink-0" />{parseError}
                  </div>
                )}

                {/* Bulk assign controls */}
                {parsedPhones.length > 0 && !bulkResult && (
                  <div className="flex items-center gap-3 mt-3">
                    <LabelSelect value={bulkLabelId} onChange={setBulkLabelId} />
                    <button
                      disabled={!bulkLabelId || bulkAssigning}
                      onClick={() => { if (!bulkLabelId || bulkAssigning) return; setBulkResult(null); bulkAssign({ data: { apiToken, phoneNumberId, labelId: bulkLabelId, phoneNumbers: parsedPhones, names: parsedNames.length > 0 ? parsedNames : undefined } }); }}
                      className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl font-semibold text-sm disabled:opacity-50"
                      style={{ background: GREEN_D, color: "#fff", minWidth: 160 }}>
                      {bulkAssigning ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Upload className="w-4 h-4" />}
                      {bulkAssigning ? `Assigning ${parsedPhones.length}…` : `Assign ${parsedPhones.length} Subscribers`}
                    </button>
                  </div>
                )}

                {/* Bulk result summary */}
                {bulkResult && (
                  <div className="mt-3 rounded-2xl border p-4" style={{ borderColor: BORDER }}>
                    <p className="text-sm font-semibold mb-3" style={{ color: "#111827" }}>Bulk Assignment Complete</p>
                    <div className="grid grid-cols-4 gap-3 mb-3">
                      {[
                        { label: "Total", val: bulkResult.total, bg: "#F3F4F6",   color: "#374151" },
                        { label: "Assigned", val: bulkResult.succeeded, bg: "#DCFCE7", color: GREEN_D },
                        { label: "New Created", val: bulkResult.created ?? 0, bg: "#EFF6FF", color: BLUE },
                        { label: "Failed", val: bulkResult.failed,    bg: "#FEE2E2", color: RED       },
                      ].map(({ label, val, bg, color }) => (
                        <div key={label} className="rounded-xl p-3 text-center" style={{ background: bg }}>
                          <p className="text-lg font-bold" style={{ color }}>{val}</p>
                          <p className="text-xs font-medium" style={{ color }}>{label}</p>
                        </div>
                      ))}
                    </div>
                    {bulkResult.errors.length > 0 && (
                      <div className="rounded-xl border overflow-hidden" style={{ borderColor: BORDER }}>
                        <div className="px-3 py-2 text-xs font-semibold" style={{ background: "#FEF2F2", color: RED }}>Failed Numbers</div>
                        <div className="max-h-32 overflow-y-auto">
                          {bulkResult.errors.map(({ phone: p, reason }, i) => (
                            <div key={i} className="flex items-center gap-2 px-3 py-1.5 text-xs border-t" style={{ borderColor: BORDER }}>
                              <span className="font-mono font-semibold" style={{ color: "#111827" }}>{p}</span>
                              <span style={{ color: GRAY }}>—</span>
                              <span style={{ color: RED }}>{reason}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                    <button onClick={() => { setBulkResult(null); setParsedPhones([]); setParsedNames([]); setFileName(null); setBulkLabelId(""); }}
                      className="mt-3 text-xs font-semibold" style={{ color: BLUE }}>
                      Import another file
                    </button>
                  </div>
                )}
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}

/* ═══════════════ MESSAGE BROADCAST CARD ═══════════════ */
function MessageBroadcastCard({ apiToken, phoneNumberId }: { apiToken: string; phoneNumberId: string }) {
  const [selectedLabelName, setSelectedLabelName] = useState("");
  const [selectedTemplate, setSelectedTemplate] = useState<{ id: string; name: string; message: string } | null>(null);
  const [customMessage, setCustomMessage] = useState("");
  const [useCustom, setUseCustom] = useState(false);
  const [confirmed, setConfirmed] = useState(false);
  const [sendResult, setSendResult] = useState<{ total: number; succeeded: number; failed: number; errors: { phone: string; reason: string }[] } | null>(null);
  const [showErrors, setShowErrors] = useState(false);

  const { data: labelListData, isLoading: loadingLabels } = useGetLabelList(
    { apiToken, phoneNumberId },
    { query: { enabled: true, retry: 0 } }
  );
  const labelList = labelListData?.labels ?? [];

  const { data: labelStatsData } = useGetLabelStats(
    { apiToken, phoneNumberId },
    { query: { enabled: true, retry: 0 } }
  );
  const labelStats = labelStatsData?.labels ?? [];
  const selectedLabelStat = labelStats.find((l) => l.labelName === selectedLabelName);

  const { mutate: fetchTemplates, data: templatesData, isPending: loadingTemplates } = useGetTemplateList();
  const templates = templatesData?.templates ?? [];

  const { mutate: doSend, isPending: sending } = useSendTemplateToLabel({
    mutation: {
      onSuccess: (data) => {
        setSendResult(data);
        setConfirmed(false);
      },
    },
  });

  useEffect(() => {
    fetchTemplates({ data: { apiToken, phoneNumberId } });
  }, [apiToken, phoneNumberId]);

  const messageToSend = useCustom ? customMessage.trim() : (selectedTemplate?.message ?? "");
  const canSend = selectedLabelName && messageToSend && !sending;

  function handleSend() {
    if (!canSend) return;
    setSendResult(null);
    setShowErrors(false);
    doSend({ data: { apiToken, phoneNumberId, labelName: selectedLabelName, message: messageToSend } });
  }

  return (
    <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6 mb-6">
      <div className="flex items-center gap-3 mb-5">
        <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{ background: "#F0FDF4" }}>
          <Send size={18} style={{ color: "#16A34A" }} />
        </div>
        <div>
          <h3 className="font-semibold text-gray-900 text-base">Message Distribution</h3>
          <p className="text-xs text-gray-500 mt-0.5">Broadcast a message to all subscribers in a label</p>
        </div>
        <button
          onClick={() => fetchTemplates({ data: { apiToken, phoneNumberId } })}
          disabled={loadingTemplates}
          className="ml-auto text-xs text-blue-600 hover:text-blue-800 flex items-center gap-1 disabled:opacity-50"
        >
          <RefreshCw size={12} className={loadingTemplates ? "animate-spin" : ""} />
          Refresh
        </button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
        {/* Label select */}
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1.5">Target Label</label>
          {loadingLabels ? (
            <div className="h-9 bg-gray-100 rounded-lg animate-pulse" />
          ) : (
            <select
              value={selectedLabelName}
              onChange={(e) => { setSelectedLabelName(e.target.value); setSendResult(null); setConfirmed(false); }}
              className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-200"
            >
              <option value="">— Select a label —</option>
              {labelList.map((l) => {
                const stat = labelStats.find((s) => s.labelName === l.name);
                const count = stat?.count ?? "?";
                return (
                  <option key={l.id} value={l.name}>
                    {l.name} ({count} subscribers)
                  </option>
                );
              })}
            </select>
          )}
          {selectedLabelStat && (
            <p className="text-xs text-gray-400 mt-1">
              {selectedLabelStat.count} active · {selectedLabelStat.dormantCount} dormant
            </p>
          )}
        </div>

        {/* Template select */}
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1.5">Message Template</label>
          {loadingTemplates ? (
            <div className="h-9 bg-gray-100 rounded-lg animate-pulse" />
          ) : (
            <select
              value={selectedTemplate?.id ?? ""}
              onChange={(e) => {
                if (e.target.value === "__custom__") {
                  setSelectedTemplate(null);
                  setUseCustom(true);
                  setSendResult(null);
                  setConfirmed(false);
                } else {
                  const t = templates.find((t) => t.id === e.target.value) ?? null;
                  setSelectedTemplate(t);
                  setUseCustom(false);
                  setSendResult(null);
                  setConfirmed(false);
                }
              }}
              className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-200"
            >
              <option value="">— Select a template —</option>
              {templates.map((t) => (
                <option key={t.id} value={t.id}>{t.name}</option>
              ))}
              <option value="__custom__">✏️ Custom message…</option>
            </select>
          )}
          {templates.length === 0 && !loadingTemplates && (
            <p className="text-xs text-amber-500 mt-1">No templates found — you can still type a custom message.</p>
          )}
        </div>
      </div>

      {/* Custom message toggle */}
      {(selectedTemplate?.id === "__custom__" || useCustom) ? null : (
        <button
          className="text-xs text-blue-500 hover:text-blue-700 mb-2 underline"
          onClick={() => { setUseCustom(true); setSelectedTemplate(null); }}
        >
          Or type a custom message instead
        </button>
      )}

      {/* Message preview / custom input */}
      {(selectedTemplate && !useCustom) && (
        <div className="mb-4">
          <div className="flex items-center gap-1.5 text-xs font-medium text-gray-600 mb-1.5">
            <Eye size={12} /> Message Preview
          </div>
          <div className="bg-gray-50 border border-gray-200 rounded-xl p-3 text-sm text-gray-700 whitespace-pre-wrap leading-relaxed max-h-36 overflow-y-auto">
            {selectedTemplate.message || <span className="text-gray-400 italic">No preview available</span>}
          </div>
        </div>
      )}

      {useCustom && (
        <div className="mb-4">
          <label className="block text-xs font-medium text-gray-600 mb-1.5">Custom Message</label>
          <textarea
            value={customMessage}
            onChange={(e) => setCustomMessage(e.target.value)}
            rows={4}
            placeholder="Type your message here…"
            className="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-200 resize-none"
          />
          <button className="text-xs text-gray-400 hover:text-gray-600 mt-1 underline" onClick={() => { setUseCustom(false); setCustomMessage(""); }}>
            Cancel custom message
          </button>
        </div>
      )}

      {/* Confirm & Send */}
      {!sendResult && (
        <div className="flex items-center gap-3 flex-wrap">
          {!confirmed ? (
            <button
              disabled={!canSend}
              onClick={() => setConfirmed(true)}
              className="flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-medium text-white disabled:opacity-40 disabled:cursor-not-allowed transition-opacity"
              style={{ background: canSend ? "#16A34A" : "#9CA3AF" }}
            >
              <Send size={14} />
              {selectedLabelStat ? `Send to ${selectedLabelStat.count} subscribers` : "Send to label"}
            </button>
          ) : (
            <div className="flex items-center gap-2">
              <span className="text-sm text-gray-600">
                Send <strong>{useCustom ? "custom message" : (selectedTemplate?.name ?? "message")}</strong> to <strong>{selectedLabelStat?.count ?? "all"}</strong> subscribers in <strong>{selectedLabelName}</strong>?
              </span>
              <button
                onClick={handleSend}
                disabled={sending}
                className="px-3 py-1.5 rounded-lg text-sm font-medium text-white flex items-center gap-1.5 disabled:opacity-50"
                style={{ background: "#16A34A" }}
              >
                {sending ? <RefreshCw size={12} className="animate-spin" /> : <CheckCircle2 size={14} />}
                {sending ? "Sending…" : "Confirm"}
              </button>
              <button
                onClick={() => setConfirmed(false)}
                className="px-3 py-1.5 rounded-lg text-sm font-medium text-gray-600 border border-gray-200 hover:bg-gray-50"
              >
                Cancel
              </button>
            </div>
          )}
        </div>
      )}

      {/* Results */}
      {sendResult && (
        <div className="mt-4 rounded-xl border p-4" style={{ borderColor: sendResult.failed === 0 ? "#BBF7D0" : "#FED7AA", background: sendResult.failed === 0 ? "#F0FDF4" : "#FFF7ED" }}>
          <div className="flex items-center gap-2 mb-2">
            {sendResult.failed === 0 ? (
              <CheckCircle2 size={16} style={{ color: "#16A34A" }} />
            ) : (
              <AlertCircle size={16} style={{ color: "#EA580C" }} />
            )}
            <span className="text-sm font-semibold" style={{ color: sendResult.failed === 0 ? "#15803D" : "#9A3412" }}>
              Broadcast {sendResult.failed === 0 ? "complete" : "finished with errors"}
            </span>
          </div>
          <div className="grid grid-cols-3 gap-3 text-center mb-3">
            <div className="bg-white rounded-lg py-2">
              <p className="text-lg font-bold text-gray-800">{sendResult.total}</p>
              <p className="text-xs text-gray-500">Total</p>
            </div>
            <div className="bg-white rounded-lg py-2">
              <p className="text-lg font-bold" style={{ color: "#16A34A" }}>{sendResult.succeeded}</p>
              <p className="text-xs text-gray-500">Sent</p>
            </div>
            <div className="bg-white rounded-lg py-2">
              <p className="text-lg font-bold" style={{ color: sendResult.failed > 0 ? "#EA580C" : "#9CA3AF" }}>{sendResult.failed}</p>
              <p className="text-xs text-gray-500">Failed</p>
            </div>
          </div>
          {sendResult.failed > 0 && (
            <button
              className="text-xs text-orange-600 underline hover:text-orange-800"
              onClick={() => setShowErrors(!showErrors)}
            >
              {showErrors ? "Hide" : "Show"} {sendResult.failed} failed numbers
            </button>
          )}
          {showErrors && sendResult.errors.length > 0 && (
            <div className="mt-2 space-y-1 max-h-32 overflow-y-auto">
              {sendResult.errors.map((e, i) => (
                <div key={i} className="text-xs text-gray-600 flex gap-2">
                  <span className="font-mono text-orange-700">{e.phone}</span>
                  <span className="text-gray-400">—</span>
                  <span>{e.reason}</span>
                </div>
              ))}
            </div>
          )}
          <button
            className="mt-3 text-xs text-blue-600 hover:text-blue-800 underline"
            onClick={() => { setSendResult(null); setConfirmed(false); setSelectedLabelName(""); setSelectedTemplate(null); setCustomMessage(""); setUseCustom(false); }}
          >
            Start a new broadcast
          </button>
        </div>
      )}
    </div>
  );
}

/* ═══════════════ DASHBOARD CONTENT ═══════════════ */
function DashboardContent() {
  const { credentials, clearCredentials } = useCredentials();
  const params = { apiToken: credentials!.apiToken, phoneNumberId: credentials!.phoneNumberId };
  const queryOpts = { query: { enabled: true, retry: 0 } };

  const { data: summary,       isLoading: loadingSummary,  error: errSummary   } = useGetDashboardSummary(params, queryOpts);
  const { data: labelsData,    isLoading: loadingLabels,   error: errLabels    } = useGetLabelStats(params, queryOpts);
  const { data: agentsData,    isLoading: loadingAgents,   error: errAgents    } = useGetAgentStats(params, queryOpts);
  const { data: sequencesData, isLoading: loadingSeqs,     error: errSeqs      } = useGetSequenceStats(params, queryOpts);
  const { data: subsData,      isLoading: loadingSubs,     error: errSubs      } = useGetSubscribers(params, queryOpts);
  const { data: buttonData,    isLoading: loadingButtons,  error: errButtons   } = useGetButtonStats(params, queryOpts);

  // Detect auth failures (401) from any query
  const anyError = errSummary || errLabels || errAgents || errSeqs || errSubs || errButtons;
  const isAuthError = anyError && (anyError as { status?: number }).status === 401;

  if (isAuthError) {
    return (
      <div className="flex flex-col items-center justify-center py-24 gap-6">
        <div className="w-14 h-14 rounded-2xl flex items-center justify-center" style={{ background: "#FEF2F2", border: "1.5px solid #FECACA" }}>
          <AlertTriangle className="w-7 h-7" style={{ color: RED }} />
        </div>
        <div className="text-center max-w-sm">
          <h2 className="text-lg font-bold mb-1" style={{ color: "#111827" }}>Invalid Credentials</h2>
          <p className="text-sm" style={{ color: GRAY }}>
            Your API Token or Phone Number ID was rejected by The Wise Parrot. Please disconnect and re-enter your credentials.
          </p>
        </div>
        <button
          onClick={clearCredentials}
          className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl font-semibold text-sm transition-colors"
          style={{ background: RED, color: "#fff" }}
        >
          <LogOut className="w-4 h-4" />
          Disconnect &amp; Re-enter Credentials
        </button>
      </div>
    );
  }

  const agents      = agentsData?.agents || [];
  const sequences   = sequencesData?.sequences || [];
  const sortedSeqs  = [...sequences].sort((a, b) => b.reactivationRate - a.reactivationRate);
  const bestSeq     = sortedSeqs[0] || null;
  const worstSeq    = sortedSeqs.length > 1 ? sortedSeqs[sortedSeqs.length - 1] : null;
  const subscribers = subsData?.subscribers || [];
  const activeRate  = summary && summary.totalLeads > 0 ? (summary.activeLeads / summary.totalLeads) * 100 : 0;

  return (
    <>
      {/* ══ S1: Overview ══ */}
      <Section id="overview" icon={<BarChart2 className="w-4 h-4" />} title="Overview"
        subtitle="High-level lead engagement metrics" color={BLUE}
        onDownload={() => summary && downloadCSV("overview.csv", [{ ...summary }])}>
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
          <KpiCard label="Total Leads" value={summary?.totalLeads?.toLocaleString()} isLoading={loadingSummary}
            icon={<Users className="w-5 h-5" />} iconBg="#EFF6FF" iconColor={BLUE} accent={BLUE} />
          <KpiCard label="Active Leads" value={summary?.activeLeads?.toLocaleString()} isLoading={loadingSummary}
            icon={<Activity className="w-5 h-5" />} iconBg="#F0FDF4" iconColor={GREEN_D} accent={GREEN}
            badge={activeRate > 0 ? `${activeRate.toFixed(0)}%` : undefined} badgeClass="ch-badge-green" />
          <KpiCard label="Dormant Leads" value={summary?.dormantLeads?.toLocaleString()} isLoading={loadingSummary}
            icon={<Moon className="w-5 h-5" />} iconBg="#FEF2F2" iconColor={RED} accent={RED} />
          <KpiCard label="Reactivation Rate" value={summary ? `${summary.reactivationRate.toFixed(1)}%` : "0%"} isLoading={loadingSummary}
            icon={<TrendingUp className="w-5 h-5" />} iconBg="#FFF7ED" iconColor={ORANGE_D} accent={ORANGE} highlight />
        </div>
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mt-4">
          {[
            { label: "User Replies",        value: summary?.totalUserReplies?.toLocaleString(),             icon: <MessageSquare className="w-4 h-4" />, color: BLUE_L  },
            { label: "TWP Replies",         value: summary?.totalTwpReplies?.toLocaleString(),              icon: <MessageSquare className="w-4 h-4" />, color: BLUE    },
            { label: "Sequences Sent",      value: summary?.totalSequencesSent?.toLocaleString(),           icon: <Zap className="w-4 h-4" />,           color: ORANGE  },
            { label: "Reactivated Post-Seq",value: summary?.leadsReactivatedAfterSequence?.toLocaleString(),icon: <Flame className="w-4 h-4" />,          color: GREEN   },
          ].map(({ label, value, icon, color }) => (
            <MiniStat key={label} label={label} value={value} isLoading={loadingSummary} icon={icon} color={color} />
          ))}
        </div>
      </Section>

      {/* ══ S2: Distribution ══ */}
      <Section id="distribution" icon={<BarChart2 className="w-4 h-4" />} title="Distribution"
        subtitle="Label breakdown and reply volume" color={ORANGE}
        onDownload={() => labelsData && downloadCSV("label-distribution.csv", labelsData.labels)}>
        <LabelManagementCard apiToken={credentials!.apiToken} phoneNumberId={credentials!.phoneNumberId} />
        <MessageBroadcastCard apiToken={credentials!.apiToken} phoneNumberId={credentials!.phoneNumberId} />
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
          <ChartCard title="Reply Volume" subtitle="User vs TWP messages">
            <div className="h-[230px]">
              {loadingSummary ? <SkeletonChart /> : (
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie data={[
                      { name: "User Replies", value: summary?.totalUserReplies || 0 },
                      { name: "TWP Replies",  value: summary?.totalTwpReplies  || 0 },
                    ]} cx="50%" cy="44%" innerRadius={52} outerRadius={78} paddingAngle={4} dataKey="value">
                      <Cell fill={BLUE_L} />
                      <Cell fill={ORANGE} />
                    </Pie>
                    <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 12, color: GRAY }} />
                    <RechartsTooltip content={<ChTooltip />} />
                  </PieChart>
                </ResponsiveContainer>
              )}
            </div>
          </ChartCard>

          <ChartCard title="Lead Distribution by Label" subtitle="Total vs Dormant per pipeline stage" className="lg:col-span-2">
            <div className="h-[230px]">
              {loadingLabels ? <SkeletonChart /> : (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={labelsData?.labels || []} margin={{ top: 5, right: 5, left: -20, bottom: 40 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={BORDER} />
                    <XAxis dataKey="labelName" tick={{ fill: GRAY, fontSize: 11 }} tickLine={false} axisLine={false} angle={-30} textAnchor="end" height={55} />
                    <YAxis tick={{ fill: GRAY, fontSize: 11 }} tickLine={false} axisLine={false} />
                    <RechartsTooltip content={<ChTooltip />} />
                    <Legend wrapperStyle={{ fontSize: 12, color: GRAY }} />
                    <Bar dataKey="count"        name="Total"   fill={BLUE} radius={[4,4,0,0]} />
                    <Bar dataKey="dormantCount" name="Dormant" fill={RED}  radius={[4,4,0,0]} />
                  </BarChart>
                </ResponsiveContainer>
              )}
            </div>
          </ChartCard>
        </div>
      </Section>

      {/* ══ S3: Agents ══ */}
      <Section id="agents" icon={<Headset className="w-4 h-4" />} title="Agent Performance"
        subtitle="Workload and effectiveness by agent" color={GREEN_D}
        onDownload={() => agents.length && downloadCSV("agents.csv", agents)}>
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-5">
          {[
            { label: "Total Agents",    value: agents.length,                                                         color: BLUE,  icon: <Headset className="w-4 h-4" /> },
            { label: "Total Assigned",  value: agents.reduce((s,a) => s+a.leadsAssigned, 0).toLocaleString(),         color: BLUE_L,icon: <Users className="w-4 h-4" /> },
            { label: "Active Portfolio",value: agents.reduce((s,a) => s+a.activeLeads, 0).toLocaleString(),           color: GREEN, icon: <Activity className="w-4 h-4" /> },
            { label: "Dormant Risk",    value: agents.reduce((s,a) => s+a.dormantLeads, 0).toLocaleString(),          color: RED,   icon: <Moon className="w-4 h-4" /> },
          ].map(({ label, value, color, icon }) => (
            <MiniStat key={label} label={label} value={value} isLoading={loadingAgents} icon={icon} color={color} />
          ))}
        </div>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
          <ChartCard title="Workload Distribution" subtitle="Active vs Dormant per agent">
            <div className="h-[280px]">
              {loadingAgents ? <SkeletonChart /> : agents.length === 0 ? <EmptyChart /> : (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={[...agents].sort((a,b) => b.leadsAssigned-a.leadsAssigned)} margin={{ top:5, right:5, left:-22, bottom:50 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={BORDER} />
                    <XAxis dataKey="agentName" tick={{ fill:GRAY, fontSize:11 }} tickLine={false} axisLine={false} angle={-35} textAnchor="end" height={65} />
                    <YAxis tick={{ fill:GRAY, fontSize:11 }} tickLine={false} axisLine={false} />
                    <RechartsTooltip content={<ChTooltip />} />
                    <Legend wrapperStyle={{ fontSize:12, color:GRAY, paddingTop:8 }} />
                    <Bar dataKey="activeLeads"  name="Active"  stackId="a" fill={GREEN} radius={[0,0,0,0]} />
                    <Bar dataKey="dormantLeads" name="Dormant" stackId="a" fill={RED}   radius={[4,4,0,0]} />
                  </BarChart>
                </ResponsiveContainer>
              )}
            </div>
          </ChartCard>

          <ChartCard title="Agent Leaderboard" subtitle="Ranked by leads assigned" noPadding>
            <DataTable isLoading={loadingAgents} isEmpty={agents.length === 0} emptyText="No agent data yet"
              headers={["#","Agent","Active","Dormant","Avg"]}
              rows={[...agents].sort((a,b)=>b.leadsAssigned-a.leadsAssigned).map((a,i) => {
                const dormPct = a.leadsAssigned>0 ? (a.dormantLeads/a.leadsAssigned)*100 : 0;
                return [
                  <span key="r" className="text-xs" style={{color:GRAY}}>{i===0?"🥇":i===1?"🥈":i===2?"🥉":i+1}</span>,
                  <span key="n" className="font-semibold text-sm" style={{color:"#111827"}}>{a.agentName}</span>,
                  <span key="ac" className="font-mono font-bold text-sm" style={{color:GREEN_D}}>{a.activeLeads}</span>,
                  <span key="d" className="font-mono text-sm" style={{color:RED}}>{a.dormantLeads}</span>,
                  <span key="s" className={dormPct>50?"ch-badge-red":dormPct>30?"ch-badge-orange":"ch-badge-green"}>{a.avgUserReplies.toFixed(1)}</span>,
                ];
              })}
            />
          </ChartCard>
        </div>
      </Section>

      {/* ══ S4: Sequences ══ */}
      <Section id="sequences" icon={<Zap className="w-4 h-4" />} title="Sequence Analytics"
        subtitle="Effectiveness of automated message sequences" color={ORANGE}
        onDownload={() => sortedSeqs.length && downloadCSV("sequences.csv", sortedSeqs)}>
        {(loadingSeqs || bestSeq) && (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-5">
            {loadingSeqs ? (
              <><Skeleton className="h-32 rounded-xl" /><Skeleton className="h-32 rounded-xl" /></>
            ) : bestSeq && (
              <>
                <HighlightCard type="top" title={bestSeq.sequenceName} stats={[
                  { label:"Conv. Rate", value:`${bestSeq.reactivationRate.toFixed(1)}%` },
                  { label:"Sent", value:String(bestSeq.totalSent) },
                  { label:"Replies", value:String(bestSeq.repliesAfterSequence) },
                ]} />
                {worstSeq && worstSeq.sequenceName!==bestSeq.sequenceName && (
                  <HighlightCard type="worst" title={worstSeq.sequenceName} stats={[
                    { label:"Conv. Rate", value:`${worstSeq.reactivationRate.toFixed(1)}%` },
                    { label:"Sent", value:String(worstSeq.totalSent) },
                    { label:"Replies", value:String(worstSeq.repliesAfterSequence) },
                  ]} />
                )}
              </>
            )}
          </div>
        )}

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
          <ChartCard title="Reactivation Rates" subtitle="% of leads replying after each sequence">
            <div className="h-[240px]">
              {loadingSeqs ? <SkeletonChart /> : sortedSeqs.length===0 ? <EmptyChart /> : (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={sortedSeqs} layout="vertical" margin={{ top:5, right:40, left:10, bottom:5 }}>
                    <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke={BORDER} />
                    <XAxis type="number" domain={[0,100]} tickFormatter={(v)=>`${v}%`} tick={{ fill:GRAY, fontSize:11 }} tickLine={false} axisLine={false} />
                    <YAxis dataKey="sequenceName" type="category" tick={{ fill:"#374151", fontSize:11 }} tickLine={false} axisLine={false} width={140} />
                    <RechartsTooltip content={<ChTooltip pct />} />
                    <Bar dataKey="reactivationRate" name="Reactivation Rate" radius={[0,6,6,0]}>
                      {sortedSeqs.map((_,i) => <Cell key={i} fill={i===0?ORANGE:ORANGE_L} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              )}
            </div>
          </ChartCard>

          <ChartCard title="All Sequences" subtitle="Ranked by conversion rate" noPadding>
            <DataTable isLoading={loadingSeqs} isEmpty={sortedSeqs.length===0} emptyText="No sequence data"
              headers={["Sequence","Sent","Replies","Rate"]}
              rows={sortedSeqs.map((seq,i) => [
                <div key="n" className="flex items-center gap-2">
                  {i===0&&<span className="text-orange-500 shrink-0">⭐</span>}
                  <span className="font-medium text-sm truncate max-w-[150px]" style={{color:"#111827"}}>{seq.sequenceName}</span>
                </div>,
                <span key="s" className="font-mono text-sm" style={{color:GRAY}}>{seq.totalSent}</span>,
                <span key="r" className="font-mono text-sm" style={{color:GRAY}}>{seq.repliesAfterSequence}</span>,
                <span key="rt" className="font-mono font-bold text-sm" style={{color:i===0?ORANGE_D:"#374151"}}>{seq.reactivationRate.toFixed(1)}%</span>,
              ])}
            />
          </ChartCard>
        </div>
      </Section>

      {/* ══ S5: Button Clicks ══ */}
      <Section id="buttons" icon={<MousePointerClick className="w-4 h-4" />} title="Button Click Analytics"
        subtitle="Interactive button engagement from conversations" color={BLUE}
        onDownload={() => buttonData?.buttons?.length && downloadCSV("button-clicks.csv", buttonData.buttons)}>
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-5">
          <KpiCard label="Total Clicks"   value={buttonData?.totalClicks?.toLocaleString()} isLoading={loadingButtons}
            icon={<MousePointerClick className="w-5 h-5" />} iconBg="#FFF7ED" iconColor={ORANGE_D} accent={ORANGE} />
          <KpiCard label="Unique Buttons" value={buttonData?.buttons?.length}              isLoading={loadingButtons}
            icon={<Zap className="w-5 h-5" />}              iconBg="#EFF6FF" iconColor={BLUE}     accent={BLUE} />
          <KpiCard label="Most Clicked"   value={buttonData?.insights?.mostClicked ?? "—"} isLoading={loadingButtons}
            icon={<Trophy className="w-5 h-5" />}           iconBg="#FEFCE8" iconColor="#CA8A04"  accent="#EAB308" />
          <KpiCard label="Best Conversion" value={buttonData?.insights?.bestConversion ?? "—"} isLoading={loadingButtons}
            icon={<Award className="w-5 h-5" />}            iconBg="#F0FDF4" iconColor={GREEN_D}  accent={GREEN} />
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
          <ChartCard title="Clicks per Button" subtitle="Total clicks vs unique lead reach">
            <div className="h-[280px]">
              {loadingButtons ? <SkeletonChart /> : !buttonData?.buttons?.length ? (
                <EmptyChart text="No interactive button clicks found" />
              ) : (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={buttonData.buttons} margin={{ top:5, right:5, left:-22, bottom:55 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={BORDER} />
                    <XAxis dataKey="buttonName" tick={{ fill:GRAY, fontSize:11 }} tickLine={false} axisLine={false} angle={-35} textAnchor="end" height={75} />
                    <YAxis tick={{ fill:GRAY, fontSize:11 }} tickLine={false} axisLine={false} />
                    <RechartsTooltip content={<ChTooltip />} />
                    <Legend wrapperStyle={{ fontSize:12, color:GRAY, paddingTop:8 }} />
                    <Bar dataKey="totalClicks" name="Total Clicks" fill={BLUE}   radius={[4,4,0,0]} />
                    <Bar dataKey="uniqueLeads" name="Unique Leads" fill={ORANGE} radius={[4,4,0,0]} />
                  </BarChart>
                </ResponsiveContainer>
              )}
            </div>
          </ChartCard>

          {/* Paginated button performance */}
          <ButtonPerformanceTable buttons={buttonData?.buttons || []} isLoading={loadingButtons} insights={buttonData?.insights} />
        </div>
      </Section>

      {/* ══ S6: Subscribers ══ */}
      <Section id="subscribers" icon={<Users className="w-4 h-4" />} title="Subscribers"
        subtitle="Full lead database with conversation insights" color={BLUE_D}
        onDownload={() => subscribers.length && downloadCSV("subscribers.csv", subscribers)}>
        <SubscriberTable subscribers={subscribers} isLoading={loadingSubs} />
      </Section>
    </>
  );
}

/* ═══════════════ BUTTON PERFORMANCE TABLE (paginated) ═══════════════ */
function ButtonPerformanceTable({ buttons, isLoading, insights }: {
  buttons: Array<{ buttonName: string; totalClicks: number; uniqueLeads: number; clickRate: number; subscribers?: Array<{ name: string; phoneNumber: string }> }>;
  isLoading: boolean;
  insights?: { mostClicked: string | null; leastClicked: string | null; bestConversion: string | null };
}) {
  const pg = usePagination(buttons, PAGE_SIZE);
  const [expanded, setExpanded] = useState<string | null>(null);

  return (
    <ChartCard title="Button Performance" subtitle={`${buttons.length} buttons · ${PAGE_SIZE} per page`} noPadding>
      {isLoading ? (
        <div className="p-5 space-y-4">
          {Array.from({ length: PAGE_SIZE }).map((_, i) => <Skeleton key={i} className="h-14 rounded-xl" />)}
        </div>
      ) : !buttons.length ? (
        <div className="p-8 text-center text-sm" style={{ color: GRAY }}>No button data found</div>
      ) : (
        <>
          <div className="divide-y" style={{ borderColor: BORDER }}>
            {pg.slice.map((btn, relIdx) => {
              const absoluteIdx = (pg.page * PAGE_SIZE) + relIdx;
              const isFirst = absoluteIdx === 0;
              const isLast  = absoluteIdx === buttons.length - 1 && buttons.length > 1;
              const subs = btn.subscribers ?? [];
              const isOpen = expanded === btn.buttonName;

              return (
                <div key={btn.buttonName} className={`transition-colors ${isFirst ? "bg-blue-50/40" : ""}`}>
                  {/* Row header */}
                  <div className="px-5 py-4">
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-2.5 min-w-0">
                        {isFirst
                          ? <Trophy className="w-4 h-4 shrink-0" style={{ color: "#CA8A04" }} />
                          : isLast
                          ? <Minus className="w-4 h-4 shrink-0" style={{ color: "#D1D5DB" }} />
                          : <span className="w-5 text-xs font-mono text-center shrink-0" style={{ color: GRAY }}>{absoluteIdx + 1}</span>}
                        <span className="font-semibold text-sm truncate" style={{ color: "#111827" }}>{btn.buttonName}</span>
                      </div>
                      <div className="flex items-center gap-2 shrink-0 ml-3">
                        {/* People count badge */}
                        <button
                          onClick={() => setExpanded(isOpen ? null : btn.buttonName)}
                          className="inline-flex items-center gap-1 px-2 py-1 rounded-lg text-xs font-medium transition-colors"
                          style={{ background: isOpen ? "#DBEAFE" : "#EFF6FF", color: BLUE, border: `1px solid #BFDBFE` }}
                          title={`${subs.length} people clicked · click to ${isOpen ? "hide" : "view"}`}
                        >
                          <Users className="w-3 h-3" />
                          {subs.length} people
                        </button>
                        {/* Download CSV for this button */}
                        {subs.length > 0 && (
                          <button
                            onClick={() => downloadCSV(`button-${btn.buttonName.replace(/\s+/g, "_")}.csv`, subs.map((s) => ({ Name: s.name, "Phone Number": s.phoneNumber })))}
                            className="inline-flex items-center gap-1 px-2 py-1 rounded-lg text-xs font-medium transition-colors"
                            style={{ background: "#F0FDF4", color: GREEN_D, border: `1px solid #BBF7D0` }}
                            title="Download subscriber list as CSV"
                          >
                            <Download className="w-3 h-3" />
                            CSV
                          </button>
                        )}
                        <span className="font-bold text-sm font-mono" style={{ color: isFirst ? ORANGE_D : "#374151" }}>
                          {btn.totalClicks.toLocaleString()} clicks
                        </span>
                      </div>
                    </div>
                    <div className="flex items-center gap-3">
                      <div className="flex-1 h-2 rounded-full overflow-hidden" style={{ background: "#F3F4F6" }}>
                        <div className="h-full rounded-full animate-bar-grow" style={{
                          width: `${btn.clickRate}%`,
                          background: isFirst ? `linear-gradient(90deg, ${ORANGE}, ${ORANGE_L})` : `linear-gradient(90deg, ${BLUE}, ${BLUE_L})`,
                        }} />
                      </div>
                      <span className="text-xs font-mono w-10 text-right shrink-0" style={{ color: GRAY }}>{btn.clickRate.toFixed(1)}%</span>
                    </div>
                  </div>

                  {/* Expanded subscriber list */}
                  {isOpen && subs.length > 0 && (
                    <div className="border-t mx-5 mb-4" style={{ borderColor: BORDER }}>
                      <div className="pt-3 pb-1">
                        <div className="flex items-center justify-between mb-2">
                          <span className="text-xs font-semibold uppercase tracking-wider" style={{ color: GRAY }}>Subscribers who clicked</span>
                        </div>
                        <div className="rounded-xl overflow-hidden border" style={{ borderColor: BORDER }}>
                          <table className="w-full text-sm">
                            <thead>
                              <tr style={{ background: "#F9FAFB" }}>
                                <th className="text-left px-4 py-2 text-xs font-semibold" style={{ color: GRAY }}>#</th>
                                <th className="text-left px-4 py-2 text-xs font-semibold" style={{ color: GRAY }}>Name</th>
                                <th className="text-left px-4 py-2 text-xs font-semibold" style={{ color: GRAY }}>Phone Number</th>
                              </tr>
                            </thead>
                            <tbody>
                              {subs.map((s, i) => (
                                <tr key={s.phoneNumber} className={i % 2 === 0 ? "" : ""} style={{ borderTop: `1px solid ${BORDER}` }}>
                                  <td className="px-4 py-2 text-xs font-mono" style={{ color: GRAY }}>{i + 1}</td>
                                  <td className="px-4 py-2 font-medium" style={{ color: "#111827" }}>{s.name}</td>
                                  <td className="px-4 py-2 font-mono text-xs" style={{ color: GRAY }}>{s.phoneNumber}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>

          {/* Summary footer */}
          {insights && (
            <div className="px-5 py-2.5 border-t flex flex-wrap gap-3 text-xs" style={{ borderColor: BORDER, background: "#F9FAFB", color: GRAY }}>
              <span>🏆 <strong style={{ color: "#374151" }}>{insights.mostClicked}</strong></span>
              <span>📉 Least: <strong style={{ color: "#374151" }}>{insights.leastClicked}</strong></span>
            </div>
          )}

          <Paginator page={pg.page} totalPages={pg.totalPages} from={pg.from} to={pg.to} total={pg.total}
            onPrev={pg.prevPage} onNext={pg.nextPage} />
        </>
      )}
    </ChartCard>
  );
}

/* ═══════════════ SUBSCRIBER TABLE (paginated) ═══════════════ */
interface SubRow {
  name: string; phoneNumber: string; labelName: string;
  lastMessageTime?: string | null; assignedAgent?: string | null;
  assignedSequence?: string | null; userReplyCount: number;
  twpReplyCount: number; isDormant: boolean; postSequenceReplies: number;
}

function SubscriberTable({ subscribers, isLoading }: { subscribers: SubRow[]; isLoading: boolean }) {
  const [search,       setSearch]       = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [agentFilter,  setAgentFilter]  = useState("all");
  const [labelFilter,  setLabelFilter]  = useState("all");

  const agentList = useMemo(() => Array.from(new Set(subscribers.map((s) => s.assignedAgent).filter(Boolean) as string[])), [subscribers]);
  const labelList = useMemo(() => Array.from(new Set(subscribers.map((s) => s.labelName).filter(Boolean) as string[])), [subscribers]);

  const filtered = useMemo(() => subscribers.filter((s) => {
    const q = search.toLowerCase();
    if (q && !s.name.toLowerCase().includes(q) && !s.phoneNumber.includes(q)) return false;
    if (statusFilter === "active"  &&  s.isDormant) return false;
    if (statusFilter === "dormant" && !s.isDormant) return false;
    if (agentFilter !== "all" && s.assignedAgent !== agentFilter) return false;
    if (labelFilter !== "all" && s.labelName     !== labelFilter) return false;
    return true;
  }), [subscribers, search, statusFilter, agentFilter, labelFilter]);

  const pg = usePagination(filtered, PAGE_SIZE);

  // Reset page whenever filters change
  useEffect(() => { pg.resetPage(); }, [search, statusFilter, agentFilter, labelFilter]);

  const activeCount  = subscribers.filter((s) => !s.isDormant).length;
  const dormantCount = subscribers.filter((s) =>  s.isDormant).length;

  return (
    <div className="space-y-4">
      {/* Filter bar */}
      <div className="flex flex-col sm:flex-row gap-3 p-4 rounded-xl border" style={{ background: "#fff", borderColor: BORDER }}>
        <div className="relative flex-1 min-w-[180px]">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4" style={{ color: "#9CA3AF" }} />
          <input
            className="w-full rounded-xl pl-9 pr-4 py-2 text-sm border outline-none transition-all"
            style={{ background: "#F9FAFB", borderColor: BORDER, color: "#111827" }}
            placeholder="Search name or phone..."
            value={search} onChange={(e) => setSearch(e.target.value)}
            onFocus={(e) => { e.target.style.borderColor = BLUE; e.target.style.boxShadow = "0 0 0 3px rgba(37,99,235,0.1)"; }}
            onBlur={(e)  => { e.target.style.borderColor = BORDER; e.target.style.boxShadow = "none"; }}
          />
        </div>
        <div className="flex flex-wrap gap-2">
          {([
            { value: statusFilter, onChange: setStatusFilter, opts: [["all","All Status"],["active","Active"],["dormant","Dormant"]] },
            { value: agentFilter,  onChange: setAgentFilter,  opts: [["all","All Agents"], ...agentList.map((a) => [a, a])] },
            { value: labelFilter,  onChange: setLabelFilter,  opts: [["all","All Labels"], ...labelList.map((l) => [l, l])] },
          ] as Array<{ value: string; onChange: (v: string)=>void; opts: string[][] }>).map((sel, i) => (
            <select key={i} value={sel.value} onChange={(e) => sel.onChange(e.target.value)}
              className="rounded-xl px-3 py-2 text-sm border outline-none transition-all"
              style={{ background:"#F9FAFB", borderColor:BORDER, color:"#374151", minWidth:130, cursor:"pointer" }}>
              {sel.opts.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
            </select>
          ))}
        </div>
      </div>

      {/* Table */}
      <div className="ch-card overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr style={{ borderBottom:`1px solid ${BORDER}`, background:"#F9FAFB" }}>
                {["Subscriber","Status","Label","Agent","Sequence","Replies (U/T)","Post-Seq","Last Active"].map((h) => (
                  <th key={h} className="px-4 py-3 text-left font-semibold text-xs uppercase tracking-wider whitespace-nowrap" style={{ color:GRAY }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {isLoading ? (
                Array.from({ length: PAGE_SIZE }).map((_, i) => (
                  <tr key={i} style={{ borderBottom:`1px solid ${BORDER}` }}>
                    {Array.from({ length: 8 }).map((__, j) => <td key={j} className="px-4 py-3"><Skeleton className="h-4 w-full" /></td>)}
                  </tr>
                ))
              ) : pg.slice.length === 0 ? (
                <tr>
                  <td colSpan={8} className="px-4 py-10 text-center text-sm" style={{ color:GRAY }}>
                    No subscribers match the current filters.
                  </td>
                </tr>
              ) : pg.slice.map((sub, idx) => (
                <tr key={`${sub.phoneNumber}-${idx}`} style={{ borderBottom:`1px solid ${BORDER}` }} className="hover:bg-gray-50 transition-colors">
                  <td className="px-4 py-3.5">
                    <div className="font-semibold text-sm" style={{ color:"#111827" }}>{sub.name}</div>
                    <div className="text-xs font-mono mt-0.5" style={{ color:GRAY }}>+{sub.phoneNumber}</div>
                  </td>
                  <td className="px-4 py-3.5">
                    <span className={sub.isDormant ? "ch-badge-red" : "ch-badge-green"}>
                      <span className={`w-1.5 h-1.5 rounded-full ${sub.isDormant ? "bg-red-500" : "bg-green-500 animate-pulse-dot"}`} />
                      {sub.isDormant ? "Dormant" : "Active"}
                    </span>
                  </td>
                  <td className="px-4 py-3.5">
                    <span className="ch-badge-blue max-w-[130px] truncate inline-block" title={sub.labelName}>{sub.labelName}</span>
                  </td>
                  <td className="px-4 py-3.5 text-sm" style={{ color:"#374151" }}>
                    {sub.assignedAgent || <span className="text-xs italic" style={{ color:"#D1D5DB" }}>Unassigned</span>}
                  </td>
                  <td className="px-4 py-3.5 text-sm max-w-[140px]">
                    <span className="truncate block" title={sub.assignedSequence || undefined} style={{ color:"#374151" }}>
                      {sub.assignedSequence || <span className="text-xs italic" style={{ color:"#D1D5DB" }}>None</span>}
                    </span>
                  </td>
                  <td className="px-4 py-3.5 text-center font-mono text-sm">
                    <span style={{ color:GREEN_D }}>{sub.userReplyCount}</span>
                    <span className="mx-1" style={{ color:"#D1D5DB" }}>/</span>
                    <span style={{ color:BLUE }}>{sub.twpReplyCount}</span>
                  </td>
                  <td className="px-4 py-3.5 text-center font-mono text-sm">
                    {sub.postSequenceReplies > 0
                      ? <span className="ch-badge-orange">{sub.postSequenceReplies}</span>
                      : <span style={{ color:"#D1D5DB" }}>—</span>}
                  </td>
                  <td className="px-4 py-3.5 text-sm whitespace-nowrap" style={{ color:GRAY }}>
                    {sub.lastMessageTime
                      ? format(new Date(sub.lastMessageTime), "MMM d, HH:mm")
                      : <span className="text-xs italic" style={{ color:"#D1D5DB" }}>Never</span>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Footer with stats + paginator */}
        {!isLoading && filtered.length > 0 && (
          <div className="border-t" style={{ borderColor: BORDER }}>
            <Paginator page={pg.page} totalPages={pg.totalPages} from={pg.from} to={pg.to} total={pg.total}
              onPrev={pg.prevPage} onNext={pg.nextPage}
              extra={
                <div className="flex gap-3">
                  <span className="ch-badge-green">
                    <span className="w-1.5 h-1.5 rounded-full bg-green-500 animate-pulse-dot" />{activeCount} active
                  </span>
                  <span className="ch-badge-red">{dormantCount} dormant</span>
                </div>
              }
            />
          </div>
        )}
      </div>
    </div>
  );
}

/* ═══════════════ SHARED COMPONENTS ═══════════════ */

function Paginator({ page, totalPages, from, to, total, onPrev, onNext, extra }: {
  page: number; totalPages: number; from: number; to: number; total: number;
  onPrev: () => void; onNext: () => void; extra?: React.ReactNode;
}) {
  const btnBase: React.CSSProperties = {
    display: "inline-flex", alignItems: "center", gap: 4,
    padding: "5px 12px", borderRadius: 8, fontSize: 12, fontWeight: 600,
    border: `1px solid ${BORDER}`, cursor: "pointer", transition: "all 0.15s",
  };
  return (
    <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-3" style={{ background: "#F9FAFB" }}>
      <div className="flex items-center gap-3">
        <span className="text-xs" style={{ color: GRAY }}>
          {total === 0 ? "No results" : `Showing ${from}–${to} of ${total}`}
        </span>
        {extra}
      </div>
      <div className="flex items-center gap-2">
        <button onClick={onPrev} disabled={page === 0}
          style={{ ...btnBase, background: page===0?"#F3F4F6":"#fff", color: page===0?"#D1D5DB":GRAY, cursor: page===0?"not-allowed":"pointer" }}>
          <ChevronLeft className="w-3.5 h-3.5" /> Prev
        </button>
        <span className="text-xs font-mono px-2" style={{ color: GRAY }}>{page + 1} / {totalPages}</span>
        <button onClick={onNext} disabled={page >= totalPages - 1}
          style={{ ...btnBase, background: page>=totalPages-1?"#F3F4F6":"#fff", color: page>=totalPages-1?"#D1D5DB":GRAY, cursor: page>=totalPages-1?"not-allowed":"pointer" }}>
          Next <ChevronRight className="w-3.5 h-3.5" />
        </button>
      </div>
    </div>
  );
}

function Section({ id, icon, title, subtitle, color, onDownload, children }: {
  id: string; icon: React.ReactNode; title: string; subtitle: string;
  color: string; onDownload?: () => void; children: React.ReactNode;
}) {
  return (
    <section id={id}>
      <div className="flex items-center gap-3 mb-6">
        <div className="w-9 h-9 rounded-xl flex items-center justify-center shrink-0"
          style={{ background:`${color}15`, border:`1.5px solid ${color}30`, color }}>
          {icon}
        </div>
        <div className="flex-1 min-w-0">
          <h2 className="text-lg font-bold leading-tight truncate" style={{ color:"#111827" }}>{title}</h2>
          <p className="text-xs mt-0.5" style={{ color:GRAY }}>{subtitle}</p>
        </div>
        {onDownload && (
          <button onClick={onDownload}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold shrink-0 transition-all"
            style={{ background:"#F9FAFB", color:GRAY, border:`1px solid ${BORDER}` }}
            onMouseEnter={(e) => { const el=e.currentTarget as HTMLElement; el.style.background="#EFF6FF"; el.style.color=BLUE; el.style.borderColor="#DBEAFE"; }}
            onMouseLeave={(e) => { const el=e.currentTarget as HTMLElement; el.style.background="#F9FAFB"; el.style.color=GRAY; el.style.borderColor=BORDER; }}>
            <Download className="w-3.5 h-3.5" /> Export CSV
          </button>
        )}
        <div className="hidden sm:block h-px w-16 ml-1" style={{ background:`linear-gradient(90deg, ${color}40, transparent)` }} />
      </div>
      {children}
    </section>
  );
}

function KpiCard({ label, value, isLoading, icon, iconBg, iconColor, accent, badge, badgeClass, highlight }: {
  label: string; value?: string | number; isLoading: boolean;
  icon: React.ReactNode; iconBg: string; iconColor: string; accent: string;
  badge?: string; badgeClass?: string; highlight?: boolean;
}) {
  return (
    <div className="ch-card p-5 hover:shadow-md transition-shadow relative overflow-hidden">
      <div className="absolute top-0 left-0 right-0 h-0.5 rounded-t-xl" style={{ background: accent }} />
      <div className="flex items-start justify-between mb-4 mt-1">
        <div className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0" style={{ background:iconBg, color:iconColor }}>{icon}</div>
        {badge && <span className={badgeClass}>{badge}</span>}
      </div>
      <p className="text-xs font-semibold uppercase tracking-wide mb-1" style={{ color:GRAY }}>{label}</p>
      {isLoading
        ? <Skeleton className="h-8 w-20 mt-1" />
        : <p className="text-2xl font-bold tracking-tight animate-count-in" style={{ color: highlight ? accent : "#111827" }}>{value ?? 0}</p>}
    </div>
  );
}

function MiniStat({ label, value, isLoading, icon, color }: {
  label: string; value?: string | number; isLoading: boolean; icon: React.ReactNode; color: string;
}) {
  return (
    <div className="ch-card flex items-center gap-3 px-4 py-3 hover:shadow-md transition-shadow">
      <div className="shrink-0 w-8 h-8 rounded-lg flex items-center justify-center" style={{ background:`${color}15`, color }}>{icon}</div>
      <div className="min-w-0">
        <p className="text-[10px] font-semibold uppercase tracking-wide truncate" style={{ color:GRAY }}>{label}</p>
        {isLoading ? <Skeleton className="h-5 w-14 mt-0.5" /> : <p className="text-base font-bold font-mono" style={{ color:"#111827" }}>{value ?? 0}</p>}
      </div>
    </div>
  );
}

function ChartCard({ title, subtitle, children, className, noPadding }: {
  title: string; subtitle?: string; children: React.ReactNode; className?: string; noPadding?: boolean;
}) {
  return (
    <div className={`ch-card overflow-hidden ${className ?? ""}`}>
      <div className="px-5 pt-5 pb-3">
        <p className="font-bold text-sm" style={{ color:"#111827" }}>{title}</p>
        {subtitle && <p className="text-xs mt-0.5" style={{ color:GRAY }}>{subtitle}</p>}
      </div>
      <div className={noPadding ? "" : "px-5 pb-5"}>{children}</div>
    </div>
  );
}

function DataTable({ headers, rows, isLoading, isEmpty, emptyText }: {
  headers: string[]; rows: React.ReactNode[][];
  isLoading: boolean; isEmpty: boolean; emptyText?: string;
}) {
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-sm">
        <thead>
          <tr style={{ borderBottom:`1px solid ${BORDER}`, background:"#F9FAFB" }}>
            {headers.map((h, i) => (
              <th key={h} className={`px-4 py-3 text-xs font-semibold uppercase tracking-wider ${i>0?"text-right":"text-left"}`} style={{ color:GRAY }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {isLoading ? (
            Array.from({ length: 5 }).map((_,i) => (
              <tr key={i} style={{ borderBottom:`1px solid ${BORDER}` }}>
                {headers.map((__,j) => <td key={j} className="px-4 py-3"><Skeleton className="h-4 w-full" /></td>)}
              </tr>
            ))
          ) : isEmpty ? (
            <tr><td colSpan={headers.length} className="px-4 py-8 text-center text-sm" style={{ color:GRAY }}>{emptyText}</td></tr>
          ) : rows.map((row,i) => (
            <tr key={i} style={{ borderBottom:`1px solid ${BORDER}` }} className="hover:bg-gray-50 transition-colors">
              {row.map((cell,j) => <td key={j} className={`px-4 py-3 ${j>0?"text-right":"text-left"}`}>{cell}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function HighlightCard({ type, title, stats }: {
  type:"top"|"worst"; title:string; stats:Array<{label:string;value:string}>;
}) {
  const isTop = type==="top";
  return (
    <div className="rounded-2xl p-5 relative overflow-hidden border" style={{
      background: isTop?"linear-gradient(135deg,#EFF6FF,#DBEAFE30)":"linear-gradient(135deg,#FEF2F2,#FECACA30)",
      borderColor: isTop?"#BFDBFE":"#FECACA",
    }}>
      <div className="flex items-center gap-2 mb-3">
        {isTop ? <TrendingUp className="w-4 h-4" style={{color:BLUE}} /> : <AlertTriangle className="w-4 h-4" style={{color:RED}} />}
        <span className="text-xs font-bold uppercase tracking-widest" style={{color:isTop?BLUE:RED}}>
          {isTop?"Top Performing":"Needs Improvement"}
        </span>
      </div>
      <p className="font-bold text-base mb-4 truncate" style={{color:"#111827"}}>{title}</p>
      <div className="grid grid-cols-3 gap-3">
        {stats.map(({label,value}) => (
          <div key={label}>
            <p className="text-[10px] font-semibold uppercase tracking-wide mb-1" style={{color:GRAY}}>{label}</p>
            <p className="text-lg font-bold font-mono" style={{color:isTop?BLUE_D:RED}}>{value}</p>
          </div>
        ))}
      </div>
    </div>
  );
}

function SkeletonChart() {
  return (
    <div className="w-full h-full flex items-end gap-2 px-4 pb-4">
      {Array.from({length:6}).map((_,i) => (
        <div key={i} className="flex-1 rounded-t-md" style={{height:`${30+(i*11)%60}%`,background:"#F3F4F6"}} />
      ))}
    </div>
  );
}

function EmptyChart({text}:{text?:string}) {
  return (
    <div className="w-full h-full flex flex-col items-center justify-center gap-2" style={{color:GRAY}}>
      <span className="text-3xl opacity-30">📭</span>
      <p className="text-sm text-center max-w-[200px]">{text||"No data available yet"}</p>
    </div>
  );
}
