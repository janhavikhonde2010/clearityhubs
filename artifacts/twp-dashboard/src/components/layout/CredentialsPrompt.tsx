import { useState } from "react";
import { useCredentials } from "@/contexts/CredentialsContext";

export function CredentialsPrompt() {
  const { setCredentials } = useCredentials();
  const [apiToken, setApiToken] = useState("");
  const [phoneNumberId, setPhoneNumberId] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!apiToken.trim() || !phoneNumberId.trim()) return;
    setLoading(true);
    setTimeout(() => {
      setCredentials({ apiToken: apiToken.trim(), phoneNumberId: phoneNumberId.trim() });
    }, 350);
  };

  const inputStyle = {
    background: "#F9FAFB",
    border: "1.5px solid #E5E7EB",
    color: "#111827",
    borderRadius: "12px",
    padding: "10px 14px",
    fontSize: "14px",
    outline: "none",
    width: "100%",
    transition: "border-color 0.15s, box-shadow 0.15s",
    fontFamily: "inherit",
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-4" style={{ background: "#F8FAFC" }}>
      {/* Ambient decoration */}
      <div className="absolute inset-0 overflow-hidden pointer-events-none">
        <div className="absolute -top-32 -right-32 w-[500px] h-[500px] rounded-full"
          style={{ background: "radial-gradient(circle, rgba(37,99,235,0.06), transparent)" }} />
        <div className="absolute -bottom-32 -left-32 w-[500px] h-[500px] rounded-full"
          style={{ background: "radial-gradient(circle, rgba(249,115,22,0.05), transparent)" }} />
      </div>

      <div className="relative w-full max-w-sm animate-fade-up">
        {/* Logo + branding */}
        <div className="text-center mb-8">
          <div className="inline-flex items-center justify-center w-20 h-20 rounded-2xl mb-5 shadow-md"
            style={{ background: "#EFF6FF", border: "1.5px solid #DBEAFE" }}>
            <img src="/logo.png" alt="Clarity Hub" className="w-14 h-14 object-contain" />
          </div>
          <h1 className="text-2xl font-bold mb-1" style={{ color: "#111827" }}>Clarity Hub</h1>
          <p className="text-sm" style={{ color: "#6B7280" }}>WhatsApp Analytics Dashboard</p>
        </div>

        {/* Card */}
        <div className="ch-card p-7">
          <p className="text-sm font-bold mb-5" style={{ color: "#374151" }}>Connect your account</p>

          <form onSubmit={handleSubmit} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold mb-1.5" style={{ color: "#374151" }}>
                API Token <span style={{ color: "#EF4444" }}>*</span>
              </label>
              <input
                type="password"
                placeholder="Enter your API token"
                value={apiToken}
                onChange={(e) => setApiToken(e.target.value)}
                required
                style={{ ...inputStyle, fontFamily: "JetBrains Mono, monospace" }}
                onFocus={(e) => { e.target.style.borderColor = "#2563EB"; e.target.style.boxShadow = "0 0 0 3px rgba(37,99,235,0.1)"; }}
                onBlur={(e) => { e.target.style.borderColor = "#E5E7EB"; e.target.style.boxShadow = "none"; }}
              />
            </div>

            <div>
              <label className="block text-xs font-semibold mb-1.5" style={{ color: "#374151" }}>
                Phone Number ID <span style={{ color: "#EF4444" }}>*</span>
              </label>
              <input
                type="text"
                placeholder="e.g. 1234567890"
                value={phoneNumberId}
                onChange={(e) => setPhoneNumberId(e.target.value)}
                required
                style={{ ...inputStyle, fontFamily: "JetBrains Mono, monospace" }}
                onFocus={(e) => { e.target.style.borderColor = "#2563EB"; e.target.style.boxShadow = "0 0 0 3px rgba(37,99,235,0.1)"; }}
                onBlur={(e) => { e.target.style.borderColor = "#E5E7EB"; e.target.style.boxShadow = "none"; }}
              />
            </div>

            <button
              type="submit"
              disabled={!apiToken || !phoneNumberId || loading}
              className="w-full py-3 rounded-xl text-sm font-semibold text-white transition-all mt-1"
              style={{
                background: !apiToken || !phoneNumberId || loading ? "#93C5FD" : "#2563EB",
                cursor: !apiToken || !phoneNumberId || loading ? "not-allowed" : "pointer",
              }}
              onMouseEnter={(e) => { if (apiToken && phoneNumberId && !loading) (e.target as HTMLButtonElement).style.background = "#1D4ED8"; }}
              onMouseLeave={(e) => { if (apiToken && phoneNumberId && !loading) (e.target as HTMLButtonElement).style.background = "#2563EB"; }}
            >
              {loading ? "Connecting…" : "Connect to Dashboard →"}
            </button>
          </form>
        </div>

        <p className="text-center text-xs mt-4" style={{ color: "#9CA3AF" }}>
          Credentials are stored locally in your browser only.
        </p>
      </div>
    </div>
  );
}
