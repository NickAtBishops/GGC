"use client";

// Lets a visitor supply their own Anthropic API key instead of using the
// server's default. Mirrors index.html's API key card: stored only in this
// browser's localStorage, sent with each /api/analyze request, never
// persisted server-side. No sign-in is involved — this app currently has no
// auth gate (see AuthGate.tsx, unused; backend.py REQUIRE_AUTH defaults off).

interface ApiKeyCardProps {
  apiKey: string;
  defaultKeyPresent: boolean;
  onApiKeyChange: (key: string) => void;
}

export default function ApiKeyCard({ apiKey, defaultKeyPresent, onApiKeyChange }: ApiKeyCardProps) {
  const connected = Boolean(apiKey) || defaultKeyPresent;
  const statusLabel = apiKey ? "Using your key" : defaultKeyPresent ? "Using shared key" : "No API key";

  return (
    <div className="card rounded-xl p-5 mb-6">
      <div className="flex items-center justify-between mb-3">
        <label htmlFor="api-key" className="flex items-center gap-2 text-sm font-semibold">
          <svg className="w-4 h-4 text-blue-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={2}
              d="M15 7a2 2 0 012 2m4 0a6 6 0 01-7.743 5.743L11 17H9v2H7v2H4a1 1 0 01-1-1v-2.586a1 1 0 01.293-.707l5.964-5.964A6 6 0 1121 9z"
            />
          </svg>
          Anthropic API Key
        </label>
        <div
          className={`flex items-center gap-2 px-3 py-1.5 rounded-full border ${
            connected ? "bg-green-500/10 border-green-500/30" : "bg-red-500/10 border-red-500/30"
          }`}
        >
          <div className={`w-2 h-2 rounded-full ${connected ? "bg-green-400" : "bg-red-400"}`} />
          <span className={`text-xs ${connected ? "text-green-400" : "text-red-400"}`}>{statusLabel}</span>
        </div>
      </div>
      <input
        id="api-key"
        type="password"
        autoComplete="off"
        value={apiKey}
        onChange={(e) => onApiKeyChange(e.target.value)}
        placeholder={defaultKeyPresent ? "sk-ant-api03-... (optional)" : "sk-ant-api03-..."}
        className="input-field w-full rounded-lg px-4 py-2.5 text-sm"
      />
      <p className="text-xs text-slate-500 mt-2">
        Stored only in this browser and sent with each analysis request — never saved on the server.
        {defaultKeyPresent && " Leave blank to use the shared key instead."}
      </p>
    </div>
  );
}
