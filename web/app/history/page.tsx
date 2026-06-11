import type { Metadata } from "next";
import RunHistory from "@/components/RunHistory";

export const metadata: Metadata = {
  title: "Run History | GGC Deal Engine",
};

export default function HistoryPage() {
  return (
    <>
      <div className="mb-4">
        <h2 className="text-lg font-semibold">Run History</h2>
        <p className="text-sm text-slate-400">Your last 50 underwriting runs.</p>
      </div>
      <RunHistory />
    </>
  );
}
