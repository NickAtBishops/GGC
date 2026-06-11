import type { Metadata } from "next";
import type { ReactNode } from "react";
import "./globals.css";
import Header from "@/components/Header";

export const metadata: Metadata = {
  title: "GGC Deal Engine | AI-Powered MHP Underwriting",
  description:
    "AI underwriting for manufactured-housing communities and RV parks. " +
    "Upload seller financials, get a populated 16-tab GGC underwriting model.",
};

export default function RootLayout({
  children,
}: Readonly<{ children: ReactNode }>) {
  return (
    <html lang="en">
      <body className="gradient-bg min-h-screen text-slate-100 antialiased">
        <Header />
        <main className="max-w-6xl mx-auto px-6 pb-12">{children}</main>
      </body>
    </html>
  );
}
