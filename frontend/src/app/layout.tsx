import type { Metadata } from "next";
import { IBM_Plex_Sans_Thai, Fira_Code } from "next/font/google";
import "./globals.css";

const ibmPlexThai = IBM_Plex_Sans_Thai({
  subsets: ["thai", "latin"],
  weight: ["300", "400", "500", "600", "700"],
  variable: "--font-sans",
});

const firaCode = Fira_Code({
  subsets: ["latin"],
  weight: ["400", "500", "600"],
  variable: "--font-mono",
});

export const metadata: Metadata = {
  title: "Thailand Oil Price Structure & Retail Tracker",
  description: "Auto-synced EPPO Oil Price Structure & Daily Retail Price Dashboard (PTTOR, Bangchak, Caltex)",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="th" className={`${ibmPlexThai.variable} ${firaCode.variable}`}>
      <body className="bg-slate-50 text-slate-900 dark:bg-slate-950 dark:text-slate-100 font-sans antialiased">
        {children}
      </body>
    </html>
  );
}
