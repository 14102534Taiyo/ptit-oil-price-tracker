import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Thailand Oil Price Structure & Retail Tracker",
  description: "Auto-synced EPPO Oil Price Structure & Daily Retail Price Dashboard (PTTOR, Bangchak, Caltex)",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="th">
      <head>
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="anonymous" />
        <link
          href="https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500;600&family=IBM+Plex+Sans+Thai:wght@300;400;500;600;700&display=swap"
          rel="stylesheet"
        />
      </head>
      <body className="bg-slate-50 text-slate-900 dark:bg-slate-950 dark:text-slate-100 font-sans antialiased">
        {children}
      </body>
    </html>
  );
}
