import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Transaction Generator",
  description: "Generate statement-style transactions from a guided form.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
