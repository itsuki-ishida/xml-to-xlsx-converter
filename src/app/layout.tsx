import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "XML → XLSX Converter",
  description: "XMLファイルをExcel（.xlsx）形式に変換するツール",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="ja">
      <body className="bg-gray-50 text-gray-900 antialiased">{children}</body>
    </html>
  );
}
