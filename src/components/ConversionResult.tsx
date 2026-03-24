"use client";

import { useState } from "react";
import type { ParsedXmlResult } from "@/lib/types";

interface ConversionResultProps {
  result: ParsedXmlResult;
}

export default function ConversionResult({ result }: ConversionResultProps) {
  const [activeSheet, setActiveSheet] = useState(0);

  if (result.sheets.length === 0) {
    return (
      <div className="rounded-xl border border-gray-200 bg-white p-6 text-center text-gray-500">
        データが見つかりませんでした
      </div>
    );
  }

  const sheet = result.sheets[activeSheet];

  return (
    <div className="rounded-xl border border-gray-200 bg-white shadow-sm overflow-hidden">
      {/* シートタブ */}
      <div className="flex overflow-x-auto border-b border-gray-200 bg-gray-50">
        {result.sheets.map((s, i) => (
          <button
            key={i}
            onClick={() => setActiveSheet(i)}
            className={`whitespace-nowrap px-4 py-2.5 text-xs font-medium transition-colors border-b-2 ${
              i === activeSheet
                ? "border-blue-500 text-blue-600 bg-white"
                : "border-transparent text-gray-500 hover:text-gray-700 hover:bg-gray-100"
            }`}
          >
            {s.name}
            <span className="ml-1.5 text-gray-400">({s.rows.length})</span>
          </button>
        ))}
      </div>

      {/* テーブル */}
      <div className="overflow-x-auto max-h-[400px] overflow-y-auto">
        <table className="min-w-full divide-y divide-gray-200 text-sm">
          <thead className="bg-gray-50 sticky top-0">
            <tr>
              <th className="px-3 py-2 text-left text-xs font-semibold text-gray-500 uppercase tracking-wider">
                #
              </th>
              {sheet.headers.map((h) => (
                <th
                  key={h}
                  className="px-3 py-2 text-left text-xs font-semibold text-gray-500 uppercase tracking-wider whitespace-nowrap"
                >
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100 bg-white">
            {sheet.rows.map((row, ri) => (
              <tr key={ri} className="hover:bg-blue-50/50 transition-colors">
                <td className="px-3 py-2 text-gray-400 whitespace-nowrap">
                  {ri + 1}
                </td>
                {sheet.headers.map((h) => (
                  <td
                    key={h}
                    className="px-3 py-2 text-gray-700 whitespace-nowrap max-w-xs truncate"
                    title={row[h] ?? ""}
                  >
                    {row[h] ?? ""}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
