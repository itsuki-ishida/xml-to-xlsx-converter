"use client";

import { useState } from "react";
import type { ParsedXmlResult, SheetData } from "@/lib/types";

interface ConversionResultProps {
  result: ParsedXmlResult;
}

export default function ConversionResult({ result }: ConversionResultProps) {
  const [activeTab, setActiveTab] = useState<"overview" | "details">(
    "overview"
  );

  if (result.sheets.length === 0) {
    return (
      <div className="rounded-xl border border-gray-200 bg-white p-6 text-center text-gray-500">
        データが見つかりませんでした
      </div>
    );
  }

  const singleSheets = result.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = result.sheets.filter((s) => s.rows.length >= 2);

  const tabs: { key: "overview" | "details"; label: string; count: number }[] =
    [];
  if (singleSheets.length > 0) {
    tabs.push({ key: "overview", label: "概要", count: singleSheets.length });
  }
  if (multiSheets.length > 0) {
    tabs.push({ key: "details", label: "明細", count: multiSheets.length });
  }

  // タブが1つしかない場合はそれを表示
  const effectiveTab =
    tabs.length === 1 ? tabs[0].key : activeTab;

  return (
    <div className="rounded-xl border border-gray-200 bg-white shadow-sm overflow-hidden">
      {/* タブ */}
      {tabs.length > 1 && (
        <div className="flex border-b border-gray-200 bg-gray-50">
          {tabs.map((tab) => (
            <button
              key={tab.key}
              onClick={() => setActiveTab(tab.key)}
              className={`px-5 py-2.5 text-sm font-medium transition-colors border-b-2 ${
                effectiveTab === tab.key
                  ? "border-blue-500 text-blue-600 bg-white"
                  : "border-transparent text-gray-500 hover:text-gray-700 hover:bg-gray-100"
              }`}
            >
              {tab.label}
              <span className="ml-1.5 text-xs text-gray-400">
                ({tab.count}セクション)
              </span>
            </button>
          ))}
        </div>
      )}

      {/* コンテンツ */}
      <div className="overflow-x-auto max-h-[500px] overflow-y-auto">
        {effectiveTab === "overview" && singleSheets.length > 0 && (
          <OverviewView sheets={singleSheets} />
        )}
        {effectiveTab === "details" && multiSheets.length > 0 && (
          <DetailsView sheets={multiSheets} />
        )}
      </div>
    </div>
  );
}

/** 概要ビュー: セクション+キーバリュー */
function OverviewView({ sheets }: { sheets: SheetData[] }) {
  return (
    <div className="divide-y divide-gray-100">
      {sheets.map((sheet, si) => {
        const row = sheet.rows[0];
        const dataHeaders = sheet.headers.filter((h) => !h.startsWith("@"));
        return (
          <div key={si} className="px-4 py-3">
            <h4 className="mb-2 text-xs font-bold text-blue-700 uppercase tracking-wider">
              {sheet.name}
            </h4>
            <div className="grid grid-cols-[auto_1fr] gap-x-4 gap-y-0.5 text-sm">
              {dataHeaders.map((h) => (
                <div key={h} className="contents">
                  <span className="text-gray-500 whitespace-nowrap">{h}</span>
                  <span
                    className="text-gray-800 truncate"
                    title={row[h] ?? ""}
                  >
                    {row[h] ?? ""}
                  </span>
                </div>
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/** 明細ビュー: セクション区切りのテーブル */
function DetailsView({ sheets }: { sheets: SheetData[] }) {
  return (
    <div className="space-y-4 p-4">
      {sheets.map((sheet, si) => {
        const dataHeaders = sheet.headers.filter((h) => !h.startsWith("@"));
        return (
          <div key={si}>
            <h4 className="mb-2 text-xs font-bold text-blue-700 uppercase tracking-wider">
              {sheet.name}{" "}
              <span className="text-gray-400 font-normal">
                ({sheet.rows.length}件)
              </span>
            </h4>
            <div className="overflow-x-auto rounded-lg border border-gray-200">
              <table className="min-w-full divide-y divide-gray-200 text-sm">
                <thead className="bg-gray-50">
                  <tr>
                    {dataHeaders.map((h) => (
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
                    <tr
                      key={ri}
                      className="hover:bg-blue-50/50 transition-colors"
                    >
                      {dataHeaders.map((h) => (
                        <td
                          key={h}
                          className="px-3 py-1.5 text-gray-700 whitespace-nowrap max-w-xs truncate"
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
      })}
    </div>
  );
}
