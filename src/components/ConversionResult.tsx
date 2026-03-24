"use client";

import { useState } from "react";
import type { ParsedXmlResult, SheetData } from "@/lib/types";
import { getXlsxSummary } from "@/lib/xlsx-generator";

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
  const summary = getXlsxSummary(result);

  const tabs: { key: "overview" | "details"; label: string; desc: string }[] =
    [];
  if (singleSheets.length > 0) {
    tabs.push({
      key: "overview",
      label: "概要",
      desc: `${singleSheets.length}セクション`,
    });
  }
  if (multiSheets.length > 0) {
    tabs.push({
      key: "details",
      label: "明細",
      desc: `${multiSheets.length}テーブル・${summary.detailRows}行`,
    });
  }

  // タブが1つしかない場合はそれを表示
  const effectiveTab = tabs.length === 1 ? tabs[0].key : activeTab;

  return (
    <div className="space-y-2">
      {/* XLSX構成情報 */}
      <div className="text-xs text-gray-500">
        XLSX出力: {summary.sheetCount}シート構成
        {summary.overviewSections > 0 &&
          ` / 概要 ${summary.overviewSections}セクション`}
        {summary.detailSections > 0 &&
          ` / 明細 ${summary.detailSections}テーブル (${summary.detailRows}行)`}
      </div>

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
                  ({tab.desc})
                </span>
              </button>
            ))}
          </div>
        )}

        {/* シート名ラベル（タブが1つの場合） */}
        {tabs.length === 1 && (
          <div className="border-b border-gray-200 bg-gray-50 px-5 py-2 text-sm font-medium text-gray-600">
            {tabs[0].label}
            <span className="ml-1.5 text-xs text-gray-400">
              ({tabs[0].desc})
            </span>
          </div>
        )}

        {/* コンテンツ */}
        <div className="overflow-x-auto max-h-[500px] overflow-y-auto">
          {effectiveTab === "overview" && singleSheets.length > 0 && (
            <OverviewView sheets={singleSheets} parsed={result} />
          )}
          {effectiveTab === "details" && multiSheets.length > 0 && (
            <DetailsView sheets={multiSheets} />
          )}
        </div>
      </div>
    </div>
  );
}

/**
 * 表示用ヘッダーを取得（全行同一の@属性を除外）
 * xlsx-generator.ts の getDisplayHeaders と同じロジック
 */
function getDisplayHeaders(sheet: SheetData): string[] {
  const headers: string[] = [];
  for (const h of sheet.headers) {
    if (h.startsWith("@")) {
      const values = new Set(sheet.rows.map((r) => r[h] ?? ""));
      if (values.size <= 1) continue;
    }
    headers.push(h);
  }
  return headers;
}

function displayHeaderName(h: string): string {
  return h.startsWith("@") ? h.substring(1) : h;
}

/** 概要ビュー: ファイル情報 + セクション+キーバリュー */
function OverviewView({
  sheets,
  parsed,
}: {
  sheets: SheetData[];
  parsed: ParsedXmlResult;
}) {
  return (
    <div className="divide-y divide-gray-100">
      {/* ファイル情報セクション */}
      <div className="px-4 py-3">
        <h4 className="mb-2 text-xs font-bold text-gray-500 uppercase tracking-wider">
          ファイル情報
        </h4>
        <div className="grid grid-cols-[auto_1fr] gap-x-4 gap-y-0.5 text-sm">
          <span className="text-gray-500 whitespace-nowrap">ファイル名</span>
          <span className="text-gray-800">{parsed.fileName}</span>
          <span className="text-gray-500 whitespace-nowrap">ルート要素</span>
          <span className="text-gray-800">{parsed.rootElement}</span>
        </div>
      </div>

      {/* 各セクション */}
      {sheets.map((sheet, si) => {
        const row = sheet.rows[0];
        const displayHeaders = getDisplayHeaders(sheet);
        return (
          <div key={si} className="px-4 py-3">
            <h4 className="mb-2 text-xs font-bold text-blue-700 uppercase tracking-wider">
              {sheet.name}
            </h4>
            <div className="grid grid-cols-[auto_1fr] gap-x-4 gap-y-0.5 text-sm">
              {displayHeaders.map((h) => (
                <div key={h} className="contents">
                  <span className="text-gray-500 whitespace-nowrap">
                    {displayHeaderName(h)}
                  </span>
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

/** 明細ビュー: セクション区切りの行番号付きテーブル */
function DetailsView({ sheets }: { sheets: SheetData[] }) {
  return (
    <div className="space-y-4 p-4">
      {sheets.map((sheet, si) => {
        const displayHeaders = getDisplayHeaders(sheet);
        const displayNames = displayHeaders.map(displayHeaderName);
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
                    <th className="px-2 py-2 text-center text-xs font-semibold text-gray-400 w-10">
                      #
                    </th>
                    {displayNames.map((h) => (
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
                      <td className="px-2 py-1.5 text-center text-xs text-gray-400">
                        {ri + 1}
                      </td>
                      {displayHeaders.map((h) => (
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
