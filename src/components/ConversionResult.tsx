"use client";

import { useState } from "react";
import type { ParsedXmlResult, SheetData } from "@/lib/types";
import { getXlsxSummary, getDisplayHeaders } from "@/lib/xlsx-generator";
import {
  translateSection,
  translateFieldShort,
  stripFieldPrefix,
  TRANSLATION_SOURCE,
  analyzeTranslationCoverage,
} from "@/lib/translations";

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

  const effectiveTab = tabs.length === 1 ? tabs[0].key : activeTab;

  // SAP翻訳適用状況を分析
  const sectionNames = result.sheets.map((s) => s.name);
  const fieldNames = result.sheets.flatMap((s) => s.headers);
  const coverage = analyzeTranslationCoverage(sectionNames, fieldNames);

  return (
    <div className="space-y-2">
      {/* XLSX構成情報 + 形式検出インジケーター */}
      <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-gray-500">
        <span>
          XLSX出力: {summary.sheetCount}シート構成
          {summary.overviewSections > 0 &&
            ` / 概要 ${summary.overviewSections}セクション`}
          {summary.detailSections > 0 &&
            ` / 明細 ${summary.detailSections}テーブル (${summary.detailRows}行)`}
        </span>
        {coverage.isSapDetected ? (
          <span className="inline-flex items-center gap-1 rounded-full bg-blue-50 px-2.5 py-0.5 text-blue-600 border border-blue-200">
            <svg className="h-3 w-3" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
            SAP形式を検出 — {TRANSLATION_SOURCE}
            （{coverage.translatedSections}/{coverage.totalSections}セクション,{" "}
            {coverage.translatedFields}/{coverage.totalFields}フィールド翻訳済み）
          </span>
        ) : (
          <span className="inline-flex items-center gap-1 rounded-full bg-gray-50 px-2.5 py-0.5 text-gray-500 border border-gray-200">
            <svg className="h-3 w-3" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v3.75m9-.75a9 9 0 11-18 0 9 9 0 0118 0zm-9 3.75h.008v.008H12v-.008z" />
            </svg>
            汎用XMLとして変換（SAP翻訳辞書に該当するフィールドがありません）
          </span>
        )}
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

/** 概要ビュー: ファイル情報 + 3列構成（日本語名 / 技術名 / 値） */
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

      {/* 各セクション（日本語名 + 技術名を両方表示） */}
      {sheets.map((sheet, si) => {
        const row = sheet.rows[0];
        const displayHeaders = getDisplayHeaders(sheet);
        return (
          <div key={si} className="px-4 py-3">
            <h4 className="mb-2 text-xs font-bold text-blue-700 tracking-wider">
              {translateSection(sheet.name)}
            </h4>
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-xs text-gray-400">
                  <th className="pb-1 pr-3 font-medium w-1/4">日本語名</th>
                  <th className="pb-1 pr-3 font-medium w-1/4">技術名</th>
                  <th className="pb-1 font-medium">値</th>
                </tr>
              </thead>
              <tbody>
                {displayHeaders.map((h) => {
                  const raw = stripFieldPrefix(h);
                  const ja = translateFieldShort(h);
                  return (
                    <tr key={h} className="border-t border-gray-50">
                      <td className="py-0.5 pr-3 text-gray-600 whitespace-nowrap">
                        {ja}
                      </td>
                      <td className="py-0.5 pr-3 text-gray-400 font-mono text-xs whitespace-nowrap">
                        {raw}
                      </td>
                      <td
                        className="py-0.5 text-gray-800 truncate max-w-xs"
                        title={row[h] ?? ""}
                      >
                        {row[h] ?? ""}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        );
      })}
    </div>
  );
}

/** 明細ビュー: 2行ヘッダー（日本語名 + 技術名）付きテーブル */
function DetailsView({ sheets }: { sheets: SheetData[] }) {
  return (
    <div className="space-y-4 p-4">
      {sheets.map((sheet, si) => {
        const displayHeaders = getDisplayHeaders(sheet);
        return (
          <div key={si}>
            <h4 className="mb-2 text-xs font-bold text-blue-700 tracking-wider">
              {translateSection(sheet.name)}{" "}
              <span className="text-gray-400 font-normal">
                ({sheet.rows.length}件)
              </span>
            </h4>
            <div className="overflow-x-auto rounded-lg border border-gray-200">
              <table className="min-w-full divide-y divide-gray-200 text-sm">
                <thead className="bg-gray-50">
                  {/* 1行目: 日本語名 */}
                  <tr>
                    <th
                      rowSpan={2}
                      className="px-2 py-2 text-center text-xs font-semibold text-gray-400 w-10 align-middle"
                    >
                      #
                    </th>
                    {displayHeaders.map((h) => (
                      <th
                        key={h}
                        className="px-3 py-1 text-left text-xs font-semibold text-gray-600 whitespace-nowrap"
                      >
                        {translateFieldShort(h)}
                      </th>
                    ))}
                  </tr>
                  {/* 2行目: 技術名（英名） */}
                  <tr>
                    {displayHeaders.map((h) => (
                      <th
                        key={h}
                        className="px-3 py-1 text-left text-[10px] font-normal font-mono text-gray-400 whitespace-nowrap border-b border-gray-200"
                      >
                        {stripFieldPrefix(h)}
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
