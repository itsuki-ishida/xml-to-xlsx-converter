"use client";

import type { UploadedFile } from "@/lib/types";
import { getXlsxSummary } from "@/lib/xlsx-generator";

interface FileListProps {
  files: UploadedFile[];
  onRemoveFile: (id: string) => void;
  onDownload: (id: string) => void;
}

const statusConfig: Record<
  UploadedFile["status"],
  { label: string; color: string }
> = {
  pending: { label: "待機中", color: "text-gray-500 bg-gray-100" },
  converting: {
    label: "変換中...",
    color: "text-blue-600 bg-blue-100",
  },
  done: {
    label: "変換完了",
    color: "text-green-600 bg-green-100",
  },
  error: {
    label: "エラー",
    color: "text-red-600 bg-red-100",
  },
};

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export default function FileList({
  files,
  onRemoveFile,
  onDownload,
}: FileListProps) {
  if (files.length === 0) return null;

  return (
    <div className="space-y-2">
      <h3 className="text-sm font-medium text-gray-600">
        ファイル ({files.length})
      </h3>
      <div className="space-y-2">
        {files.map((f) => {
          const status = statusConfig[f.status];
          const summary = f.result ? getXlsxSummary(f.result) : null;

          return (
            <div
              key={f.id}
              className="flex items-center justify-between rounded-xl border border-gray-200 bg-white px-4 py-3 shadow-sm"
            >
              <div className="flex items-center gap-3 min-w-0">
                {/* XMLアイコン */}
                <div className="flex-shrink-0 rounded-lg bg-orange-100 p-2">
                  <svg
                    className="h-5 w-5 text-orange-600"
                    fill="none"
                    viewBox="0 0 24 24"
                    stroke="currentColor"
                    strokeWidth={1.5}
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z"
                    />
                  </svg>
                </div>

                <div className="min-w-0">
                  <p className="truncate text-sm font-medium text-gray-800">
                    {f.file.name}
                  </p>
                  <p className="text-xs text-gray-400">
                    {formatFileSize(f.file.size)}
                    {summary && (
                      <>
                        {" — "}
                        {summary.sheetCount}シート
                        {summary.overviewSections > 0 &&
                          ` (概要: ${summary.overviewSections}セクション`}
                        {summary.overviewSections > 0 &&
                          summary.detailSections > 0 &&
                          ", "}
                        {summary.overviewSections === 0 &&
                          summary.detailSections > 0 &&
                          " ("}
                        {summary.detailSections > 0 &&
                          `明細: ${summary.detailRows}行`}
                        {(summary.overviewSections > 0 ||
                          summary.detailSections > 0) &&
                          ")"}
                      </>
                    )}
                  </p>
                  {f.error && (
                    <p className="mt-0.5 text-xs text-red-500">{f.error}</p>
                  )}
                </div>
              </div>

              <div className="flex items-center gap-2 flex-shrink-0">
                {/* ステータスバッジ */}
                {f.status === "converting" ? (
                  <span className="inline-flex items-center gap-1.5 rounded-full px-2.5 py-0.5 text-xs font-medium text-blue-600 bg-blue-100">
                    <svg
                      className="h-3 w-3 animate-spin"
                      fill="none"
                      viewBox="0 0 24 24"
                    >
                      <circle
                        className="opacity-25"
                        cx="12"
                        cy="12"
                        r="10"
                        stroke="currentColor"
                        strokeWidth="4"
                      />
                      <path
                        className="opacity-75"
                        fill="currentColor"
                        d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                      />
                    </svg>
                    {status.label}
                  </span>
                ) : (
                  <span
                    className={`inline-flex items-center gap-1 rounded-full px-2.5 py-0.5 text-xs font-medium ${status.color}`}
                  >
                    {status.label}
                  </span>
                )}

                {/* ダウンロードボタン（完了時のみ） */}
                {f.status === "done" && f.result && (
                  <button
                    onClick={() => onDownload(f.id)}
                    className="inline-flex items-center gap-1.5 rounded-lg bg-green-600 px-3 py-1.5 text-xs font-medium text-white transition-colors hover:bg-green-700"
                    title="XLSXをダウンロード"
                  >
                    <svg
                      className="h-3.5 w-3.5"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={2}
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3"
                      />
                    </svg>
                    .xlsx
                  </button>
                )}

                {/* 削除ボタン */}
                {f.status !== "converting" && (
                  <button
                    onClick={() => onRemoveFile(f.id)}
                    className="rounded-lg p-1.5 text-gray-400 hover:bg-red-50 hover:text-red-500 transition-colors"
                    title="削除"
                  >
                    <svg
                      className="h-4 w-4"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={2}
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M6 18L18 6M6 6l12 12"
                      />
                    </svg>
                  </button>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
