"use client";

import { useState, useCallback } from "react";
import { saveAs } from "file-saver";
import JSZip from "jszip";
import FileUploader from "@/components/FileUploader";
import FileList from "@/components/FileList";
import ConversionResult from "@/components/ConversionResult";
import { parseXml } from "@/lib/xml-parser";
import { generateXlsx, changeExtension } from "@/lib/xlsx-generator";
import type { UploadedFile } from "@/lib/types";

export default function Home() {
  const [files, setFiles] = useState<UploadedFile[]>([]);
  const [selectedPreview, setSelectedPreview] = useState<string | null>(null);

  // ファイル選択 → 即座に変換開始（変換ボタン不要）
  const handleFilesSelected = useCallback((newFiles: File[]) => {
    const entries: UploadedFile[] = newFiles.map((file) => ({
      id: `${file.name}-${Date.now()}-${Math.random().toString(36).slice(2)}`,
      file,
      status: "converting" as const,
    }));

    setFiles((prev) => [...prev, ...entries]);

    // 初回アップロード時は最初のファイルをプレビュー用に自動選択
    if (entries.length > 0) {
      setSelectedPreview((prev) => prev ?? entries[0].id);
    }

    // 各ファイルを非同期で変換
    for (const entry of entries) {
      entry.file
        .text()
        .then((text) => {
          const result = parseXml(text, entry.file.name);
          setFiles((prev) =>
            prev.map((f) =>
              f.id === entry.id
                ? { ...f, status: "done" as const, result }
                : f
            )
          );
        })
        .catch((err) => {
          const errorMessage =
            err instanceof Error ? err.message : "不明なエラーが発生しました";
          setFiles((prev) =>
            prev.map((f) =>
              f.id === entry.id
                ? { ...f, status: "error" as const, error: errorMessage }
                : f
            )
          );
        });
    }
  }, []);

  const handleRemoveFile = useCallback((id: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== id));
    setSelectedPreview((prev) => (prev === id ? null : prev));
  }, []);

  const handleDownload = useCallback(
    (id: string) => {
      const fileEntry = files.find((f) => f.id === id);
      if (!fileEntry?.result) return;

      const blob = generateXlsx(fileEntry.result);
      const xlsxName = changeExtension(fileEntry.file.name, "xlsx");
      saveAs(blob, xlsxName);
    },
    [files]
  );

  const handleDownloadAll = useCallback(async () => {
    const doneFiles = files.filter((f) => f.status === "done" && f.result);
    if (doneFiles.length === 0) return;

    if (doneFiles.length === 1) {
      handleDownload(doneFiles[0].id);
      return;
    }

    const zip = new JSZip();
    for (const f of doneFiles) {
      if (!f.result) continue;
      const blob = generateXlsx(f.result);
      const xlsxName = changeExtension(f.file.name, "xlsx");
      const safeName = xlsxName.replace(/[<>:"/\\|?*]/g, "_");
      zip.file(safeName, blob);
    }
    const zipBlob = await zip.generateAsync({
      type: "blob",
      platform: "DOS",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    saveAs(zipBlob, "converted_xlsx_files.zip");
  }, [files, handleDownload]);

  const doneFiles = files.filter((f) => f.status === "done");
  const isConverting = files.some((f) => f.status === "converting");
  const hasFiles = files.length > 0;

  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-50 to-blue-50/30 flex flex-col">
      {/* ヘッダー */}
      <header className="border-b border-gray-200 bg-white/80 backdrop-blur-sm">
        <div className="mx-auto max-w-5xl px-4 py-5">
          <div className="flex items-center gap-3">
            <div className="rounded-xl bg-blue-600 p-2.5">
              <svg
                className="h-6 w-6 text-white"
                fill="none"
                viewBox="0 0 24 24"
                stroke="currentColor"
                strokeWidth={1.5}
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M7.5 21L3 16.5m0 0L7.5 12M3 16.5h13.5m0-13.5L21 7.5m0 0L16.5 12M21 7.5H7.5"
                />
              </svg>
            </div>
            <div>
              <h1 className="text-xl font-bold text-gray-900">
                XML → XLSX Converter
              </h1>
              <p className="text-sm text-gray-500">
                XMLファイルをExcel形式に厳密に変換
              </p>
            </div>
          </div>
        </div>
      </header>

      {/* メインコンテンツ */}
      <main className="mx-auto max-w-5xl w-full px-4 py-8 flex-1">
        <div className="space-y-6">
          {/* アップロードエリア */}
          <FileUploader
            onFilesSelected={handleFilesSelected}
            disabled={isConverting}
          />

          {/* ファイルリスト + アクション */}
          {hasFiles && (
            <div className="space-y-3">
              <FileList
                files={files}
                onRemoveFile={handleRemoveFile}
                onDownload={handleDownload}
              />

              {/* アクションバー */}
              <div className="flex flex-wrap items-center gap-3">
                {/* 一括ダウンロード（2ファイル以上の場合） */}
                {doneFiles.length >= 2 && (
                  <button
                    onClick={handleDownloadAll}
                    className="inline-flex items-center gap-2 rounded-xl border border-green-300 bg-green-50 px-5 py-2.5 text-sm font-semibold text-green-700 shadow-sm transition-all duration-200 hover:bg-green-100 hover:shadow-md active:scale-[0.98]"
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
                        d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3"
                      />
                    </svg>
                    全{doneFiles.length}ファイルをZIPでダウンロード
                  </button>
                )}

                {/* 全クリア */}
                {!isConverting && (
                  <button
                    onClick={() => {
                      setFiles([]);
                      setSelectedPreview(null);
                    }}
                    className="inline-flex items-center gap-1.5 rounded-xl px-4 py-2.5 text-sm font-medium text-gray-500 transition-colors hover:bg-red-50 hover:text-red-600"
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
                        d="M14.74 9l-.346 9m-4.788 0L9.26 9m9.968-3.21c.342.052.682.107 1.022.166m-1.022-.165L18.16 19.673a2.25 2.25 0 01-2.244 2.077H8.084a2.25 2.25 0 01-2.244-2.077L4.772 5.79m14.456 0a48.108 48.108 0 00-3.478-.397m-12 .562c.34-.059.68-.114 1.022-.165m0 0a48.11 48.11 0 013.478-.397m7.5 0v-.916c0-1.18-.91-2.164-2.09-2.201a51.964 51.964 0 00-3.32 0c-1.18.037-2.09 1.022-2.09 2.201v.916m7.5 0a48.667 48.667 0 00-7.5 0"
                      />
                    </svg>
                    全てクリア
                  </button>
                )}
              </div>
            </div>
          )}

          {/* 変換結果プレビュー */}
          {doneFiles.length > 0 && (
            <div className="space-y-3">
              <h2 className="text-base font-semibold text-gray-800">
                変換結果プレビュー
              </h2>

              {/* ファイル選択タブ（複数完了時） */}
              {doneFiles.length > 1 && (
                <div className="flex gap-2 overflow-x-auto pb-1">
                  {doneFiles.map((f) => (
                    <button
                      key={f.id}
                      onClick={() => setSelectedPreview(f.id)}
                      className={`whitespace-nowrap rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                        selectedPreview === f.id
                          ? "bg-blue-600 text-white"
                          : "bg-gray-100 text-gray-600 hover:bg-gray-200"
                      }`}
                    >
                      {f.file.name}
                    </button>
                  ))}
                </div>
              )}

              {/* プレビュー表示 */}
              {(() => {
                const previewFile =
                  doneFiles.length === 1
                    ? doneFiles[0]
                    : doneFiles.find((f) => f.id === selectedPreview);

                if (!previewFile?.result) {
                  if (doneFiles.length > 1 && !selectedPreview) {
                    return (
                      <p className="text-sm text-gray-500">
                        上のタブからファイルを選択するとプレビューが表示されます
                      </p>
                    );
                  }
                  return null;
                }

                return <ConversionResult result={previewFile.result} />;
              })()}
            </div>
          )}
        </div>
      </main>

      {/* フッター */}
      <footer className="border-t border-gray-200 bg-white/50 py-4 text-center text-xs text-gray-400">
        XML → XLSX Converter — ブラウザ上で厳密に変換。データは外部に送信されません。
      </footer>
    </div>
  );
}
