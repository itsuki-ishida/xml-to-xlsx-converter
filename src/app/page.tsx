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
  const [isConverting, setIsConverting] = useState(false);
  const [selectedPreview, setSelectedPreview] = useState<string | null>(null);

  const handleFilesSelected = useCallback((newFiles: File[]) => {
    const uploaded: UploadedFile[] = newFiles.map((file) => ({
      id: `${file.name}-${Date.now()}-${Math.random().toString(36).slice(2)}`,
      file,
      status: "pending" as const,
    }));
    setFiles((prev) => [...prev, ...uploaded]);
  }, []);

  const handleRemoveFile = useCallback((id: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== id));
    setSelectedPreview((prev) => (prev === id ? null : prev));
  }, []);

  const handleConvert = useCallback(async () => {
    setIsConverting(true);

    // 最新のstateからpendingファイルを取得するためにPromiseで同期
    const filesToConvert = await new Promise<UploadedFile[]>((resolve) => {
      setFiles((prev) => {
        const pending = prev.filter((f) => f.status === "pending");
        resolve(pending);
        // ステータスを converting に更新
        return prev.map((f) =>
          f.status === "pending"
            ? { ...f, status: "converting" as const }
            : f
        );
      });
    });

    if (filesToConvert.length === 0) {
      setIsConverting(false);
      return;
    }

    // 各ファイルを順次変換
    for (const fileEntry of filesToConvert) {
      try {
        const text = await fileEntry.file.text();
        const result = parseXml(text, fileEntry.file.name);

        setFiles((prev) =>
          prev.map((f) =>
            f.id === fileEntry.id
              ? { ...f, status: "done" as const, result }
              : f
          )
        );
      } catch (err) {
        const errorMessage =
          err instanceof Error ? err.message : "不明なエラーが発生しました";
        setFiles((prev) =>
          prev.map((f) =>
            f.id === fileEntry.id
              ? { ...f, status: "error" as const, error: errorMessage }
              : f
          )
        );
      }
    }

    setIsConverting(false);
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

    // 複数ファイルはZIPにまとめてダウンロード
    const zip = new JSZip();
    for (const f of doneFiles) {
      if (!f.result) continue;
      const blob = generateXlsx(f.result);
      const xlsxName = changeExtension(f.file.name, "xlsx");
      // Windows禁止文字をファイル名から除去
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
  const pendingFiles = files.filter((f) => f.status === "pending");
  const hasFiles = files.length > 0;

  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-50 to-blue-50/30">
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
      <main className="mx-auto max-w-5xl px-4 py-8">
        <div className="space-y-6">
          {/* アップロードエリア */}
          <FileUploader
            onFilesSelected={handleFilesSelected}
            disabled={isConverting}
          />

          {/* ファイルリスト */}
          {hasFiles && (
            <FileList files={files} onRemoveFile={handleRemoveFile} />
          )}

          {/* アクションボタン */}
          {hasFiles && (
            <div className="flex flex-wrap items-center gap-3">
              {/* 変換ボタン */}
              {pendingFiles.length > 0 && (
                <button
                  onClick={handleConvert}
                  disabled={isConverting}
                  className={`
                    inline-flex items-center gap-2 rounded-xl px-6 py-3 text-sm font-semibold
                    shadow-sm transition-all duration-200
                    ${
                      isConverting
                        ? "bg-gray-400 text-white cursor-not-allowed"
                        : "bg-blue-600 text-white hover:bg-blue-700 hover:shadow-md active:scale-[0.98]"
                    }
                  `}
                >
                  {isConverting ? (
                    <>
                      <svg
                        className="h-4 w-4 animate-spin"
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
                      変換中...
                    </>
                  ) : (
                    <>
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
                          d="M7.5 21L3 16.5m0 0L7.5 12M3 16.5h13.5m0-13.5L21 7.5m0 0L16.5 12M21 7.5H7.5"
                        />
                      </svg>
                      .xlsx形式に変換する
                    </>
                  )}
                </button>
              )}

              {/* 一括ダウンロード */}
              {doneFiles.length > 0 && (
                <button
                  onClick={handleDownloadAll}
                  className="inline-flex items-center gap-2 rounded-xl border border-green-300 bg-green-50 px-6 py-3 text-sm font-semibold text-green-700 shadow-sm transition-all duration-200 hover:bg-green-100 hover:shadow-md active:scale-[0.98]"
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
                  {doneFiles.length === 1
                    ? "XLSXをダウンロード"
                    : `全${doneFiles.length}ファイルをZIPでダウンロード`}
                </button>
              )}

              {/* 全クリア */}
              {!isConverting && (
                <button
                  onClick={() => {
                    setFiles([]);
                    setSelectedPreview(null);
                  }}
                  className="inline-flex items-center gap-1.5 rounded-xl px-4 py-3 text-sm font-medium text-gray-500 transition-colors hover:bg-red-50 hover:text-red-600"
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
          )}

          {/* 変換結果プレビュー */}
          {doneFiles.length > 0 && (
            <div className="space-y-4">
              <h2 className="text-lg font-semibold text-gray-800">
                変換結果プレビュー
              </h2>

              {/* ファイル選択タブ（複数完了時） */}
              {doneFiles.length > 1 && (
                <div className="flex gap-2 overflow-x-auto">
                  {doneFiles.map((f) => (
                    <button
                      key={f.id}
                      onClick={() =>
                        setSelectedPreview(
                          selectedPreview === f.id ? null : f.id
                        )
                      }
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

                return (
                  <div className="space-y-3">
                    <div className="flex items-center justify-between">
                      <p className="text-sm text-gray-600">
                        <span className="font-medium">
                          {previewFile.file.name}
                        </span>
                        {" — "}
                        {previewFile.result.sheets.length} シート,{" "}
                        {previewFile.result.sheets.reduce(
                          (acc, s) => acc + s.rows.length,
                          0
                        )}{" "}
                        行
                      </p>
                      <button
                        onClick={() => handleDownload(previewFile.id)}
                        className="inline-flex items-center gap-1.5 rounded-lg bg-green-600 px-3 py-1.5 text-xs font-medium text-white transition-colors hover:bg-green-700"
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
                        ダウンロード
                      </button>
                    </div>
                    <ConversionResult result={previewFile.result} />
                  </div>
                );
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
