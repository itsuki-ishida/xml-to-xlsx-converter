"use client";

import type { UploadedFile } from "@/lib/types";

interface FileListProps {
  files: UploadedFile[];
  onRemoveFile: (id: string) => void;
}

const statusConfig: Record<
  UploadedFile["status"],
  { label: string; color: string; icon: string }
> = {
  pending: { label: "待機中", color: "text-gray-500 bg-gray-100", icon: "⏳" },
  converting: {
    label: "変換中",
    color: "text-blue-600 bg-blue-100",
    icon: "⚙️",
  },
  done: {
    label: "変換完了",
    color: "text-green-600 bg-green-100",
    icon: "✅",
  },
  error: {
    label: "エラー",
    color: "text-red-600 bg-red-100",
    icon: "❌",
  },
};

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export default function FileList({ files, onRemoveFile }: FileListProps) {
  if (files.length === 0) return null;

  return (
    <div className="space-y-2">
      <h3 className="text-sm font-medium text-gray-600">
        アップロード済みファイル ({files.length})
      </h3>
      <div className="space-y-2">
        {files.map((f) => {
          const status = statusConfig[f.status];
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
                    {f.result &&
                      ` — ${f.result.sheets.length} シート`}
                  </p>
                  {f.error && (
                    <p className="mt-0.5 text-xs text-red-500">{f.error}</p>
                  )}
                </div>
              </div>

              <div className="flex items-center gap-3 flex-shrink-0">
                {/* ステータスバッジ */}
                <span
                  className={`inline-flex items-center gap-1 rounded-full px-2.5 py-0.5 text-xs font-medium ${status.color}`}
                >
                  <span>{status.icon}</span>
                  {status.label}
                </span>

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
