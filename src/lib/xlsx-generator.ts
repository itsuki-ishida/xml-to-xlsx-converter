import * as XLSX from "xlsx";
import type { ParsedXmlResult } from "./types";

/**
 * ParsedXmlResult から XLSX ワークブックを生成しダウンロード用の Blob を返す
 */
export function generateXlsx(parsed: ParsedXmlResult): Blob {
  const wb = XLSX.utils.book_new();

  for (const sheet of parsed.sheets) {
    // ヘッダー行 + データ行の2次元配列を構築
    const aoa: string[][] = [];

    // ヘッダー行
    aoa.push(sheet.headers);

    // データ行
    for (const row of sheet.rows) {
      const rowData: string[] = sheet.headers.map((h) => row[h] ?? "");
      aoa.push(rowData);
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // 列幅の自動調整
    const colWidths: number[] = sheet.headers.map((h) => {
      let maxLen = h.length;
      for (const row of sheet.rows) {
        const val = row[h] ?? "";
        maxLen = Math.max(maxLen, val.length);
      }
      // 最大幅を60文字に制限、最小幅は8文字
      return Math.min(Math.max(maxLen + 2, 8), 60);
    });
    ws["!cols"] = colWidths.map((w) => ({ wch: w }));

    // ヘッダー行をフリーズ（固定表示）
    if (!ws["!views"]) ws["!views"] = [];
    (ws["!views"] as Array<Record<string, unknown>>).push({
      state: "frozen",
      ySplit: 1,
    });

    // オートフィルターを設定
    if (sheet.headers.length > 0 && sheet.rows.length > 0) {
      ws["!autofilter"] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: sheet.rows.length, c: sheet.headers.length - 1 },
        }),
      };
    }

    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
  }

  // ワークブックが空の場合、空のシートを追加
  if (parsed.sheets.length === 0) {
    const ws = XLSX.utils.aoa_to_sheet([["（データなし）"]]);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  }

  // Blobとして出力
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([wbout], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

/**
 * ファイル名から拡張子を変更する
 */
export function changeExtension(
  fileName: string,
  newExt: string
): string {
  const dotIndex = fileName.lastIndexOf(".");
  if (dotIndex < 0) return `${fileName}.${newExt}`;
  return `${fileName.substring(0, dotIndex)}.${newExt}`;
}
