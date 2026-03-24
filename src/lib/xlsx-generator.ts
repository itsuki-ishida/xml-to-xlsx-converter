import * as XLSX from "xlsx";
import type { ParsedXmlResult, SheetData } from "./types";

/**
 * ParsedXmlResult から XLSX ワークブックを生成しダウンロード用の Blob を返す
 *
 * 2シート構成:
 *   シート1「概要」: ファイル情報 + 単一インスタンス要素をセクション+キーバリュー形式で表示
 *   シート2「明細」: 複数インスタンス要素をセクション区切りのテーブル形式で表示（行番号付き）
 *
 * 全て単一の場合は「概要」のみ、全て複数の場合は「明細」のみ生成
 * 属性は全行で同一値の場合のみ除外（メタデータノイズ除去）
 */
export function generateXlsx(parsed: ParsedXmlResult): Blob {
  const wb = XLSX.utils.book_new();

  const singleSheets = parsed.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = parsed.sheets.filter((s) => s.rows.length >= 2);

  // --- シート1: 概要（ファイル情報 + 単一インスタンス要素） ---
  if (singleSheets.length > 0) {
    const aoa = buildOverviewSheet(singleSheets, parsed);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 30 }, { wch: 55 }];
    ws["!views"] = [{ state: "frozen", ySplit: 1 }];
    XLSX.utils.book_append_sheet(wb, ws, "概要");
  }

  // --- シート2: 明細（複数インスタンス要素） ---
  if (multiSheets.length > 0) {
    const aoa = buildDetailsSheet(multiSheets);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    // 列幅: # 列 + 最大データ列数
    const maxDataCols = Math.max(
      ...multiSheets.map((s) => getDisplayHeaders(s).length)
    );
    const colWidths: XLSX.ColInfo[] = [{ wch: 5 }]; // # 列
    for (let i = 0; i < maxDataCols; i++) {
      colWidths.push({ wch: 22 });
    }
    ws["!cols"] = colWidths;
    XLSX.utils.book_append_sheet(wb, ws, "明細");
  }

  // ワークブックが空の場合
  if (parsed.sheets.length === 0) {
    const ws = XLSX.utils.aoa_to_sheet([["（データなし）"]]);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  }

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([wbout], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

/**
 * 表示用ヘッダーを取得する
 * 全行で同一値の @属性列は除外（メタデータノイズ）
 * 意味のある属性（値がバラバラのもの）は保持
 */
function getDisplayHeaders(sheet: SheetData): string[] {
  const headers: string[] = [];
  for (const h of sheet.headers) {
    if (h.startsWith("@")) {
      const values = new Set(sheet.rows.map((r) => r[h] ?? ""));
      if (values.size <= 1) continue; // 全行同一値 → 除外
    }
    headers.push(h);
  }
  return headers;
}

/**
 * ヘッダー表示名を返す（@プレフィックスを除去）
 */
function displayHeaderName(h: string): string {
  return h.startsWith("@") ? h.substring(1) : h;
}

/**
 * 概要シートの2次元配列を構築
 * ファイル情報 + セクション名 + キーバリュー形式
 */
function buildOverviewSheet(
  sheets: SheetData[],
  parsed: ParsedXmlResult
): string[][] {
  const aoa: string[][] = [];

  // ヘッダー行
  aoa.push(["セクション / 項目名", "値"]);

  // ファイル情報セクション
  aoa.push(["■ ファイル情報", ""]);
  aoa.push(["  ファイル名", parsed.fileName]);
  aoa.push(["  ルート要素", parsed.rootElement]);
  aoa.push([]);

  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];
    const row = sheet.rows[0];
    const displayHeaders = getDisplayHeaders(sheet);

    // セクションヘッダー
    aoa.push([`■ ${sheet.name}`, ""]);

    // 各フィールドをキーバリューで出力
    for (const header of displayHeaders) {
      aoa.push([`  ${displayHeaderName(header)}`, row[header] ?? ""]);
    }

    // 次のセクションとの間に空行
    if (si < sheets.length - 1) {
      aoa.push([]);
    }
  }

  return aoa;
}

/**
 * 明細シートの2次元配列を構築
 * セクション名 + 行番号付きテーブル形式
 */
function buildDetailsSheet(sheets: SheetData[]): string[][] {
  const aoa: string[][] = [];

  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];

    // セクション区切り（最初以外は空行を挿入）
    if (si > 0) {
      aoa.push([]);
    }

    const displayHeaders = getDisplayHeaders(sheet);
    const displayNames = displayHeaders.map(displayHeaderName);

    // セクションヘッダー
    aoa.push([
      `■ ${sheet.name} (${sheet.rows.length}件)`,
      ...Array(displayNames.length).fill(""),
    ]);

    // テーブルヘッダー（# + データヘッダー）
    aoa.push(["#", ...displayNames]);

    // データ行（行番号付き）
    for (let ri = 0; ri < sheet.rows.length; ri++) {
      const row = sheet.rows[ri];
      aoa.push([
        String(ri + 1),
        ...displayHeaders.map((h) => row[h] ?? ""),
      ]);
    }
  }

  return aoa;
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

/**
 * パース結果からXLSXの構造概要を返す
 */
export function getXlsxSummary(parsed: ParsedXmlResult): {
  overviewSections: number;
  detailSections: number;
  detailRows: number;
  sheetCount: number;
} {
  const singleSheets = parsed.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = parsed.sheets.filter((s) => s.rows.length >= 2);
  return {
    overviewSections: singleSheets.length,
    detailSections: multiSheets.length,
    detailRows: multiSheets.reduce((acc, s) => acc + s.rows.length, 0),
    sheetCount:
      (singleSheets.length > 0 ? 1 : 0) +
      (multiSheets.length > 0 ? 1 : 0),
  };
}
