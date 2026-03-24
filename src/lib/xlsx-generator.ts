import * as XLSX from "xlsx";
import type { ParsedXmlResult, SheetData } from "./types";
import {
  translateSection,
  translateField,
  translateFieldShort,
} from "./translations";

/**
 * ParsedXmlResult から XLSX ワークブックを生成しダウンロード用の Blob を返す
 *
 * 2シート構成:
 *   シート1「概要」: ファイル情報 + 単一インスタンス要素をセクション+キーバリュー形式で表示
 *   シート2「明細」: 複数インスタンス要素をセクション区切りのテーブル形式で表示（行番号付き）
 *
 * フィールド名・セクション名は翻訳辞書に基づいて日本語に翻訳される。
 * 翻訳が存在しないフィールドは元の技術名をそのまま表示する。
 */
export function generateXlsx(parsed: ParsedXmlResult): Blob {
  const wb = XLSX.utils.book_new();

  const singleSheets = parsed.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = parsed.sheets.filter((s) => s.rows.length >= 2);

  // --- シート1: 概要（ファイル情報 + 単一インスタンス要素） ---
  if (singleSheets.length > 0) {
    const aoa = buildOverviewSheet(singleSheets, parsed);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 36 }, { wch: 55 }];
    ws["!views"] = [{ state: "frozen", ySplit: 1 }];
    XLSX.utils.book_append_sheet(wb, ws, "概要");
  }

  // --- シート2: 明細（複数インスタンス要素） ---
  if (multiSheets.length > 0) {
    const aoa = buildDetailsSheet(multiSheets);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
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
 */
export function getDisplayHeaders(sheet: SheetData): string[] {
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

/**
 * 概要シートの2次元配列を構築
 * セクション名・フィールド名は翻訳付き
 */
function buildOverviewSheet(
  sheets: SheetData[],
  parsed: ParsedXmlResult
): string[][] {
  const aoa: string[][] = [];

  aoa.push(["セクション / 項目名", "値"]);

  // ファイル情報
  aoa.push(["■ ファイル情報", ""]);
  aoa.push(["  ファイル名", parsed.fileName]);
  aoa.push(["  ルート要素", parsed.rootElement]);
  aoa.push([]);

  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];
    const row = sheet.rows[0];
    const displayHeaders = getDisplayHeaders(sheet);

    // セクションヘッダー（翻訳付き）
    aoa.push([`■ ${translateSection(sheet.name)}`, ""]);

    // 各フィールド（翻訳付き）
    for (const header of displayHeaders) {
      aoa.push([`  ${translateField(header)}`, row[header] ?? ""]);
    }

    if (si < sheets.length - 1) {
      aoa.push([]);
    }
  }

  return aoa;
}

/**
 * 明細シートの2次元配列を構築
 * セクション名は翻訳付きフル表記、列ヘッダーは短縮翻訳
 */
function buildDetailsSheet(sheets: SheetData[]): string[][] {
  const aoa: string[][] = [];

  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];

    if (si > 0) {
      aoa.push([]);
    }

    const displayHeaders = getDisplayHeaders(sheet);

    // セクションヘッダー（翻訳付き）
    aoa.push([
      `■ ${translateSection(sheet.name)} (${sheet.rows.length}件)`,
      ...Array(displayHeaders.length).fill(""),
    ]);

    // テーブルヘッダー（短縮翻訳: 列幅節約のため翻訳名のみ）
    aoa.push(["#", ...displayHeaders.map(translateFieldShort)]);

    // データ行
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
