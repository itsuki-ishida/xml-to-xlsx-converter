import * as XLSX from "xlsx";
import type { ParsedXmlResult, SheetData } from "./types";

/**
 * ParsedXmlResult から XLSX ワークブックを生成しダウンロード用の Blob を返す
 *
 * 2シート構成:
 *   シート1「概要」: 単一インスタンス要素をセクション+キーバリュー形式で表示
 *   シート2「明細」: 複数インスタンス要素をセクション区切りのテーブル形式で表示
 *
 * 全て単一の場合は「概要」のみ、全て複数の場合は「明細」のみ生成
 */
export function generateXlsx(parsed: ParsedXmlResult): Blob {
  const wb = XLSX.utils.book_new();

  // シートを単一(1行)と複数(2行以上)に分類
  const singleSheets = parsed.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = parsed.sheets.filter((s) => s.rows.length >= 2);

  // --- シート1: 概要（単一インスタンス要素） ---
  if (singleSheets.length > 0) {
    const aoa = buildOverviewSheet(singleSheets);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 28 }, { wch: 50 }];
    // 1行目をフリーズ
    ws["!views"] = [{ state: "frozen", ySplit: 1 }];
    XLSX.utils.book_append_sheet(wb, ws, "概要");
  }

  // --- シート2: 明細（複数インスタンス要素） ---
  if (multiSheets.length > 0) {
    const aoa = buildDetailsSheet(multiSheets);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    // 列幅: 最大列数に基づいて設定
    const maxCols = Math.max(...multiSheets.map((s) => s.headers.length));
    const colWidths: XLSX.ColInfo[] = [];
    for (let i = 0; i < maxCols; i++) {
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
 * 概要シートの2次元配列を構築
 * セクション名 + キーバリュー形式
 */
function buildOverviewSheet(sheets: SheetData[]): string[][] {
  const aoa: string[][] = [];

  // ヘッダー行
  aoa.push(["セクション / 項目名", "値"]);

  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];
    const row = sheet.rows[0];

    // セクション区切り（最初以外は空行を挿入）
    if (si > 0) {
      aoa.push([]);
    }

    // セクションヘッダー
    aoa.push([`■ ${sheet.name}`, ""]);

    // 各フィールドをキーバリューで出力（@属性列は除外）
    for (const header of sheet.headers) {
      if (header.startsWith("@")) continue;
      aoa.push([`  ${header}`, row[header] ?? ""]);
    }
  }

  return aoa;
}

/**
 * 明細シートの2次元配列を構築
 * セクション名 + テーブル形式
 */
function buildDetailsSheet(sheets: SheetData[]): string[][] {
  const aoa: string[][] = [];

  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];

    // セクション区切り（最初以外は空行を挿入）
    if (si > 0) {
      aoa.push([]);
    }

    // @属性列を除外したヘッダー
    const dataHeaders = sheet.headers.filter((h) => !h.startsWith("@"));

    // セクションヘッダー
    aoa.push([`■ ${sheet.name} (${sheet.rows.length}件)`, ...Array(dataHeaders.length - 1).fill("")]);

    // テーブルヘッダー
    aoa.push(dataHeaders);

    // データ行
    for (const row of sheet.rows) {
      aoa.push(dataHeaders.map((h) => row[h] ?? ""));
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
