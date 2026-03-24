import * as XLSX from "xlsx";
import type { ParsedXmlResult, SheetData } from "./types";
import {
  translateSectionShort,
  translateFieldShort,
  stripFieldPrefix,
  TRANSLATION_SOURCE,
} from "./translations";

/**
 * ParsedXmlResult から XLSX ワークブックを生成しダウンロード用の Blob を返す
 *
 * 2シート構成:
 *   シート1「概要」: 3列構成（日本語名 / 技術名 / 値）
 *   シート2「明細」: セクション区切り、2行ヘッダー（日本語名 + 技術名）、セル結合付き
 *
 * フィールド名・セクション名はSAP公式ドキュメントに基づいて日本語に翻訳される。
 * 翻訳が存在しないフィールドは元の技術名をそのまま表示する。
 */
export function generateXlsx(parsed: ParsedXmlResult): Blob {
  const wb = XLSX.utils.book_new();

  const singleSheets = parsed.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = parsed.sheets.filter((s) => s.rows.length >= 2);

  // --- シート1: 概要（3列: 日本語名 / 技術名 / 値） ---
  if (singleSheets.length > 0) {
    const aoa = buildOverviewSheet(singleSheets, parsed);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 28 }, { wch: 28 }, { wch: 55 }];
    ws["!views"] = [{ state: "frozen", ySplit: 1 }];
    XLSX.utils.book_append_sheet(wb, ws, "概要");
  }

  // --- シート2: 明細（2行ヘッダー + セクション見出しセル結合） ---
  if (multiSheets.length > 0) {
    const { aoa, merges } = buildDetailsSheet(multiSheets);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const maxDataCols = Math.max(
      ...multiSheets.map((s) => getDisplayHeaders(s).length)
    );
    const colWidths: XLSX.ColInfo[] = [{ wch: 5 }]; // # 列
    for (let i = 0; i < maxDataCols; i++) {
      colWidths.push({ wch: 22 });
    }
    ws["!cols"] = colWidths;
    if (merges.length > 0) {
      ws["!merges"] = merges;
    }
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
 * 3列構成: 項目名（日本語） / 項目名（技術名） / 値
 */
function buildOverviewSheet(
  sheets: SheetData[],
  parsed: ParsedXmlResult
): string[][] {
  const aoa: string[][] = [];

  aoa.push(["項目名（日本語）", "項目名（技術名）", "値"]);

  // ファイル情報
  aoa.push(["■ ファイル情報", "", ""]);
  aoa.push(["  ファイル名", "", parsed.fileName]);
  aoa.push(["  ルート要素", "", parsed.rootElement]);
  aoa.push([]);

  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];
    const row = sheet.rows[0];
    const displayHeaders = getDisplayHeaders(sheet);

    // セクションヘッダー（日本語名 + 技術名を別列に）
    aoa.push([`■ ${translateSectionShort(sheet.name)}`, sheet.name, ""]);

    // 各フィールド（日本語名 + 技術名を別列に）
    for (const header of displayHeaders) {
      const raw = stripFieldPrefix(header);
      const ja = translateFieldShort(header);
      aoa.push([`  ${ja}`, raw, row[header] ?? ""]);
    }

    if (si < sheets.length - 1) {
      aoa.push([]);
    }
  }

  // 翻訳出典注記
  aoa.push([]);
  aoa.push([`※ ${TRANSLATION_SOURCE}`, "", ""]);

  return aoa;
}

/**
 * 明細シートの2次元配列を構築
 * セクション見出しはセル結合、列ヘッダーは2行（日本語名 + 技術名）
 */
function buildDetailsSheet(
  sheets: SheetData[]
): { aoa: string[][]; merges: XLSX.Range[] } {
  const aoa: string[][] = [];
  const merges: XLSX.Range[] = [];

  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];

    if (si > 0) {
      aoa.push([]);
    }

    const displayHeaders = getDisplayHeaders(sheet);
    const totalCols = 1 + displayHeaders.length; // # 列 + データ列

    // セクションヘッダー（セル結合で全幅表示）
    const sectionRowIdx = aoa.length;
    const sectionJa = translateSectionShort(sheet.name);
    const sectionLabel =
      sectionJa !== sheet.name
        ? `■ ${sectionJa} / ${sheet.name} (${sheet.rows.length}件)`
        : `■ ${sheet.name} (${sheet.rows.length}件)`;
    aoa.push([sectionLabel, ...Array(displayHeaders.length).fill("")]);
    if (totalCols > 1) {
      merges.push({
        s: { r: sectionRowIdx, c: 0 },
        e: { r: sectionRowIdx, c: totalCols - 1 },
      });
    }

    // 列ヘッダー1行目: 日本語名
    aoa.push(["#", ...displayHeaders.map(translateFieldShort)]);

    // 列ヘッダー2行目: 技術名（英名）
    aoa.push(["", ...displayHeaders.map(stripFieldPrefix)]);

    // データ行
    for (let ri = 0; ri < sheet.rows.length; ri++) {
      const row = sheet.rows[ri];
      aoa.push([
        String(ri + 1),
        ...displayHeaders.map((h) => row[h] ?? ""),
      ]);
    }
  }

  // 翻訳出典注記
  aoa.push([]);
  aoa.push([`※ ${TRANSLATION_SOURCE}`]);

  return { aoa, merges };
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
