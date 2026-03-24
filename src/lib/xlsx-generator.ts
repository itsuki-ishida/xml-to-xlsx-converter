import XLSX from "xlsx-js-style";
import JSZip from "jszip";
import type { ParsedXmlResult, SheetData } from "./types";
import {
  translateSectionShort,
  translateFieldShort,
  stripFieldPrefix,
  TRANSLATION_SOURCE,
} from "./translations";

// ─── スタイル定義 ──────────────────────────────────────────────

const BORDER_THIN = {
  top: { style: "thin", color: { rgb: "D9D9D9" } },
  bottom: { style: "thin", color: { rgb: "D9D9D9" } },
  left: { style: "thin", color: { rgb: "D9D9D9" } },
  right: { style: "thin", color: { rgb: "D9D9D9" } },
} as const;

const S = {
  /** 概要シート: ヘッダー行（ダークブルー + 白太字） */
  overviewHeader: {
    font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
    fill: { fgColor: { rgb: "2F5496" } },
    border: BORDER_THIN,
    alignment: { vertical: "center" as const },
  },
  /** 概要シート: セクション見出し（ライトブルー + 太字） */
  overviewSection: {
    font: { bold: true, sz: 11, color: { rgb: "1F3864" } },
    fill: { fgColor: { rgb: "D6E4F0" } },
    border: BORDER_THIN,
  },
  /** 概要シート: フィールド日本語名 */
  overviewFieldJa: {
    font: { sz: 10 },
    border: BORDER_THIN,
  },
  /** 概要シート: フィールド技術名（グレー） */
  overviewFieldTech: {
    font: { sz: 9, color: { rgb: "808080" } },
    border: BORDER_THIN,
  },
  /** 概要シート: フィールド値 */
  overviewFieldValue: {
    font: { sz: 10 },
    border: BORDER_THIN,
  },
  /** 明細: セクション見出し（ダークブルー + 白太字） */
  detailSection: {
    font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
    fill: { fgColor: { rgb: "2F5496" } },
    border: BORDER_THIN,
    alignment: { vertical: "center" as const },
  },
  /** 明細: 日本語ヘッダー行（ミディアムブルー + 白太字） */
  detailHeaderJa: {
    font: { bold: true, color: { rgb: "FFFFFF" }, sz: 10 },
    fill: { fgColor: { rgb: "4472C4" } },
    border: BORDER_THIN,
  },
  /** 明細: 技術名ヘッダー行（ライトグレー + グレー文字） */
  detailHeaderTech: {
    font: { sz: 9, color: { rgb: "808080" } },
    fill: { fgColor: { rgb: "F2F2F2" } },
    border: BORDER_THIN,
  },
  /** 明細: データセル */
  detailData: {
    font: { sz: 10 },
    border: BORDER_THIN,
  },
  /** 明細: データセル（偶数行: 薄い背景） */
  detailDataAlt: {
    font: { sz: 10 },
    fill: { fgColor: { rgb: "F7F9FC" } },
    border: BORDER_THIN,
  },
  /** 明細: 行番号セル */
  detailRowNum: {
    font: { sz: 9, color: { rgb: "999999" } },
    border: BORDER_THIN,
    alignment: { horizontal: "center" as const },
  },
  /** 明細: 行番号セル（偶数行） */
  detailRowNumAlt: {
    font: { sz: 9, color: { rgb: "999999" } },
    fill: { fgColor: { rgb: "F7F9FC" } },
    border: BORDER_THIN,
    alignment: { horizontal: "center" as const },
  },
  /** 出典注記 */
  sourceNote: {
    font: { italic: true, sz: 9, color: { rgb: "999999" } },
  },
} as const;

// ─── スタイル適用ヘルパー ─────────────────────────────────────

type CellStyle = typeof S[keyof typeof S];

function applyRowStyle(
  ws: XLSX.WorkSheet,
  row: number,
  colCount: number,
  style: CellStyle
) {
  for (let c = 0; c < colCount; c++) {
    const ref = XLSX.utils.encode_cell({ r: row, c });
    if (!ws[ref]) ws[ref] = { v: "", t: "s" };
    ws[ref].s = style;
  }
}

function applyCellStyle(
  ws: XLSX.WorkSheet,
  row: number,
  col: number,
  style: CellStyle
) {
  const ref = XLSX.utils.encode_cell({ r: row, c: col });
  if (!ws[ref]) ws[ref] = { v: "", t: "s" };
  ws[ref].s = style;
}

// ─── シート分割閾値 ─────────────────────────────────────────

/**
 * この行数以上のテーブルは「明細」統合シートから分離し、独立シートに出力する。
 * 30行 ≈ Excelで1画面に収まる量。これを超えるとスクロールが煩わしくなるため分離。
 */
export const LARGE_TABLE_THRESHOLD = 30;

// ─── メイン生成関数 ───────────────────────────────────────────

/**
 * ParsedXmlResult から XLSX ワークブックを生成しダウンロード用の Blob を返す
 *
 * 動的シート構成:
 *   シート1「概要」: 単一行セクション → 3列構成（日本語名 / 技術名 / 値）
 *   シート2「明細」: 小規模テーブル（< 閾値行）→ 統合表示
 *   シート3+: 大規模テーブル（>= 閾値行）→ 各テーブルが独立シート
 */
export async function generateXlsx(parsed: ParsedXmlResult): Promise<Blob> {
  const wb = XLSX.utils.book_new();

  const singleSheets = parsed.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = parsed.sheets.filter((s) => s.rows.length >= 2);
  const smallMultiSheets = multiSheets.filter(
    (s) => s.rows.length < LARGE_TABLE_THRESHOLD
  );
  const largeMultiSheets = multiSheets.filter(
    (s) => s.rows.length >= LARGE_TABLE_THRESHOLD
  );

  // フローズンペイン設定を記録（シート番号 → ySplit行数）
  // xlsx-js-styleは!viewsを無視するため、後処理でXMLに直接注入する
  const frozenPanes: Map<number, number> = new Map();
  let sheetIndex = 0;

  // --- シート1: 概要 ---
  if (singleSheets.length > 0) {
    const { aoa, rowMeta } = buildOverviewSheet(singleSheets, parsed);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 28 }, { wch: 28 }, { wch: 55 }];
    applyOverviewStyles(ws, rowMeta);
    XLSX.utils.book_append_sheet(wb, ws, "概要");
    frozenPanes.set(sheetIndex, 1);
    sheetIndex++;
  }

  // --- シート2: 明細（小規模テーブル統合） ---
  if (smallMultiSheets.length > 0) {
    const { aoa, merges, rowMeta, rowCols, maxCols } =
      buildDetailsSheet(smallMultiSheets);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const maxDataCols = Math.max(
      ...smallMultiSheets.map((s) => getDisplayHeaders(s).length)
    );
    const colWidths: XLSX.ColInfo[] = [{ wch: 5 }];
    for (let i = 0; i < maxDataCols; i++) {
      colWidths.push({ wch: 22 });
    }
    ws["!cols"] = colWidths;
    if (merges.length > 0) ws["!merges"] = merges;
    applyDetailStyles(ws, rowMeta, rowCols);
    XLSX.utils.book_append_sheet(wb, ws, "明細");
    sheetIndex++;
  }

  // --- 個別シート: 大規模テーブル（閾値以上） ---
  const usedSheetNames = new Set(["概要", "明細"]);
  for (const sheet of largeMultiSheets) {
    const { aoa, merges, rowMeta, rowCols } = buildDetailsSheet([sheet]);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const displayHeaders = getDisplayHeaders(sheet);
    const colWidths: XLSX.ColInfo[] = [{ wch: 5 }];
    for (let i = 0; i < displayHeaders.length; i++) {
      colWidths.push({ wch: 22 });
    }
    ws["!cols"] = colWidths;
    if (merges.length > 0) ws["!merges"] = merges;
    applyDetailStyles(ws, rowMeta, rowCols);
    const sheetName = makeUniqueSheetName(sheet.name, usedSheetNames);
    usedSheetNames.add(sheetName);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    // フローズンペイン: セクションヘッダー + 2行ヘッダーの下で固定
    frozenPanes.set(sheetIndex, 3);
    sheetIndex++;
  }

  // ワークブックが空の場合
  if (parsed.sheets.length === 0) {
    const ws = XLSX.utils.aoa_to_sheet([["（データなし）"]]);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  }

  // SheetJS公式パターン: type "binary" → ArrayBuffer変換
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
  const buf = new ArrayBuffer(wbout.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < wbout.length; i++) {
    view[i] = wbout.charCodeAt(i) & 0xff;
  }

  // xlsx-js-styleはフローズンペイン(!views)未対応のため、
  // JSZipでXLSX(zip)を展開し、シートXMLにpane要素を直接注入する
  if (frozenPanes.size > 0) {
    const zip = await JSZip.loadAsync(buf);
    for (const [idx, ySplit] of frozenPanes) {
      const sheetPath = `xl/worksheets/sheet${idx + 1}.xml`;
      const xml = await zip.file(sheetPath)?.async("string");
      if (xml) {
        const topLeftCell = `A${ySplit + 1}`;
        const paneXml =
          `<pane ySplit="${ySplit}" topLeftCell="${topLeftCell}" activePane="bottomLeft" state="frozen"/>` +
          `<selection pane="bottomLeft" activeCell="${topLeftCell}" sqref="${topLeftCell}"/>`;
        // <sheetView workbookViewId="0"/> → <sheetView workbookViewId="0"> + pane + </sheetView>
        const patched = xml.replace(
          /<sheetView workbookViewId="0"\/>/,
          `<sheetView workbookViewId="0">${paneXml}</sheetView>`
        );
        zip.file(sheetPath, patched);
      }
    }
    const patchedBuf = await zip.generateAsync({
      type: "arraybuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    return new Blob([patchedBuf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
  }

  return new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

// ─── 表示ヘッダー ─────────────────────────────────────────────

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

// ─── 概要シート構築 ───────────────────────────────────────────

type OverviewRowType = "header" | "section" | "field" | "empty" | "source";

function buildOverviewSheet(
  sheets: SheetData[],
  parsed: ParsedXmlResult
): { aoa: string[][]; rowMeta: OverviewRowType[] } {
  const aoa: string[][] = [];
  const rowMeta: OverviewRowType[] = [];

  aoa.push(["項目名（日本語）", "項目名（技術名）", "値"]);
  rowMeta.push("header");

  aoa.push(["■ ファイル情報", "", ""]);
  rowMeta.push("section");
  aoa.push(["  ファイル名", "", parsed.fileName]);
  rowMeta.push("field");
  aoa.push(["  ルート要素", "", parsed.rootElement]);
  rowMeta.push("field");
  aoa.push([]);
  rowMeta.push("empty");

  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];
    const row = sheet.rows[0];
    const displayHeaders = getDisplayHeaders(sheet);

    aoa.push([`■ ${translateSectionShort(sheet.name)}`, sheet.name, ""]);
    rowMeta.push("section");

    for (const header of displayHeaders) {
      const raw = stripFieldPrefix(header);
      const ja = translateFieldShort(header);
      aoa.push([`  ${ja}`, raw, row[header] ?? ""]);
      rowMeta.push("field");
    }

    if (si < sheets.length - 1) {
      aoa.push([]);
      rowMeta.push("empty");
    }
  }

  aoa.push([]);
  rowMeta.push("empty");
  aoa.push([`※ ${TRANSLATION_SOURCE}`, "", ""]);
  rowMeta.push("source");

  return { aoa, rowMeta };
}

function applyOverviewStyles(
  ws: XLSX.WorkSheet,
  rowMeta: OverviewRowType[]
) {
  const colCount = 3;
  const rowHeights: XLSX.RowInfo[] = [];

  for (let r = 0; r < rowMeta.length; r++) {
    switch (rowMeta[r]) {
      case "header":
        applyRowStyle(ws, r, colCount, S.overviewHeader);
        rowHeights[r] = { hpt: 24 };
        break;
      case "section":
        applyRowStyle(ws, r, colCount, S.overviewSection);
        rowHeights[r] = { hpt: 22 };
        break;
      case "field":
        applyCellStyle(ws, r, 0, S.overviewFieldJa);
        applyCellStyle(ws, r, 1, S.overviewFieldTech);
        applyCellStyle(ws, r, 2, S.overviewFieldValue);
        break;
      case "source":
        applyCellStyle(ws, r, 0, S.sourceNote);
        break;
    }
  }

  ws["!rows"] = rowHeights;
}

// ─── 明細シート構築 ───────────────────────────────────────────

type DetailRowType =
  | "section"
  | "headerJa"
  | "headerTech"
  | "data"
  | "dataAlt"
  | "empty"
  | "source";

function buildDetailsSheet(sheets: SheetData[]): {
  aoa: string[][];
  merges: XLSX.Range[];
  rowMeta: DetailRowType[];
  rowCols: number[];
  maxCols: number;
} {
  const aoa: string[][] = [];
  const merges: XLSX.Range[] = [];
  const rowMeta: DetailRowType[] = [];
  const rowCols: number[] = [];
  let maxCols = 0;

  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];

    if (si > 0) {
      aoa.push([]);
      rowMeta.push("empty");
      rowCols.push(0);
    }

    const displayHeaders = getDisplayHeaders(sheet);
    const totalCols = 1 + displayHeaders.length;
    if (totalCols > maxCols) maxCols = totalCols;

    // セクションヘッダー
    const sectionRowIdx = aoa.length;
    const sectionJa = translateSectionShort(sheet.name);
    const sectionLabel =
      sectionJa !== sheet.name
        ? `■ ${sectionJa} / ${sheet.name} (${sheet.rows.length}件)`
        : `■ ${sheet.name} (${sheet.rows.length}件)`;
    aoa.push([sectionLabel, ...Array(displayHeaders.length).fill("")]);
    rowMeta.push("section");
    rowCols.push(totalCols);
    if (totalCols > 1) {
      merges.push({
        s: { r: sectionRowIdx, c: 0 },
        e: { r: sectionRowIdx, c: totalCols - 1 },
      });
    }

    // 日本語ヘッダー
    aoa.push(["#", ...displayHeaders.map(translateFieldShort)]);
    rowMeta.push("headerJa");
    rowCols.push(totalCols);

    // 技術名ヘッダー
    aoa.push(["", ...displayHeaders.map(stripFieldPrefix)]);
    rowMeta.push("headerTech");
    rowCols.push(totalCols);

    // データ行
    for (let ri = 0; ri < sheet.rows.length; ri++) {
      const row = sheet.rows[ri];
      aoa.push([
        String(ri + 1),
        ...displayHeaders.map((h) => row[h] ?? ""),
      ]);
      rowMeta.push(ri % 2 === 0 ? "data" : "dataAlt");
      rowCols.push(totalCols);
    }
  }

  aoa.push([]);
  rowMeta.push("empty");
  rowCols.push(0);
  aoa.push([`※ ${TRANSLATION_SOURCE}`]);
  rowMeta.push("source");
  rowCols.push(1);

  return { aoa, merges, rowMeta, rowCols, maxCols };
}

function applyDetailStyles(
  ws: XLSX.WorkSheet,
  rowMeta: DetailRowType[],
  rowCols: number[]
) {
  const rowHeights: XLSX.RowInfo[] = [];

  for (let r = 0; r < rowMeta.length; r++) {
    const cols = rowCols[r] || 0;
    switch (rowMeta[r]) {
      case "section":
        applyRowStyle(ws, r, cols, S.detailSection);
        rowHeights[r] = { hpt: 24 };
        break;
      case "headerJa":
        applyRowStyle(ws, r, cols, S.detailHeaderJa);
        rowHeights[r] = { hpt: 20 };
        break;
      case "headerTech":
        applyRowStyle(ws, r, cols, S.detailHeaderTech);
        break;
      case "data":
        applyCellStyle(ws, r, 0, S.detailRowNum);
        for (let c = 1; c < cols; c++) {
          applyCellStyle(ws, r, c, S.detailData);
        }
        break;
      case "dataAlt":
        applyCellStyle(ws, r, 0, S.detailRowNumAlt);
        for (let c = 1; c < cols; c++) {
          applyCellStyle(ws, r, c, S.detailDataAlt);
        }
        break;
      case "source":
        applyCellStyle(ws, r, 0, S.sourceNote);
        break;
    }
  }

  ws["!rows"] = rowHeights;
}

// ─── ユーティリティ ───────────────────────────────────────────

/**
 * セクション名からExcelシート名を生成する
 * 日本語翻訳名があればそれを使い、なければ技術名を使用。
 * 31文字制限・禁止文字除去・重複回避を行う。
 */
function makeUniqueSheetName(
  sectionName: string,
  usedNames: Set<string>
): string {
  const ja = translateSectionShort(sectionName);
  let base = ja !== sectionName ? ja : sectionName;
  // Excel禁止文字を除去
  base = base.replace(/[\\/?*[\]:]/g, "_");
  if (base.length > 31) base = base.substring(0, 31);

  let name = base;
  let suffix = 2;
  while (usedNames.has(name)) {
    const suffixStr = `_${suffix}`;
    name = base.substring(0, 31 - suffixStr.length) + suffixStr;
    suffix++;
  }
  return name;
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
  separateSheets: { name: string; jaName: string; rows: number }[];
  sheetCount: number;
} {
  const singleSheets = parsed.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = parsed.sheets.filter((s) => s.rows.length >= 2);
  const smallMulti = multiSheets.filter(
    (s) => s.rows.length < LARGE_TABLE_THRESHOLD
  );
  const largeMulti = multiSheets.filter(
    (s) => s.rows.length >= LARGE_TABLE_THRESHOLD
  );
  return {
    overviewSections: singleSheets.length,
    detailSections: smallMulti.length,
    detailRows: smallMulti.reduce((acc, s) => acc + s.rows.length, 0),
    separateSheets: largeMulti.map((s) => ({
      name: s.name,
      jaName: translateSectionShort(s.name),
      rows: s.rows.length,
    })),
    sheetCount:
      (singleSheets.length > 0 ? 1 : 0) +
      (smallMulti.length > 0 ? 1 : 0) +
      largeMulti.length,
  };
}
