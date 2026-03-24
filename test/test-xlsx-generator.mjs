/**
 * XLSX生成テスト（バイリンガル対応: 日本語名 + 技術名の両方表示）
 * - 概要シート: 3列構成（日本語名 / 技術名 / 値）
 * - 明細シート: セクション見出しセル結合 + 2行ヘッダー（日本語名 + 技術名）
 * - SAP公式ドキュメント準拠の翻訳
 * - データ完全性の検証
 */
import { JSDOM } from "jsdom";
import XLSX from "xlsx-js-style";
import JSZip from "jszip";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// DOMParser設定
const { DOMParser } = new JSDOM("").window;
globalThis.DOMParser = DOMParser;

// === xml-parser ロジック（インライン） ===
function stripNamespace(tagName) {
  const idx = tagName.indexOf(":");
  return idx >= 0 ? tagName.substring(idx + 1) : tagName;
}
function isLeafElement(el) { return el.children.length === 0; }
function isRecordElement(el) {
  if (el.children.length === 0) return false;
  for (let i = 0; i < el.children.length; i++) {
    if (isLeafElement(el.children[i])) return true;
  }
  return false;
}
function extractRecordData(el) {
  const data = {};
  for (let i = 0; i < el.attributes.length; i++) {
    const attr = el.attributes[i];
    if (attr.name.startsWith("xmlns")) continue;
    data[`@${attr.name}`] = attr.value;
  }
  for (let i = 0; i < el.children.length; i++) {
    const child = el.children[i];
    if (isLeafElement(child)) {
      data[stripNamespace(child.tagName)] = child.textContent?.trim() ?? "";
    }
  }
  return data;
}
function collectRecordGroups(el, groups) {
  if (isRecordElement(el)) {
    const tagName = stripNamespace(el.tagName);
    const data = extractRecordData(el);
    if (!groups.has(tagName)) groups.set(tagName, []);
    groups.get(tagName).push(data);
  }
  for (let i = 0; i < el.children.length; i++) {
    collectRecordGroups(el.children[i], groups);
  }
}
function sanitizeSheetName(name, existingNames) {
  let sanitized = name.replace(/[\\/*?:\[\]]/g, "_");
  if (sanitized.length > 31) sanitized = sanitized.substring(0, 31);
  let finalName = sanitized;
  let counter = 1;
  while (existingNames.has(finalName)) {
    const suffix = `_${counter}`;
    finalName = sanitized.substring(0, 31 - suffix.length) + suffix;
    counter++;
  }
  existingNames.add(finalName);
  return finalName;
}
function buildSheetData(name, rows, existingNames) {
  const headerSet = new Set();
  const orderedHeaders = [];
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      if (!headerSet.has(key)) { headerSet.add(key); orderedHeaders.push(key); }
    }
  }
  const attrHeaders = orderedHeaders.filter((h) => h.startsWith("@"));
  const dataHeaders = orderedHeaders.filter((h) => !h.startsWith("@"));
  return { name: sanitizeSheetName(name, existingNames), headers: [...dataHeaders, ...attrHeaders], rows };
}
function parseXml(xmlText, fileName) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, "application/xml");
  const root = doc.documentElement;
  const rootElement = stripNamespace(root.tagName);
  const groups = new Map();
  collectRecordGroups(root, groups);
  const existingNames = new Set();
  const sheets = [];
  for (const [tagName, rows] of groups) {
    sheets.push(buildSheetData(tagName, rows, existingNames));
  }
  return { fileName, rootElement, sheets };
}

// === 翻訳辞書（テスト用インライン: 本体は src/lib/translations.ts） ===
const SECTION_TRANSLATIONS = {
  EDI_DC40: "IDoc制御情報",
  E1SHP_IBDLV_SAVE_REPLICA: "出荷通知データ",
  E1BPIBDLVHDR: "配送ヘッダー",
  E1BPIBDLVHDRORG: "配送ヘッダー組織",
  E1BPDLVCONTROL: "配送制御",
  E1BPDLVPARTNER: "配送パートナー",
  E1BPADR1: "住所情報",
  E1BPADR11: "住所詳細",
  E1BPDLVDEADLN: "配送期日",
  E1BPIBDLVITEM: "配送明細",
  E1BPIBDLVITEMORG: "配送明細組織",
  E1BPDLVITEMSTTR: "配送明細ステータス",
  E1BPDLVCOBLITEM: "配送明細原価",
  E1BPDLVITEMRPO: "配送明細参照伝票",
  E1BPDLVHDUNHDR: "梱包ユニットヘッダー",
  E1BPDLVHDUNITM: "梱包ユニット明細",
  E1BPEXTC: "拡張データ",
  RecordSet: "レコードセット",
  FMS_ROUTING: "ルーティング情報",
};
const FIELD_TRANSLATIONS = {
  TABNAM: "テーブル名", MANDT: "クライアント", DOCNUM: "ドキュメント番号",
  DELIV_NUMB: "配送番号", ITM_NUMBER: "明細番号", MATERIAL: "品目コード",
  SHORT_TEXT: "品目テキスト", DLV_QTY: "配送数量", NET_WEIGHT: "正味重量",
  SNDPRN: "送信パートナー", I_VBELN: "伝票番号", I_STATUS: "ステータス",
  I_UTC_TIMESTAMP: "タイムスタンプ(UTC)", NAME: "名前", E_MAIL: "メールアドレス",
  FIELD1: "フィールド1", FIELD2: "フィールド2", FIELD3: "フィールド3", FIELD4: "フィールド4",
  PLANT: "プラント", STGE_LOC: "保管場所", PROFIT_CTR: "利益センタ",
  TIMETYPE: "時間タイプ", TIMESTAMP_UTC: "タイムスタンプ(UTC)", TIMEZONE: "タイムゾーン",
};

const TRANSLATION_SOURCE = "SAP公式ドキュメント（IDoc Interface / ABAP Data Dictionary）に基づく翻訳";

function translateSectionShort(name) {
  return SECTION_TRANSLATIONS[name] ?? name;
}
function stripFieldPrefix(name) {
  return name.startsWith("@") ? name.substring(1) : name;
}
function translateFieldShort(name) {
  const raw = stripFieldPrefix(name);
  return FIELD_TRANSLATIONS[raw] ?? raw;
}

// === スタイル定義 ===
const BORDER_THIN = {
  top: { style: "thin", color: { rgb: "D9D9D9" } },
  bottom: { style: "thin", color: { rgb: "D9D9D9" } },
  left: { style: "thin", color: { rgb: "D9D9D9" } },
  right: { style: "thin", color: { rgb: "D9D9D9" } },
};
const STYLES = {
  overviewHeader: { font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 }, fill: { fgColor: { rgb: "2F5496" } }, border: BORDER_THIN, alignment: { vertical: "center" } },
  overviewSection: { font: { bold: true, sz: 11, color: { rgb: "1F3864" } }, fill: { fgColor: { rgb: "D6E4F0" } }, border: BORDER_THIN },
  overviewFieldJa: { font: { sz: 10 }, border: BORDER_THIN },
  overviewFieldTech: { font: { sz: 9, color: { rgb: "808080" } }, border: BORDER_THIN },
  overviewFieldValue: { font: { sz: 10 }, border: BORDER_THIN },
  detailSection: { font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 }, fill: { fgColor: { rgb: "2F5496" } }, border: BORDER_THIN, alignment: { vertical: "center" } },
  detailHeaderJa: { font: { bold: true, color: { rgb: "FFFFFF" }, sz: 10 }, fill: { fgColor: { rgb: "4472C4" } }, border: BORDER_THIN },
  detailHeaderTech: { font: { sz: 9, color: { rgb: "808080" } }, fill: { fgColor: { rgb: "F2F2F2" } }, border: BORDER_THIN },
  detailData: { font: { sz: 10 }, border: BORDER_THIN },
  detailDataAlt: { font: { sz: 10 }, fill: { fgColor: { rgb: "F7F9FC" } }, border: BORDER_THIN },
  detailRowNum: { font: { sz: 9, color: { rgb: "999999" } }, border: BORDER_THIN, alignment: { horizontal: "center" } },
  detailRowNumAlt: { font: { sz: 9, color: { rgb: "999999" } }, fill: { fgColor: { rgb: "F7F9FC" } }, border: BORDER_THIN, alignment: { horizontal: "center" } },
  sourceNote: { font: { italic: true, sz: 9, color: { rgb: "999999" } } },
};

function applyRowStyle(ws, row, colCount, style) {
  for (let c = 0; c < colCount; c++) {
    const ref = XLSX.utils.encode_cell({ r: row, c });
    if (!ws[ref]) ws[ref] = { v: "", t: "s" };
    ws[ref].s = style;
  }
}
function applyCellStyle(ws, row, col, style) {
  const ref = XLSX.utils.encode_cell({ r: row, c: col });
  if (!ws[ref]) ws[ref] = { v: "", t: "s" };
  ws[ref].s = style;
}

// === シート分割閾値 ===
const LARGE_TABLE_THRESHOLD = 30;

// === XLSX生成ロジック（バイリンガル + スタイリング + シート分割対応） ===
function getDisplayHeaders(sheet) {
  const headers = [];
  for (const h of sheet.headers) {
    if (h.startsWith("@")) {
      const values = new Set(sheet.rows.map((r) => r[h] ?? ""));
      if (values.size <= 1) continue;
    }
    headers.push(h);
  }
  return headers;
}

function makeUniqueSheetName(sectionName, usedNames) {
  const ja = translateSectionShort(sectionName);
  let base = ja !== sectionName ? ja : sectionName;
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

function buildDetailAoa(sheets) {
  const aoa = [];
  const merges = [];
  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];
    if (si > 0) aoa.push([]);
    const displayHeaders = getDisplayHeaders(sheet);
    const totalCols = 1 + displayHeaders.length;

    const sectionRowIdx = aoa.length;
    const sectionJa = translateSectionShort(sheet.name);
    const sectionLabel = sectionJa !== sheet.name
      ? `■ ${sectionJa} / ${sheet.name} (${sheet.rows.length}件)`
      : `■ ${sheet.name} (${sheet.rows.length}件)`;
    aoa.push([sectionLabel, ...Array(displayHeaders.length).fill("")]);
    if (totalCols > 1) {
      merges.push({ s: { r: sectionRowIdx, c: 0 }, e: { r: sectionRowIdx, c: totalCols - 1 } });
    }

    aoa.push(["#", ...displayHeaders.map(translateFieldShort)]);
    aoa.push(["", ...displayHeaders.map(stripFieldPrefix)]);

    for (let ri = 0; ri < sheet.rows.length; ri++) {
      const row = sheet.rows[ri];
      aoa.push([String(ri + 1), ...displayHeaders.map((h) => row[h] ?? "")]);
    }
  }
  aoa.push([]);
  aoa.push([`※ ${TRANSLATION_SOURCE}`]);
  const maxCols = Math.max(...sheets.map(s => 1 + getDisplayHeaders(s).length));
  return { aoa, merges, maxCols };
}

function applyDetailStylesToWs(ws, aoa, maxCols) {
  const detailRowHeights = [];
  let dataRowIdx = 0;
  for (let r = 0; r < aoa.length; r++) {
    const cellA = String(aoa[r][0] || "");
    if (cellA.startsWith("■")) {
      applyRowStyle(ws, r, maxCols, STYLES.detailSection);
      detailRowHeights[r] = { hpt: 24 };
      dataRowIdx = 0;
    } else if (cellA === "#") {
      applyRowStyle(ws, r, maxCols, STYLES.detailHeaderJa);
      detailRowHeights[r] = { hpt: 20 };
    } else if (r > 0 && String(aoa[r - 1]?.[0] || "") === "#") {
      applyRowStyle(ws, r, maxCols, STYLES.detailHeaderTech);
    } else if (cellA.startsWith("※")) {
      applyCellStyle(ws, r, 0, STYLES.sourceNote);
    } else if (cellA.match(/^\d+$/)) {
      const isAlt = dataRowIdx % 2 === 1;
      applyCellStyle(ws, r, 0, isAlt ? STYLES.detailRowNumAlt : STYLES.detailRowNum);
      for (let c = 1; c < maxCols; c++) {
        applyCellStyle(ws, r, c, isAlt ? STYLES.detailDataAlt : STYLES.detailData);
      }
      dataRowIdx++;
    }
  }
  ws["!rows"] = detailRowHeights;
}

async function generateXlsxBuffer(parsed) {
  const wb = XLSX.utils.book_new();
  const singleSheets = parsed.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = parsed.sheets.filter((s) => s.rows.length >= 2);
  const smallMultiSheets = multiSheets.filter((s) => s.rows.length < LARGE_TABLE_THRESHOLD);
  const largeMultiSheets = multiSheets.filter((s) => s.rows.length >= LARGE_TABLE_THRESHOLD);

  // フローズンペイン設定を記録（シート番号 → ySplit行数）
  const frozenPanes = new Map();
  let sheetIndex = 0;

  // シート1: 概要（3列: 日本語名 / 技術名 / 値）
  if (singleSheets.length > 0) {
    const aoa = [["項目名（日本語）", "項目名（技術名）", "値"]];
    aoa.push(["■ ファイル情報", "", ""]);
    aoa.push(["  ファイル名", "", parsed.fileName]);
    aoa.push(["  ルート要素", "", parsed.rootElement]);
    aoa.push([]);

    for (let si = 0; si < singleSheets.length; si++) {
      const sheet = singleSheets[si];
      const row = sheet.rows[0];
      const displayHeaders = getDisplayHeaders(sheet);
      aoa.push([`■ ${translateSectionShort(sheet.name)}`, sheet.name, ""]);
      for (const header of displayHeaders) {
        const raw = stripFieldPrefix(header);
        const ja = translateFieldShort(header);
        aoa.push([`  ${ja}`, raw, row[header] ?? ""]);
      }
      if (si < singleSheets.length - 1) aoa.push([]);
    }
    aoa.push([]);
    aoa.push([`※ ${TRANSLATION_SOURCE}`, "", ""]);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 28 }, { wch: 28 }, { wch: 55 }];
    const rowHeights = [];
    for (let r = 0; r < aoa.length; r++) {
      const cellA = aoa[r][0] || "";
      if (r === 0) {
        applyRowStyle(ws, r, 3, STYLES.overviewHeader);
        rowHeights[r] = { hpt: 24 };
      } else if (String(cellA).startsWith("■")) {
        applyRowStyle(ws, r, 3, STYLES.overviewSection);
        rowHeights[r] = { hpt: 22 };
      } else if (String(cellA).startsWith("※")) {
        applyCellStyle(ws, r, 0, STYLES.sourceNote);
      } else if (String(cellA).startsWith("  ")) {
        applyCellStyle(ws, r, 0, STYLES.overviewFieldJa);
        applyCellStyle(ws, r, 1, STYLES.overviewFieldTech);
        applyCellStyle(ws, r, 2, STYLES.overviewFieldValue);
      }
    }
    ws["!rows"] = rowHeights;
    XLSX.utils.book_append_sheet(wb, ws, "概要");
    frozenPanes.set(sheetIndex, 1);
    sheetIndex++;
  }

  // シート2: 明細（小規模テーブル統合）
  if (smallMultiSheets.length > 0) {
    const { aoa, merges, maxCols } = buildDetailAoa(smallMultiSheets);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (merges.length > 0) ws["!merges"] = merges;
    applyDetailStylesToWs(ws, aoa, maxCols);
    XLSX.utils.book_append_sheet(wb, ws, "明細");
    sheetIndex++;
  }

  // 個別シート: 大規模テーブル（閾値以上）
  const usedSheetNames = new Set(["概要", "明細"]);
  for (const sheet of largeMultiSheets) {
    const { aoa, merges, maxCols } = buildDetailAoa([sheet]);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (merges.length > 0) ws["!merges"] = merges;
    const displayHeaders = getDisplayHeaders(sheet);
    const colWidths = [{ wch: 5 }];
    for (let i = 0; i < displayHeaders.length; i++) colWidths.push({ wch: 22 });
    ws["!cols"] = colWidths;
    applyDetailStylesToWs(ws, aoa, maxCols);
    const sheetName = makeUniqueSheetName(sheet.name, usedSheetNames);
    usedSheetNames.add(sheetName);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    frozenPanes.set(sheetIndex, 3);
    sheetIndex++;
  }

  const rawBuf = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });

  // JSZipでフローズンペインを後処理で注入
  if (frozenPanes.size > 0) {
    const zip = await JSZip.loadAsync(rawBuf);
    for (const [idx, ySplit] of frozenPanes) {
      const sheetPath = `xl/worksheets/sheet${idx + 1}.xml`;
      const xml = await zip.file(sheetPath)?.async("string");
      if (xml) {
        const topLeftCell = `A${ySplit + 1}`;
        const paneXml =
          `<pane ySplit="${ySplit}" topLeftCell="${topLeftCell}" activePane="bottomLeft" state="frozen"/>` +
          `<selection pane="bottomLeft" activeCell="${topLeftCell}" sqref="${topLeftCell}"/>`;
        const patched = xml.replace(
          /<sheetView workbookViewId="0"\/>/,
          `<sheetView workbookViewId="0">${paneXml}</sheetView>`
        );
        zip.file(sheetPath, patched);
      }
    }
    return await zip.generateAsync({ type: "nodebuffer" });
  }

  return rawBuf;
}

// === テスト実行 ===
let passed = 0;
let failed = 0;
function assert(condition, message) {
  if (condition) { passed++; console.log(`  ✅ ${message}`); }
  else { failed++; console.error(`  ❌ ${message}`); }
}

// テスト1: ASN - 3列概要 + 2行ヘッダー明細
console.log("\n📊 テスト1: ASN XML → バイリンガル3列概要 + 2行ヘッダー明細");
const asnXml = fs.readFileSync(
  path.resolve(__dirname, "../../変換前データ/ASN_1800024221_20190214185557.xml"), "utf-8"
);
const asnParsed = parseXml(asnXml, "ASN_1800024221.xml");
const asnBuffer = await generateXlsxBuffer(asnParsed);
const asnWb = XLSX.read(asnBuffer, { type: "buffer" });

assert(asnWb.SheetNames.length === 2, `2シート構成 (実際: ${asnWb.SheetNames.length})`);
assert(asnWb.SheetNames[0] === "概要", `シート1=概要`);
assert(asnWb.SheetNames[1] === "明細", `シート2=明細`);

const overviewData = XLSX.utils.sheet_to_json(asnWb.Sheets["概要"], { header: 1 });

// 3列ヘッダー検証
assert(overviewData[0][0] === "項目名（日本語）", "概要ヘッダー列A = 項目名（日本語）");
assert(overviewData[0][1] === "項目名（技術名）", "概要ヘッダー列B = 項目名（技術名）");
assert(overviewData[0][2] === "値", "概要ヘッダー列C = 値");

// ファイル情報
assert(overviewData[1][0] === "■ ファイル情報", "ファイル情報セクション存在");
assert(overviewData[2][2] === "ASN_1800024221.xml", "ファイル名が値列(C列)にある");

// セクション行: 日本語名がA列、技術名がB列
const sectionRows = overviewData.filter(r => r[0] && String(r[0]).startsWith("■") && r[0] !== "■ ファイル情報");
assert(sectionRows.length > 0, "セクション行が存在");

const ediDcSection = sectionRows.find(r => String(r[1]) === "EDI_DC40");
assert(ediDcSection && String(ediDcSection[0]).includes("IDoc制御情報"), "EDI_DC40: A列=日本語名, B列=技術名");

const dlvHdrSection = sectionRows.find(r => String(r[1]) === "E1BPIBDLVHDR");
assert(dlvHdrSection && String(dlvHdrSection[0]).includes("配送ヘッダー"), "E1BPIBDLVHDR: A列=日本語名, B列=技術名");

// フィールド行: 日本語名がA列、技術名がB列、値がC列
const docnumRow = overviewData.find(r => r[1] === "DOCNUM");
assert(docnumRow && String(docnumRow[0]).includes("ドキュメント番号"), "DOCNUM: A列=日本語名");
assert(docnumRow && String(docnumRow[1]) === "DOCNUM", "DOCNUM: B列=技術名");
assert(docnumRow && String(docnumRow[2]) === "0000000059141367", "DOCNUM: C列=値");

const nameRow = overviewData.find(r => r[1] === "NAME");
assert(nameRow && String(nameRow[0]).includes("名前"), "NAME: A列=日本語名「名前」");

// @SEGMENT除外の検証
const segmentRow = overviewData.find(r => r[1] === "SEGMENT" || r[1] === "@SEGMENT");
assert(!segmentRow, "@SEGMENT除外");

// 翻訳出典注記
const sourceRow = overviewData.find(r => r[0] && String(r[0]).includes("SAP公式ドキュメント"));
assert(sourceRow, "概要シートに翻訳出典注記あり");

// --- 明細シートの検証 ---
const detailData = XLSX.utils.sheet_to_json(asnWb.Sheets["明細"], { header: 1 });

// セクション見出しの検証（日本語名 / 技術名の両方含む）
const detailText = detailData.map(r => String(r[0] || "")).join("\n");
assert(detailText.includes("配送明細 / E1BPIBDLVITEM"), "明細セクション見出しに日本語名と技術名の両方");
assert(detailText.includes("拡張データ / E1BPEXTC"), "拡張データセクション見出しに両名");
assert(detailText.includes("配送期日 / E1BPDLVDEADLN"), "配送期日セクション見出しに両名");

// 2行ヘッダーの検証（日本語名 + 技術名）
const itemHeaderIdx = detailData.findIndex(r => r[0] && String(r[0]).includes("配送明細 / E1BPIBDLVITEM") && !String(r[0]).includes("ORG"));
if (itemHeaderIdx >= 0) {
  const jaHeaders = detailData[itemHeaderIdx + 1];  // 日本語名行
  const enHeaders = detailData[itemHeaderIdx + 2];  // 技術名行

  // 日本語名ヘッダー
  assert(jaHeaders[0] === "#", "行番号(#)ヘッダー");
  assert(jaHeaders.includes("明細番号"), "日本語ヘッダーに「明細番号」");
  assert(jaHeaders.includes("品目コード"), "日本語ヘッダーに「品目コード」");
  assert(jaHeaders.includes("品目テキスト"), "日本語ヘッダーに「品目テキスト」");
  assert(jaHeaders.includes("配送数量"), "日本語ヘッダーに「配送数量」");

  // 技術名ヘッダー
  assert(enHeaders.includes("ITM_NUMBER"), "技術名ヘッダーに「ITM_NUMBER」");
  assert(enHeaders.includes("MATERIAL"), "技術名ヘッダーに「MATERIAL」");
  assert(enHeaders.includes("SHORT_TEXT"), "技術名ヘッダーに「SHORT_TEXT」");
  assert(enHeaders.includes("DLV_QTY"), "技術名ヘッダーに「DLV_QTY」");

  // データ行（ヘッダー2行の後）
  const row1 = detailData[itemHeaderIdx + 3];
  assert(String(row1[0]) === "1", "データ行の行番号 = 1");
  const itmIdx = jaHeaders.indexOf("明細番号");
  assert(String(row1[itmIdx]) === "000010", "データ値は翻訳されない (000010)");
}

// セル結合の検証
const detailWs = asnWb.Sheets["明細"];
assert(detailWs["!merges"] && detailWs["!merges"].length > 0, "明細シートにセル結合あり");

// 明細の翻訳出典
const detailSourceRow = detailData.find(r => r[0] && String(r[0]).includes("SAP公式ドキュメント"));
assert(detailSourceRow, "明細シートに翻訳出典注記あり");

// テスト2: OBDS - 3列概要
console.log("\n📊 テスト2: OBDS XML → バイリンガル3列概要シート");
const obdsXml = fs.readFileSync(
  path.resolve(__dirname, "../../変換前データ/OBDS_8000580368_20180426103512.xml"), "utf-8"
);
const obdsParsed = parseXml(obdsXml, "OBDS_8000580368.xml");
const obdsBuffer = await generateXlsxBuffer(obdsParsed);
const obdsWb = XLSX.read(obdsBuffer, { type: "buffer" });

assert(obdsWb.SheetNames.length === 1, `OBDSは1シート`);
const obdsOverview = XLSX.utils.sheet_to_json(obdsWb.Sheets["概要"], { header: 1 });

// 3列構成の検証
assert(obdsOverview[0][0] === "項目名（日本語）", "OBDS概要ヘッダーA列");
assert(obdsOverview[0][1] === "項目名（技術名）", "OBDS概要ヘッダーB列");
assert(obdsOverview[0][2] === "値", "OBDS概要ヘッダーC列");

const rsSection = obdsOverview.find(r => r[1] === "RecordSet");
assert(rsSection && String(rsSection[0]).includes("レコードセット"), "RecordSetのA列に日本語名");

const routingSection = obdsOverview.find(r => r[1] === "FMS_ROUTING");
assert(routingSection && String(routingSection[0]).includes("ルーティング情報"), "FMS_ROUTINGのA列に日本語名");

const vbelnRow = obdsOverview.find(r => r[1] === "I_VBELN");
assert(vbelnRow && String(vbelnRow[0]).includes("伝票番号"), "I_VBELNのA列=「伝票番号」");
assert(vbelnRow && String(vbelnRow[2]) === "8000580368", "I_VBELNのC列=値");

// テスト3: 翻訳なしの汎用XML
console.log("\n📊 テスト3: 汎用XML → 翻訳なしで元の名前表示");
const genericXml = `<?xml version="1.0"?>
<root>
  <product><sku>ABC-123</sku><price>9.99</price></product>
  <product><sku>DEF-456</sku><price>19.99</price></product>
</root>`;
const genericParsed = parseXml(genericXml, "generic.xml");
const genericBuffer = await generateXlsxBuffer(genericParsed);
const genericWb = XLSX.read(genericBuffer, { type: "buffer" });

const genericDetail = XLSX.utils.sheet_to_json(genericWb.Sheets["明細"], { header: 1 });
// セクション見出し: 翻訳なしなので技術名のみ
assert(String(genericDetail[0][0]).includes("product"), "翻訳なし: セクション名はそのまま");
// 日本語ヘッダー行 = 技術名（翻訳がないため）
const genericJaHeaders = genericDetail[1];
assert(genericJaHeaders.includes("sku"), "翻訳なし: skuはそのまま");
assert(genericJaHeaders.includes("price"), "翻訳なし: priceはそのまま");
// 技術名ヘッダー行
const genericEnHeaders = genericDetail[2];
assert(genericEnHeaders.includes("sku"), "技術名行にskuあり");
assert(genericEnHeaders.includes("price"), "技術名行にpriceあり");
// データ値
assert(genericDetail[3][1] === "ABC-123", "翻訳なし: データ値そのまま");

// テスト4: スマート属性フィルタリング
console.log("\n📊 テスト4: スマート属性フィルタリング");
const attrXml = `<?xml version="1.0"?>
<root>
  <item status="active" type="A"><name>Item 1</name></item>
  <item status="inactive" type="A"><name>Item 2</name></item>
</root>`;
const attrParsed = parseXml(attrXml, "attr.xml");
const attrBuffer = await generateXlsxBuffer(attrParsed);
const attrWb = XLSX.read(attrBuffer, { type: "buffer" });
const attrDetail = XLSX.utils.sheet_to_json(attrWb.Sheets["明細"], { header: 1 });
const attrJaHeaders = attrDetail[1];
assert(attrJaHeaders.includes("status"), "値バラバラの@statusは保持");
assert(!attrJaHeaders.includes("type") && !attrJaHeaders.includes("@type"), "全行同値の@typeは除外");

// テスト5: データ完全性
console.log("\n📊 テスト5: ASNデータ完全性");
const allValues = new Set();
for (const row of overviewData) {
  // C列（値列）をチェック
  if (row[2] !== undefined && row[2] !== "") allValues.add(String(row[2]));
}
for (const row of detailData) {
  for (let i = 1; i < (row.length || 0); i++) {
    if (row[i] !== undefined && row[i] !== "") allValues.add(String(row[i]));
  }
}

assert(allValues.has("0000000059141367"), "DOCNUM値");
assert(allValues.has("1800024221"), "DELIV_NUMB値");
assert(allValues.has("000000010245236005"), "MATERIAL値");
assert(allValues.has("MULT-PACK CREW 3PK, 0090, S"), "SHORT_TEXT値");
assert(allValues.has("887749824742"), "EAN_UPC値");
assert(allValues.has("juily@waltknit.com"), "E_MAIL値");
assert(allValues.has("WALT TECHNOLOGY GROUP CO.,LTD."), "NAME値");

// テスト6: スタイリング検証（セル背景色 + 行高 — xlsx-js-styleの読み戻しで検証可能な項目）
// ※ 太字・フォント色・ボーダーはXLSXファイルに正しく書き込まれるが、
//    XLSX.read()の制約で読み戻し時に取得できないため、背景色と行高で検証する。
//    Excelで出力ファイルを開くと全スタイルが正しく反映されていることを確認済み。
console.log("\n📊 テスト6: セルスタイリング検証");
const styledWb = XLSX.read(asnBuffer, { type: "buffer", cellStyles: true });

// 概要シートの背景色
const ovWs = styledWb.Sheets["概要"];
const ovA1 = ovWs["A1"];
assert(ovA1?.s?.fgColor?.rgb === "2F5496", "概要ヘッダーA1: ダークブルー背景(#2F5496)");

const ovA2 = ovWs["A2"];
assert(ovA2?.s?.fgColor?.rgb === "D6E4F0", "概要セクション行: ライトブルー背景(#D6E4F0)");

// 行高の設定
assert(ovWs["!rows"]?.[0]?.hpt === 24, "概要ヘッダー行高=24pt");
assert(ovWs["!rows"]?.[1]?.hpt === 22, "概要セクション行高=22pt");

// 明細シートの背景色
const dtWs = styledWb.Sheets["明細"];
const dtA1 = dtWs["A1"];
assert(dtA1?.s?.fgColor?.rgb === "2F5496", "明細セクション見出し: ダークブルー背景(#2F5496)");

const dtA2 = dtWs["A2"];
assert(dtA2?.s?.fgColor?.rgb === "4472C4", "明細日本語ヘッダー行: ミディアムブルー背景(#4472C4)");

const dtA3 = dtWs["A3"];
assert(dtA3?.s?.fgColor?.rgb === "F2F2F2", "明細技術名行: ライトグレー背景(#F2F2F2)");

// 明細行高
assert(dtWs["!rows"]?.[0]?.hpt === 24, "明細セクション行高=24pt");
assert(dtWs["!rows"]?.[1]?.hpt === 20, "明細日本語ヘッダー行高=20pt");

// テスト7: シート分割（閾値テスト: 大規模テーブルが独立シートに分離）
console.log("\n📊 テスト7: シート分割 — 30行以上のテーブルが独立シートに分離");

// 35行のテーブル（閾値超過）+ 5行のテーブル（閾値未満）+ 1行セクション
function generateLargeXml(largeRowCount, smallRowCount) {
  let xml = `<?xml version="1.0"?>\n<root>\n  <header><title>Test</title><version>1.0</version></header>\n`;
  for (let i = 0; i < largeRowCount; i++) {
    xml += `  <E1BPEXTC><FIELD1>VAL_${i + 1}</FIELD1><FIELD2>DATA_${i + 1}</FIELD2></E1BPEXTC>\n`;
  }
  for (let i = 0; i < smallRowCount; i++) {
    xml += `  <E1BPIBDLVITEM><ITM_NUMBER>${String(i + 1).padStart(6, "0")}</ITM_NUMBER><MATERIAL>MAT_${i + 1}</MATERIAL></E1BPIBDLVITEM>\n`;
  }
  xml += `</root>`;
  return xml;
}

// 7a: 35行テーブル → 独立シート、5行テーブル → 明細に統合
const splitXml = generateLargeXml(35, 5);
const splitParsed = parseXml(splitXml, "split_test.xml");
const splitBuffer = await generateXlsxBuffer(splitParsed);
const splitWb = XLSX.read(splitBuffer, { type: "buffer", cellStyles: true });

assert(splitWb.SheetNames.length === 3, `3シート構成: 概要+明細+個別 (実際: ${splitWb.SheetNames.length})`);
assert(splitWb.SheetNames[0] === "概要", "シート1=概要");
assert(splitWb.SheetNames[1] === "明細", "シート2=明細（小規模テーブル）");
assert(splitWb.SheetNames[2] === "拡張データ", "シート3=拡張データ（日本語名で独立シート）");

// 明細シートには5行テーブルのみ
const splitDetail = XLSX.utils.sheet_to_json(splitWb.Sheets["明細"], { header: 1 });
const splitDetailText = splitDetail.map(r => String(r[0] || "")).join("\n");
assert(splitDetailText.includes("配送明細 / E1BPIBDLVITEM"), "明細に小規模テーブル(配送明細)あり");
assert(!splitDetailText.includes("拡張データ / E1BPEXTC"), "明細に大規模テーブル(拡張データ)なし");

// 独立シートに35行テーブル
const extSheet = splitWb.Sheets["拡張データ"];
const extData = XLSX.utils.sheet_to_json(extSheet, { header: 1 });
assert(String(extData[0][0]).includes("拡張データ / E1BPEXTC"), "独立シートにセクション見出しあり");
assert(extData[1][0] === "#", "独立シートに日本語ヘッダーあり");

// データ行数の検証（セクション1行 + ヘッダー2行 + 35データ行 + 空行 + 出典 = 40行）
const extDataRows = extData.filter(r => String(r[0] || "").match(/^\d+$/));
assert(extDataRows.length === 35, `独立シートに35データ行 (実際: ${extDataRows.length})`);
assert(String(extDataRows[0][1]) === "VAL_1", "独立シート1行目のデータ正しい");
assert(String(extDataRows[34][1]) === "VAL_35", "独立シート35行目のデータ正しい");

// フローズンペイン: JSZipで実際のXMLを検証
assert(extSheet["!cols"] && extSheet["!cols"].length > 0, "独立シートに列幅設定あり");
{
  const splitZip = await JSZip.loadAsync(splitBuffer);
  // 概要シート(sheet1): ySplit=1
  const sheet1Xml = await splitZip.file("xl/worksheets/sheet1.xml")?.async("string");
  assert(sheet1Xml?.includes('ySplit="1"') && sheet1Xml?.includes('state="frozen"'), "概要シートにフローズンペイン(ySplit=1)あり");
  // 独立シート(sheet3): ySplit=3
  const sheet3Xml = await splitZip.file("xl/worksheets/sheet3.xml")?.async("string");
  assert(sheet3Xml?.includes('ySplit="3"') && sheet3Xml?.includes('state="frozen"'), "独立シートにフローズンペイン(ySplit=3)あり");
}

// 7b: 29行テーブル（閾値未満）→ 全て明細に統合
const noSplitXml = generateLargeXml(29, 5);
const noSplitParsed = parseXml(noSplitXml, "no_split_test.xml");
const noSplitBuffer = await generateXlsxBuffer(noSplitParsed);
const noSplitWb = XLSX.read(noSplitBuffer, { type: "buffer" });

assert(noSplitWb.SheetNames.length === 2, `29行テーブルは分割なし: 2シート (実際: ${noSplitWb.SheetNames.length})`);
assert(noSplitWb.SheetNames[1] === "明細", "29行テーブルは明細に統合");

// 7c: 複数の大規模テーブル → それぞれ独立シート
function generateMultiLargeXml() {
  let xml = `<?xml version="1.0"?>\n<root>\n  <header><title>Multi</title></header>\n`;
  for (let i = 0; i < 40; i++) {
    xml += `  <E1BPEXTC><FIELD1>EXT_${i}</FIELD1></E1BPEXTC>\n`;
  }
  for (let i = 0; i < 50; i++) {
    xml += `  <E1BPIBDLVITEM><ITM_NUMBER>${i}</ITM_NUMBER></E1BPIBDLVITEM>\n`;
  }
  for (let i = 0; i < 3; i++) {
    xml += `  <E1BPDLVDEADLN><TIMETYPE>TYPE_${i}</TIMETYPE></E1BPDLVDEADLN>\n`;
  }
  xml += `</root>`;
  return xml;
}
const multiLargeXml = generateMultiLargeXml();
const multiLargeParsed = parseXml(multiLargeXml, "multi_large.xml");
const multiLargeBuffer = await generateXlsxBuffer(multiLargeParsed);
const multiLargeWb = XLSX.read(multiLargeBuffer, { type: "buffer" });

// 概要(header 1行) + 明細(3行テーブル) + 拡張データ(40行) + 配送明細(50行) = 4シート
assert(multiLargeWb.SheetNames.length === 4, `複数大規模テーブル: 4シート (実際: ${multiLargeWb.SheetNames.length})`);
assert(multiLargeWb.SheetNames.includes("拡張データ"), "拡張データが独立シート");
assert(multiLargeWb.SheetNames.includes("配送明細"), "配送明細が独立シート");
assert(multiLargeWb.SheetNames.includes("明細"), "小規模テーブルは明細に統合");

// 7d: 全テーブルが大規模 → 明細シートなし
function generateAllLargeXml() {
  let xml = `<?xml version="1.0"?>\n<root>\n  <header><title>AllLarge</title></header>\n`;
  for (let i = 0; i < 30; i++) {
    xml += `  <E1BPEXTC><FIELD1>F_${i}</FIELD1></E1BPEXTC>\n`;
  }
  for (let i = 0; i < 30; i++) {
    xml += `  <E1BPIBDLVITEM><ITM_NUMBER>${i}</ITM_NUMBER></E1BPIBDLVITEM>\n`;
  }
  xml += `</root>`;
  return xml;
}
const allLargeXml = generateAllLargeXml();
const allLargeParsed = parseXml(allLargeXml, "all_large.xml");
const allLargeBuffer = await generateXlsxBuffer(allLargeParsed);
const allLargeWb = XLSX.read(allLargeBuffer, { type: "buffer" });

assert(allLargeWb.SheetNames.length === 3, `全て大規模テーブル: 3シート=概要+個別2 (実際: ${allLargeWb.SheetNames.length})`);
assert(!allLargeWb.SheetNames.includes("明細"), "全て大規模の場合、明細シートなし");
assert(allLargeWb.SheetNames.includes("拡張データ"), "拡張データが独立シート");
assert(allLargeWb.SheetNames.includes("配送明細"), "配送明細が独立シート");

// 7e: 翻訳なしの大規模テーブル → 技術名がシート名
function generateGenericLargeXml() {
  let xml = `<?xml version="1.0"?>\n<root>\n`;
  for (let i = 0; i < 35; i++) {
    xml += `  <bigdata><col1>v${i}</col1><col2>d${i}</col2></bigdata>\n`;
  }
  xml += `</root>`;
  return xml;
}
const genericLargeXml = generateGenericLargeXml();
const genericLargeParsed = parseXml(genericLargeXml, "generic_large.xml");
const genericLargeBuffer = await generateXlsxBuffer(genericLargeParsed);
const genericLargeWb = XLSX.read(genericLargeBuffer, { type: "buffer" });

assert(genericLargeWb.SheetNames.includes("bigdata"), "翻訳なし大規模テーブルは技術名がシート名");
const bigdataRows = XLSX.utils.sheet_to_json(genericLargeWb.Sheets["bigdata"], { header: 1 })
  .filter(r => String(r[0] || "").match(/^\d+$/));
assert(bigdataRows.length === 35, `翻訳なし大規模テーブル: 35データ行 (実際: ${bigdataRows.length})`);

// テスト8: ファイル書き出し
console.log("\n📊 テスト8: ファイル書き出し");
const outputDir = path.resolve(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(path.join(outputDir, "ASN_styled.xlsx"), asnBuffer);
fs.writeFileSync(path.join(outputDir, "OBDS_styled.xlsx"), obdsBuffer);
fs.writeFileSync(path.join(outputDir, "SPLIT_test.xlsx"), splitBuffer);
fs.writeFileSync(path.join(outputDir, "MULTI_LARGE_test.xlsx"), multiLargeBuffer);
assert(true, "スタイル付きXLSXファイル出力済み（通常+分割テスト）");

// 結果
console.log(`\n${"=".repeat(50)}`);
console.log(`テスト結果: ${passed} passed, ${failed} failed`);
if (failed > 0) process.exit(1);
else console.log("✅ 全テスト合格！");
console.log(`\n📁 出力: ${outputDir}`);
