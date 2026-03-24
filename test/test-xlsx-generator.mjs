/**
 * XLSX生成テスト（バイリンガル対応: 日本語名 + 技術名の両方表示）
 * - 概要シート: 3列構成（日本語名 / 技術名 / 値）
 * - 明細シート: セクション見出しセル結合 + 2行ヘッダー（日本語名 + 技術名）
 * - SAP公式ドキュメント準拠の翻訳
 * - データ完全性の検証
 */
import { JSDOM } from "jsdom";
import XLSX from "xlsx";
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

// === XLSX生成ロジック（バイリンガル対応） ===
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

function generateXlsxBuffer(parsed) {
  const wb = XLSX.utils.book_new();
  const singleSheets = parsed.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = parsed.sheets.filter((s) => s.rows.length >= 2);

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
    XLSX.utils.book_append_sheet(wb, ws, "概要");
  }

  // シート2: 明細（セクション見出しセル結合 + 2行ヘッダー）
  if (multiSheets.length > 0) {
    const aoa = [];
    const merges = [];
    for (let si = 0; si < multiSheets.length; si++) {
      const sheet = multiSheets[si];
      if (si > 0) aoa.push([]);
      const displayHeaders = getDisplayHeaders(sheet);
      const totalCols = 1 + displayHeaders.length;

      // セクションヘッダー（セル結合）
      const sectionRowIdx = aoa.length;
      const sectionJa = translateSectionShort(sheet.name);
      const sectionLabel = sectionJa !== sheet.name
        ? `■ ${sectionJa} / ${sheet.name} (${sheet.rows.length}件)`
        : `■ ${sheet.name} (${sheet.rows.length}件)`;
      aoa.push([sectionLabel, ...Array(displayHeaders.length).fill("")]);
      if (totalCols > 1) {
        merges.push({ s: { r: sectionRowIdx, c: 0 }, e: { r: sectionRowIdx, c: totalCols - 1 } });
      }

      // 列ヘッダー1行目: 日本語名
      aoa.push(["#", ...displayHeaders.map(translateFieldShort)]);

      // 列ヘッダー2行目: 技術名
      aoa.push(["", ...displayHeaders.map(stripFieldPrefix)]);

      // データ行
      for (let ri = 0; ri < sheet.rows.length; ri++) {
        const row = sheet.rows[ri];
        aoa.push([String(ri + 1), ...displayHeaders.map((h) => row[h] ?? "")]);
      }
    }
    aoa.push([]);
    aoa.push([`※ ${TRANSLATION_SOURCE}`]);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (merges.length > 0) ws["!merges"] = merges;
    XLSX.utils.book_append_sheet(wb, ws, "明細");
  }

  return XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
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
const asnBuffer = generateXlsxBuffer(asnParsed);
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
const obdsBuffer = generateXlsxBuffer(obdsParsed);
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
const genericBuffer = generateXlsxBuffer(genericParsed);
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
const attrBuffer = generateXlsxBuffer(attrParsed);
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

// テスト6: ファイル書き出し
console.log("\n📊 テスト6: ファイル書き出し");
const outputDir = path.resolve(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(path.join(outputDir, "ASN_bilingual.xlsx"), asnBuffer);
fs.writeFileSync(path.join(outputDir, "OBDS_bilingual.xlsx"), obdsBuffer);
assert(true, "XLSXファイル出力済み");

// 結果
console.log(`\n${"=".repeat(50)}`);
console.log(`テスト結果: ${passed} passed, ${failed} failed`);
if (failed > 0) process.exit(1);
else console.log("✅ 全テスト合格！");
console.log(`\n📁 出力: ${outputDir}`);
