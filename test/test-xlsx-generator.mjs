/**
 * XLSX生成テスト（2シート構成 + 翻訳機能）
 * - ファイル情報ヘッダー
 * - スマート属性フィルタリング
 * - 明細テーブルに行番号(#)付き
 * - フィールド名・セクション名の翻訳
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

// === 翻訳辞書（インライン: 本体は src/lib/translations.ts） ===
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
};

function translateSection(name) {
  const ja = SECTION_TRANSLATIONS[name];
  return ja ? `${ja} (${name})` : name;
}
function translateField(name) {
  const raw = name.startsWith("@") ? name.substring(1) : name;
  const ja = FIELD_TRANSLATIONS[raw];
  return ja ? `${ja} (${raw})` : raw;
}
function translateFieldShort(name) {
  const raw = name.startsWith("@") ? name.substring(1) : name;
  return FIELD_TRANSLATIONS[raw] ?? raw;
}

// === XLSX生成ロジック（翻訳付き） ===
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

  // シート1: 概要（翻訳付き）
  if (singleSheets.length > 0) {
    const aoa = [["セクション / 項目名", "値"]];
    aoa.push(["■ ファイル情報", ""]);
    aoa.push(["  ファイル名", parsed.fileName]);
    aoa.push(["  ルート要素", parsed.rootElement]);
    aoa.push([]);

    for (let si = 0; si < singleSheets.length; si++) {
      const sheet = singleSheets[si];
      const row = sheet.rows[0];
      const displayHeaders = getDisplayHeaders(sheet);
      aoa.push([`■ ${translateSection(sheet.name)}`, ""]);
      for (const header of displayHeaders) {
        aoa.push([`  ${translateField(header)}`, row[header] ?? ""]);
      }
      if (si < singleSheets.length - 1) aoa.push([]);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 36 }, { wch: 55 }];
    XLSX.utils.book_append_sheet(wb, ws, "概要");
  }

  // シート2: 明細（翻訳付き）
  if (multiSheets.length > 0) {
    const aoa = [];
    for (let si = 0; si < multiSheets.length; si++) {
      const sheet = multiSheets[si];
      if (si > 0) aoa.push([]);
      const displayHeaders = getDisplayHeaders(sheet);
      aoa.push([`■ ${translateSection(sheet.name)} (${sheet.rows.length}件)`, ...Array(displayHeaders.length).fill("")]);
      aoa.push(["#", ...displayHeaders.map(translateFieldShort)]);
      for (let ri = 0; ri < sheet.rows.length; ri++) {
        const row = sheet.rows[ri];
        aoa.push([String(ri + 1), ...displayHeaders.map((h) => row[h] ?? "")]);
      }
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
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

// テスト1: ASN - 2シート構成 + 翻訳
console.log("\n📊 テスト1: ASN XML → 翻訳付き2シート構成XLSX");
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
const overviewText = overviewData.map(r => r[0] || "").join("\n");

// ファイル情報
assert(overviewData[1][0] === "■ ファイル情報", "ファイル情報セクション存在");
assert(overviewData[2][1] === "ASN_1800024221.xml", "ファイル名正しい");

// 翻訳付きセクション名の検証
assert(overviewText.includes("IDoc制御情報 (EDI_DC40)"), "EDI_DC40が翻訳付きで表示");
assert(overviewText.includes("配送ヘッダー (E1BPIBDLVHDR)"), "E1BPIBDLVHDRが翻訳付きで表示");
assert(overviewText.includes("住所情報 (E1BPADR1)"), "E1BPADR1が翻訳付きで表示");
assert(overviewText.includes("住所詳細 (E1BPADR11)"), "E1BPADR11が翻訳付きで表示");

// 翻訳付きフィールド名の検証（概要: フル表記）
const docnumRow = overviewData.find(r => r[0] && r[0].includes("DOCNUM"));
assert(docnumRow && docnumRow[0].includes("ドキュメント番号"), "DOCNUMが「ドキュメント番号」と翻訳表示");
assert(docnumRow && docnumRow[1] === "0000000059141367", "DOCNUM値正しい");

const nameRow = overviewData.find(r => r[0] && r[0].includes("NAME") && !r[0].includes("TABNAM"));
assert(nameRow && nameRow[0].includes("名前"), "NAMEが「名前」と翻訳表示");

// @SEGMENT除外の検証
assert(!overviewText.includes("SEGMENT"), "@SEGMENT除外");

// 明細シートの検証
const detailData = XLSX.utils.sheet_to_json(asnWb.Sheets["明細"], { header: 1 });
const detailText = detailData.map(r => r[0] || "").join("\n");

// 翻訳付きセクション名（明細）
assert(detailText.includes("配送明細 (E1BPIBDLVITEM)"), "明細のE1BPIBDLVITEMが翻訳付き");
assert(detailText.includes("拡張データ (E1BPEXTC)"), "明細のE1BPEXTCが翻訳付き");
assert(detailText.includes("配送期日 (E1BPDLVDEADLN)"), "明細のE1BPDLVDEADLNが翻訳付き");

// 翻訳付き列ヘッダー（明細: 短縮表記）
const itemHeaderIdx = detailData.findIndex(r => r[0] && String(r[0]).includes("配送明細 (E1BPIBDLVITEM)") && !String(r[0]).includes("ORG"));
if (itemHeaderIdx >= 0) {
  const colHeaders = detailData[itemHeaderIdx + 1];
  assert(colHeaders[0] === "#", "行番号(#)ヘッダー");
  assert(colHeaders.includes("明細番号"), "ITM_NUMBERが「明細番号」と翻訳表示");
  assert(colHeaders.includes("品目コード"), "MATERIALが「品目コード」と翻訳表示");
  assert(colHeaders.includes("品目テキスト"), "SHORT_TEXTが「品目テキスト」と翻訳表示");
  assert(colHeaders.includes("配送数量"), "DLV_QTYが「配送数量」と翻訳表示");

  // データ値は翻訳されない（元の値のまま）
  const row1 = detailData[itemHeaderIdx + 2];
  const itmIdx = colHeaders.indexOf("明細番号");
  assert(String(row1[itmIdx]) === "000010", "データ値は翻訳されない (000010)");
}

// テスト2: OBDS - 翻訳付き
console.log("\n📊 テスト2: OBDS XML → 翻訳付き概要シート");
const obdsXml = fs.readFileSync(
  path.resolve(__dirname, "../../変換前データ/OBDS_8000580368_20180426103512.xml"), "utf-8"
);
const obdsParsed = parseXml(obdsXml, "OBDS_8000580368.xml");
const obdsBuffer = generateXlsxBuffer(obdsParsed);
const obdsWb = XLSX.read(obdsBuffer, { type: "buffer" });

assert(obdsWb.SheetNames.length === 1, `OBDSは1シート`);
const obdsOverview = XLSX.utils.sheet_to_json(obdsWb.Sheets["概要"], { header: 1 });
const obdsText = obdsOverview.map(r => r[0] || "").join("\n");

assert(obdsText.includes("レコードセット (RecordSet)"), "RecordSetが翻訳付き");
assert(obdsText.includes("ルーティング情報 (FMS_ROUTING)"), "FMS_ROUTINGが翻訳付き");

const vbelnRow = obdsOverview.find(r => r[0] && r[0].includes("I_VBELN"));
assert(vbelnRow && vbelnRow[0].includes("伝票番号"), "I_VBELNが「伝票番号」と翻訳表示");
assert(vbelnRow && String(vbelnRow[1]) === "8000580368", "I_VBELN値正しい");

// テスト3: 翻訳なしの汎用XML → 元の名前がそのまま表示
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
// 翻訳辞書にない "product" はそのまま表示
assert(genericDetail[0][0].includes("product"), "翻訳なし: セクション名はそのまま");
const genericHeaders = genericDetail[1];
assert(genericHeaders.includes("sku"), "翻訳なし: skuはそのまま");
assert(genericHeaders.includes("price"), "翻訳なし: priceはそのまま");
// データ値もそのまま
assert(genericDetail[2][1] === "ABC-123", "翻訳なし: データ値そのまま");

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
const attrHeaders = attrDetail[1];
assert(attrHeaders.includes("status"), "値バラバラの@statusは保持");
assert(!attrHeaders.includes("type") && !attrHeaders.includes("@type"), "全行同値の@typeは除外");

// テスト5: データ完全性
console.log("\n📊 テスト5: ASNデータ完全性");
const allValues = new Set();
for (const row of overviewData) {
  if (row[1] !== undefined && row[1] !== "") allValues.add(String(row[1]));
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
fs.writeFileSync(path.join(outputDir, "ASN_2sheet.xlsx"), asnBuffer);
fs.writeFileSync(path.join(outputDir, "OBDS_2sheet.xlsx"), obdsBuffer);
assert(true, "XLSXファイル出力済み");

// 結果
console.log(`\n${"=".repeat(50)}`);
console.log(`テスト結果: ${passed} passed, ${failed} failed`);
if (failed > 0) process.exit(1);
else console.log("✅ 全テスト合格！");
console.log(`\n📁 出力: ${outputDir}`);
