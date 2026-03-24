/**
 * XLSX生成テスト（2シート構成フォーマット v2）
 * - ファイル情報ヘッダー
 * - スマート属性フィルタリング（全行同一値の@属性は除外）
 * - 明細テーブルに行番号(#)付き
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

// xml-parser のロジック（インライン再現）
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

// --- 新しいXLSX生成ロジック（v2: ファイル情報+行番号+スマート属性フィルタ） ---
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

function displayHeaderName(h) {
  return h.startsWith("@") ? h.substring(1) : h;
}

function generateXlsxBuffer(parsed) {
  const wb = XLSX.utils.book_new();
  const singleSheets = parsed.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = parsed.sheets.filter((s) => s.rows.length >= 2);

  // シート1: 概要
  if (singleSheets.length > 0) {
    const aoa = [["セクション / 項目名", "値"]];
    // ファイル情報
    aoa.push(["■ ファイル情報", ""]);
    aoa.push(["  ファイル名", parsed.fileName]);
    aoa.push(["  ルート要素", parsed.rootElement]);
    aoa.push([]);

    for (let si = 0; si < singleSheets.length; si++) {
      const sheet = singleSheets[si];
      const row = sheet.rows[0];
      const displayHeaders = getDisplayHeaders(sheet);
      aoa.push([`■ ${sheet.name}`, ""]);
      for (const header of displayHeaders) {
        aoa.push([`  ${displayHeaderName(header)}`, row[header] ?? ""]);
      }
      if (si < singleSheets.length - 1) aoa.push([]);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 30 }, { wch: 55 }];
    XLSX.utils.book_append_sheet(wb, ws, "概要");
  }

  // シート2: 明細
  if (multiSheets.length > 0) {
    const aoa = [];
    for (let si = 0; si < multiSheets.length; si++) {
      const sheet = multiSheets[si];
      if (si > 0) aoa.push([]);
      const displayHeaders = getDisplayHeaders(sheet);
      const displayNames = displayHeaders.map(displayHeaderName);
      aoa.push([`■ ${sheet.name} (${sheet.rows.length}件)`, ...Array(displayNames.length).fill("")]);
      aoa.push(["#", ...displayNames]);
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

// テスト1: ASNファイル - 2シート構成
console.log("\n📊 テスト1: ASN XML → 2シート構成XLSX (v2)");
const asnXml = fs.readFileSync(
  path.resolve(__dirname, "../../変換前データ/ASN_1800024221_20190214185557.xml"),
  "utf-8"
);
const asnParsed = parseXml(asnXml, "ASN_1800024221.xml");
const asnBuffer = generateXlsxBuffer(asnParsed);
const asnWb = XLSX.read(asnBuffer, { type: "buffer" });

assert(asnWb.SheetNames.length === 2, `2シート構成 (実際: ${asnWb.SheetNames.length})`);
assert(asnWb.SheetNames[0] === "概要", `シート1の名前が「概要」`);
assert(asnWb.SheetNames[1] === "明細", `シート2の名前が「明細」`);

// 概要シートの検証
const overviewWs = asnWb.Sheets["概要"];
const overviewData = XLSX.utils.sheet_to_json(overviewWs, { header: 1 });
assert(overviewData[0][0] === "セクション / 項目名", "概要シートのヘッダーが正しい");

// ファイル情報セクション
assert(overviewData[1][0] === "■ ファイル情報", "ファイル情報セクションが存在");
assert(overviewData[2][0] === "  ファイル名" && overviewData[2][1] === "ASN_1800024221.xml", "ファイル名が正しい");
assert(overviewData[3][0] === "  ルート要素" && overviewData[3][1] === "SHP_IBDLV_SAVE_REPLICA04", "ルート要素が正しい");

// セクションヘッダーが含まれていることを確認
const overviewText = overviewData.map(r => r[0] || "").join("\n");
assert(overviewText.includes("■ EDI_DC40"), "概要に EDI_DC40 セクションが含まれる");
assert(overviewText.includes("■ E1BPIBDLVHDR"), "概要に E1BPIBDLVHDR セクションが含まれる");
assert(overviewText.includes("■ E1BPADR1"), "概要に E1BPADR1 セクションが含まれる");

// DOCNUMの値が概要に含まれることを確認
const docnumRow = overviewData.find(r => r[0] && r[0].includes("DOCNUM"));
assert(docnumRow && docnumRow[1] === "0000000059141367", "概要にDOCNUM値が正しく含まれる");

// @SEGMENT属性が除外されていることを確認（全行同一値"1"なので）
const segmentInOverview = overviewData.find(r => r[0] && r[0].includes("SEGMENT"));
assert(!segmentInOverview, "@SEGMENT属性が概要から除外されている（全行同一値）");

// 明細シートの検証
const detailWs = asnWb.Sheets["明細"];
const detailData = XLSX.utils.sheet_to_json(detailWs, { header: 1 });
const detailText = detailData.map(r => r[0] || "").join("\n");
assert(detailText.includes("■ E1BPIBDLVITEM"), "明細に E1BPIBDLVITEM セクションが含まれる");
assert(detailText.includes("■ E1BPEXTC"), "明細に E1BPEXTC セクションが含まれる");

// E1BPIBDLVITEM テーブルの行番号検証
const itemHeaderIdx = detailData.findIndex(r => r[0] && String(r[0]).includes("E1BPIBDLVITEM") && !String(r[0]).includes("ORG"));
if (itemHeaderIdx >= 0) {
  const colHeaders = detailData[itemHeaderIdx + 1];
  assert(colHeaders[0] === "#", "明細テーブルに行番号(#)ヘッダーがある");

  const row1 = detailData[itemHeaderIdx + 2];
  const row2 = detailData[itemHeaderIdx + 3];
  assert(String(row1[0]) === "1", "明細の1行目 # = 1");
  assert(String(row2[0]) === "2", "明細の2行目 # = 2");

  const itmIdx = colHeaders.indexOf("ITM_NUMBER");
  assert(String(row1[itmIdx]) === "000010", "明細の1行目 ITM_NUMBER = 000010");
  assert(String(row2[itmIdx]) === "000020", "明細の2行目 ITM_NUMBER = 000020");
}

// @SEGMENT属性が明細からも除外されていることを確認
const detailHeaders = detailData.find(r => r[0] === "#" && r.includes("ITM_NUMBER"));
if (detailHeaders) {
  assert(!detailHeaders.includes("SEGMENT") && !detailHeaders.includes("@SEGMENT"),
    "@SEGMENT属性が明細テーブルからも除外されている");
}

// テスト2: OBDSファイル - 概要のみ（全て単一インスタンス）
console.log("\n📊 テスト2: OBDS XML → 概要シートのみ");
const obdsXml = fs.readFileSync(
  path.resolve(__dirname, "../../変換前データ/OBDS_8000580368_20180426103512.xml"),
  "utf-8"
);
const obdsParsed = parseXml(obdsXml, "OBDS_8000580368.xml");
const obdsBuffer = generateXlsxBuffer(obdsParsed);
const obdsWb = XLSX.read(obdsBuffer, { type: "buffer" });

assert(obdsWb.SheetNames.length === 1, `OBDSは1シートのみ (実際: ${obdsWb.SheetNames.length})`);
assert(obdsWb.SheetNames[0] === "概要", "OBDSは概要シートのみ");

const obdsOverview = XLSX.utils.sheet_to_json(obdsWb.Sheets["概要"], { header: 1 });
// ファイル情報
const obdsFileNameRow = obdsOverview.find(r => r[0] === "  ファイル名");
assert(obdsFileNameRow && obdsFileNameRow[1] === "OBDS_8000580368.xml", "OBDSファイル情報が正しい");

const vbelnRow = obdsOverview.find(r => r[0] && r[0].includes("I_VBELN"));
assert(vbelnRow && String(vbelnRow[1]) === "8000580368", "概要にI_VBELN値が正しく含まれる");

// テスト3: 全て複数インスタンスのXML → 明細のみ
console.log("\n📊 テスト3: 全て複数インスタンス → 明細シートのみ");
const multiXml = `<?xml version="1.0"?>
<root>
  <item><a>1</a><b>x</b></item>
  <item><a>2</a><b>y</b></item>
  <item><a>3</a><b>z</b></item>
</root>`;
const multiParsed = parseXml(multiXml, "multi.xml");
const multiBuffer = generateXlsxBuffer(multiParsed);
const multiWb = XLSX.read(multiBuffer, { type: "buffer" });

assert(multiWb.SheetNames.length === 1, `全複数は1シートのみ (実際: ${multiWb.SheetNames.length})`);
assert(multiWb.SheetNames[0] === "明細", "全複数は明細シートのみ");

// 行番号の検証
const multiDetail = XLSX.utils.sheet_to_json(multiWb.Sheets["明細"], { header: 1 });
assert(multiDetail[1][0] === "#", "明細テーブルに#ヘッダー");
assert(String(multiDetail[2][0]) === "1", "行番号1");
assert(String(multiDetail[3][0]) === "2", "行番号2");
assert(String(multiDetail[4][0]) === "3", "行番号3");

// テスト4: スマート属性フィルタリング（値がバラバラの属性は保持）
console.log("\n📊 テスト4: スマート属性フィルタリング");
const attrXml = `<?xml version="1.0"?>
<root>
  <item status="active" type="A"><name>Item 1</name></item>
  <item status="inactive" type="A"><name>Item 2</name></item>
  <item status="active" type="A"><name>Item 3</name></item>
</root>`;
const attrParsed = parseXml(attrXml, "attr.xml");
const attrBuffer = generateXlsxBuffer(attrParsed);
const attrWb = XLSX.read(attrBuffer, { type: "buffer" });

const attrDetail = XLSX.utils.sheet_to_json(attrWb.Sheets["明細"], { header: 1 });
const attrHeaders = attrDetail[1]; // [#, name, status] - type should be excluded
assert(attrHeaders.includes("status"), "値がバラバラの@status属性が保持される");
assert(!attrHeaders.includes("type") && !attrHeaders.includes("@type"), "全行同一値の@type属性が除外される");

// テスト5: データ完全性 - ASNの全リーフテキスト値が出力に含まれること
console.log("\n📊 テスト5: ASNデータ完全性検証");
// 概要シートの全値を収集
const allOverviewValues = new Set();
for (const row of overviewData) {
  if (row[1] !== undefined && row[1] !== "") allOverviewValues.add(String(row[1]));
}
// 明細シートの全値を収集
const allDetailValues = new Set();
for (const row of detailData) {
  for (let i = 1; i < (row.length || 0); i++) {
    if (row[i] !== undefined && row[i] !== "") allDetailValues.add(String(row[i]));
  }
}
const allValues = new Set([...allOverviewValues, ...allDetailValues]);

// 重要なビジネスデータが含まれていることを確認
assert(allValues.has("0000000059141367"), "DOCNUM値が出力に含まれる");
assert(allValues.has("1800024221"), "DELIV_NUMB値が出力に含まれる");
assert(allValues.has("000000010245236005"), "MATERIAL値が出力に含まれる");
assert(allValues.has("MULT-PACK CREW 3PK, 0090, S"), "SHORT_TEXT値が出力に含まれる");
assert(allValues.has("887749824742"), "EAN_UPC値が出力に含まれる");
assert(allValues.has("juily@waltknit.com"), "E_MAIL値が出力に含まれる");
assert(allValues.has("WALT TECHNOLOGY GROUP CO.,LTD."), "NAME値が出力に含まれる");
assert(allValues.has("EXTWHAAG01"), "SNDPRN/RECV_SYS値が出力に含まれる");
assert(allValues.has("FOBPRICE"), "E1BPEXTC FIELD1値が出力に含まれる");

// テスト6: ファイル書き出し
console.log("\n📊 テスト6: ファイル書き出し");
const outputDir = path.resolve(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

fs.writeFileSync(path.join(outputDir, "ASN_2sheet.xlsx"), asnBuffer);
fs.writeFileSync(path.join(outputDir, "OBDS_2sheet.xlsx"), obdsBuffer);
assert(true, "XLSXファイルを出力済み");

// 結果
console.log(`\n${"=".repeat(50)}`);
console.log(`テスト結果: ${passed} passed, ${failed} failed`);
if (failed > 0) process.exit(1);
else console.log("✅ 全テスト合格！");
console.log(`\n📁 出力: ${outputDir}`);
