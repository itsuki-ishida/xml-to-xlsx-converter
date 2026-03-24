/**
 * XLSX生成テスト（2シート構成フォーマット）
 * 実際にXLSXファイルを生成し、読み戻して内容を検証する
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

// --- 新しい2シート構成のXLSX生成ロジック ---
function generateXlsxBuffer(parsed) {
  const wb = XLSX.utils.book_new();
  const singleSheets = parsed.sheets.filter((s) => s.rows.length === 1);
  const multiSheets = parsed.sheets.filter((s) => s.rows.length >= 2);

  // シート1: 概要
  if (singleSheets.length > 0) {
    const aoa = [["セクション / 項目名", "値"]];
    for (let si = 0; si < singleSheets.length; si++) {
      const sheet = singleSheets[si];
      const row = sheet.rows[0];
      if (si > 0) aoa.push([]);
      aoa.push([`■ ${sheet.name}`, ""]);
      for (const header of sheet.headers) {
        if (header.startsWith("@")) continue;
        aoa.push([`  ${header}`, row[header] ?? ""]);
      }
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 28 }, { wch: 50 }];
    XLSX.utils.book_append_sheet(wb, ws, "概要");
  }

  // シート2: 明細
  if (multiSheets.length > 0) {
    const aoa = [];
    for (let si = 0; si < multiSheets.length; si++) {
      const sheet = multiSheets[si];
      if (si > 0) aoa.push([]);
      const dataHeaders = sheet.headers.filter((h) => !h.startsWith("@"));
      aoa.push([`■ ${sheet.name} (${sheet.rows.length}件)`, ...Array(dataHeaders.length - 1).fill("")]);
      aoa.push(dataHeaders);
      for (const row of sheet.rows) {
        aoa.push(dataHeaders.map((h) => row[h] ?? ""));
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
console.log("\n📊 テスト1: ASN XML → 2シート構成XLSX");
const asnXml = fs.readFileSync(
  path.resolve(__dirname, "../../変換前データ/ASN_1800024221_20190214185557.xml"),
  "utf-8"
);
const asnParsed = parseXml(asnXml, "ASN.xml");
const asnBuffer = generateXlsxBuffer(asnParsed);
const asnWb = XLSX.read(asnBuffer, { type: "buffer" });

assert(asnWb.SheetNames.length === 2, `2シート構成 (実際: ${asnWb.SheetNames.length})`);
assert(asnWb.SheetNames[0] === "概要", `シート1の名前が「概要」`);
assert(asnWb.SheetNames[1] === "明細", `シート2の名前が「明細」`);

// 概要シートの検証
const overviewWs = asnWb.Sheets["概要"];
const overviewData = XLSX.utils.sheet_to_json(overviewWs, { header: 1 });
assert(overviewData[0][0] === "セクション / 項目名", "概要シートのヘッダーが正しい");

// セクションヘッダーが含まれていることを確認
const overviewText = overviewData.map(r => r[0] || "").join("\n");
assert(overviewText.includes("■ EDI_DC40"), "概要に EDI_DC40 セクションが含まれる");
assert(overviewText.includes("■ E1BPIBDLVHDR"), "概要に E1BPIBDLVHDR セクションが含まれる");
assert(overviewText.includes("■ E1BPADR1"), "概要に E1BPADR1 セクションが含まれる");

// DOCNUMの値が概要に含まれることを確認
const docnumRow = overviewData.find(r => r[0] && r[0].includes("DOCNUM"));
assert(docnumRow && docnumRow[1] === "0000000059141367", "概要にDOCNUM値が正しく含まれる");

// 明細シートの検証
const detailWs = asnWb.Sheets["明細"];
const detailData = XLSX.utils.sheet_to_json(detailWs, { header: 1 });
const detailText = detailData.map(r => r[0] || "").join("\n");
assert(detailText.includes("■ E1BPIBDLVITEM"), "明細に E1BPIBDLVITEM セクションが含まれる");
assert(detailText.includes("■ E1BPEXTC"), "明細に E1BPEXTC セクションが含まれる");

// E1BPIBDLVITEM テーブルデータの検証
const itemHeaderIdx = detailData.findIndex(r => r[0] && String(r[0]).includes("E1BPIBDLVITEM"));
if (itemHeaderIdx >= 0) {
  const colHeaders = detailData[itemHeaderIdx + 1];
  const row1 = detailData[itemHeaderIdx + 2];
  const row2 = detailData[itemHeaderIdx + 3];
  assert(colHeaders.includes("ITM_NUMBER"), "明細テーブルにITM_NUMBERヘッダーがある");
  const itmIdx = colHeaders.indexOf("ITM_NUMBER");
  assert(String(row1[itmIdx]) === "000010", "明細の1行目 ITM_NUMBER = 000010");
  assert(String(row2[itmIdx]) === "000020", "明細の2行目 ITM_NUMBER = 000020");
}

// テスト2: OBDSファイル - 概要のみ（全て単一インスタンス）
console.log("\n📊 テスト2: OBDS XML → 概要シートのみ");
const obdsXml = fs.readFileSync(
  path.resolve(__dirname, "../../変換前データ/OBDS_8000580368_20180426103512.xml"),
  "utf-8"
);
const obdsParsed = parseXml(obdsXml, "OBDS.xml");
const obdsBuffer = generateXlsxBuffer(obdsParsed);
const obdsWb = XLSX.read(obdsBuffer, { type: "buffer" });

assert(obdsWb.SheetNames.length === 1, `OBDSは1シートのみ (実際: ${obdsWb.SheetNames.length})`);
assert(obdsWb.SheetNames[0] === "概要", "OBDSは概要シートのみ");

const obdsOverview = XLSX.utils.sheet_to_json(obdsWb.Sheets["概要"], { header: 1 });
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

// テスト4: ファイル書き出し
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
