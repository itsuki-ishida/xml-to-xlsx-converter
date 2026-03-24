/**
 * XLSX生成テスト
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

// xml-parser のロジック（前のテストと同じ）
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

// xlsx-generator のロジック
function generateXlsxBuffer(parsed) {
  const wb = XLSX.utils.book_new();
  for (const sheet of parsed.sheets) {
    const aoa = [];
    aoa.push(sheet.headers);
    for (const row of sheet.rows) {
      aoa.push(sheet.headers.map((h) => row[h] ?? ""));
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const colWidths = sheet.headers.map((h) => {
      let maxLen = h.length;
      for (const row of sheet.rows) {
        maxLen = Math.max(maxLen, (row[h] ?? "").length);
      }
      return Math.min(Math.max(maxLen + 2, 8), 60);
    });
    ws["!cols"] = colWidths.map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
  }
  return XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
}

// テスト
let passed = 0;
let failed = 0;
function assert(condition, message) {
  if (condition) { passed++; console.log(`  ✅ ${message}`); }
  else { failed++; console.error(`  ❌ ${message}`); }
}

// テスト1: ASNファイルのXLSX生成と読み戻し
console.log("\n📊 テスト1: ASN XMLからXLSXを生成し読み戻し検証");
const asnXml = fs.readFileSync(
  path.resolve(__dirname, "../../変換前データ/ASN_1800024221_20190214185557.xml"),
  "utf-8"
);
const asnParsed = parseXml(asnXml, "ASN.xml");
const asnBuffer = generateXlsxBuffer(asnParsed);

// バッファから読み戻し
const asnWb = XLSX.read(asnBuffer, { type: "buffer" });
assert(asnWb.SheetNames.length === asnParsed.sheets.length, `シート数が一致 (${asnWb.SheetNames.length})`);
console.log(`  📋 XLSXシート名: ${asnWb.SheetNames.join(", ")}`);

// EDI_DC40 シートの内容検証
const ediWs = asnWb.Sheets["EDI_DC40"];
const ediData = XLSX.utils.sheet_to_json(ediWs);
assert(ediData.length === 1, "EDI_DC40は1行");
assert(ediData[0]["DOCNUM"] === "0000000059141367", "XLSXのDOCNUM値が正しい");

// E1BPIBDLVITEM シートの内容検証
const itemWs = asnWb.Sheets["E1BPIBDLVITEM"];
const itemData = XLSX.utils.sheet_to_json(itemWs);
assert(itemData.length === 2, "E1BPIBDLVITEMは2行");
assert(String(itemData[0]["ITM_NUMBER"]) === "000010", "ITM_NUMBERが正しい");

// E1BPEXTC シートの内容検証
const extcWs = asnWb.Sheets["E1BPEXTC"];
const extcData = XLSX.utils.sheet_to_json(extcWs);
assert(extcData.length === 22, `E1BPEXTCは22行 (実際: ${extcData.length})`);

// テスト2: OBDSファイルのXLSX生成と読み戻し
console.log("\n📊 テスト2: OBDS XMLからXLSXを生成し読み戻し検証");
const obdsXml = fs.readFileSync(
  path.resolve(__dirname, "../../変換前データ/OBDS_8000580368_20180426103512.xml"),
  "utf-8"
);
const obdsParsed = parseXml(obdsXml, "OBDS.xml");
const obdsBuffer = generateXlsxBuffer(obdsParsed);
const obdsWb = XLSX.read(obdsBuffer, { type: "buffer" });

assert(obdsWb.SheetNames.length > 0, `OBDSシートが生成された (${obdsWb.SheetNames.length})`);
console.log(`  📋 XLSXシート名: ${obdsWb.SheetNames.join(", ")}`);

const rsWs = obdsWb.Sheets["RecordSet"];
if (rsWs) {
  const rsData = XLSX.utils.sheet_to_json(rsWs);
  assert(rsData.length === 1, "RecordSetは1行");
  assert(String(rsData[0]["I_VBELN"]) === "8000580368", "I_VBELN値が正しい");
}

// テスト3: ファイルをディスクに書き出して確認
const outputDir = path.resolve(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

const asnOutPath = path.join(outputDir, "ASN_converted.xlsx");
fs.writeFileSync(asnOutPath, asnBuffer);
assert(fs.existsSync(asnOutPath), `ASN XLSXファイルが生成された: ${asnOutPath}`);
assert(fs.statSync(asnOutPath).size > 0, "ASN XLSXファイルサイズが0より大きい");

const obdsOutPath = path.join(outputDir, "OBDS_converted.xlsx");
fs.writeFileSync(obdsOutPath, obdsBuffer);
assert(fs.existsSync(obdsOutPath), `OBDS XLSXファイルが生成された: ${obdsOutPath}`);

// 結果
console.log(`\n${"=".repeat(50)}`);
console.log(`テスト結果: ${passed} passed, ${failed} failed`);
if (failed > 0) process.exit(1);
else console.log("✅ 全テスト合格！");
console.log(`\n📁 テスト出力ファイル: ${outputDir}`);
