/**
 * XML Parser のテスト
 * DOMParser を jsdom で提供してNode.jsで実行
 */
import { JSDOM } from "jsdom";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// グローバルに DOMParser を設定（ブラウザAPI互換）
const { DOMParser } = new JSDOM("").window;
globalThis.DOMParser = DOMParser;

// xml-parser.ts のロジックをインラインで再現（ESM互換のため）
function stripNamespace(tagName) {
  const idx = tagName.indexOf(":");
  return idx >= 0 ? tagName.substring(idx + 1) : tagName;
}

function isLeafElement(el) {
  return el.children.length === 0;
}

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
      const key = stripNamespace(child.tagName);
      const value = child.textContent?.trim() ?? "";
      data[key] = value;
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

function parseXml(xmlText, fileName) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, "application/xml");
  const parseError = doc.querySelector("parsererror");
  if (parseError) {
    throw new Error(`XMLパースエラー: ${parseError.textContent?.substring(0, 200)}`);
  }
  const root = doc.documentElement;
  const rootElement = stripNamespace(root.tagName);
  const groups = new Map();
  collectRecordGroups(root, groups);

  if (groups.size === 0) {
    // フォールバック: フラット化
    const flatData = flattenElement(root, "");
    if (Object.keys(flatData).length > 0) {
      groups.set(rootElement, [flatData]);
    }
  }

  const existingNames = new Set();
  const sheets = [];
  for (const [tagName, rows] of groups) {
    sheets.push(buildSheetData(tagName, rows, existingNames));
  }
  return { fileName, rootElement, sheets };
}

function flattenElement(el, prefix) {
  const data = {};
  const currentPrefix = prefix ? `${prefix}.${stripNamespace(el.tagName)}` : stripNamespace(el.tagName);
  if (el.children.length === 0) {
    const value = el.textContent?.trim() ?? "";
    if (value) data[currentPrefix] = value;
  } else {
    for (let i = 0; i < el.children.length; i++) {
      Object.assign(data, flattenElement(el.children[i], currentPrefix));
    }
  }
  return data;
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
      if (!headerSet.has(key)) {
        headerSet.add(key);
        orderedHeaders.push(key);
      }
    }
  }
  const attrHeaders = orderedHeaders.filter((h) => h.startsWith("@"));
  const dataHeaders = orderedHeaders.filter((h) => !h.startsWith("@"));
  const headers = [...dataHeaders, ...attrHeaders];
  return { name: sanitizeSheetName(name, existingNames), headers, rows };
}

// === テスト実行 ===
let passed = 0;
let failed = 0;

function assert(condition, message) {
  if (condition) {
    passed++;
    console.log(`  ✅ ${message}`);
  } else {
    failed++;
    console.error(`  ❌ ${message}`);
  }
}

// テスト 1: ASN ファイル
console.log("\n📄 テスト1: ASN XMLファイル (SAP IDoc)");
const asnXml = fs.readFileSync(
  path.resolve(__dirname, "../../変換前データ/ASN_1800024221_20190214185557.xml"),
  "utf-8"
);
const asnResult = parseXml(asnXml, "ASN_1800024221_20190214185557.xml");

assert(asnResult.rootElement === "SHP_IBDLV_SAVE_REPLICA04", "ルート要素名が正しい");
assert(asnResult.sheets.length > 0, `シートが生成された (${asnResult.sheets.length}シート)`);

const sheetNames = asnResult.sheets.map((s) => s.name);
console.log(`  📋 生成されたシート: ${sheetNames.join(", ")}`);

assert(sheetNames.includes("EDI_DC40"), "EDI_DC40 シートが存在する");
assert(sheetNames.includes("E1BPIBDLVHDR"), "E1BPIBDLVHDR シートが存在する");
assert(sheetNames.includes("E1BPIBDLVITEM"), "E1BPIBDLVITEM シートが存在する");
assert(sheetNames.includes("E1BPEXTC"), "E1BPEXTC シートが存在する");

// EDI_DC40 のデータ検証
const ediSheet = asnResult.sheets.find((s) => s.name === "EDI_DC40");
assert(ediSheet.rows.length === 1, "EDI_DC40は1行");
assert(ediSheet.rows[0]["DOCNUM"] === "0000000059141367", "DOCNUM値が正しい");
assert(ediSheet.rows[0]["CREDAT"] === "20190214", "CREDAT値が正しい");

// E1BPIBDLVITEM のデータ検証
const itemSheet = asnResult.sheets.find((s) => s.name === "E1BPIBDLVITEM");
assert(itemSheet.rows.length === 2, "E1BPIBDLVITEMは2行");
assert(itemSheet.rows[0]["ITM_NUMBER"] === "000010", "1行目のITM_NUMBERが正しい");
assert(itemSheet.rows[1]["ITM_NUMBER"] === "000020", "2行目のITM_NUMBERが正しい");
assert(itemSheet.rows[0]["SHORT_TEXT"] === "MULT-PACK CREW 3PK, 0090, S", "SHORT_TEXTが正しい");
assert(itemSheet.rows[0]["DLV_QTY"] === "39.000", "DLV_QTYが正しい");

// E1BPEXTC のデータ検証
const extcSheet = asnResult.sheets.find((s) => s.name === "E1BPEXTC");
assert(extcSheet.rows.length === 22, `E1BPEXTCは22行 (実際: ${extcSheet.rows.length})`);
assert(extcSheet.headers.includes("FIELD1"), "FIELD1ヘッダーが存在する");

// テスト 2: OBDS ファイル
console.log("\n📄 テスト2: OBDS XMLファイル (名前空間付き)");
const obdsXml = fs.readFileSync(
  path.resolve(__dirname, "../../変換前データ/OBDS_8000580368_20180426103512.xml"),
  "utf-8"
);
const obdsResult = parseXml(obdsXml, "OBDS_8000580368_20180426103512.xml");

assert(obdsResult.rootElement === "MT_PickingSts", "ルート要素名が正しい（名前空間除去済み）");
assert(obdsResult.sheets.length > 0, `シートが生成された (${obdsResult.sheets.length}シート)`);

console.log(`  📋 生成されたシート: ${obdsResult.sheets.map((s) => s.name).join(", ")}`);

// RecordSet のデータ検証
const recordSetSheet = obdsResult.sheets.find((s) => s.name === "RecordSet");
if (recordSetSheet) {
  assert(true, "RecordSetシートが存在する");
  assert(recordSetSheet.rows[0]["I_VBELN"] === "8000580368", "I_VBELN値が正しい");
  assert(recordSetSheet.rows[0]["I_STATUS"] === "Released", "I_STATUS値が正しい");
} else {
  // FMS_ROUTINGとして分割されたかもしれない
  const anySheet = obdsResult.sheets[0];
  console.log(`  📋 最初のシート: ${anySheet.name}`);
  console.log(`  📋 ヘッダー: ${anySheet.headers.join(", ")}`);
  console.log(`  📋 行数: ${anySheet.rows.length}`);
  assert(false, "RecordSetまたは適切なシートが存在する - 要調整");
}

// テスト 3: エッジケース - 空のXML
console.log("\n📄 テスト3: エッジケース");
try {
  parseXml("not valid xml", "invalid.xml");
  assert(false, "不正なXMLでエラーが発生するべき");
} catch (e) {
  assert(e.message.includes("XMLパースエラー"), "不正なXMLでエラーメッセージが表示される");
}

// テスト 4: 空要素のXML
const emptyXml = '<?xml version="1.0"?><root><item><name>test</name><value></value></item></root>';
const emptyResult = parseXml(emptyXml, "empty.xml");
assert(emptyResult.sheets.length > 0, "空要素を含むXMLでもシートが生成される");
const itemSheetEmpty = emptyResult.sheets.find((s) => s.name === "item");
if (itemSheetEmpty) {
  assert(itemSheetEmpty.rows[0]["name"] === "test", "空要素混在でもデータが正しい");
  assert(itemSheetEmpty.rows[0]["value"] === "", "空要素は空文字列になる");
}

// テスト 5: 深くネストされたXML
const nestedXml = `<?xml version="1.0"?>
<root>
  <level1>
    <level2>
      <record><a>1</a><b>2</b></record>
      <record><a>3</a><b>4</b></record>
    </level2>
  </level1>
</root>`;
const nestedResult = parseXml(nestedXml, "nested.xml");
const recordSheet = nestedResult.sheets.find((s) => s.name === "record");
assert(recordSheet !== undefined, "深いネストのレコードが検出される");
if (recordSheet) {
  assert(recordSheet.rows.length === 2, "2つのrecordが正しく検出される");
  assert(recordSheet.rows[0]["a"] === "1", "深いネスト内のデータが正しい");
}

// 結果
console.log(`\n${"=".repeat(50)}`);
console.log(`テスト結果: ${passed} passed, ${failed} failed`);
if (failed > 0) {
  process.exit(1);
} else {
  console.log("✅ 全テスト合格！");
}
