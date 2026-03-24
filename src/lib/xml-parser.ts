import type { SheetData, ParsedXmlResult } from "./types";

/**
 * 名前空間プレフィックスを除去してローカル名を返す
 * 例: "ns0:MT_PickingSts" → "MT_PickingSts"
 */
function stripNamespace(tagName: string): string {
  const idx = tagName.indexOf(":");
  return idx >= 0 ? tagName.substring(idx + 1) : tagName;
}

/**
 * 要素が「リーフ要素」か判定する
 * リーフ要素 = 子要素を持たず、テキストコンテンツのみを持つ要素
 */
function isLeafElement(el: Element): boolean {
  // 子要素が1つもない
  return el.children.length === 0;
}

/**
 * 要素が「レコード要素」か判定する
 * レコード要素 = 少なくとも1つのリーフ子要素を持つ要素
 */
function isRecordElement(el: Element): boolean {
  if (el.children.length === 0) return false;
  for (let i = 0; i < el.children.length; i++) {
    if (isLeafElement(el.children[i])) {
      return true;
    }
  }
  return false;
}

/**
 * 要素からレコードデータ（key-value）を抽出する
 * リーフ子要素のタグ名をキー、テキストコンテンツを値とする
 * 属性も @属性名 として含める
 */
function extractRecordData(el: Element): Record<string, string> {
  const data: Record<string, string> = {};

  // 属性を抽出
  for (let i = 0; i < el.attributes.length; i++) {
    const attr = el.attributes[i];
    // xmlns系の属性は除外
    if (attr.name.startsWith("xmlns")) continue;
    data[`@${attr.name}`] = attr.value;
  }

  // リーフ子要素を抽出
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

/**
 * シート名をExcelの制限（31文字）に合わせて切り詰める
 * 重複する場合はサフィックスを付加
 */
function sanitizeSheetName(name: string, existingNames: Set<string>): string {
  // Excelで禁止されている文字を除去
  let sanitized = name.replace(/[\\/*?:\[\]]/g, "_");

  // 31文字に切り詰め
  if (sanitized.length > 31) {
    sanitized = sanitized.substring(0, 31);
  }

  // 重複チェック
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

/**
 * XMLツリーを再帰的に走査し、レコード要素をタグ名ごとにグループ化する
 */
function collectRecordGroups(
  el: Element,
  groups: Map<string, Record<string, string>[]>
): void {
  if (isRecordElement(el)) {
    const tagName = stripNamespace(el.tagName);
    const data = extractRecordData(el);

    if (!groups.has(tagName)) {
      groups.set(tagName, []);
    }
    groups.get(tagName)!.push(data);
  }

  // 子要素も再帰的に走査
  for (let i = 0; i < el.children.length; i++) {
    collectRecordGroups(el.children[i], groups);
  }
}

/**
 * レコードグループからSheetDataを生成する
 * 全行のキーを結合してヘッダーを生成（出現順を保持）
 */
function buildSheetData(
  name: string,
  rows: Record<string, string>[],
  existingNames: Set<string>
): SheetData {
  // 全行からキーを収集（出現順を保持）
  const headerSet = new Set<string>();
  const orderedHeaders: string[] = [];

  for (const row of rows) {
    for (const key of Object.keys(row)) {
      if (!headerSet.has(key)) {
        headerSet.add(key);
        orderedHeaders.push(key);
      }
    }
  }

  // 属性列(@始まり)は末尾に配置
  const attrHeaders = orderedHeaders.filter((h) => h.startsWith("@"));
  const dataHeaders = orderedHeaders.filter((h) => !h.startsWith("@"));
  const headers = [...dataHeaders, ...attrHeaders];

  return {
    name: sanitizeSheetName(name, existingNames),
    headers,
    rows,
  };
}

/**
 * XMLテキストをパースし、SheetDataの配列に変換する
 * @param xmlText - XML文字列
 * @param fileName - 元ファイル名
 * @returns ParsedXmlResult
 * @throws Error - XMLパースエラーの場合
 */
export function parseXml(xmlText: string, fileName: string): ParsedXmlResult {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, "application/xml");

  // パースエラーチェック
  const parseError = doc.querySelector("parsererror");
  if (parseError) {
    throw new Error(
      `XMLパースエラー: ${parseError.textContent?.substring(0, 200)}`
    );
  }

  const root = doc.documentElement;
  const rootElement = stripNamespace(root.tagName);

  // レコードグループを収集
  const groups = new Map<string, Record<string, string>[]>();
  collectRecordGroups(root, groups);

  // グループが空の場合（全てコンテナ要素のみのXML）
  // ルート要素自体をフラットに展開する
  if (groups.size === 0) {
    const flatData = flattenElement(root, "");
    if (Object.keys(flatData).length > 0) {
      groups.set(rootElement, [flatData]);
    }
  }

  // SheetDataに変換
  const existingNames = new Set<string>();
  const sheets: SheetData[] = [];
  for (const [tagName, rows] of groups) {
    sheets.push(buildSheetData(tagName, rows, existingNames));
  }

  return {
    fileName,
    rootElement,
    sheets,
  };
}

/**
 * 要素を再帰的にフラット化する（フォールバック用）
 * 通常のレコード抽出で結果が得られない場合に使用
 */
function flattenElement(
  el: Element,
  prefix: string
): Record<string, string> {
  const data: Record<string, string> = {};
  const currentPrefix = prefix
    ? `${prefix}.${stripNamespace(el.tagName)}`
    : stripNamespace(el.tagName);

  if (el.children.length === 0) {
    // リーフ要素
    const value = el.textContent?.trim() ?? "";
    if (value) {
      data[currentPrefix] = value;
    }
  } else {
    // 子要素を再帰的に処理
    for (let i = 0; i < el.children.length; i++) {
      const childData = flattenElement(el.children[i], currentPrefix);
      Object.assign(data, childData);
    }
  }

  return data;
}
