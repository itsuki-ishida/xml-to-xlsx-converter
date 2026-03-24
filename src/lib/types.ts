/** XMLから抽出された1つのシート（レコードグループ）を表す */
export interface SheetData {
  /** シート名（要素のタグ名） */
  name: string;
  /** 列ヘッダー名の配列 */
  headers: string[];
  /** 各行のデータ（キー=ヘッダー名, 値=セルの値） */
  rows: Record<string, string>[];
}

/** 1つのXMLファイルの解析結果 */
export interface ParsedXmlResult {
  /** 元ファイル名 */
  fileName: string;
  /** ルート要素名 */
  rootElement: string;
  /** 各シートのデータ */
  sheets: SheetData[];
}

/** ファイルの変換状態 */
export type ConversionStatus = "pending" | "converting" | "done" | "error";

/** アップロードされたファイルの管理情報 */
export interface UploadedFile {
  id: string;
  file: File;
  status: ConversionStatus;
  error?: string;
  result?: ParsedXmlResult;
}
