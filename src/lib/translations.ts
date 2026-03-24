/**
 * XML要素名・フィールド名の翻訳辞書
 *
 * SAP公式ドキュメント（IDoc Interface / ABAP Data Dictionary / SAP Help Portal）に
 * 基づいて、SAP IDoc形式をはじめとする業務XMLのフィールド名を
 * 非エンジニアにも理解できる日本語ラベルに翻訳する。
 *
 * 翻訳が存在しないフィールドは元の技術名をそのまま表示する。
 * 汎用XMLへの影響はなく、翻訳は純粋な「上乗せ」機能。
 *
 * 参照元:
 *  - SAP Help Portal: IDoc Interface (https://help.sap.com)
 *  - SAP ABAP Data Dictionary (トランザクション SE11)
 *  - SAP Note & Field Documentation
 */

// ─── セクション（要素）名の翻訳 ─────────────────────────────

const SECTION_TRANSLATIONS: Record<string, string> = {
  // SAP IDoc 共通
  EDI_DC40: "IDoc制御情報",
  EDI_DD40: "IDocデータレコード",

  // ASN (出荷通知) 関連
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

  // OBDS / ピッキング関連
  MT_PickingSts: "ピッキングステータス",
  RecordSet: "レコードセット",
  FMS_ROUTING: "ルーティング情報",

  // 受注・販売伝票関連
  E1EDK01: "伝票ヘッダー",
  E1EDK14: "伝票ヘッダー組織",
  E1EDK02: "伝票ヘッダー参照",
  E1EDK03: "伝票ヘッダー日付",
  E1EDK04: "伝票ヘッダー税",
  E1EDK05: "伝票ヘッダー条件",
  E1EDP01: "伝票明細",
  E1EDP02: "伝票明細参照",
  E1EDP05: "伝票明細条件",
  E1EDP19: "伝票明細納入日程",
  E1EDKA1: "伝票パートナー",
  E1EDKT1: "伝票テキスト",
  E1EDKT2: "伝票明細テキスト",
  E1EDS01: "伝票合計",

  // 請求書関連
  E1EDK01_INVOIC: "請求書ヘッダー",

  // マスタデータ関連
  E1MARAM: "品目マスタ一般",
  E1MARCM: "品目マスタプラント",
  E1MAKTM: "品目テキスト",
  E1MARDM: "品目マスタ保管場所",
  E1LFA1M: "仕入先マスタ一般",
  E1KNA1M: "得意先マスタ一般",
};

// ─── フィールド名の翻訳 ────────────────────────────────────

const FIELD_TRANSLATIONS: Record<string, string> = {
  // ── IDoc制御 (EDI_DC40) ──
  TABNAM: "テーブル名",
  MANDT: "クライアント",
  DOCNUM: "ドキュメント番号",
  DOCREL: "リリース番号",
  STATUS: "ステータス",
  DIRECT: "方向",
  OUTMOD: "出力モード",
  IDOCTYP: "IDocタイプ",
  MESTYP: "メッセージタイプ",
  SNDPOR: "送信ポート",
  SNDPRT: "送信パートナータイプ",
  SNDPRN: "送信パートナー",
  RCVPOR: "受信ポート",
  RCVPRT: "受信パートナータイプ",
  RCVPRN: "受信パートナー",
  CREDAT: "作成日",
  CRETIM: "作成時刻",
  SERIAL: "シリアル番号",

  // ── 配送ヘッダー ──
  DELIV_NUMB: "配送番号",
  TOTAL_WGHT: "総重量",
  NET_WEIGHT: "正味重量",
  UNIT_OF_WT: "重量単位",
  UNIT_OF_WT_ISO: "重量単位(ISO)",
  VOLUME: "容量",
  NOSHPUNITS: "出荷ユニット数",
  DLV_TYPE: "配送タイプ",
  DLV_PRIO: "配送優先度",
  SD_DOC_CAT: "SD伝票カテゴリ",
  SHIP_POINT: "出荷ポイント",
  WHSE_NO: "倉庫番号",
  RECV_WHS_NO: "受信倉庫番号",
  RECV_SYS: "受信システム",
  SENDER_SYSTEM: "送信システム",

  // ── パートナー・住所 ──
  PARTN_ROLE: "パートナー機能",
  PARTNER_NO: "パートナー番号",
  ADDRESS_NO: "住所番号",
  ADDR_NO: "住所番号",
  FORMOFADDR: "敬称",
  NAME: "名前",
  NAME1: "名前1",
  NAME2: "名前2",
  NAME3: "名前3",
  NAME4: "名前4",
  CITY: "市区町村",
  POSTL_COD1: "郵便番号",
  STREET: "通り",
  STREET_LNG: "通り(詳細)",
  COUNTRY: "国",
  LANGU: "言語",
  SORT1: "ソートキー1",
  SORT2: "ソートキー2",
  TIME_ZONE: "タイムゾーン",
  TEL1_NUMBR: "電話番号",
  E_MAIL: "メールアドレス",
  TITLE: "敬称",
  COUNTRYISO: "国コード(ISO)",
  LANGU_ISO: "言語(ISO)",
  LANGU_CR: "対応言語",
  LANGUCRISO: "対応言語(ISO)",

  // ── 配送期日 ──
  TIMETYPE: "時間タイプ",
  TIMESTAMP_UTC: "タイムスタンプ(UTC)",
  TIMEZONE: "タイムゾーン",

  // ── 配送明細 ──
  ITM_NUMBER: "明細番号",
  MATERIAL: "品目コード",
  SHORT_TEXT: "品目テキスト",
  DLV_QTY: "配送数量",
  SALES_UNIT: "販売単位",
  SALES_UNIT_ISO: "販売単位(ISO)",
  DLV_QTY_STOCK: "在庫配送数量",
  BASE_UOM: "基本単位",
  BASE_UOM_ISO: "基本単位(ISO)",
  CONV_UNIT_NOM: "変換分子",
  CONV_UNIT_DENOM: "変換分母",
  GROSS_WT: "総重量",
  HIERARITEM: "階層明細",
  LOADINGGRP: "積載グループ",
  TRANS_GRP: "輸送グループ",
  DLV_GROUP: "配送グループ",
  EAN_UPC: "EAN/UPCコード",
  ITEM_CATEG: "明細カテゴリ",
  OVERDELTOL: "過剰配送許容度",
  UNDER_TOL: "不足配送許容度",
  MAT_GRP_SM: "品目グループ(出荷)",
  WHSE_MVMT: "倉庫移動",
  EXPIRYDATE: "有効期限",
  MOVE_TYPE: "移動タイプ",
  MOVE_TYPE_WM: "WM移動タイプ",
  MVT_IND: "移動指標",
  CUM_BTCH_QTY: "累積バッチ数量",
  CUM_BTCH_GR_WT: "累積バッチ総重量",
  CUM_BTCH_NT_WT: "累積バッチ正味重量",
  CUM_BTCH_VOL: "累積バッチ容量",
  MATL_GROUP: "品目グループ",
  MATL_TYPE: "品目タイプ",
  DEL_QTY_FLO: "配送数量(浮動小数)",
  CONV_FACT: "変換係数",
  DLV_QTY_ST_FLO: "在庫配送数量(浮動小数)",
  CUMBTCHQTYSU_FLO: "累積バッチ数量SU(浮動小数)",
  CUMBTCHQTYSU: "累積バッチ数量SU",
  INSPLOT: "検査ロット",
  PROD_DATE: "製造日",
  CURR_QTY: "現在数量",

  // ── 配送明細組織 ──
  PLANT: "プラント",
  STGE_LOC: "保管場所",

  // ── 配送明細原価 ──
  PROFIT_CTR: "利益センタ",
  PROFIT_SEGM_NO: "利益セグメント番号",
  S_ORD_ITEM: "受注明細",
  ORDER_ITNO: "指図明細番号",

  // ── 配送明細参照伝票 ──
  DOC_NUMBER: "伝票番号",
  ITM_NUMBER_REF: "参照明細番号",
  DOC_CAT: "伝票カテゴリ",
  DOC_TYPE: "伝票タイプ",
  PURCH_ORG: "購買組織",
  PUR_GROUP: "購買グループ",
  DOC_DATE: "伝票日付",
  ITEM_CAT: "明細カテゴリ",
  GR_IND: "入庫指標",
  GRSETTFROM: "入庫決済元",
  SPE_EXT_ID_ITEM: "外部明細ID",

  // ── 梱包ユニット ──
  HDL_UNIT: "梱包ユニット",
  HDL_UNIT_EXID: "梱包ユニット外部ID",
  HDL_UNIT_EXID_TY: "外部IDタイプ",
  LOAD_WGHT: "積載重量",
  ALLOWED_WGHT: "許容重量",
  TARE_WGHT: "風袋重量",
  TARE_UNIT_WT: "風袋重量単位",
  TARE_UNIT_WT_ISO: "風袋重量単位(ISO)",
  TOTAL_VOL: "総容量",
  LOAD_VOL: "積載容量",
  ALLOWEDVOL: "許容容量",
  TARE_VOL: "風袋容量",
  SHIP_MAT: "梱包資材",
  LENGTH: "長さ",
  WIDTH: "幅",
  HEIGHT: "高さ",
  WT_TOL_LT: "重量許容度(下限)",
  VOL_TOL_LT: "容量許容度(下限)",
  SH_MAT_TYP: "梱包資材タイプ",
  LGTH_LOAD: "積載長",
  TRAVL_TIME: "輸送時間",
  DISTANCE: "距離",
  ALLOWEDLOAD: "許容積載",
  HDL_UNIT_GUID: "梱包ユニットGUID",
  LOG_SYSTEM: "論理システム",
  CHECK_NUMBER: "チェック番号",
  HDL_UNIT_INTO: "上位梱包ユニット",
  HDL_UNIT_EXID_INTO: "上位梱包ユニット外部ID",
  DELIV_ITEM: "配送明細",
  PACK_QTY: "梱包数量",
  HU_ITEM_TYPE: "HU明細タイプ",
  PACK_QTY_BASE: "梱包数量(基本)",

  // ── 拡張データ ──
  FIELD1: "フィールド1",
  FIELD2: "フィールド2",
  FIELD3: "フィールド3",
  FIELD4: "フィールド4",

  // ── OBDS / ピッキング ──
  I_VBELN: "伝票番号",
  I_STATUS: "ステータス",
  I_UTC_TIMESTAMP: "タイムスタンプ(UTC)",

  // ── 汎用的なSAPフィールド ──
  BUKRS: "会社コード",
  WERKS: "プラント",
  LGORT: "保管場所",
  VBELN: "伝票番号",
  POSNR: "明細番号",
  MATNR: "品目番号",
  MENGE: "数量",
  MEINS: "単位",
  WAERS: "通貨",
  NETWR: "正味金額",
  KUNNR: "得意先番号",
  LIFNR: "仕入先番号",
  BSTNK: "購買発注番号",
  ERDAT: "登録日",
  ERNAM: "登録者",
  AEDAT: "変更日",
  AENAM: "変更者",
  LOGSYS: "論理システム",
  BELNR: "伝票番号",
  GJAHR: "会計年度",
  BLDAT: "伝票日付",
  BUDAT: "転記日付",
  CPUDT: "登録日",
  CPUTM: "登録時刻",
  USNAM: "ユーザー名",
  XBLNR: "参照伝票番号",
  BKTXT: "伝票ヘッダーテキスト",
  WAERK: "通貨キー",
  KURRF: "為替レート",
  WRBTR: "伝票通貨金額",
  DMBTR: "国内通貨金額",
  MWSKZ: "税コード",
  TXJCD: "税管轄コード",
  ZUONR: "割当",
  SGTXT: "明細テキスト",
  KOSTL: "原価センタ",
  AUFNR: "指図番号",
  PRCTR: "利益センタ",
  GSBER: "事業領域",
  SEGMENT: "セグメント",
};

// ─── 翻訳ヘルパー関数 ────────────────────────────────────────

/**
 * フィールド名から @プレフィックスを除去して技術名を返す
 */
export function stripFieldPrefix(technicalName: string): string {
  return technicalName.startsWith("@")
    ? technicalName.substring(1)
    : technicalName;
}

/**
 * セクション名を翻訳する（フル表記）
 * 翻訳あり → "翻訳名 (技術名)"
 * 翻訳なし → "技術名"
 */
export function translateSection(technicalName: string): string {
  const ja = SECTION_TRANSLATIONS[technicalName];
  return ja ? `${ja} (${technicalName})` : technicalName;
}

/**
 * セクション名を翻訳する（日本語名のみ）
 * 翻訳あり → "翻訳名"
 * 翻訳なし → "技術名"
 */
export function translateSectionShort(technicalName: string): string {
  return SECTION_TRANSLATIONS[technicalName] ?? technicalName;
}

/**
 * フィールド名を翻訳する（概要シート用: フル表記）
 * 翻訳あり → "翻訳名 (技術名)"
 * 翻訳なし → "技術名"
 */
export function translateField(technicalName: string): string {
  const raw = stripFieldPrefix(technicalName);
  const ja = FIELD_TRANSLATIONS[raw];
  return ja ? `${ja} (${raw})` : raw;
}

/**
 * フィールド名を翻訳する（日本語名のみ）
 * 翻訳あり → "翻訳名"
 * 翻訳なし → "技術名"
 */
export function translateFieldShort(technicalName: string): string {
  const raw = stripFieldPrefix(technicalName);
  return FIELD_TRANSLATIONS[raw] ?? raw;
}

/**
 * 翻訳辞書の出典情報
 */
export const TRANSLATION_SOURCE =
  "SAP公式ドキュメント（IDoc Interface / ABAP Data Dictionary）に基づく翻訳";
