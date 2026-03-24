/**
 * 生成されたXLSXの内容を詳しく出力して目視確認する
 */
import XLSX from "xlsx";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function inspectXlsx(filePath) {
  const wb = XLSX.read(filePath, { type: "file" });
  console.log(`\n📊 ${path.basename(filePath)}`);
  console.log(`   シート数: ${wb.SheetNames.length}`);

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const headers = data[0] || [];
    const rows = data.slice(1);

    console.log(`\n   📋 シート: "${sheetName}" (${rows.length}行 x ${headers.length}列)`);
    console.log(`   ヘッダー: ${headers.join(" | ")}`);

    for (let i = 0; i < Math.min(rows.length, 3); i++) {
      const rowStr = rows[i].map((v, j) => `${headers[j]}=${v}`).join(", ");
      console.log(`   行${i + 1}: ${rowStr.substring(0, 200)}${rowStr.length > 200 ? "..." : ""}`);
    }
    if (rows.length > 3) {
      console.log(`   ... (以降 ${rows.length - 3}行省略)`);
    }
  }
}

inspectXlsx(path.join(__dirname, "output/ASN_converted.xlsx"));
inspectXlsx(path.join(__dirname, "output/OBDS_converted.xlsx"));
